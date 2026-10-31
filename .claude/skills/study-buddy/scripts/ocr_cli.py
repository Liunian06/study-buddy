#!/usr/bin/env python3
"""把课件、参考资料、往年题、作业要求转成带页码和图片转写的Markdown草稿。

流程：Office转PDF、切分、上传OCR引擎（OXR或Paddle）、按页组装、下载图片、VLM转写图片、合并、校验。
用--engine选择引擎，默认取config.json里的engine，没写则用paddle（PaddleOCR-VL-1.6）；paddle没填key时退用oxr。
结果以JSON摘要输出到stdout，日志输出到stderr。

退出码：0全部成功，1部分失败，2参数或环境错误。
"""
import argparse
import base64
import hashlib
import io
import json
import re
import shutil
import sys
import time
from pathlib import Path

TYPE_NAMES = {"material": "课件/参考资料", "textbook": "教材", "past-exam": "往年题", "assignment": "作业要求"}
OFFICE_PPT = {".ppt", ".pptx"}
OFFICE_DOC = {".doc", ".docx"}
SKILL_DIR = Path(__file__).resolve().parent.parent
SKIP_MARK = "[跳过："


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def fail_env(msg):
    log(f"环境或参数错误：{msg}")
    sys.exit(2)


def safe_name(s):
    """去掉非法字符，空白换成下划线，避免Markdown图片路径里出现空格"""
    s = re.sub(r'[\\/:*?"<>|]', "_", s).strip()
    return re.sub(r"\s+", "_", s) or "unnamed"


def parse_pages(spec):
    """'10,17,20-22' -> [10,17,20,21,22]（1起）"""
    out = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return sorted(set(out))


# ---------------------------------------------------------------- 配置

PLACEHOLDER_KEYS = {"sk-123", "paddle-token-123"}  # config.template.json里的示例key

def load_config(data_dir):
    cfg_path = data_dir / "config.json"
    created = False
    if not cfg_path.exists():
        tpl = SKILL_DIR / "assets" / "templates" / "config.template.json"
        data_dir.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(tpl, cfg_path)
        created = True
    with open(cfg_path, encoding="utf-8") as f:
        cfg = json.load(f)
    with open(SKILL_DIR / "assets" / "templates" / "config.template.json", encoding="utf-8") as f:
        tpl = json.load(f)
    for k, v in tpl.items():  # 旧版config.json缺的字段用模板补齐（只在内存里补，不改文件）
        if isinstance(v, dict):
            sec = cfg.setdefault(k, {})
            for kk, vv in v.items():
                sec.setdefault(kk, vv)
        else:
            cfg.setdefault(k, v)
    # 模板里的示例key（sk-123等）只是格式示范，视为没填，不能触发真实请求或挡住引擎回退
    for sec in ("paddle", "vlm"):
        if cfg[sec].get("api_key") in PLACEHOLDER_KEYS:
            cfg[sec]["api_key"] = ""
    return cfg, created


# ---------------------------------------------------------------- 转换与切分


def office_to_pdf(src, out_pdf):
    try:
        import win32com.client
    except ImportError:
        fail_env("缺少pywin32，无法转换Office文件，见references/setup.md")
    src_s, out_s = str(src.resolve()), str(out_pdf.resolve())
    ext = src.suffix.lower()
    if ext in OFFICE_PPT:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            pres = app.Presentations.Open(src_s, True, False, False)
            try:
                pres.SaveAs(out_s, 32)  # ppSaveAsPDF
            finally:
                pres.Close()
        finally:
            app.Quit()
    else:
        app = win32com.client.DispatchEx("Word.Application")
        app.Visible = False
        try:
            doc = app.Documents.Open(src_s, ReadOnly=True)
            try:
                doc.ExportAsFixedFormat(out_s, 17)  # wdExportFormatPDF
            finally:
                doc.Close(False)
        finally:
            app.Quit()
    if not out_pdf.exists():
        raise RuntimeError("Office未生成PDF")


def make_chunk_pdf(fitz, src_pdf, pages, out_path):
    """pages为1起的原始页码列表"""
    src = fitz.open(str(src_pdf))
    new = fitz.open()
    for p in pages:
        new.insert_pdf(src, from_page=p - 1, to_page=p - 1)
    new.save(str(out_path), garbage=3, deflate=True)
    new.close()
    src.close()


def _page_count(fitz, pdf):
    d = fitz.open(str(pdf))
    n = d.page_count
    d.close()
    return n


def plan_chunks(fitz, src_pdf, pages, max_pages, max_bytes, work_dir):
    """按页数切分；某块超过大小上限时对半再切。返回块的页码列表"""
    result = []
    # 扫描件（如整本教材）每页体积大，先按平均每页体积估一个块的页数，减少对半重切的次数
    avg = max(src_pdf.stat().st_size / max(_page_count(fitz, src_pdf), 1), 1)
    est = max(1, min(max_pages, int(max_bytes / avg * 0.8)))
    queue = [pages[i:i + est] for i in range(0, len(pages), est)]
    probe = work_dir / "_probe.pdf"
    while queue:
        pg = queue.pop(0)
        make_chunk_pdf(fitz, src_pdf, pg, probe)
        if probe.stat().st_size > max_bytes and len(pg) > 1:
            mid = len(pg) // 2
            queue[0:0] = [pg[:mid], pg[mid:]]
            continue
        result.append(pg)
    probe.unlink(missing_ok=True)
    return result


# ---------------------------------------------------------------- OXR


class FatalEngineError(Exception):
    """额度用尽、鉴权失败等重试也没用的错误，遇到后停止提交后面的块"""


def oxr_parse(cfg, pdf_path):
    """OXR同步接口。429表示免费接口上已有任务，等久一点再试；网络错误短间隔重试"""
    import requests
    url = cfg["base_url"].rstrip("/") + "/parse/sync"
    last = None
    for attempt in range(6):
        try:
            with open(pdf_path, "rb") as fh:
                r = requests.post(
                    url,
                    data=[("api_key", cfg.get("api_key", "dummy")), ("output_formats", "json")],
                    files={"file": (pdf_path.name, fh, "application/pdf")},
                    timeout=cfg.get("timeout_seconds", 300),
                )
            if r.status_code == 429:
                last = RuntimeError("HTTP 429：接口上已有任务在进行")
                time.sleep(cfg.get("retry_429_seconds", 30))
                continue
            if r.status_code >= 500:
                raise RuntimeError(f"HTTP {r.status_code}")
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}：{r.text[:200]}")
            body = r.json()
            if body.get("code") != 0:
                raise RuntimeError(f"OXR返回code={body.get('code')}：{body.get('message')}")
            return body["data"]["result"]["outputs"]["json"]
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (attempt + 1))
    raise RuntimeError(str(last))


_IMG_TAG = re.compile(r'<img\b[^>]*?src="([^"]+)"[^>]*?>', re.I)
_IMG_MD = re.compile(r"!\[[^\]]*\]\(([^)\s]+)[^)]*\)")
_DIV = re.compile(r"</?div[^>]*>", re.I)
_BOX = re.compile(r"_box_(\d+)_(\d+)_(\d+)_(\d+)")
# Paddle的块类型 -> 我们的类型。页眉、页码、页脚不进正文，单独存为页的元数据
_PADDLE_SKIP = {"header", "header_image", "footer", "footer_image", "aside_text", "number"}
_PADDLE_PICTURE = {"image", "chart", "seal"}


def _bbox_key(bbox):
    return tuple(int(round(x)) for x in bbox)


def _anchor_y(b, pr):
    """图块插入阅读顺序时用的纵坐标：有图注就挨着图注（放在图注上方），没有就用图自己的位置"""
    if b["block_label"] in _PADDLE_PICTURE:
        x0, y0, x1, y1 = b["block_bbox"]
        for c in pr:
            if c["block_label"] == "figure_title":
                cx0, cy0, cx1, _ = c["block_bbox"]
                if 0 <= cy0 - y1 <= 80 and min(x1, cx1) - max(x0, cx0) > 0:
                    return cy0 - 0.5
    return b["block_bbox"][1]


def paddle_page_to_blocks(res):
    """把Paddle一页的块列表(parsing_res_list)转成与OXR一致的{type, format, content}块。

    页眉、页码、页脚不进正文，页码（number块）存成printed_page，用来对应书上印的页码。
    parsing_res_list缺失时退回按Markdown拆块。
    """
    pr = (res.get("prunedResult") or {}).get("parsing_res_list")
    if pr is None:
        return _paddle_page_from_markdown(res)
    imgs = (res.get("markdown") or {}).get("images") or {}
    box_url = {}
    for path, url in imgs.items():
        m = _BOX.search(path)
        if m:
            box_url[tuple(int(x) for x in m.groups())] = url
    ordered = sorted(
        enumerate(pr), key=lambda ib: (ib[1].get("block_order") is None, ib[1].get("block_order") or 0, ib[0]))
    # block_order为None的块（图、图注）按版面位置并入阅读顺序：放在上方最近的有序块之后
    seq = [b for _, b in ordered if b.get("block_order") is not None]
    for _, b in ordered:
        if b.get("block_order") is None and b["block_label"] not in _PADDLE_SKIP:
            y = _anchor_y(b, pr)
            at = len(seq)
            for i, o in enumerate(seq):
                if o["block_bbox"][1] > y:
                    at = i
                    break
            seq.insert(at, b)
    # 公式编号按纵向位置并入对应的公式（编号常和公式隔着别的块，不能只看相邻）
    merged_ids = set()
    tags = {}
    for fn in [b for b in pr if b["block_label"] == "formula_number"]:
        fy0, fy1 = fn["block_bbox"][1], fn["block_bbox"][3]
        best = None
        for f in pr:
            if f["block_label"] not in ("display_formula", "inline_formula"):
                continue
            y0, y1 = f["block_bbox"][1], f["block_bbox"][3]
            if min(y1, fy1) - max(y0, fy0) > 0.3 * (fy1 - fy0) and f["block_bbox"][2] <= fn["block_bbox"][2]:
                if best is None or f["block_bbox"][2] > best["block_bbox"][2]:
                    best = f
        if best is not None and id(best) not in tags:
            tags[id(best)] = (fn.get("block_content") or "").strip().strip("()（）")
            merged_ids.add(id(fn))
    meta = {"printed_page": None, "header": None, "labels": []}
    for b in pr:
        if b["block_label"] == "number" and meta["printed_page"] is None:
            meta["printed_page"] = (b.get("block_content") or "").strip() or None
        elif b["block_label"] == "header" and meta["header"] is None:
            meta["header"] = (b.get("block_content") or "").strip() or None
    blocks = []
    for b in seq:
        lab, text = b["block_label"], (b.get("block_content") or "").strip()
        if id(b) in merged_ids:
            continue
        if id(b) in tags and text.endswith("$$"):
            text = text[:-2].rstrip() + " \\tag{" + tags[id(b)] + "} $$"
        meta["labels"].append(lab)
        if lab in _PADDLE_PICTURE:
            blocks.append({"type": "Picture", "img_url": box_url.get(_bbox_key(b["block_bbox"]), ""), "label": lab})
        elif lab == "paragraph_title" or lab == "doc_title":
            blocks.append({"type": "Title", "format": "text", "content": text, "label": lab})
        elif text:
            fmt = "latex" if lab in ("display_formula", "inline_formula") else "markdown"
            blocks.append({"type": "Formula" if fmt == "latex" else "Text", "format": fmt, "content": text,
                           "label": lab})
    pg = {"blocks": blocks}
    pg.update({k: v for k, v in meta.items() if k != "labels"})
    pg["labels"] = meta["labels"]
    return pg


def _paddle_page_from_markdown(res):
    md = res["markdown"]["text"]
    imgs = res["markdown"].get("images") or {}
    blocks, pos = [], 0
    hits = sorted((m.start(), m.end(), m.group(1)) for rx in (_IMG_TAG, _IMG_MD) for m in rx.finditer(md))
    for st, en, path in hits:
        if st < pos:
            continue
        seg = _DIV.sub("", md[pos:st]).strip()
        if seg:
            blocks.append({"type": "Text", "format": "markdown", "content": seg})
        blocks.append({"type": "Picture", "img_url": imgs.get(path) or (path if path.startswith("http") else "")})
        pos = en
    seg = _DIV.sub("", md[pos:]).strip()
    if seg:
        blocks.append({"type": "Text", "format": "markdown", "content": seg})
    return {"blocks": blocks}


def paddle_parse(cfg, pdf_path):
    """Paddle异步接口：提交、轮询、取结果。返回与OXR相同结构的{"pages": [...]}"""
    import requests
    job_url = cfg["base_url"].rstrip("/")
    headers = {"Authorization": f"bearer {cfg['api_key']}"}
    payload = {"useDocOrientationClassify": False, "useDocUnwarping": False, "useChartRecognition": False}
    job_id, last = None, None
    for attempt in range(6):  # 提交：限流等待，网络错误重试
        try:
            with open(pdf_path, "rb") as fh:
                r = requests.post(job_url, headers=headers,
                                  data={"model": cfg["model"], "optionalPayload": json.dumps(payload)},
                                  files={"file": (pdf_path.name, fh, "application/pdf")},
                                  timeout=cfg.get("timeout_seconds", 300))
            if r.status_code == 429:
                last = RuntimeError("HTTP 429：请求太频繁")
                time.sleep(cfg.get("retry_429_seconds", 30))
                continue
            if r.status_code in (401, 403):
                raise FatalEngineError(f"HTTP {r.status_code}（key无效或今日额度已用完）：{r.text[:150]}")
            if r.status_code != 200:
                raise RuntimeError(f"提交失败 HTTP {r.status_code}：{r.text[:150]}")
            job_id = r.json()["data"]["jobId"]
            break
        except FatalEngineError:
            raise
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (attempt + 1))
    if not job_id:
        raise RuntimeError(f"提交未成功：{last}")
    deadline = time.time() + cfg.get("job_timeout_seconds", 1800)
    jsonl_url, fails = None, 0
    while time.time() < deadline:  # 轮询：单次查询失败只重试这一次查询
        time.sleep(cfg.get("poll_seconds", 5))
        try:
            r = requests.get(f"{job_url}/{job_id}", headers=headers, timeout=60)
            if r.status_code != 200:
                raise RuntimeError(f"查询 HTTP {r.status_code}")
            d = r.json()["data"]
            fails = 0
        except Exception as e:  # noqa: BLE001
            fails += 1
            if fails >= 10:
                raise RuntimeError(f"连续查询失败：{e}")
            continue
        if d["state"] == "done":
            jsonl_url = d["resultUrl"]["jsonUrl"]
            break
        if d["state"] == "failed":
            raise RuntimeError(f"任务失败：{d.get('errorMsg')}")
    if not jsonl_url:
        raise RuntimeError("任务超时")
    rr = None
    for attempt in range(4):
        try:
            rr = requests.get(jsonl_url, timeout=120)
            rr.raise_for_status()
            break
        except Exception as e:  # noqa: BLE001
            last, rr = e, None
            time.sleep(3 * (attempt + 1))
    if rr is None:
        raise RuntimeError(f"下载结果失败：{last}")
    pages = []
    for line in rr.text.splitlines():
        line = line.strip()
        if not line:
            continue
        for res in json.loads(line)["result"]["layoutParsingResults"]:
            pg = paddle_page_to_blocks(res)
            pg["page_num"] = len(pages)
            pages.append(pg)
    return {"pages": pages}


ENGINES = {"oxr": oxr_parse, "paddle": paddle_parse}


def fetch_image(url, dest_base):
    """下载图片，按真实格式确定扩展名。返回保存的路径"""
    import requests
    from PIL import Image
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    img = Image.open(io.BytesIO(r.content))
    ext = {"PNG": ".png", "JPEG": ".jpg", "GIF": ".gif", "WEBP": ".webp", "BMP": ".bmp"}.get(img.format, ".png")
    dest = dest_base.with_suffix(ext)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)
    return dest


# ---------------------------------------------------------------- VLM


class Vlm:
    def __init__(self, cfg):
        self.cfg = cfg
        self.enabled = bool(cfg.get("base_url") and cfg.get("model") and cfg.get("api_key"))
        self.prompt = (SKILL_DIR / "assets" / "prompts" / "vlm-image.md").read_text(encoding="utf-8")

    def transcribe(self, path):
        import requests
        raw = path.read_bytes()
        mime = {".png": "image/png", ".jpg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp"}.get(
            path.suffix.lower(), "image/png")
        b64 = base64.b64encode(raw).decode("ascii")
        payload = {
            "model": self.cfg["model"],
            "temperature": 0,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": self.prompt},
                    {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
                ],
            }],
        }
        url = self.cfg["base_url"].rstrip("/") + "/chat/completions"
        headers = {"Authorization": f"Bearer {self.cfg['api_key']}"}
        last = None
        for attempt in range(3):
            try:
                r = requests.post(url, json=payload, headers=headers, timeout=self.cfg.get("timeout_seconds", 120))
                if r.status_code != 200:
                    raise RuntimeError(f"HTTP {r.status_code}")
                text = r.json()["choices"][0]["message"]["content"]
                return (text or "").strip()
            except Exception as e:  # noqa: BLE001
                last = e
                time.sleep(2 * (attempt + 1))
        raise RuntimeError(str(last))


def too_small(path, min_px):
    from PIL import Image
    try:
        with Image.open(path) as im:
            return min(im.size) < min_px
    except Exception:  # noqa: BLE001
        return False


# ---------------------------------------------------------------- 组装


def heading_level(text):
    m = re.match(r"\s*(\d+(?:\.\d+){0,3})[\s.、)）]", text)
    if m:
        return min(m.group(1).count(".") + 1, 4)
    return 2


def render_block(b, ctx):
    t = (b.get("type") or "").lower()
    fmt = (b.get("format") or "").lower()
    content = (b.get("content") or "").strip()
    if t == "picture":
        return ctx["picture"](b)
    if not content:
        return ""
    if t == "title":
        return "#" * heading_level(content) + " " + re.sub(r"\s+", " ", content)
    if fmt == "latex" or t in ("formula", "equation"):
        return content if content.startswith("$$") else f"$$\n{content}\n$$"
    return content


def vlm_text_block(entry):
    """entry: manifest里的图片记录。返回草稿里图片下方的转写文字"""
    st = entry.get("vlm")
    if st == "ok":
        txt = entry["text"]
        if txt.startswith(SKIP_MARK):
            return ""
        return "> 图片转写：\n" + "\n".join("> " + ln for ln in txt.splitlines())
    if st == "skipped":
        return ""
    return "> 图片转写：未转写"


def assemble(mat, course_dir, manifest_imgs):
    """用缓存的块JSON+图片记录组装草稿。返回(markdown, 页数, 失败页列表)"""
    out, failed_pages, n_pages = [], [], 0
    media_rel = f"../媒体/{mat['dirname']}"
    for ci, ch in enumerate(mat["chunks"]):
        if ch["status"] != "ok":
            for p in ch["pages"]:
                out.append(f"<!-- p.{p:03d} -->\n\n【资料缺失】本页解析失败，未能取得内容。")
                failed_pages.append(p)
            continue
        data = json.loads((course_dir / ch["cache"]).read_text(encoding="utf-8"))
        for pg in data["pages"]:
            orig = ch["pages"][pg["page_num"]]
            n_pages += 1
            pp = pg.get("printed_page")
            parts = [f"<!-- p.{orig:03d}" + (f" | 书p.{pp}" if pp else "") + " -->"]
            pic_i = 0

            def picture(b, orig=orig):
                nonlocal pic_i
                pic_i += 1
                key = f"p{orig:03d}_{pic_i:02d}"
                e = manifest_imgs.get(key)
                if not e or not e.get("file"):
                    return f"【图片缺失：{key}】"
                s = f"![{key}]({media_rel}/{e['file']})"
                tx = vlm_text_block(e)
                return s + ("\n\n" + tx if tx else "")

            ctx = {"picture": picture}
            for b in pg["blocks"]:
                s = render_block(b, ctx)
                if s:
                    parts.append(s)
            out.append("\n\n".join(parts))
    return "\n\n".join(out) + "\n", n_pages, failed_pages


# ---------------------------------------------------------------- 单份资料


def process_pdf_material(mat, src_pdf, pages, course_dir, work_dir, cfg, vlm, fitz, engine):
    ecfg = cfg[engine]
    parse = ENGINES[engine]
    media_dir = course_dir / "媒体" / mat["dirname"]
    cache_dir = course_dir / ".缓存" / mat["id"]
    cache_dir.mkdir(parents=True, exist_ok=True)
    max_pages = ecfg["max_pages_per_chunk"]

    def new_chunk(p):
        return {"pages": p, "status": "pending", "cache": f".缓存/{mat['id']}/p{p[0]:04d}.json", "error": None}

    if not mat.get("chunks"):
        pl = plan_chunks(fitz, src_pdf, pages, max_pages, int(ecfg["max_mb_per_chunk"] * 1024 * 1024 * 0.98), work_dir)
        mat["chunks"] = [new_chunk(p) for p in pl]
    else:  # 之前用别的引擎切的块，超过当前引擎上限的未完成块要重切
        new = []
        for ch in mat["chunks"]:
            if ch["status"] != "ok" and len(ch["pages"]) > max_pages:
                new += [new_chunk(ch["pages"][i:i + max_pages]) for i in range(0, len(ch["pages"]), max_pages)]
            else:
                new.append(ch)
        mat["chunks"] = new
    imgs = mat.setdefault("images", {})

    stop = False
    for ci, ch in enumerate(mat["chunks"]):
        cache_file = course_dir / ch["cache"]
        if ch["status"] == "ok" and cache_file.exists():
            continue
        if stop:
            ch["status"], ch["error"] = "failed", "已停止：前面的块遇到额度或鉴权错误"
            continue
        log(f"  {engine}解析：第{ch['pages'][0]}–{ch['pages'][-1]}页（块{ci + 1}/{len(mat['chunks'])}）")
        pdf_chunk = work_dir / f"chunk_{ci:02d}.pdf"
        t0 = time.time()
        try:
            make_chunk_pdf(fitz, src_pdf, ch["pages"], pdf_chunk)
            res = parse(ecfg, pdf_chunk)
            if len(res.get("pages", [])) != len(ch["pages"]):
                raise RuntimeError(f"返回{len(res.get('pages', []))}页，应为{len(ch['pages'])}页")
            cache_file.write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
            ch["status"], ch["error"] = "ok", None
            ch["engine"], ch["seconds"] = engine, round(time.time() - t0, 1)
        except FatalEngineError as e:
            ch["status"], ch["error"] = "failed", str(e)[:300]
            log(f"    失败：{ch['error']}")
            stop = True
        except Exception as e:  # noqa: BLE001
            ch["status"], ch["error"] = "failed", str(e)[:300]
            log(f"    失败：{ch['error']}")

    # 图片：下载并登记
    for ch in mat["chunks"]:
        if ch["status"] != "ok":
            continue
        cache_file = course_dir / ch["cache"]
        data = json.loads(cache_file.read_text(encoding="utf-8"))
        for pg in data["pages"]:
            orig = ch["pages"][pg["page_num"]]
            i = 0
            for b in pg["blocks"]:
                if (b.get("type") or "").lower() != "picture":
                    continue
                i += 1
                key = f"p{orig:03d}_{i:02d}"
                e = imgs.setdefault(key, {})
                if e.get("file") and (media_dir / e["file"]).exists():
                    continue
                url = b.get("img_url") or ""
                if not url:
                    m = re.search(r"https?://[^\s)\"']+", b.get("content") or "")
                    url = m.group(0) if m else ""
                if not url:
                    e["file"], e["error"] = None, "OXR未返回图片链接"
                    continue
                try:
                    dest = fetch_image(url, media_dir / key)
                    e["file"], e["error"] = dest.name, None
                except Exception as ex:  # noqa: BLE001
                    e["file"], e["error"] = None, f"下载失败：{str(ex)[:150]}"
    return None


def process_ipynb_material(mat, src, course_dir, cfg):
    """ipynb本地解析，图像输出另存到media。返回草稿正文，同时登记mat['images']"""
    media_dir = course_dir / "媒体" / mat["dirname"]
    nb = json.loads(src.read_text(encoding="utf-8"))
    cells = nb.get("cells", [])
    imgs = mat.setdefault("images", {})
    lang = nb.get("metadata", {}).get("language_info", {}).get("name", "python")
    mat["pages"] = len(cells)
    parts = []
    for ci, c in enumerate(cells, 1):
        src_text = "".join(c.get("source", []))
        parts.append(f"<!-- p.{ci:03d} -->")
        if c.get("cell_type") == "markdown":
            parts.append(src_text.strip())
        elif c.get("cell_type") == "code":
            parts.append(f"```{lang}\n{src_text.rstrip()}\n```")
            n = 0
            for o in c.get("outputs", []):
                if o.get("output_type") == "stream":
                    parts.append("输出：\n```\n" + "".join(o.get("text", [])).rstrip() + "\n```")
                    continue
                d = o.get("data", {})
                if "image/png" in d:
                    n += 1
                    key = f"p{ci:03d}_{n:02d}"
                    media_dir.mkdir(parents=True, exist_ok=True)
                    raw = d["image/png"]
                    raw = "".join(raw) if isinstance(raw, list) else raw
                    (media_dir / f"{key}.png").write_bytes(base64.b64decode(raw))
                    imgs.setdefault(key, {})["file"] = f"{key}.png"
                    parts.append(f"![{key}](../媒体/{mat['dirname']}/{key}.png)\n\n@@IMG:{key}@@")
                elif "text/plain" in d:
                    txt = d["text/plain"]
                    parts.append("输出：\n```\n" + ("".join(txt) if isinstance(txt, list) else txt).rstrip() + "\n```")
                elif o.get("output_type") == "error":
                    parts.append("报错：\n```\n" + "\n".join(o.get("traceback", []))[:1500] + "\n```")
    return "\n\n".join(p for p in parts if p) + "\n"


def run_vlm(mat, course_dir, vlm, manifest, cfg):
    media_dir = course_dir / "媒体" / mat["dirname"]
    cache = manifest.setdefault("vlm_cache", {})
    min_px = cfg["vlm"].get("min_image_pixels", 64)
    for key, e in mat.get("images", {}).items():
        if not e.get("file"):
            continue
        if e.get("vlm") in ("ok", "skipped"):
            continue
        path = media_dir / e["file"]
        if too_small(path, min_px):
            e["vlm"] = "skipped"
            continue
        if not vlm.enabled:
            e["vlm"] = "pending"
            continue
        h = hashlib.sha256(path.read_bytes()).hexdigest()
        e["hash"] = h
        if h in cache:
            e["vlm"], e["text"] = "ok", cache[h]
            continue
        log(f"  VLM转写：{key}")
        try:
            txt = vlm.transcribe(path)
            if not txt:
                raise RuntimeError("空结果")
            cache[h] = txt
            e["vlm"], e["text"] = "ok", txt
        except Exception as ex:  # noqa: BLE001
            e["vlm"] = "failed"
            e["vlm_error"] = str(ex)[:150]


def verify_and_summarize(mat, course_dir, expected_pages):
    from PIL import Image
    media_dir = course_dir / "媒体" / mat["dirname"]
    problems = []
    imgs = mat.get("images", {})
    broken = 0
    for key, e in imgs.items():
        if not e.get("file"):
            broken += 1
            continue
        try:
            with Image.open(media_dir / e["file"]) as im:
                im.verify()
        except Exception:  # noqa: BLE001
            broken += 1
            e["error"] = "图片无法打开"
    failed_pages = sorted(p for ch in mat["chunks"] if ch["status"] != "ok" for p in ch["pages"])
    untranscribed = sum(1 for e in imgs.values() if e.get("vlm") in ("pending", "failed"))
    transcribed = sum(1 for e in imgs.values() if e.get("vlm") == "ok")
    vlm_failed = sum(1 for e in imgs.values() if e.get("vlm") == "failed")
    status = "ok"
    if failed_pages or broken or vlm_failed:
        status = "partial" if (len(failed_pages) < expected_pages) else "failed"
    mat["status"] = status
    return {
        "id": mat["id"], "name": mat["name"], "type": TYPE_NAMES.get(mat["type"], mat["type"]),
        "status": status, "pages": expected_pages, "failed_pages": ranges(failed_pages),
        "images": len(imgs), "images_broken": broken,
        "transcribed": transcribed, "untranscribed": untranscribed, "vlm_failed": vlm_failed,
        "draft": f"草稿/{mat['dirname']}.md",
        "page_offset": estimate_offset(mat, course_dir),
        "engines": sorted({c.get("engine", "oxr") for c in mat["chunks"] if c["status"] == "ok"}),
        "ocr_seconds": round(sum(c.get("seconds", 0) for c in mat["chunks"])),
        "problems": problems,
    }


def estimate_offset(mat, course_dir):
    """用各页识别到的书页码（number块）估计“PDF页码-书页码”的偏移。只统计书页码是纯数字的页"""
    from collections import Counter
    cnt = Counter()
    for ch in mat.get("chunks", []):
        if ch["status"] != "ok":
            continue
        f = course_dir / ch["cache"]
        if not f.exists():
            continue
        for pg in json.loads(f.read_text(encoding="utf-8"))["pages"]:
            pp = pg.get("printed_page")
            if pp and pp.isdigit():
                cnt[ch["pages"][pg["page_num"]] - int(pp)] += 1
    if not cnt:
        return None
    off, n = cnt.most_common(1)[0]
    return {"offset": off, "pages_agree": n, "pages_with_number": sum(cnt.values())}


def ranges(pages):
    if not pages:
        return []
    out, s, p = [], pages[0], pages[0]
    for x in pages[1:]:
        if x == p + 1:
            p = x
            continue
        out.append(f"{s}-{p}" if s != p else str(s))
        s = p = x
    out.append(f"{s}-{p}" if s != p else str(s))
    return out


# ---------------------------------------------------------------- 主流程


def main():
    ap = argparse.ArgumentParser(description="课件、资料转Markdown草稿（OXR或Paddle，加VLM）")
    ap.add_argument("files", nargs="*", help="一个或多个文件，按传入顺序编号；--init-config时不需要")
    ap.add_argument("--course-dir", help="课程文件夹，如study-buddy-data/0101_微观经济学；--init-config时不需要")
    ap.add_argument("--data-dir", help="数据目录，放config.json；默认为课程文件夹的上一级，--init-config时默认为./study-buddy-data")
    ap.add_argument("--init-config", action="store_true",
                    help="只在数据目录里生成config.json（已存在则不动）然后退出，用于首次使用时先让用户填key")
    ap.add_argument("--types", default="material",
                    help="material/textbook/past-exam/assignment，逗号分隔与files一一对应，或只写一个应用到全部")
    ap.add_argument("--pages", help="只处理这些页（1起，如10,17,20-22），用于测试；多份文件时应用到每份")
    ap.add_argument("--engine", choices=sorted(ENGINES), help="OCR引擎；默认取config.json的engine，没写则用paddle，paddle没填key时退用oxr")
    ap.add_argument("--keep-temp", action="store_true", help="保留临时PDF和切块")
    args = ap.parse_args()

    if args.init_config:
        data_dir = Path(args.data_dir).resolve() if args.data_dir else Path("study-buddy-data").resolve()
        _, created = load_config(data_dir)
        print(json.dumps({"data_dir": str(data_dir), "config": str(data_dir / "config.json"), "created": created},
                         ensure_ascii=False, indent=2))
        sys.exit(0)
    if not args.files or not args.course_dir:
        fail_env("需要传入资料文件和--course-dir（只生成配置文件时用--init-config）")

    course_dir = Path(args.course_dir).resolve()
    data_dir = Path(args.data_dir).resolve() if args.data_dir else course_dir.parent
    files = [Path(f).resolve() for f in args.files]
    for f in files:
        if not f.exists():
            fail_env(f"文件不存在：{f}")
        if f.suffix.lower() not in OFFICE_PPT | OFFICE_DOC | {".pdf", ".ipynb"}:
            fail_env(f"不支持的格式：{f.suffix}（支持ppt/pptx/doc/docx/pdf/ipynb）")
    types = [t.strip() for t in args.types.split(",")]
    if len(types) == 1:
        types *= len(files)
    if len(types) != len(files) or any(t not in TYPE_NAMES for t in types):
        fail_env("--types应为material/textbook/past-exam/assignment，数量与files一致或只写一个")

    try:
        import pymupdf as fitz
        import requests  # noqa: F401
        from PIL import Image  # noqa: F401
    except ImportError as e:
        fail_env(f"缺少依赖：{e.name}，见references/setup.md")

    cfg, created = load_config(data_dir)
    engine = args.engine or cfg.get("engine") or "paddle"
    has_key = bool(cfg["paddle"].get("api_key"))
    warnings = []
    if engine == "paddle" and not has_key:
        if args.engine == "paddle":  # 用户显式要求paddle但没有key，不悄悄换引擎
            fail_env(f"使用paddle需要key：到 https://aistudio.baidu.com/paddleocr 申请，填进{data_dir / 'config.json'}的paddle.api_key")
        engine = "oxr"
        warnings.append("默认引擎是paddle，但paddle.api_key还没填，本次退用oxr（慢、同一时间只处理一个任务）；"
                        "到 https://aistudio.baidu.com/paddleocr 申请key填进config.json即可")
    vlm = Vlm(cfg["vlm"])
    if created:
        warnings.append(f"已生成{data_dir / 'config.json'}，请填写vlm的base_url、model、api_key")
    if not vlm.enabled:
        warnings.append("VLM未配置，图片全部标为“未转写”；填好config.json后重新运行即可补转写")

    course_dir.mkdir(parents=True, exist_ok=True)
    for sub in ("原始资料", "课程资料", "草稿", "媒体", "讲义"):
        (course_dir / sub).mkdir(exist_ok=True)
    mpath = course_dir / "manifest.json"
    manifest = json.loads(mpath.read_text(encoding="utf-8")) if mpath.exists() else {"version": 1, "materials": []}
    sel = parse_pages(args.pages) if args.pages else None

    summary, any_problem = [], False
    for f, tp in zip(files, types):
        mat = next((m for m in manifest["materials"] if m["source"] == str(f)), None)
        if mat is None:
            nid = max([int(m["id"]) for m in manifest["materials"]] + [0]) + 1
            mat = {"id": f"{nid:03d}", "name": f.stem.strip(), "type": tp, "source": str(f), "chunks": [], "images": {}}
            mat["dirname"] = f"{mat['id']}_{safe_name(f.stem)}"
            manifest["materials"].append(mat)
        mat["type"] = tp
        log(f"[{mat['id']}] {f.name}（{TYPE_NAMES[tp]}）")
        draft_path = course_dir / "草稿" / f"{mat['dirname']}.md"
        work_dir = course_dir / ".临时" / mat["id"]
        work_dir.mkdir(parents=True, exist_ok=True)

        stale = {"failed", "pending"} if vlm.enabled else {"failed"}
        done = (mat.get("status") == "ok" and draft_path.exists()
                and not any(e.get("vlm") in stale for e in mat.get("images", {}).values()))
        if done:
            log("  已完成，跳过")
            summary.append({"id": mat["id"], "name": mat["name"], "status": "skipped（已完成）",
                            "draft": f"草稿/{mat['dirname']}.md"})
            continue

        try:
            if f.suffix.lower() == ".ipynb":
                body = process_ipynb_material(mat, f, course_dir, cfg)
                run_vlm(mat, course_dir, vlm, manifest, cfg)
                for key, e in mat["images"].items():
                    body = body.replace(f"@@IMG:{key}@@", vlm_text_block(e))
                draft_path.write_text(body, encoding="utf-8")
                mat["chunks"] = []
                expected = mat["pages"]
                info = verify_and_summarize(mat, course_dir, expected)
            else:
                if f.suffix.lower() == ".pdf":
                    src_pdf = f
                else:
                    src_pdf = work_dir / (f.stem + ".pdf")
                    if not src_pdf.exists():
                        log("  转PDF（Office COM）")
                        office_to_pdf(f, src_pdf)
                doc = fitz.open(str(src_pdf))
                total = doc.page_count
                doc.close()
                pages = [p for p in (sel or range(1, total + 1)) if 1 <= p <= total]
                if not pages:
                    raise RuntimeError(f"没有可处理的页（共{total}页）")
                mat["pages"], mat["selected_pages"] = len(pages), (sel and pages)
                process_pdf_material(mat, src_pdf, pages, course_dir, work_dir, cfg, vlm, fitz, engine)
                run_vlm(mat, course_dir, vlm, manifest, cfg)
                md, n_pages, _ = assemble(mat, course_dir, mat["images"])
                draft_path.write_text(md, encoding="utf-8")
                info = verify_and_summarize(mat, course_dir, len(pages))
        except SystemExit:
            raise
        except Exception as e:  # noqa: BLE001
            mat["status"] = "failed"
            info = {"id": mat["id"], "name": mat["name"], "status": "failed", "error": str(e)[:300]}
            log(f"  失败：{info['error']}")

        if info["status"] == "ok" and not args.keep_temp:
            shutil.rmtree(work_dir, ignore_errors=True)
        if info["status"] != "ok":
            any_problem = True
        summary.append(info)
        mpath.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")

    mpath.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    work_root = course_dir / ".临时"
    if work_root.exists() and not any(work_root.iterdir()):
        work_root.rmdir()
    out = {"course_dir": str(course_dir), "warnings": warnings, "materials": summary}
    print(json.dumps(out, ensure_ascii=False, indent=2))
    sys.exit(1 if any_problem else 0)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    try:
        main()
        code = 0
    except SystemExit as e:
        code = e.code
    # 部分失败（退出码1）也提示：重跑前更新可能正好修掉问题
    if code in (None, 0, 1):
        try:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            from check_update import notify_if_update
            notify_if_update()
        except Exception:
            pass
    sys.exit(code)
