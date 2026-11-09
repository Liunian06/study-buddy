#!/usr/bin/env python3
"""把课件、参考资料、往年题、作业要求转成带页码和图片转写的Markdown草稿。

流程：Office转PDF（Windows用Office COM，没有则LibreOffice；Mac、Linux用LibreOffice）、切分、上传OCR引擎（OXR、Paddle或MinerU）、按页组装、下载图片、VLM转写图片、合并、校验。
用--engine选择引擎，默认取config.json里的engine，没写则用paddle（PaddleOCR-VL-1.6）。
优先级：paddle key → oxr key（商业版）→ mineru key（精准解析）→ oxr免费接口 → mineru免登录轻量接口。paddle没填key时自动在后几档里选；mineru失败时退用oxr（有oxr key用商业版）。
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
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from concurrent.futures import wait as futures_wait
from pathlib import Path

TYPE_NAMES = {"material": "课件/参考资料", "textbook": "教材", "past-exam": "往年题", "assignment": "作业要求"}
OFFICE_PPT = {".ppt", ".pptx"}
OFFICE_DOC = {".doc", ".docx"}
SKILL_DIR = Path(__file__).resolve().parent.parent
SKIP_MARK = "[跳过："


def log(msg):
    print(msg, file=sys.stderr, flush=True)


HEARTBEAT_SECONDS = 60  # 耗时较长的操作期间，每隔这么久打一行“仍在处理”


def progress(msg):
    """进度行统一加[进度]前缀，方便助手只取这几行转告用户"""
    log(f"[进度] {msg}")


def fmt_dur(sec):
    sec = int(round(sec))
    if sec < 60:
        return f"{sec}秒"
    m, s = divmod(sec, 60)
    if m < 60:
        return f"{m}分{s}秒" if s else f"{m}分钟"
    h, m = divmod(m, 60)
    return f"{h}小时{m}分"


class Heartbeat:
    """with块里的操作耗时较长（OCR排队、轮询）时，每隔一段时间打一行“仍在处理”，让人看得出没有卡住。
    用独立线程打印，所以不用改各引擎的函数"""

    def __init__(self, label):
        self.label = label
        self._stop = threading.Event()
        self._t = None

    def __enter__(self):
        t0, interval = time.time(), HEARTBEAT_SECONDS

        def beat():
            while not self._stop.wait(interval):
                progress(f"{self.label}仍在处理，已等待{fmt_dur(time.time() - t0)}")
        self._t = threading.Thread(target=beat, daemon=True)
        self._t.start()
        return self

    def __exit__(self, *exc):
        self._stop.set()
        self._t.join(timeout=2)
        return False


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

PLACEHOLDER_KEYS = {"sk-123", "paddle-token-123", "MinerU_API_KEY", "OXR_API_KEY", "dummy"}  # config.template.json里的示例key；dummy是旧版oxr免费接口的占位值

def load_config(data_dir, create=True):
    """create=False时config.json不存在也不生成文件，直接按模板默认值返回（--dry-run用）"""
    cfg_path = data_dir / "config.json"
    created = False
    if not cfg_path.exists():
        tpl = SKILL_DIR / "assets" / "templates" / "config.template.json"
        if create:
            data_dir.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(tpl, cfg_path)
            created = True
        else:
            cfg_path = tpl
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
    for sec in ("paddle", "oxr", "mineru", "vlm"):
        if cfg[sec].get("api_key") in PLACEHOLDER_KEYS:
            cfg[sec]["api_key"] = ""
    return cfg, created


# ---------------------------------------------------------------- 转换与切分


def office_to_pdf(src, out_pdf):
    """Office转PDF：Windows优先Office COM，没有就退到LibreOffice；macOS、Linux直接LibreOffice。细节见office_compat.py"""
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from office_compat import OfficeUnavailable, convert_to_pdf
    try:
        convert_to_pdf(src, out_pdf, log)
    except OfficeUnavailable as e:
        fail_env(f"无法转换{src.name}：{e}；也可以手动另存为PDF再交给脚本，见references/setup.md")


def cached_office_pdf(src, course_dir):
    """Office文件转PDF，结果缓存在课程文件夹的.临时/预转换/。
    缓存键含路径、修改时间和大小，文件改过就重新转；--dry-run和正式运行共用，避免转两遍"""
    st = src.stat()
    key = hashlib.sha1(f"{src.resolve()}|{st.st_mtime_ns}|{st.st_size}".encode("utf-8")).hexdigest()[:16]
    pdf = course_dir / ".临时" / "预转换" / f"{key}.pdf"
    if pdf.exists() and pdf.stat().st_size > 0:
        log(f"  复用已转换的PDF：{src.name}")
        return pdf
    pdf.parent.mkdir(parents=True, exist_ok=True)
    log(f"  转PDF：{src.name}")
    tmp = pdf.with_suffix(".part.pdf")
    office_to_pdf(src, tmp)
    tmp.replace(pdf)  # 转换成功才落盘，中断不会留下半截缓存
    return pdf


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


def build_chain(cfg):
    """引擎档位，按优先级从高到低：paddle key → oxr key → mineru key → oxr免费接口 → mineru免登录轻量接口。
    paddle、oxr、mineru是三家不同的供应商，key互不通用；每档由“供应商 + 有没有自己的key”决定。
    没有key的档位用api_key置空的配置副本，run_engine/oxr_parse看到空key就走免费或免登录的路径。
    oxr免费接口和mineru轻量接口不需要key，永远可用；paddle没填key则没有这一档"""
    def blank(sec):
        return {**cfg[sec], "api_key": ""}

    chain = []
    if cfg["paddle"].get("api_key"):
        chain.append({"engine": "paddle", "cfg": cfg["paddle"], "label": "paddle"})
    if cfg["oxr"].get("api_key"):
        chain.append({"engine": "oxr", "cfg": cfg["oxr"], "label": "oxr商业版"})
    if cfg["mineru"].get("api_key"):
        chain.append({"engine": "mineru", "cfg": cfg["mineru"], "label": "mineru精准解析"})
    chain.append({"engine": "oxr", "cfg": blank("oxr"), "label": "oxr免费接口"})
    chain.append({"engine": "mineru", "cfg": blank("mineru"), "label": "mineru轻量接口"})
    return chain


def parse_with_tier(fitz, src_pdf, pages, tier, course_dir, mat_id, work_dir, tag):
    """用指定档位解析一个块。各档的单块上限不同（例如oxr每块只有20页、10MB，比paddle、mineru精准解析的块小），
    所以按这一档的上限重新拆，各小块的结果按顺序拼回，页号重新编成0..n-1，和原来的块对齐"""
    max_pages, max_bytes = chunk_limits(tier["engine"], tier["cfg"])
    subs = plan_chunks(fitz, src_pdf, pages, max_pages, max_bytes, work_dir)
    merged = []
    for si, sub in enumerate(subs):
        part = work_dir / f"fallback_{tag}_{si:02d}.pdf"
        make_chunk_pdf(fitz, src_pdf, sub, part)
        res = run_engine(tier["engine"], tier["cfg"], part, course_dir, mat_id, f"{tag}f{si}")
        part.unlink(missing_ok=True)
        if len(res.get("pages", [])) != len(sub):
            raise RuntimeError(f"{tier['label']}返回{len(res.get('pages', []))}页，应为{len(sub)}页")
        for pg in res["pages"]:
            if pg.get("unpaged"):  # 没有分页信息的输出，合并范围是这个小块而不是整块
                pg["unpaged_range"] = [sub[0], sub[-1]]
        merged.extend(res["pages"])
    for i, pg in enumerate(merged):
        pg["page_num"] = i
    return {"pages": merged}


# ---------------------------------------------------------------- OXR


class FatalEngineError(Exception):
    """额度用尽、鉴权失败等重试也没用的错误，遇到后停止提交后面的块"""


def oxr_parse(cfg, pdf_path):
    """OXR同步接口。429表示免费接口上已有任务，等久一点再试；网络错误短间隔重试"""
    import requests
    key = cfg.get("api_key") or ""  # 有key走商业版地址，没有走免费接口（api_key传dummy）
    base = (cfg.get("paid_base_url") or "https://somark.ai/api/v1") if key else cfg["base_url"]
    url = base.rstrip("/") + "/parse/sync"
    last = None
    for attempt in range(6):
        try:
            with open(pdf_path, "rb") as fh:
                r = requests.post(
                    url,
                    data=[("api_key", key or "dummy"), ("output_formats", "json")],
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


ENGINES = {"oxr": oxr_parse, "paddle": paddle_parse}  # mineru见下，签名不同，由run_engine分发

# ---------------------------------------------------------------- MinerU

_MINERU_FATAL = {"A0202", "A0211", -60018}  # token错误、token过期、每日任务数已达上限：重试没用
_MINERU_SKIP = {"header", "footer", "page_number", "aside_text"}  # 不进正文；page_number单独存成页码
_MINERU_PLACEHOLDER = re.compile(r"^<[a-z]+>$")  # 如PPT页脚里的“<number>”占位符，不是正文


def _mineru_json(r):
    """检查一次MinerU响应，返回JSON。鉴权失败、额度用尽抛FatalEngineError，其余错误抛RuntimeError供重试"""
    if r.status_code in (401, 403):
        raise FatalEngineError(f"HTTP {r.status_code}（token无效或已过期）：{r.text[:150]}")
    if r.status_code != 200:
        raise RuntimeError(f"HTTP {r.status_code}：{r.text[:150]}")
    body = r.json()
    code = body.get("code", body.get("msgCode"))
    if code == 0:
        return body
    if code in _MINERU_FATAL:
        raise FatalEngineError(f"MinerU返回{code}：{body.get('msg')}")
    raise RuntimeError(f"MinerU返回{code}：{body.get('msg')}")


def _mineru_emphasis(text):
    """MinerU的文本里混用<strong>/<em>标签和**；统一成Markdown"""
    return (text.replace("<strong>", "**").replace("</strong>", "**")
            .replace("<em>", "*").replace("</em>", "*"))


def mineru_items_to_pages(items, n_pages, img_rel):
    """content_list.json（按阅读顺序的扁平列表，每项带page_idx，从0起）转成与其他引擎一致的{"pages": [...]}。
    页数以块PDF的真实页数为准，没有内容的页也返回一个空页，保证和切块的页数一一对应"""
    pages = [{"page_num": i, "blocks": [], "printed_page": None} for i in range(n_pages)]
    for it in items:
        idx = it.get("page_idx")
        if not isinstance(idx, int) or not 0 <= idx < n_pages:
            continue
        pg, blocks, t = pages[idx], pages[idx]["blocks"], it.get("type")
        text = _mineru_emphasis((it.get("text") or "").strip())
        if t == "page_number":
            pg["printed_page"] = pg["printed_page"] or text or None
        elif t in _MINERU_SKIP:
            continue
        elif t == "text" or (t is None and text):
            if text and it.get("text_level"):
                blocks.append({"type": "Title", "format": "text", "content": text, "label": "title"})
            elif text:
                fmt = "latex" if it.get("text_format") == "latex" else "markdown"
                blocks.append({"type": "Formula" if fmt == "latex" else "Text", "format": fmt, "content": text})
        elif t == "equation":
            if text:
                blocks.append({"type": "Formula", "format": "latex", "content": text})
        elif t in ("image", "chart") or (it.get("img_path") and t not in ("table", "list", "code")):
            # chart（图表，如分布曲线图、柱状图）和image一样当图片；其他带图的未知类型也留着，宁可多不能漏
            if it.get("img_path"):
                blocks.append({"type": "Picture", "img_local": f"{img_rel}/{Path(it['img_path']).name}"})
            caps = (it.get(f"{t}_caption") or []) + (it.get(f"{t}_footnote") or [])
            if t == "chart" and (it.get("content") or "").strip():
                caps.append(it["content"].strip())  # 图表识别出的文字（如有）
            for cap in caps:
                if cap.strip():
                    blocks.append({"type": "Text", "format": "markdown", "content": _mineru_emphasis(cap.strip())})
        elif t == "table":
            for cap in it.get("table_caption") or []:
                if cap.strip():
                    blocks.append({"type": "Text", "format": "markdown", "content": _mineru_emphasis(cap.strip())})
            if (it.get("table_body") or "").strip():
                blocks.append({"type": "Text", "format": "markdown", "content": it["table_body"].strip()})
            elif it.get("img_path"):  # 没有结构化的表格内容，只有表格截图：当图片处理，交给VLM转写
                blocks.append({"type": "Picture", "img_local": f"{img_rel}/{Path(it['img_path']).name}"})
            for cap in it.get("table_footnote") or []:
                if cap.strip():
                    blocks.append({"type": "Text", "format": "markdown", "content": _mineru_emphasis(cap.strip())})
        elif t == "list":
            lines = [_mineru_emphasis(x.strip()) for x in it.get("list_items") or [] if x.strip()]
            if lines:
                blocks.append({"type": "Text", "format": "markdown", "content": "\n".join(lines)})
        elif t == "code":
            body = (it.get("code_body") or "").rstrip()
            if body:
                blocks.append({"type": "Text", "format": "markdown", "content": f"```\n{body}\n```"})
        elif text and not _MINERU_PLACEHOLDER.match(text):  # 不认识的类型：有文字就留着，宁可多不能漏
            blocks.append({"type": "Text", "format": "markdown", "content": text})
    return {"pages": pages}


def mineru_parse(cfg, pdf_path, img_dir=None, img_rel=None):
    """MinerU精准解析（批量上传接口，一次一个文件）：申请上传链接、PUT上传、轮询、下载结果包。
    结果包里的content_list.json转成与其他引擎一致的{"pages": [...]}，图片解压到img_dir，
    块里记相对课程文件夹的路径img_rel/<文件名>（缓存里不存绝对路径，课程文件夹搬家后仍可用）"""
    import requests
    import zipfile
    base = cfg["base_url"].rstrip("/")
    headers = {"Content-Type": "application/json", "Authorization": f"Bearer {cfg['api_key']}"}
    img_dir = Path(img_dir) if img_dir else pdf_path.parent / f"{pdf_path.stem}_img"
    img_rel = img_rel or img_dir.name
    timeout = cfg.get("timeout_seconds", 300)

    upload_url, batch_id, last = None, None, None
    for attempt in range(6):  # 申请上传链接：限流等待，网络错误重试
        try:
            r = requests.post(f"{base}/api/v4/file-urls/batch", headers=headers, timeout=timeout,
                              json={"files": [{"name": pdf_path.name, "data_id": pdf_path.stem[:100]}],
                                    "model_version": cfg.get("model_version", "vlm")})
            if r.status_code == 429:
                last = RuntimeError("HTTP 429：请求太频繁")
                time.sleep(cfg.get("retry_429_seconds", 30))
                continue
            data = _mineru_json(r)["data"]
            upload_url, batch_id = data["file_urls"][0], data["batch_id"]
            break
        except FatalEngineError:
            raise
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (attempt + 1))
    if not batch_id:
        raise RuntimeError(f"提交未成功：{last}")
    for attempt in range(4):  # 上传（PUT，不设Content-Type）
        try:
            with open(pdf_path, "rb") as fh:
                pr = requests.put(upload_url, data=fh, timeout=timeout)
            if pr.status_code != 200:
                raise RuntimeError(f"上传失败 HTTP {pr.status_code}")
            break
        except Exception as e:  # noqa: BLE001
            last = e
            if attempt == 3:
                raise RuntimeError(f"上传未成功：{last}")
            time.sleep(3 * (attempt + 1))

    deadline = time.time() + cfg.get("job_timeout_seconds", 1800)
    zip_url, fails = None, 0
    while time.time() < deadline:  # 轮询：单次查询失败只重试这一次查询
        time.sleep(cfg.get("poll_seconds", 3))
        try:
            rr = requests.get(f"{base}/api/v4/extract-results/batch/{batch_id}", headers=headers, timeout=60)
            res = (_mineru_json(rr).get("data") or {}).get("extract_result") or []
            fails = 0
        except FatalEngineError:
            raise
        except Exception as e:  # noqa: BLE001
            fails += 1
            if fails >= 10:
                raise RuntimeError(f"连续查询失败：{e}")
            continue
        if not res:
            continue
        st = res[0].get("state")
        if st == "done":
            zip_url = res[0].get("full_zip_url")
            break
        if st == "failed":
            raise RuntimeError(f"任务失败：{res[0].get('err_msg')}")
    if not zip_url:
        raise RuntimeError("任务超时")
    blob = None
    for attempt in range(4):
        try:
            zr = requests.get(zip_url, timeout=120)
            zr.raise_for_status()
            blob = zr.content
            break
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (attempt + 1))
    if blob is None:
        raise RuntimeError(f"下载结果失败：{last}")

    try:
        zf = zipfile.ZipFile(io.BytesIO(blob))
    except zipfile.BadZipFile as e:
        raise RuntimeError("结果包损坏") from e
    with zf:
        names = zf.namelist()
        cl = next((n for n in names if n.endswith("_content_list.json")), None)
        if cl is None:
            raise RuntimeError("结果包里没有content_list.json")
        items = json.loads(zf.read(cl).decode("utf-8"))
        img_dir.mkdir(parents=True, exist_ok=True)
        for n in names:
            if n.startswith("images/") and not n.endswith("/"):
                (img_dir / Path(n).name).write_bytes(zf.read(n))  # 只取文件名，不信任压缩包里的路径
        lay = next((n for n in names if n.endswith("layout.json")), None)
        got = len(json.loads(zf.read(lay).decode("utf-8")).get("pdf_info") or []) if lay else None
    import pymupdf
    with pymupdf.open(str(pdf_path)) as doc:
        n_pages = doc.page_count
    if got is not None and got != n_pages:
        raise RuntimeError(f"返回{got}页，应为{n_pages}页")
    return mineru_items_to_pages(items, n_pages, img_rel)


class _NoRetry(RuntimeError):
    """重试也不会变好的错误（超出大小、页数、文件类型不支持等）"""


def _agent_markdown_to_blocks(md):
    """轻量接口只给一份Markdown：按图片语法拆成文字块和图片块。图片是http链接才能下载，相对路径拿不到图"""
    blocks, pos = [], 0
    for m in _IMG_MD.finditer(md):
        seg = md[pos:m.start()].strip()
        if seg:
            blocks.append({"type": "Text", "format": "markdown", "content": _mineru_emphasis(seg)})
        url = m.group(1)
        blocks.append({"type": "Picture", "img_url": url if url.startswith("http") else ""})
        pos = m.end()
    seg = md[pos:].strip()
    if seg:
        blocks.append({"type": "Text", "format": "markdown", "content": _mineru_emphasis(seg)})
    return blocks


def mineru_agent_parse(cfg, pdf_path, img_dir=None, img_rel=None):
    """MinerU Agent轻量接口：不需要token，按IP限频，固定用pipeline模型。每份最多20页、10MB，一次一个文件，
    只返回一份Markdown，没有分页信息。所以整块内容放在块的第一页，其余页留空，assemble在那里标注
    “本块合并输出”，页码只能标到块级。参数img_dir、img_rel与mineru_parse对齐，这里用不上"""
    import requests
    base = cfg["base_url"].rstrip("/") + "/api/v1/agent"
    timeout = cfg.get("timeout_seconds", 300)
    task_id = upload_url = last = None
    for attempt in range(6):  # 申请上传链接：限频等待，网络错误重试
        try:
            r = requests.post(f"{base}/parse/file", json={"file_name": pdf_path.name, "language": "ch"},
                              timeout=timeout)
            if r.status_code == 429:
                last = RuntimeError("HTTP 429：轻量接口限频")
                time.sleep(cfg.get("retry_429_seconds", 30))
                continue
            if r.status_code != 200:
                raise RuntimeError(f"HTTP {r.status_code}：{r.text[:150]}")
            body = r.json()
            code = body.get("code")
            if code in (-30001, -30002, -30003, -30004):  # 超10MB、类型不支持、超页数、参数错误
                raise _NoRetry(f"轻量接口返回{code}：{body.get('msg')}")
            if code != 0:
                raise RuntimeError(f"轻量接口返回{code}：{body.get('msg')}")
            task_id, upload_url = body["data"]["task_id"], body["data"]["file_url"]
            break
        except _NoRetry:
            raise
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (attempt + 1))
    if not task_id:
        raise RuntimeError(f"提交未成功：{last}")
    for attempt in range(4):  # 上传（PUT）
        try:
            with open(pdf_path, "rb") as fh:
                pr = requests.put(upload_url, data=fh, timeout=timeout)
            if pr.status_code != 200:
                raise RuntimeError(f"上传失败 HTTP {pr.status_code}")
            break
        except Exception as e:  # noqa: BLE001
            last = e
            if attempt == 3:
                raise RuntimeError(f"上传未成功：{last}")
            time.sleep(3 * (attempt + 1))

    # 轻量接口没有队列优先级，排队时间不可控，等待上限单独设，超过就当作失败，由上层退到下一档（它是链上最后一档，则这一块失败）
    deadline = time.time() + cfg.get("agent_job_timeout_seconds", 120)
    md_url, fails = None, 0
    while time.time() < deadline:
        time.sleep(cfg.get("poll_seconds", 3))
        try:
            g = requests.get(f"{base}/parse/{task_id}", timeout=60)
            if g.status_code != 200:
                raise RuntimeError(f"查询 HTTP {g.status_code}")
            d = g.json().get("data") or {}
            fails = 0
        except Exception as e:  # noqa: BLE001
            fails += 1
            if fails >= 10:
                raise RuntimeError(f"连续查询失败：{e}")
            continue
        if d.get("state") == "done":
            md_url = d.get("markdown_url")
            break
        if d.get("state") == "failed":
            raise RuntimeError(f"任务失败：{d.get('err_code')} {d.get('err_msg')}")
    if not md_url:
        raise RuntimeError("轻量接口排队或处理超时")
    md = None
    for attempt in range(4):
        try:
            mr = requests.get(md_url, timeout=120)
            mr.raise_for_status()
            md = mr.content.decode("utf-8")
            break
        except Exception as e:  # noqa: BLE001
            last = e
            time.sleep(3 * (attempt + 1))
    if md is None:
        raise RuntimeError(f"下载结果失败：{last}")
    import pymupdf
    with pymupdf.open(str(pdf_path)) as doc:
        n_pages = doc.page_count
    pages = [{"page_num": i, "blocks": [], "printed_page": None} for i in range(n_pages)]
    pages[0]["blocks"] = _agent_markdown_to_blocks(md)
    if n_pages > 1:
        pages[0]["unpaged"] = True
    return {"pages": pages}


ENGINES["mineru"] = mineru_parse


def run_engine(engine, ecfg, pdf, course_dir, mat_id, tag):
    """调用引擎解析一个块。mineru的图片随结果包下来，要给它一个存放目录（.缓存/<资料序号>/<tag>_img/）"""
    if engine == "mineru":
        rel = f".缓存/{mat_id}/{tag}_img"
        if ecfg.get("api_key"):  # 有token：精准解析，vlm模型，有分页和图片
            return mineru_parse(ecfg, pdf, course_dir / rel, rel)
        return mineru_agent_parse(ecfg, pdf, course_dir / rel, rel)  # 没token：轻量接口，pipeline模型
    return ENGINES[engine](ecfg, pdf)


AGENT_MAX_PAGES, AGENT_MAX_MB = 20, 10  # MinerU轻量接口的单文件上限


def chunk_limits(engine, ecfg):
    """切块上限（页数，字节数）。mineru没填token走轻量接口时，上限收紧到20页、10MB"""
    pages, mb = ecfg["max_pages_per_chunk"], ecfg["max_mb_per_chunk"]
    if engine == "mineru" and not ecfg.get("api_key"):
        pages, mb = min(pages, AGENT_MAX_PAGES), min(mb, AGENT_MAX_MB)
    return pages, int(mb * 1024 * 1024 * 0.98)


DL_PERMANENT = {401, 403, 404, 410}  # 这几种状态说明链接已失效，重试没有意义，直接补OCR换新链接


class ImageFetchError(Exception):
    """图片下载失败。permanent为真表示链接已失效（403、404等），重试无用"""

    def __init__(self, msg, permanent=False):
        super().__init__(msg)
        self.permanent = permanent


LOCAL_PREFIX = "local:"  # MinerU的图片随结果包一起下来，存在课程文件夹的.缓存/里，“下载”就是复制并校验


def _fetch_once(url, dest_base):
    import requests
    from PIL import Image
    if url.startswith(LOCAL_PREFIX):
        src = Path(url[len(LOCAL_PREFIX):])
        if not src.exists():  # 缓存里的图被删了：重试没用，走补OCR重新生成
            raise ImageFetchError(f"本地缓存图片不存在：{src.name}", permanent=True)
        content = src.read_bytes()
    else:
        try:
            r = requests.get(url, timeout=60)
        except requests.RequestException as e:
            raise ImageFetchError(f"网络错误：{type(e).__name__}") from e
        if r.status_code in DL_PERMANENT:
            raise ImageFetchError(f"HTTP {r.status_code}（链接已失效）", permanent=True)
        if r.status_code != 200:
            raise ImageFetchError(f"HTTP {r.status_code}")
        content = r.content
    try:
        img = Image.open(io.BytesIO(content))
        img.verify()
        fmt = img.format
    except Exception as e:  # noqa: BLE001
        raise ImageFetchError("下载内容不是有效图片") from e
    ext = {"PNG": ".png", "JPEG": ".jpg", "GIF": ".gif", "WEBP": ".webp", "BMP": ".bmp"}.get(fmt, ".png")
    dest = dest_base.with_suffix(ext)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(content)
    return dest


def fetch_image(url, dest_base, attempts=3, wait=2):
    """下载图片，按真实格式确定扩展名。返回(保存的路径, 重试次数)。
    超时、5xx、连接错误等重试，间隔wait、2*wait秒递增；403、404等链接失效直接抛出，不重试"""
    last = None
    for i in range(max(1, attempts)):
        try:
            return _fetch_once(url, dest_base), i
        except ImageFetchError as e:
            last = e
            if e.permanent or i + 1 >= attempts:
                break
            time.sleep(wait * (i + 1))
    raise last


# ---------------------------------------------------------------- VLM


class Vlm:
    def __init__(self, cfg):
        self.cfg = cfg
        self.enabled = bool(cfg.get("base_url") and cfg.get("model") and cfg.get("api_key"))
        try:
            self.max_concurrency = max(1, int(cfg.get("max_concurrency", 10)))
        except (TypeError, ValueError):
            self.max_concurrency = 10
        try:
            self.max_attempts = max(1, int(cfg.get("max_attempts", 3)))
        except (TypeError, ValueError):
            self.max_attempts = 3
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
        for attempt in range(self.max_attempts):
            try:
                r = requests.post(url, json=payload, headers=headers, timeout=self.cfg.get("timeout_seconds", 120))
                if r.status_code != 200:
                    raise RuntimeError(f"HTTP {r.status_code}")
                text = r.json()["choices"][0]["message"]["content"]
                return (text or "").strip()
            except Exception as e:  # noqa: BLE001
                last = e
                if attempt + 1 < self.max_attempts:
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
            if pg.get("unpaged"):  # 轻量接口没有分页信息：整块内容都在这一页，页码只能标到块级
                rng = pg.get("unpaged_range") or [ch["pages"][0], ch["pages"][-1]]
                parts.append(f"<!-- 以下是第{rng[0]}–{rng[1]}页的合并输出，引擎没有给出分页，"
                             f"出处页码只能标到这个范围 -->")
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


def _dl_cfg(cfg):
    """download配置：重试次数、并发数、连续几页补OCR后仍下载失败就停止补OCR。缺失或写错时用默认值"""
    d = cfg.get("download") or {}

    def num(k, default):
        try:
            return max(1, int(d.get(k, default)))
        except (TypeError, ValueError):
            return default
    return {"max_attempts": num("max_attempts", 3), "max_concurrency": num("max_concurrency", 4),
            "max_reocr_failures": num("max_reocr_failures", 3)}


def _pictures(pg, orig, base=None):
    """一页里的图片：产出(编号, 链接)。编号p<页码>_<序号>；没有链接时链接为空串。
    base是课程文件夹：块里有img_local（相对它的路径，MinerU的图随结果包存在.缓存/里）时拼成本地链接"""
    i = 0
    for b in pg["blocks"]:
        if (b.get("type") or "").lower() != "picture":
            continue
        i += 1
        url = b.get("img_url") or ""
        if not url and base is not None and b.get("img_local"):
            url = LOCAL_PREFIX + str(Path(base) / b["img_local"])
        if not url:
            m = re.search(r"https?://[^\s)\"']+", b.get("content") or "")
            url = m.group(0) if m else ""
        yield f"p{orig:03d}_{i:02d}", url


def _drain(futures, imgs, media_dir):
    """等线程池里的下载全部结束，把结果登记进manifest（只在主线程里改）。返回本批重试的次数"""
    retried = 0
    for key, fut in list(futures.items()):
        e = imgs.setdefault(key, {})
        try:
            dest, n = fut.result()
            e["file"], e["error"] = dest.name, None
            e.pop("dl_failed", None)
            retried += n
        except ImageFetchError as ex:
            e["file"], e["error"], e["dl_failed"] = None, f"下载失败：{str(ex)[:150]}", True
        except Exception as ex:  # noqa: BLE001
            e["file"], e["error"], e["dl_failed"] = None, f"下载失败：{str(ex)[:150]}", True
    futures.clear()
    return retried


def _reocr_failed_pages(mat, src_pdf, course_dir, work_dir, cfg, fitz, tier, stats):
    """图片下载（含重试）仍失败的页，单页重新请求OCR引擎拿新链接，再下载。
    每页最多补一次；连续几页补完仍失败，说明多半是网络问题而不是链接过期，停止补OCR以免白白消耗额度。
    在主线程里串行做：OXR免费接口同一时间只处理一个任务"""
    imgs = mat.setdefault("images", {})
    media_dir = course_dir / "媒体" / mat["dirname"]
    dl = _dl_cfg(cfg)
    bad_pages = sorted({int(k[1:4]) for k, e in imgs.items() if e.get("dl_failed")})
    consec = 0
    for n, pno in enumerate(bad_pages, 1):
        if consec >= dl["max_reocr_failures"]:
            stats["reocr_stopped"] = True
            log(f"  连续{consec}页补OCR后图片仍下载失败，疑似网络问题而不是链接过期，停止补OCR")
            break
        ch = next((c for c in mat["chunks"] if c["status"] == "ok" and pno in c["pages"]), None)
        if ch is None:
            continue
        idx = ch["pages"].index(pno)
        progress(f"资料{mat['id']} 第{pno}页图片下载失败，单页补OCR换新链接（{n}/{len(bad_pages)}）")
        stats["reocr_pages"].append(pno)
        try:
            one = work_dir / f"redo_{pno:04d}.pdf"
            make_chunk_pdf(fitz, src_pdf, [pno], one)
            with Heartbeat(f"资料{mat['id']} 第{pno}页补OCR"):
                res = run_engine(tier["engine"], tier["cfg"], one, course_dir, mat["id"], f"redo{pno:04d}")
            one.unlink(missing_ok=True)
            if len(res.get("pages", [])) != 1:
                raise RuntimeError(f"返回{len(res.get('pages', []))}页，应为1页")
            cache_file = course_dir / ch["cache"]
            data = json.loads(cache_file.read_text(encoding="utf-8"))
            res["pages"][0]["page_num"] = idx
            data["pages"][idx] = res["pages"][0]
            # 新结果里图片的数量和顺序可能与第一次不同：先清掉这页的旧记录和旧文件，再按新结果登记
            for k in [k for k in imgs if k.startswith(f"p{pno:03d}_")]:
                f = imgs.pop(k).get("file")
                if f:
                    (media_dir / f).unlink(missing_ok=True)
            cache_file.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
            futures = {}
            with ThreadPoolExecutor(max_workers=dl["max_concurrency"]) as pool:
                for key, url in _pictures(res["pages"][0], pno, course_dir):
                    e = imgs.setdefault(key, {})
                    if not url:
                        e["file"], e["error"] = None, "引擎未返回图片链接"
                        continue
                    futures[key] = pool.submit(fetch_image, url, media_dir / key, dl["max_attempts"])
                stats["retried"] += _drain(futures, imgs, media_dir)
            still = any(e.get("dl_failed") for k, e in imgs.items() if k.startswith(f"p{pno:03d}_"))
            consec = consec + 1 if still else 0
        except FatalEngineError as e:
            log(f"    补OCR失败（额度或鉴权）：{str(e)[:150]}")
            stats["reocr_stopped"] = True
            break
        except Exception as e:  # noqa: BLE001
            log(f"    补OCR失败：{str(e)[:150]}")
            consec += 1


def _existing_image(media_dir, key):
    """磁盘上已有的这张图（上次下载完但没来得及登记进manifest就被中断）。能正常打开才算，被截断的不要"""
    from PIL import Image
    for p in sorted(media_dir.glob(f"{key}.*")):
        try:
            with Image.open(p) as im:
                im.verify()
            return p
        except Exception:  # noqa: BLE001
            p.unlink(missing_ok=True)
    return None


def process_pdf_material(mat, src_pdf, pages, course_dir, work_dir, cfg, vlm, fitz, chain, start, save=None):
    """chain：build_chain给出的档位列表，start：本次从哪一档开始。某一档失败时依次退到后面的档位。
    save：把manifest写盘的函数。切块计划生成后、每个块处理完（成功或失败）后各调一次，
    这样进程被强杀时，已经解析完的块下次不用重新OCR，只丢弃正在进行的那一块"""
    engine, ecfg = chain[start]["engine"], chain[start]["cfg"]
    media_dir = course_dir / "媒体" / mat["dirname"]
    cache_dir = course_dir / ".缓存" / mat["id"]
    cache_dir.mkdir(parents=True, exist_ok=True)
    media_dir.mkdir(parents=True, exist_ok=True)
    max_pages, max_bytes = chunk_limits(engine, ecfg)

    def new_chunk(p):
        return {"pages": p, "status": "pending", "cache": f".缓存/{mat['id']}/p{p[0]:04d}.json", "error": None}

    if not mat.get("chunks"):
        pl = plan_chunks(fitz, src_pdf, pages, max_pages, max_bytes, work_dir)
        mat["chunks"] = [new_chunk(p) for p in pl]
    else:  # 之前用别的引擎切的块，超过当前引擎上限的未完成块要重切
        new = []
        for ch in mat["chunks"]:
            if ch["status"] != "ok" and len(ch["pages"]) > max_pages:
                new += [new_chunk(ch["pages"][i:i + max_pages]) for i in range(0, len(ch["pages"]), max_pages)]
            else:
                new.append(ch)
        mat["chunks"] = new
    if save:
        save()
    imgs = mat.setdefault("images", {})
    dl = _dl_cfg(cfg)
    stats = {"retried": 0, "reocr_pages": [], "reocr_stopped": False}
    pool = ThreadPoolExecutor(max_workers=dl["max_concurrency"])
    futures = {}  # key -> Future，key是图片编号，保证同一张图只下载一次

    def queue_downloads(ch):
        """把一个已解析的块里还没下载的图片交给线程池。线程只下载，不碰manifest"""
        data = json.loads((course_dir / ch["cache"]).read_text(encoding="utf-8"))
        for pg in data["pages"]:
            orig = ch["pages"][pg["page_num"]]
            for key, url in _pictures(pg, orig, course_dir):
                e = imgs.setdefault(key, {})
                if key in futures or (e.get("file") and (media_dir / e["file"]).exists()):
                    continue
                have = _existing_image(media_dir, key)
                if have:  # 上次下载完了但没登记就被中断：直接认，不用旧链接再下载
                    e["file"], e["error"] = have.name, None
                    e.pop("dl_failed", None)
                    continue
                if not url:
                    e["file"], e["error"] = None, "引擎未返回图片链接"
                    continue
                futures[key] = pool.submit(fetch_image, url, media_dir / key, dl["max_attempts"])

    # 上次已解析的块：缓存里的链接最旧，先下载（上次中断留下的缺图也在这里补），再去解析剩下的块
    for ch in mat["chunks"]:
        if ch["status"] == "ok" and (course_dir / ch["cache"]).exists():
            queue_downloads(ch)

    if futures:
        progress(f"资料{mat['id']} 上次遗留{len(futures)}张图片待下载，已开始补下载")
    pending_chunks = [c for c in mat["chunks"] if not (c["status"] == "ok" and (course_dir / c["cache"]).exists())]
    todo_pages = sum(len(c["pages"]) for c in pending_chunks)
    done_pages = failed_run = 0
    t_loop = time.time()
    if todo_pages:
        progress(f"资料{mat['id']} 待解析{todo_pages}页，共{len(pending_chunks)}块（{chain[start]['label']}）")
    stop = False
    level = start  # 当前档位：某一档出现致命错误或排队超时后，这份资料后面的块都从下一档开始
    last_tier = chain[start]  # 最后一次成功解析所用的档位，单页补OCR用它
    for ci, ch in enumerate(mat["chunks"]):
        cache_file = course_dir / ch["cache"]
        if ch["status"] == "ok" and cache_file.exists():
            continue
        if stop:
            ch["status"], ch["error"] = "failed", "已停止：前面的块在所有档位上都遇到额度或鉴权错误"
            continue
        log(f"  {chain[level]['label']}解析：第{ch['pages'][0]}–{ch['pages'][-1]}页（块{ci + 1}/{len(mat['chunks'])}）")
        pdf_chunk = work_dir / f"chunk_{ci:02d}.pdf"
        t0 = time.time()
        try:
            make_chunk_pdf(fitz, src_pdf, ch["pages"], pdf_chunk)
            tag = f"p{ch['pages'][0]:04d}"
            hb = f"资料{mat['id']} 块{ci + 1}/{len(mat['chunks'])}"
            try_level, res, used = level, None, None
            while True:
                tier = chain[try_level]
                try:
                    with Heartbeat(hb + ("" if try_level == start else f"（{tier['label']}）")):
                        if try_level == start:  # 本次的起始档，块已经按它的上限切好
                            res = run_engine(tier["engine"], tier["cfg"], pdf_chunk, course_dir, mat["id"], tag)
                        else:  # 退下来的档位，单块上限不同，按它的上限重新拆
                            res = parse_with_tier(fitz, src_pdf, ch["pages"], tier, course_dir, mat["id"],
                                                  work_dir, f"{tag}t{try_level}")
                    used, last_tier = tier["label"], tier
                    break
                except Exception as e:  # noqa: BLE001
                    if try_level + 1 >= len(chain):
                        raise  # 所有档位都试过了
                    # 致命错误（key无效、额度用尽）和排队超时说明后面的块多半也一样，整份资料后面都从下一档开始；
                    # 其他错误只让这一块退到下一档
                    permanent = isinstance(e, FatalEngineError) or "超时" in str(e)
                    nxt = chain[try_level + 1]["label"]
                    note = f"{tier['label']}失败（{str(e)[:100]}），" + \
                           (f"这份资料后面的块都改用{nxt}" if permanent else f"这一块改用{nxt}")
                    mat.setdefault("engine_fallback", []).append(note)
                    progress(f"资料{mat['id']} {note}")
                    try_level += 1
                    if permanent:
                        level = try_level
            if len(res.get("pages", [])) != len(ch["pages"]):
                raise RuntimeError(f"返回{len(res.get('pages', []))}页，应为{len(ch['pages'])}页")
            cache_file.write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
            ch["status"], ch["error"] = "ok", None
            ch["engine"], ch["seconds"] = used, round(time.time() - t0, 1)
            queue_downloads(ch)  # 图片链接刚拿到，马上下载，和后面块的OCR同时进行
        except FatalEngineError as e:  # 所有档位都试过了，最后一档还是致命错误：后面的块不用再试
            ch["status"], ch["error"] = "failed", str(e)[:300]
            log(f"    失败：{ch['error']}")
            stop = True
        except Exception as e:  # noqa: BLE001
            ch["status"], ch["error"] = "failed", str(e)[:300]
            log(f"    失败：{ch['error']}")
        if save:
            save()  # 这一块的结果落盘后马上登记，之后被强杀也不用重做
        done_pages += len(ch["pages"])
        if ch["status"] != "ok":
            failed_run += len(ch["pages"])
        elapsed = time.time() - t_loop
        line = f"资料{mat['id']} 已处理{done_pages}/{todo_pages}页（{done_pages * 100 // todo_pages}%）"
        if failed_run:
            line += f"，其中失败{failed_run}页"
        line += f"，已用{fmt_dur(elapsed)}"
        if done_pages < todo_pages:
            line += f"，预计剩余约{fmt_dur(elapsed / done_pages * (todo_pages - done_pages))}"
        if futures:
            line += f"；图片下载{sum(f.done() for f in futures.values())}/{len(futures)}"
        progress(line)

    # 收尾：等流水线里的下载全部结束，下载（含重试）仍失败的页单页补OCR换新链接
    dl_keys = list(futures)
    while futures:  # 等下载期间也定时报一次进度
        _, pending = futures_wait(list(futures.values()), timeout=HEARTBEAT_SECONDS)
        if not pending:
            break
        progress(f"资料{mat['id']} 图片下载{len(dl_keys) - len(pending)}/{len(dl_keys)}")
    stats["retried"] += _drain(futures, imgs, media_dir)
    pool.shutdown(wait=True)
    if dl_keys:
        bad = sum(1 for k in dl_keys if imgs.get(k, {}).get("dl_failed"))
        progress(f"资料{mat['id']} 图片下载结束：{len(dl_keys) - bad}/{len(dl_keys)}张成功"
                 + (f"，{bad}张失败，将对这些页补OCR" if bad else ""))
    _reocr_failed_pages(mat, src_pdf, course_dir, work_dir, cfg, fitz, last_tier, stats)
    mat["dl_stats"] = stats
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
    # 第一步（主线程）：分类，筛出真正需要请求的图片，相同内容的图片只请求一次
    todo = {}  # hash -> [(key, entry, path)]
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
        todo.setdefault(h, []).append((key, e, path))
    if not todo:
        return

    def work(path):
        txt = vlm.transcribe(path)
        if not txt:
            raise RuntimeError("空结果")
        return txt

    # 第二步：线程池并发请求，最多max_concurrency个请求同时在途；manifest只在主线程里改
    workers = max(1, min(vlm.max_concurrency, len(todo)))
    n_imgs = sum(len(v) for v in todo.values())
    progress(f"资料{mat['id']} 图片转写开始：{n_imgs}张"
             + (f"（相同的合并后{len(todo)}次请求）" if n_imgs != len(todo) else "") + f"，并发{workers}")
    done = bad = 0
    step = max(1, len(todo) // 20)  # 最多打20行左右，图片上百张时不刷屏
    with Heartbeat(f"资料{mat['id']} 图片转写"), ThreadPoolExecutor(max_workers=workers) as pool:
        futures = {pool.submit(work, items[0][2]): (h, items) for h, items in todo.items()}
        for fut in as_completed(futures):
            h, items = futures[fut]
            try:
                txt = fut.result()
                cache[h] = txt
                for _, e, _ in items:
                    e["vlm"], e["text"] = "ok", txt
            except Exception as ex:  # noqa: BLE001
                bad += 1
                for _, e, _ in items:
                    e["vlm"] = "failed"
                    e["vlm_error"] = str(ex)[:150]
            done += 1
            if done == len(todo) or done % step == 0:
                progress(f"资料{mat['id']} 图片转写{done}/{len(todo)}" + (f"，失败{bad}" if bad else ""))


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
        "image_download": {
            "retries": (mat.get("dl_stats") or {}).get("retried", 0),
            "reocr_pages": ranges((mat.get("dl_stats") or {}).get("reocr_pages", [])),
            "reocr_stopped": (mat.get("dl_stats") or {}).get("reocr_stopped", False),
            "failed": sum(1 for e in imgs.values() if e.get("dl_failed")),
        },
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

# 各引擎每分钟处理页数（下限，上限）。paddle为用户给出的实测速度；oxr约5秒一页，取单一值
SPEED_PPM = {"paddle": (30, 50), "oxr": (12, 12), "mineru": (20, 60)}  # mineru：实测波动大，排队可达数分钟，仅供参考


def result_line(info, secs):
    """一份资料处理结束时的一行结果"""
    if info.get("error"):
        return f"资料{info['id']} 失败：{info['error'][:80]}，用时{fmt_dur(secs)}"
    bits = [f"状态{info['status']}", f"{info['pages']}页"]
    if info["failed_pages"]:
        bits.append("失败页" + ",".join(info["failed_pages"]))
    if info["images"]:
        bits.append(f"图片{info['images']}张（已转写{info['transcribed']}，未转写{info['untranscribed']}，"
                    f"下载失败{info['image_download']['failed']}）")
    bits.append(f"用时{fmt_dur(secs)}")
    return f"资料{info['id']} 结束：" + "，".join(bits)


def is_done(mat, draft_path, vlm):
    """这份资料是否已完整处理（草稿在、图片转写没有待补的）"""
    stale = {"failed", "pending"} if vlm.enabled else {"failed"}
    return (mat.get("status") == "ok" and draft_path.exists()
            and not any(e.get("vlm") in stale for e in mat.get("images", {}).values()))


def fmt_minutes(pages, engine):
    if pages <= 0:
        return "0分钟"
    lo_ppm, hi_ppm = SPEED_PPM[engine]
    lo, hi = max(1, round(pages / hi_ppm)), max(1, round(pages / lo_ppm))
    return f"约{lo}分钟" if lo >= hi else f"约{lo}–{hi}分钟"


def dry_run(files, types, course_dir, engine, vlm, fitz, sel, warnings):
    """只读预估：数每份资料的页数，按引擎速度估OCR耗时。不联网、不用key、不改manifest。
    Office文件要先转PDF才有页数，转出的PDF放进缓存，正式运行直接复用"""
    mpath = course_dir / "manifest.json"
    manifest = json.loads(mpath.read_text(encoding="utf-8")) if mpath.exists() else {"materials": []}
    rows, total = [], 0
    for f, tp in zip(files, types):
        row = {"name": f.stem.strip(), "file": f.name, "type": TYPE_NAMES[tp], "format": f.suffix.lower().lstrip(".")}
        mat = next((m for m in manifest["materials"] if m["source"] == str(f)), None)
        draft = course_dir / "草稿" / f"{mat['dirname']}.md" if mat else None
        try:
            if mat and is_done(mat, draft, vlm):
                row.update(status="跳过（已完成）", pages=mat.get("pages"), pages_to_ocr=0)
            elif f.suffix.lower() == ".ipynb":
                nb = json.loads(f.read_text(encoding="utf-8"))
                row.update(status="待处理", cells=len(nb.get("cells", [])), pages_to_ocr=0,
                           note="ipynb在本地解析，不走OCR，几乎不耗时")
            else:
                src_pdf = f if f.suffix.lower() == ".pdf" else cached_office_pdf(f, course_dir)
                doc = fitz.open(str(src_pdf))
                n = doc.page_count
                doc.close()
                pages = [p for p in (sel or range(1, n + 1)) if 1 <= p <= n]
                chunks = (mat or {}).get("chunks") or []
                if chunks:  # 之前处理过一部分，只算没完成的块
                    todo = sum(len(c["pages"]) for c in chunks
                               if c["status"] != "ok" or not (course_dir / c["cache"]).exists())
                    row.update(status="续跑", pages=len(pages), pages_to_ocr=todo)
                else:
                    row.update(status="待处理", pages=len(pages), pages_to_ocr=len(pages))
                if row["pages_to_ocr"] == 0:
                    row["note"] = "解析已完成，只差图片转写"
        except SystemExit:
            raise
        except Exception as e:  # noqa: BLE001
            row.update(status="无法读取", error=str(e)[:200], pages_to_ocr=0)
        row["estimate"] = fmt_minutes(row["pages_to_ocr"], engine)
        total += row["pages_to_ocr"]
        rows.append(row)
    lo, hi = SPEED_PPM[engine]
    out = {
        "dry_run": True, "engine": engine,
        "speed": f"{lo}–{hi}页/分钟" if lo != hi else f"约{lo}页/分钟",
        "materials": rows, "pages_to_ocr": total, "estimate": fmt_minutes(total, engine),
        "estimate_scope": "只含OCR解析；Office转PDF、图片转写（VLM）、下载图片的时间不计入",
        "warnings": warnings,
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))


def main():
    ap = argparse.ArgumentParser(description="课件、资料转Markdown草稿（OXR、Paddle或MinerU，加VLM）")
    ap.add_argument("files", nargs="*", help="一个或多个文件，按传入顺序编号；--init-config时不需要")
    ap.add_argument("--course-dir", help="课程文件夹，如study-buddy-data/0101_微观经济学；--init-config时不需要")
    ap.add_argument("--data-dir", help="数据目录，放config.json；默认为课程文件夹的上一级，--init-config时默认为./study-buddy-data")
    ap.add_argument("--init-config", action="store_true",
                    help="只在数据目录里生成config.json（已存在则不动）然后退出，用于首次使用时先让用户填key")
    ap.add_argument("--types", default="material",
                    help="material/textbook/past-exam/assignment，逗号分隔与files一一对应，或只写一个应用到全部")
    ap.add_argument("--pages", help="只处理这些页（1起，如10,17,20-22），用于测试；多份文件时应用到每份")
    ap.add_argument("--engine", choices=sorted(ENGINES), help="OCR起始引擎（paddle/oxr/mineru）；默认取config.json的engine，没写则用paddle。优先级链：paddle key→oxr key→mineru key→oxr免费接口→mineru轻量接口，没key的档位跳过，失败依次退到后面的档位")
    ap.add_argument("--keep-temp", action="store_true", help="保留临时PDF和切块")
    ap.add_argument("--dry-run", action="store_true",
                    help="只预估不处理：输出每份资料的页数、引擎和OCR耗时范围；不联网、不改manifest、不生成config.json")
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

    cfg, created = load_config(data_dir, create=not args.dry_run)
    engine = args.engine or cfg.get("engine") or "paddle"
    warnings = []
    chain = build_chain(cfg)  # paddle key → oxr key → mineru key → oxr免费接口 → mineru免登录轻量接口
    if engine == "paddle" and chain[0]["engine"] != "paddle":  # paddle没填key
        if args.engine == "paddle" and args.dry_run:  # 预估照常给，正式运行会报错
            warnings.append("显式指定了paddle但paddle.api_key还没填，正式运行会报错；以下按paddle速度预估")
        elif args.engine == "paddle":  # 用户显式要求paddle但没有key，不悄悄换引擎
            fail_env(f"使用paddle需要key：到 https://aistudio.baidu.com/paddleocr 申请，填进{data_dir / 'config.json'}的paddle.api_key")
        else:
            warnings.append(f"默认引擎是paddle，但paddle.api_key还没填，本次从“{chain[0]['label']}”开始；"
                            "到 https://aistudio.baidu.com/paddleocr 申请key填进config.json即可用回paddle")
        start = 0
    else:  # 指定的引擎从它在链里最靠前的可用档位开始；失败时依次退到后面的档位
        start = next(i for i, t in enumerate(chain) if t["engine"] == engine)
    speed_engine = "paddle" if (args.engine == "paddle" and args.dry_run) else chain[start]["engine"]  # 预估用的引擎
    if start < len(chain) - 1:
        warnings.append("本次引擎：" + " → ".join(t["label"] for t in chain[start:]) + "（前一档失败时依次退到后一档）")
    first = chain[start]
    if first["label"] == "oxr免费接口":
        warnings.append("oxr使用免费接口，同一时间只处理一个任务，速度较慢；有商业版key可填进config.json的oxr.api_key")
    elif first["label"] == "mineru轻量接口":
        warnings.append("mineru.api_key没填，使用mineru的免登录轻量接口（pipeline模型）：每块最多20页、10MB，"
                        "只返回一份Markdown，没有分页，页码只能标到块级，也没有表格和公式识别；排队久或失败时退用后面的档位。"
                        "到MinerU官网的API管理页创建token填进config.json的mineru.api_key，可用vlm模型和精确页码")
    vlm = Vlm(cfg["vlm"])
    if created:
        warnings.append(f"已生成{data_dir / 'config.json'}，请填写vlm的base_url、model、api_key")
    if not vlm.enabled:
        warnings.append("VLM未配置，图片全部标为“未转写”；填好config.json后重新运行即可补转写")

    if args.dry_run:
        if not (data_dir / "config.json").exists():
            warnings.append("数据目录里还没有config.json，按模板默认值预估")
        dry_run(files, types, course_dir, speed_engine, vlm, fitz, parse_pages(args.pages) if args.pages else None, warnings)
        sys.exit(0)

    course_dir.mkdir(parents=True, exist_ok=True)
    for sub in ("原始资料", "课程资料", "草稿", "媒体", "讲义"):
        (course_dir / sub).mkdir(exist_ok=True)
    mpath = course_dir / "manifest.json"
    manifest = json.loads(mpath.read_text(encoding="utf-8")) if mpath.exists() else {"version": 1, "materials": []}
    sel = parse_pages(args.pages) if args.pages else None

    def save_manifest():
        """先写临时文件再替换，进程在写的过程中被杀也不会留下半截的manifest"""
        tmp = mpath.with_suffix(".json.part")
        tmp.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
        tmp.replace(mpath)

    summary, any_problem = [], False
    for idx, (f, tp) in enumerate(zip(files, types), 1):
        t_mat = time.time()
        mat = next((m for m in manifest["materials"] if m["source"] == str(f)), None)
        if mat is None:
            nid = max([int(m["id"]) for m in manifest["materials"]] + [0]) + 1
            mat = {"id": f"{nid:03d}", "name": f.stem.strip(), "type": tp, "source": str(f), "chunks": [], "images": {}}
            mat["dirname"] = f"{mat['id']}_{safe_name(f.stem)}"
            manifest["materials"].append(mat)
        mat["type"] = tp
        log(f"[{mat['id']}] {f.name}（{TYPE_NAMES[tp]}）")
        progress(f"第{idx}/{len(files)}份开始：资料{mat['id']} {f.name}")
        draft_path = course_dir / "草稿" / f"{mat['dirname']}.md"
        work_dir = course_dir / ".临时" / mat["id"]
        work_dir.mkdir(parents=True, exist_ok=True)
        office_cache = None

        if is_done(mat, draft_path, vlm):
            progress(f"资料{mat['id']} 已完成，跳过")
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
                    src_pdf = office_cache = cached_office_pdf(f, course_dir)
                doc = fitz.open(str(src_pdf))
                total = doc.page_count
                doc.close()
                pages = [p for p in (sel or range(1, total + 1)) if 1 <= p <= total]
                if not pages:
                    raise RuntimeError(f"没有可处理的页（共{total}页）")
                mat["pages"], mat["selected_pages"] = len(pages), (sel and pages)
                process_pdf_material(mat, src_pdf, pages, course_dir, work_dir, cfg, vlm, fitz, chain, start,
                                     save=save_manifest)
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
            if office_cache:  # 处理成功后不再需要预转换的PDF；失败时保留，重跑直接复用
                office_cache.unlink(missing_ok=True)
        if info["status"] != "ok":
            any_problem = True
        progress(result_line(info, time.time() - t_mat))
        summary.append(info)
        save_manifest()

    save_manifest()
    work_root = course_dir / ".临时"
    if work_root.exists() and not any(work_root.iterdir()):
        work_root.rmdir()
    for m in manifest["materials"]:  # 本次运行中引擎档位间退用的记录，写进摘要警告
        for note in m.pop("engine_fallback", None) or []:
            warnings.append(f"资料{m['id']} {m['name']}：{note}")
    n_bad = sum(1 for m in summary if m.get("status") not in ("ok", "skipped（已完成）"))
    progress(f"全部{len(files)}份处理结束" + (f"，其中{n_bad}份有问题，见摘要" if n_bad else "，没有问题"))
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
