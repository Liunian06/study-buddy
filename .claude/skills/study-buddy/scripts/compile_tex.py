#!/usr/bin/env python3
r"""编译考试提纲的tex源文件为PDF（xelatex，MiKTeX或TeX Live都行）。

用法：python compile_tex.py 提纲.tex [-o 输出.pdf] [--fill] [--max-pages N] [--keep]
  --fill        编译填空版：模板里的\kw{答案}变成下划线空白；不加时\kw是加粗
  --max-pages   页数上限（如一页A4小抄填1）：超页时依次缩小字号、行距和页边距重新编译，直到不超；
                缩到最小仍超则返回overflow，需要删减内容
结果以JSON输出到stdout，进度日志输出到stderr。成功时含pages和last_page_fill（最后一页内容占页面高度的比例）。

退出码：0成功；1编译失败或超页（按errors里的行号改tex后重新编译）；2参数或环境错误（找不到xelatex等）。
"""
import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

# 缩放阶梯：(字号pt, 行距pt, 页边距mm)。模板默认字号是9pt，阶梯从比它小的开始
LADDER = [(8.5, 10.2, 10), (8, 9.6, 9), (7.5, 9.0, 8), (7, 8.4, 7), (6.5, 7.8, 6), (6, 7.2, 5)]
MIN_READABLE = 7.0  # 小于这个字号打印后可能难以辨认，给出警告
TIMEOUT = 300

def _install_hint():
    if sys.platform == "win32":
        return ("没有找到xelatex。请安装MiKTeX：winget install MiKTeX.MiKTeX（装完重开终端），"
                "或安装TeX Live，见references/setup.md")
    if sys.platform == "darwin":
        return ("没有找到xelatex。请安装MacTeX：brew install --cask mactex-no-gui（装完重开终端；"
                "也可以装体积小得多的BasicTeX再用tlmgr补宏包），见references/setup.md")
    return "没有找到xelatex。请安装TeX Live，如 sudo apt install texlive-xetex texlive-lang-chinese，见references/setup.md"


INSTALL_HINT = _install_hint()

# 模板里写的是fontset=windows（宋体、黑体、楷体）；编译时按平台换成ctex对应的字体集，模板本身不用改
FONTSET = {"win32": "windows", "darwin": "mac"}.get(sys.platform, "fandol")

ERROR_HINTS = [
    (r"Undefined control sequence", "命令未定义：检查拼写，或是否用了模板里没有的宏；不要自己新增宏和宏包"),
    (r"Misplaced alignment tab", "表格外出现了&：正文里的&要写成\\&"),
    (r"Missing \$ inserted", "公式外出现了_或^：要写成\\_、\\^{}；或者公式漏了$"),
    (r"File .* not found|\.sty.* not found", "缺宏包：MiKTeX会自动安装，TeX Live用tlmgr install；先确认不是拼错了名字"),
    (r"Runaway argument|File ended while scanning|Missing \} inserted|Extra \}", "括号不配对：检查{ }是否成对"),
    (r"Illegal parameter number|You can't use `macro parameter character #'", "正文里的#要写成\\#"),
    (r"Missing number|Illegal unit", "尺寸写法错误：检查单位和数字"),
    (r"Extra alignment tab|Too many columns", "表格列数和&个数对不上"),
]


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def emit(obj, code):
    print(json.dumps(obj, ensure_ascii=False, indent=2))
    sys.exit(code)


def find_engine():
    p = shutil.which("xelatex")
    if p:
        return p
    if sys.platform == "win32":
        # MiKTeX刚装完时当前终端的PATH可能还没刷新，按默认安装位置找一下
        local = os.environ.get("LOCALAPPDATA", "")
        prog = os.environ.get("ProgramFiles", "")
        for base in (Path(local) / "Programs" / "MiKTeX", Path(prog) / "MiKTeX"):
            exe = base / "miktex" / "bin" / "x64" / "xelatex.exe"
            if exe.exists():
                return str(exe)
    elif sys.platform == "darwin":
        # MacTeX、BasicTeX装在/Library/TeX/texbin，Homebrew的cask装完终端没重开时PATH里可能还没有
        for d in ("/Library/TeX/texbin", "/usr/local/texlive/bin", "/opt/homebrew/bin", "/usr/local/bin"):
            exe = Path(d) / "xelatex"
            if exe.exists():
                return str(exe)
    return None


def is_miktex(engine):
    try:
        r = subprocess.run([engine, "--version"], capture_output=True, timeout=60)
        return b"MiKTeX" in (r.stdout + r.stderr)
    except Exception:  # noqa: BLE001
        return False


class BuildTimeout(Exception):
    pass


def run_once(engine, miktex, work, defs):
    cmd = [engine, "-interaction=nonstopmode", "-file-line-error", "-jobname=out"]
    if miktex:
        cmd.append("--enable-installer")  # 缺宏包时自动安装
    cmd.append(defs + r"\input{main.tex}")
    try:
        r = subprocess.run(cmd, cwd=work, capture_output=True, timeout=TIMEOUT)
    except subprocess.TimeoutExpired:
        raise BuildTimeout(f"编译超过{TIMEOUT}秒仍未结束（可能死循环，或MiKTeX在等待安装宏包）")
    lp = work / "out.log"
    text = lp.read_text(encoding="utf-8", errors="replace") if lp.exists() else ""
    if not text:
        text = (r.stdout or b"").decode("utf-8", errors="replace")
    return text, r.returncode


def parse_errors(text):
    lines = text.splitlines()
    errors, seen = [], set()
    for i, ln in enumerate(lines):
        m = re.match(r"^(?:.*?\.tex:(\d+):|!)\s*(.*)$", ln)
        if not m:
            continue
        msg = m.group(2).strip()
        line = int(m.group(1)) if m.group(1) else None
        ctx = ""
        for nxt in lines[i + 1:i + 8]:
            if nxt.startswith("l."):
                ctx = nxt.strip()
                if line is None:
                    mm = re.match(r"l\.(\d+)", nxt)
                    line = int(mm.group(1)) if mm else None
                break
        key = (line, msg)
        if key in seen:
            continue
        seen.add(key)
        errors.append({"line": line, "message": msg, "context": ctx})
    return errors[:8]


def parse_warnings(text):
    warns = []
    chars = []
    for c in re.findall(r"Missing character: There is no (.+?) in font", text):
        c = c.strip()
        if c and c not in chars:
            chars.append(c)
    if chars:
        warns.append("字体里没有这些字符，PDF里会缺字：" + " ".join(chars[:10]) + "；换成常见字符")
    wide = [float(x) for x in re.findall(r"Overfull \\hbox \(([\d.]+)pt too wide", text) if float(x) > 10]
    if wide:
        warns.append(f"{len(wide)}处行宽溢出超过10pt（最大{max(wide):.0f}pt），内容可能被页边截掉：检查过长的公式或表格")
    refs = sorted(set(re.findall(r"Reference `(.+?)' on page \d+ undefined", text)))
    if refs:
        warns.append("引用的标签不存在：" + "、".join(refs[:6]))
    return warns


def hints_for(errors):
    out = []
    for e in errors:
        for pat, hint in ERROR_HINTS:
            if re.search(pat, e["message"]) and hint not in out:
                out.append(hint)
    return out


def build(engine, miktex, work, defs):
    """最多三遍：目录、引用、longtable需要多遍才稳定。返回(日志, 错误列表)"""
    text, errors = "", []
    for i in range(3):
        (work / "out.pdf").unlink(missing_ok=True) if i == 0 else None
        text, rc = run_once(engine, miktex, work, defs)
        errors = parse_errors(text)
        if not errors and rc != 0 and not (work / "out.pdf").exists():
            errors = [{"line": None, "message": f"xelatex返回码{rc}，没有生成PDF", "context": text[-300:]}]
        if errors:
            break
        if i >= 1 and not re.search(r"Rerun to get|Label\(s\) may have changed", text):
            break
    return text, errors


def page_count(text):
    m = re.search(r"Output written on .*?\((\d+) pages?", text)
    return int(m.group(1)) if m else None


def last_page_fill(pdf):
    """最后一页内容占页面高度的比例（0到1）。小抄只会缩小不会放大，agent靠这个判断要不要补内容"""
    try:
        import pymupdf
        doc = pymupdf.open(str(pdf))
        pg = doc[doc.page_count - 1]
        ys = [b[3] for b in pg.get_text("blocks") if b[4].strip()]
        h = pg.rect.height
        doc.close()
        return round(max(ys) / h, 2) if ys else 0.0
    except Exception:  # noqa: BLE001
        return None


def main():
    ap = argparse.ArgumentParser(description="编译考试提纲tex为PDF")
    ap.add_argument("tex", help="tex源文件（UTF-8）")
    ap.add_argument("-o", "--output", help="输出PDF路径；默认与tex同名，--fill时加“_填空版”后缀；同名PDF会被覆盖")
    ap.add_argument("--fill", action="store_true", help="编译填空版（\\kw{答案}变成下划线）")
    ap.add_argument("--max-pages", type=int, help="页数上限；超页时自动缩小字号、行距和页边距")
    ap.add_argument("--keep", action="store_true", help="保留临时编译目录（排查用）")
    args = ap.parse_args()

    src = Path(args.tex).resolve()
    if not src.exists() or src.suffix.lower() != ".tex":
        emit({"status": "error", "message": f"需要一个存在的.tex文件：{src}"}, 2)
    try:
        src_text = src.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        emit({"status": "error", "message": "tex文件需要是UTF-8编码"}, 2)
    if args.max_pages is not None and args.max_pages < 1:
        emit({"status": "error", "message": "--max-pages需要大于0"}, 2)

    engine = find_engine()
    if not engine:
        # tex源文件已经在用户手里，没有编译器时不算失败，装好之后再编译即可
        emit({"status": "no_engine", "message": INSTALL_HINT, "tex": str(src)}, 2)
    miktex = is_miktex(engine)

    out = Path(args.output).resolve() if args.output else (
        src.with_name(src.stem + "_填空版.pdf") if args.fill else src.with_suffix(".pdf"))
    base_defs = r"\def\SBFILL{1}" if args.fill else ""

    work = Path(tempfile.mkdtemp(prefix="sbtex_"))  # 临时目录用ASCII路径，避开中文和空格路径的麻烦
    # fontset=windows换成当前平台的字体集；Windows上原样不动
    src_text = src_text.replace("fontset=windows", f"fontset={FONTSET}")
    (work / "main.tex").write_text(src_text, encoding="utf-8")

    attempts = [(None, None)] + [(f"{s}pt", (s, l, m)) for s, l, m in LADDER]
    if args.max_pages is None:
        attempts = attempts[:1]

    result = {"status": "error", "tex": str(src), "variant": "填空版" if args.fill else "正常版"}
    code = 1
    try:
        for label, sc in attempts:
            defs = base_defs
            if sc:
                defs += rf"\def\SBSIZE{{{sc[0]}}}\def\SBLEAD{{{sc[1]}}}\def\SBMARGIN{{{sc[2]}}}"
            log(f"编译{'（字号' + label + '）' if label else ''}……")
            text, errors = build(engine, miktex, work, defs)
            if errors:
                result.update(status="error", errors=errors, hints=hints_for(errors),
                              message="编译失败：按errors里的行号改tex，改完重新编译")
                break
            pages = page_count(text)
            result.update(pages=pages, warnings=parse_warnings(text),
                          scale=({"size_pt": sc[0], "leading_pt": sc[1], "margin_mm": sc[2]} if sc else None))
            if args.max_pages is None or (pages is not None and pages <= args.max_pages):
                out.parent.mkdir(parents=True, exist_ok=True)
                try:
                    shutil.copyfile(work / "out.pdf", out)
                except PermissionError:
                    result.update(status="error", message=f"输出文件被占用，请先关闭它：{out}")
                    break
                result.update(status="ok", pdf=str(out), last_page_fill=last_page_fill(out))
                result.pop("errors", None)
                if sc and sc[0] < MIN_READABLE:
                    result["warnings"].append(f"为了不超过{args.max_pages}页，字号缩到了{sc[0]}pt，打印后可能难以辨认；"
                                              "建议删掉低权重内容再编译")
                code = 0
                break
        else:
            result.update(status="overflow", message=f"缩到最小字号（{LADDER[-1][0]}pt）仍有{result.get('pages')}页，"
                          f"超过上限{args.max_pages}页：需要删减内容（先删低权重的点）")
    except BuildTimeout as e:
        result.update(status="error", message=str(e))
    finally:
        if args.keep:
            result["build_dir"] = str(work)
        else:
            shutil.rmtree(work, ignore_errors=True)

    result["engine"] = "MiKTeX" if miktex else "xelatex"
    result["fontset"] = FONTSET
    emit(result, code)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    try:
        main()
        code = 0
    except SystemExit as e:
        code = e.code
    if code in (None, 0, 1):
        try:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            from check_update import notify_if_update
            notify_if_update()
        except Exception:  # noqa: BLE001
            pass
    sys.exit(code)
