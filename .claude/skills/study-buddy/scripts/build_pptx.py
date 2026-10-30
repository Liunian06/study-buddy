#!/usr/bin/env python3
"""从pre底稿的Markdown用PowerPoint COM生成pptx。

只做排版：每页写入标题和要点，讲稿与备注写进备注栏，把引用的本地图片放进页面。
不生成图表，不套模板和样式，用PowerPoint默认主题。

底稿格式见 assets/templates/presentation.template.md：
    ### 第N页　标题
    要点：
    - 要点
    图表：![说明](相对路径)
    讲稿：
    一段话
    备注：一段话

用法：python build_pptx.py 底稿.md [-o 输出.pptx]
退出码：0成功，2参数或环境错误。
"""
import argparse
import re
import sys
from pathlib import Path

LABELS = ("要点", "图表", "讲稿", "备注")
HEAD = re.compile(r"^###\s+(?:第\s*(\d+)\s*页)?[\s　]*(.*)$")
LABEL = re.compile(r"^(要点|图表|讲稿|备注)[：:]\s*(.*)$")
IMG = re.compile(r"!\[[^\]]*\]\(([^)]+)\)")
PLACEHOLDER = re.compile(r"\{\{.*?\}\}")


def parse_slides(md_path):
    slides, cur, label = [], None, None
    for raw in md_path.read_text(encoding="utf-8").splitlines():
        line = raw.rstrip()
        m = HEAD.match(line)
        if m:
            cur = {"title": m.group(2).strip(), "bullets": [], "images": [], "script": [], "note": []}
            slides.append(cur)
            label = None
            continue
        if cur is None:
            continue
        if line.startswith("## "):  # 进入下一个二级章节（如Q&A），幻灯片部分结束
            cur, label = None, None
            continue
        lm = LABEL.match(line)
        if lm:
            label, rest = lm.group(1), lm.group(2)
            line = rest
        if not line.strip():
            continue
        if label == "要点":
            t = re.sub(r"^[-*•]\s*", "", line).strip()
            if t:
                cur["bullets"].append(t)
        elif label == "图表":
            cur["images"] += [p.strip().strip("<>") for p in IMG.findall(line)]
        elif label == "讲稿":
            cur["script"].append(line.strip())
        elif label == "备注":
            cur["note"].append(line.strip())
    return slides


def build(slides, base_dir, out_path):
    import win32com.client
    app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        pres = app.Presentations.Add(WithWindow=False)
        try:
            width, height = pres.PageSetup.SlideWidth, pres.PageSetup.SlideHeight
            layout = pres.SlideMaster.CustomLayouts(2)  # 标题和内容
            for i, s in enumerate(slides, 1):
                sl = pres.Slides.AddSlide(i, layout)
                title = PLACEHOLDER.sub("", s["title"]).strip()
                sl.Shapes.Title.TextFrame.TextRange.Text = title
                body = sl.Shapes.Placeholders(2)
                body.TextFrame.TextRange.Text = "\r".join(s["bullets"])
                imgs = [(base_dir / p).resolve() for p in s["images"]]
                imgs = [p for p in imgs if p.exists()]
                if imgs:
                    body.Width = width * 0.45
                    left = body.Left + body.Width + 20
                    avail_w = width - left - 30
                    avail_h = body.Height
                    pic = sl.Shapes.AddPicture(str(imgs[0]), False, True, left, body.Top)
                    ratio = min(avail_w / pic.Width, avail_h / pic.Height, 1.0)
                    pic.LockAspectRatio = True
                    pic.Width, pic.Height = pic.Width * ratio, pic.Height * ratio
                notes = "\r".join(s["script"] + (["【备注】"] + s["note"] if s["note"] else []))
                if notes:
                    sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = notes
            out_path.parent.mkdir(parents=True, exist_ok=True)
            pres.SaveAs(str(out_path.resolve()))
        finally:
            pres.Close()
    finally:
        app.Quit()


def main():
    ap = argparse.ArgumentParser(description="从pre底稿Markdown生成pptx（PowerPoint COM）")
    ap.add_argument("markdown", help="pre底稿，最终提交语言的版本")
    ap.add_argument("-o", "--output", help="输出路径，默认与底稿同名的.pptx；已存在时另存为-2、-3")
    args = ap.parse_args()

    md = Path(args.markdown).resolve()
    if not md.exists():
        print(f"文件不存在：{md}", file=sys.stderr)
        sys.exit(2)
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        print("缺少pywin32，见references/setup.md", file=sys.stderr)
        sys.exit(2)

    slides = parse_slides(md)
    if not slides:
        print("底稿里没有找到“### 第N页 标题”格式的幻灯片", file=sys.stderr)
        sys.exit(2)
    out = Path(args.output).resolve() if args.output else md.with_suffix(".pptx")
    stem, n = out.stem, 2
    while out.exists():
        out = out.with_name(f"{stem}-{n}.pptx")
        n += 1
    build(slides, md.parent, out)
    missing = [p for s in slides for p in s["images"] if not (md.parent / p.strip("<>")).exists()]
    print(f"已生成：{out}（{len(slides)}页）")
    if missing:
        print("以下图片不存在，已跳过：" + "，".join(missing), file=sys.stderr)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    try:
        main()
        code = 0
    except SystemExit as e:
        code = e.code
    if code in (None, 0):
        try:
            sys.path.insert(0, str(Path(__file__).resolve().parent))
            from check_update import notify_if_update
            notify_if_update()
        except Exception:
            pass
    sys.exit(code)
