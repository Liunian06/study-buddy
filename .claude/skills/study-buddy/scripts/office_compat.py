#!/usr/bin/env python3
"""Office文件转PDF的跨平台实现，供ocr_cli.py、convert_submission.py共用。

选择顺序：
- Windows：先用Office COM（PowerPoint、Word），排版最接近原文件；没装Office、COM报错或没装pywin32时，
  退到LibreOffice；两者都没有才报错
- macOS、Linux：直接用LibreOffice（soffice --headless），不需要pywin32

LibreOffice版本的限制：字体、版式与Office有细微差别（如PPT里的特殊字体、艺术字），不影响OCR取字，
但提交用的PDF如果对版式有要求，建议在装了Office的电脑上导出。
"""
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

OFFICE_PPT = {".ppt", ".pptx"}
OFFICE_DOC = {".doc", ".docx"}
SOFFICE_TIMEOUT = 300
IS_WINDOWS = sys.platform == "win32"
IS_MAC = sys.platform == "darwin"


class OfficeUnavailable(RuntimeError):
    """当前电脑上没有可用的转换工具。信息里带安装建议"""


def install_hint():
    if IS_WINDOWS:
        return "安装Microsoft Office（PowerPoint和Word），或安装LibreOffice：winget install TheDocumentFoundation.LibreOffice"
    if IS_MAC:
        return "安装LibreOffice：brew install --cask libreoffice（或到 https://www.libreoffice.org 下载）"
    return "安装LibreOffice，如 sudo apt install libreoffice（或到 https://www.libreoffice.org 下载）"


def find_soffice():
    """找LibreOffice的命令行入口。PATH里没有时按各系统的默认安装位置找"""
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    cands = []
    if IS_WINDOWS:
        for env in ("ProgramFiles", "ProgramFiles(x86)"):
            base = os.environ.get(env)
            if base:
                cands.append(Path(base) / "LibreOffice" / "program" / "soffice.exe")
    elif IS_MAC:
        cands.append(Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"))
        cands.append(Path.home() / "Applications" / "LibreOffice.app" / "Contents" / "MacOS" / "soffice")
    for c in cands:
        if c.exists():
            return str(c)
    return None


def _profile_uri(path):
    """LibreOffice的-env:UserInstallation要file:///形式的URI"""
    return Path(path).resolve().as_uri()


def libreoffice_to_pdf(src, out_pdf, soffice=None):
    """用LibreOffice无界面转PDF。用独立的临时用户配置目录，避免和用户正在开着的LibreOffice抢锁。
    LibreOffice的输出文件名由输入文件名决定，所以先转到临时目录再改名到out_pdf"""
    soffice = soffice or find_soffice()
    if not soffice:
        raise OfficeUnavailable("没有找到LibreOffice。" + install_hint())
    src, out_pdf = Path(src).resolve(), Path(out_pdf).resolve()
    with tempfile.TemporaryDirectory(prefix="sb_lo_") as tmp:
        tmp = Path(tmp)
        outdir, profile = tmp / "out", tmp / "profile"
        outdir.mkdir()
        # 源文件复制成ASCII文件名放进临时目录：避开路径里的中文、空格和特殊字符在命令行上的差异
        local_src = tmp / f"in{src.suffix.lower()}"
        shutil.copyfile(src, local_src)
        cmd = [soffice, f"-env:UserInstallation={_profile_uri(profile)}", "--headless", "--norestore",
               "--convert-to", "pdf", "--outdir", str(outdir), str(local_src)]
        try:
            r = subprocess.run(cmd, capture_output=True, timeout=SOFFICE_TIMEOUT)
        except subprocess.TimeoutExpired:
            raise RuntimeError(f"LibreOffice转换超过{SOFFICE_TIMEOUT}秒仍未结束")
        produced = outdir / "in.pdf"
        if not produced.exists() or produced.stat().st_size == 0:
            msg = ((r.stderr or b"") + (r.stdout or b"")).decode("utf-8", errors="replace").strip()
            raise RuntimeError(f"LibreOffice未生成PDF（返回码{r.returncode}）：{msg[:200]}")
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(produced), str(out_pdf))


def _com_to_pdf(src, out_pdf, log):
    """Windows下用Office COM转PDF。直接打开失败时改用“打开并修复”再试，原文件不动"""
    import win32com.client  # ImportError由调用方处理
    src_s, out_s = str(src.resolve()), str(out_pdf.resolve())
    if src.suffix.lower() in OFFICE_PPT:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        try:
            try:
                pres = app.Presentations.Open(src_s, True, False, False)
            except Exception as ex:  # noqa: BLE001
                log(f"  PowerPoint直接打开失败（{str(ex)[:60]}），尝试自动修复后转换：{src.name}")
                pres = app.Presentations.Open2007(src_s, True, False, False, True)  # OpenAndRepair
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
            try:
                doc = app.Documents.Open(src_s, ReadOnly=True)
            except Exception as ex:  # noqa: BLE001
                log(f"  Word直接打开失败（{str(ex)[:60]}），尝试自动修复后转换：{src.name}")
                doc = app.Documents.Open(src_s, ReadOnly=True, OpenAndRepair=True)
            try:
                doc.ExportAsFixedFormat(out_s, 17)  # wdExportFormatPDF
            finally:
                doc.Close(False)
        finally:
            app.Quit()
    if not out_pdf.exists():
        raise RuntimeError("Office未生成PDF")


def convert_to_pdf(src, out_pdf, log=print):
    """把ppt、pptx、doc、docx转成PDF，写到out_pdf。
    没有任何可用工具时抛OfficeUnavailable，调用方按环境错误处理；转换本身失败抛RuntimeError"""
    src, out_pdf = Path(src), Path(out_pdf)
    if src.suffix.lower() not in OFFICE_PPT | OFFICE_DOC:
        raise RuntimeError(f"不支持的Office格式：{src.suffix}")
    com_err = None
    if IS_WINDOWS:
        try:
            _com_to_pdf(src, out_pdf, log)
            return
        except ImportError:
            com_err = "没有安装pywin32"
        except Exception as ex:  # noqa: BLE001  没装Office、COM启动失败、文件打不开等
            com_err = f"Office COM转换失败（{str(ex)[:80]}）"
        out_pdf.unlink(missing_ok=True)
        if not find_soffice():
            raise OfficeUnavailable(f"{com_err}，也没有找到LibreOffice。" + install_hint())
        log(f"  {com_err}，改用LibreOffice转换：{src.name}")
    else:
        log(f"  转PDF（LibreOffice）：{src.name}")
    libreoffice_to_pdf(src, out_pdf)
    if not out_pdf.exists():
        raise RuntimeError("LibreOffice未生成PDF")
