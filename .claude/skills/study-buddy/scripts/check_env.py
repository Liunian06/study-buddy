#!/usr/bin/env python3
"""检查运行环境，告诉助手哪些功能能用、缺什么、该装什么。Windows、macOS、Linux通用。

只读检查，不安装任何东西、不联网、不启动Office。输出JSON到stdout。
退出码：0没有缺少必需依赖（可选项缺了只列在missing_optional里），1缺少必需的Python包，2Python版本过低。

用法：python scripts/check_env.py
"""
import importlib.util
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from office_compat import IS_MAC, IS_WINDOWS, find_soffice, install_hint  # noqa: E402

# 导入名 -> pip包名。这几个所有平台都必需
REQUIRED = {"pymupdf": "pymupdf==2.0", "requests": "requests", "docx": "python-docx==1.2.0",
            "markdown_it": "markdown-it-py", "PIL": "pillow"}


def has_module(name):
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def office_installed():
    """Windows上PowerPoint、Word是否注册了COM。只查注册表，不启动Office"""
    if not IS_WINDOWS:
        return {"powerpoint": False, "word": False}
    import winreg

    def reg(progid):
        try:
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid))
            return True
        except OSError:
            return False
    return {"powerpoint": reg("PowerPoint.Application"), "word": reg("Word.Application")}


def tex_engine():
    try:
        import compile_tex
        return compile_tex.find_engine()
    except Exception:  # noqa: BLE001
        return None


def main():
    py_ok = sys.version_info >= (3, 10)
    plat = "Windows" if IS_WINDOWS else "macOS" if IS_MAC else "Linux"
    pip = "python -m pip" if IS_WINDOWS else "python3 -m pip"

    missing = [pkg for mod, pkg in REQUIRED.items() if not has_module(mod)]
    pptx_lib = has_module("pptx")
    pywin32 = IS_WINDOWS and has_module("win32com")
    office = office_installed()
    com_ppt = pywin32 and office["powerpoint"]
    com_word = pywin32 and office["word"]
    soffice = find_soffice()
    xelatex = tex_engine()

    caps = {
        # 转PDF：Windows有Office COM（需pywin32）就用它，否则LibreOffice；ppt和doc分别看PowerPoint、Word
        "ppt_to_pdf": bool(com_ppt or soffice),
        "doc_to_pdf": bool(com_word or soffice),
        "docx_to_pdf": bool(com_word or soffice),
        "build_pptx": bool(com_ppt or pptx_lib),
        "exam_notes_pdf": bool(xelatex),
    }
    optional_missing = []
    if not caps["ppt_to_pdf"] or not caps["doc_to_pdf"]:
        hint = install_hint()
        if IS_WINDOWS and (office["powerpoint"] or office["word"]) and not pywin32:
            hint = f"已装Office但缺pywin32，运行 {pip} install pywin32==312 即可；或者安装LibreOffice：winget install TheDocumentFoundation.LibreOffice"
        optional_missing.append({"item": "LibreOffice（或Windows上的Office加pywin32）",
                                 "affects": "转换PPT、DOC、导出PDF（pdf、docx格式的资料不受影响）",
                                 "install": hint})
    if not caps["build_pptx"]:
        optional_missing.append({"item": "python-pptx", "affects": "生成pptx",
                                 "install": f"{pip} install python-pptx==1.0.2"})
    if IS_WINDOWS and not pywin32 and (office["powerpoint"] or office["word"]):
        optional_missing.append({"item": "pywin32", "affects": "用Office COM转换和生成（没有它会改用LibreOffice和python-pptx，装上版式更接近原文件）",
                                 "install": f"{pip} install pywin32==312"})
    if not caps["exam_notes_pdf"]:
        optional_missing.append({"item": "TeX（xelatex）", "affects": "编译考试提纲，只有整理提纲时才需要，到时再装",
                                 "install": "Windows winget install MiKTeX.MiKTeX；macOS brew install --cask mactex-no-gui；"
                                            "Linux sudo apt install texlive-xetex texlive-lang-chinese"})

    out = {
        "platform": plat, "python": ".".join(map(str, sys.version_info[:3])), "python_ok": py_ok,
        "missing_required": ({"packages": missing, "install": f"{pip} install {' '.join(missing)}"}
                             if missing else None),
        "found": {"office_com": {"pywin32": pywin32, **office} if IS_WINDOWS else None,
                  "libreoffice": soffice, "python_pptx": pptx_lib, "xelatex": xelatex},
        "capabilities": caps,
        "missing_optional": optional_missing,
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))
    sys.exit(2 if not py_ok else 1 if missing else 0)


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
