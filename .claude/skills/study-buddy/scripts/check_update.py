#!/usr/bin/env python3
"""
检查 skill 是否有新版本。

由 ocr_cli.py、build_pptx.py、convert_submission.py 在运行结束时调用。设计原则：
- 查不到就不提示：断网、超时、被限流、版本号解析失败，一律静默
- 不拖慢脚本：请求超时 2 秒，结果缓存，成功缓存 24 小时、失败缓存 6 小时
- 有新版本时，提示里带上本地版本之后的 CHANGELOG 段落（按最新版本的 tag 取，只在有新版本时才多请求一次）
- 不污染标准输出：提示写到 stderr（ocr_cli.py 的 stdout 是 JSON 摘要）
- 可关闭：设置环境变量 STUDY_BUDDY_NO_UPDATE_CHECK=1
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import urllib.request
from pathlib import Path
from typing import Optional, Tuple

REPO = "Liunian06/study-buddy"
LATEST_API = f"https://api.github.com/repos/{REPO}/releases/latest"
UPDATE_COMMAND_GLOBAL = "npx skills update study-buddy -g"
UPDATE_COMMAND_PROJECT = "npx skills update study-buddy"

# 按最新版本的 tag 取 CHANGELOG，和发布出去的 skill 完全一致（main 上可能已经有还没发布的改动）
CHANGELOG_URL = f"https://raw.githubusercontent.com/{REPO}/{{tag}}/.claude/skills/study-buddy/CHANGELOG.md"
MAX_CHANGELOG_CHARS = 6000  # 提示里最多放这么多，跳过很多个版本时不至于刷屏

TIMEOUT_SECONDS = 2
CACHE_OK_SECONDS = 24 * 3600
CACHE_FAIL_SECONDS = 6 * 3600
CACHE_FILE = Path.home() / ".cache" / "study-buddy" / "update_check.json"
CHANGELOG_CACHE_FILE = Path.home() / ".cache" / "study-buddy" / "changelog.json"


def parse_version(text: str) -> Optional[Tuple[int, ...]]:
    """'v1.0.3' / '1.0.3' -> (1, 0, 3)；不是纯数字点分版本号则返回 None。"""
    m = re.fullmatch(r"v?(\d+(?:\.\d+)*)", text.strip())
    return tuple(int(x) for x in m.group(1).split(".")) if m else None


def read_local_version() -> Optional[str]:
    """读取本 skill 的 SKILL.md 中 frontmatter 的 version。"""
    skill_md = Path(__file__).resolve().parent.parent / "SKILL.md"
    try:
        for line in skill_md.read_text(encoding="utf-8").splitlines()[:20]:
            m = re.match(r"version:\s*(\S+)", line)
            if m:
                return m.group(1)
    except OSError:
        pass
    return None


def install_root() -> Optional[Path]:
    """skill 装在 <根>/.<agent>/skills/study-buddy/ 下时返回 <根>；目录结构不符返回 None。"""
    try:
        skill_dir = Path(__file__).resolve().parent.parent
        return skill_dir.parent.parent.parent
    except OSError:
        return None


def update_command() -> Tuple[str, Optional[Path]]:
    """按安装位置给出更新命令。根目录是用户主目录说明是全局安装（带 -g）；
    否则是项目内安装，命令要在该项目目录下运行，返回这个目录。"""
    root = install_root()
    try:
        if root is not None and not os.path.samefile(root, Path.home()):
            return UPDATE_COMMAND_PROJECT, root
    except OSError:
        pass
    return UPDATE_COMMAND_GLOBAL, None


def _load_cache() -> dict:
    try:
        return json.loads(CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_cache(latest: Optional[str]) -> None:
    try:
        CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        CACHE_FILE.write_text(
            json.dumps({"checked_at": time.time(), "latest": latest}), encoding="utf-8"
        )
    except OSError:
        pass


def fetch_latest_version() -> Optional[str]:
    """取最新 release 的 tag；缓存有效时不联网。查不到返回 None。"""
    cache = _load_cache()
    checked_at = cache.get("checked_at")
    if isinstance(checked_at, (int, float)):
        ttl = CACHE_OK_SECONDS if cache.get("latest") else CACHE_FAIL_SECONDS
        if 0 <= time.time() - checked_at < ttl:
            latest = cache.get("latest")
            return latest if isinstance(latest, str) else None

    latest = None
    try:
        req = urllib.request.Request(
            LATEST_API,
            headers={"Accept": "application/vnd.github+json", "User-Agent": "study-buddy-update-check"},
        )
        with urllib.request.urlopen(req, timeout=TIMEOUT_SECONDS) as resp:
            tag = json.loads(resp.read().decode("utf-8")).get("tag_name")
            if isinstance(tag, str):
                latest = tag
    except Exception:
        latest = None
    _save_cache(latest)
    return latest


def fetch_changelog(tag: str) -> Optional[str]:
    """取该 tag 下的 CHANGELOG.md 全文；按版本缓存，失败也缓存一段时间，免得每个脚本结束都等 2 秒。"""
    try:
        cache = json.loads(CHANGELOG_CACHE_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        cache = {}
    if cache.get("tag") == tag:
        text = cache.get("text")
        if isinstance(text, str):
            return text
        checked_at = cache.get("checked_at")
        if isinstance(checked_at, (int, float)) and 0 <= time.time() - checked_at < CACHE_FAIL_SECONDS:
            return None

    text = None
    try:
        req = urllib.request.Request(
            CHANGELOG_URL.format(tag=tag), headers={"User-Agent": "study-buddy-update-check"}
        )
        with urllib.request.urlopen(req, timeout=TIMEOUT_SECONDS) as resp:
            text = resp.read().decode("utf-8")
    except Exception:
        text = None
    try:
        CHANGELOG_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        CHANGELOG_CACHE_FILE.write_text(
            json.dumps({"tag": tag, "text": text, "checked_at": time.time()}, ensure_ascii=False),
            encoding="utf-8",
        )
    except OSError:
        pass
    return text


def changelog_between(text: str, local_t: Tuple[int, ...], latest_t: Tuple[int, ...]) -> str:
    """从 CHANGELOG 里取出版本号在 (本地版本, 最新版本] 之间的段落，新的在前。
    用户可能跳过了好几个版本，所以不止取最新一段。"""
    sections = []
    for line in text.replace("\r", "").splitlines():
        m = re.match(r"##\s*\[(\d+(?:\.\d+)*)\]", line)
        if m:
            sections.append([parse_version(m.group(1)), [line]])
        elif sections:
            sections[-1][1].append(line)
    picked = ["\n".join(body).strip() for v, body in sections if v and local_t < v <= latest_t]
    out = "\n\n".join(picked)
    if len(out) > MAX_CHANGELOG_CHARS:
        out = out[:MAX_CHANGELOG_CHARS].rstrip() + "\n……（内容较长，已截断，完整内容见上面的更新内容链接）"
    return out


def check_update(
    local_version: Optional[str] = None,
    latest_version: Optional[str] = None,
    changelog_text: Optional[str] = None,
) -> Optional[str]:
    """有新版本时返回提示文本，否则返回 None。三个参数仅用于测试。"""
    if os.environ.get("STUDY_BUDDY_NO_UPDATE_CHECK"):
        return None
    local = local_version or read_local_version()
    latest = latest_version or (fetch_latest_version() if local else None)
    if not local or not latest:
        return None
    local_t, latest_t = parse_version(local), parse_version(latest)
    if not local_t or not latest_t or latest_t <= local_t:
        return None
    command, project_dir = update_command()
    where = f"（在 {project_dir} 目录下运行，这个 skill 装在项目里）" if project_dir else ""
    if latest_t[0] == local_t[0]:
        auto = "自动更新：是。同一主版本，按 SKILL.md 的“自动更新”规则直接运行更新命令，不用问用户。"
    else:
        auto = ("自动更新：否。主版本升级，可能有不兼容的变更，不要自动更新：先把更新内容告诉用户，"
                "用户同意后再运行更新命令。")
    latest_s, local_s = ".".join(map(str, latest_t)), ".".join(map(str, local_t))
    text = changelog_text if changelog_text is not None else fetch_changelog(f"v{latest_s}")
    notes = changelog_between(text, local_t, latest_t) if text else ""
    message = (
        f"[更新提示] study-buddy 有新版本 v{latest_s}（当前 v{local_s}）。\n"
        f"更新命令：{command}{where}\n"
        f"更新内容链接：https://github.com/{REPO}/releases/tag/v{latest_s}\n"
        f"{auto}\n"
        f"更新只替换 skill 本身，工作目录里的 study-buddy-data/ 不受影响。"
    )
    if notes:  # 取不到就只给链接，和以前一样
        message += f"\n\n更新内容（v{local_s} 之后到 v{latest_s} 的 CHANGELOG，新的在前）：\n{notes}"
    return message


def notify_if_update() -> None:
    """脚本末尾调用。任何异常都吞掉，绝不影响主流程。"""
    try:
        message = check_update()
        if message:
            print(f"\n{message}", file=sys.stderr)
    except Exception:
        pass


if __name__ == "__main__":
    # 助手会直接运行本脚本读它的提示；Windows 下默认按系统代码页输出，中文会乱码
    for _stream in (sys.stdout, sys.stderr):
        if hasattr(_stream, "reconfigure"):
            _stream.reconfigure(encoding="utf-8")
    notify_if_update()
