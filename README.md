<div align="center">

# Study Buddy 学习搭子

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![Claude Code](https://img.shields.io/badge/Claude%20Code-Skill-5A4FCF.svg)](https://claude.com/claude-code)

[官网](https://liunian06.github.io/study-buddy/) · [更新日志](.claude/skills/study-buddy/CHANGELOG.md) · [下载](https://github.com/Liunian06/study-buddy/releases/latest)

</div>

## 这是什么

一个面向大学生的AI学习搭子skill：同时修好几门晦涩的课，只想用最少的时间让每门课都及格。它把老师发的一堆课件整理成按知识点排布的讲义，规划全学期每周该做什么，按讲义逐节带学，考前速通，pre、课程论文、case study和其他作业直接做出完整成品。

## 它能做什么

- 把课件、参考资料、往年题、整本教材转写成按知识点排好的讲义和目录，课件里的图片也能转写。原件放在每门课的 `原始资料/`，不会被改动；助手复制改名到 `课程资料/` 后再处理
- 根据评分构成和截止时间，算出每门课离及格差多少，排出每周清单
- 按讲义逐节带学：白话讲解、小测、错题本
- 考前根据剩余时间和往年题做速通计划、模拟卷和考前速记清单
- 完成个人或小组的presentation（大纲、讲稿、Q&A、pptx）、课程论文和其他作业，并转成docx或pdf提交

## 快速开始

### 方式一：AI 一键安装

把下面这段提示词发给你的 AI 编程助手（如 Claude Code），它会自动完成安装和环境检查：

```text
请帮我安装 study-buddy 这个 skill，仓库地址：https://github.com/Liunian06/study-buddy

1. 检查当前系统是否为 Windows、是否已安装 PowerPoint 和 Word、Node.js 版本是否 >= 22.20、Python 版本是否 >= 3.10。任何一项不满足，先告诉我，不要自行安装。
2. 全局安装：npx skills add Liunian06/study-buddy --skill study-buddy -g。如果命令让我选择 agent，请选择你自己（我当前使用的这个 agent）。
3. 安装完成后，从命令输出中确认 skill 的实际安装目录，阅读其中的 references/setup.md，按步骤安装 Python 依赖并运行验证命令。安装依赖前先问我要装到哪个 Python 环境。
4. 全部完成后，告诉我安装结果，然后阅读安装目录下的 references/introduction.md，向我介绍这个 skill 的功能和使用流程，并问我想先做什么。

注意：不要修改我的系统执行策略，不要询问、读取或复述任何 API key。
```

### 方式二：手动安装

需要 Windows 和已安装的 PowerPoint、Word，[Node.js](https://nodejs.org/) 22.20 及以上、Python 3.10 及以上，以及支持 Skills 的 AI 编程助手（如 [Claude Code](https://claude.com/claude-code)）。

1. 全局安装 skill，按提示选择你使用的 AI 编程助手：

```bash
npx skills add Liunian06/study-buddy --skill study-buddy -g
```

2. 安装 Python 依赖，详见 [setup.md](.claude/skills/study-buddy/references/setup.md)：

```bash
python -m pip install pymupdf==2.0 pywin32==312 requests python-docx==1.2.0 markdown-it-py pillow
```

3. 第一次整理资料时，脚本会在 `study-buddy-data/` 下生成 `config.json`。按 setup.md 填好 OCR 引擎和 VLM 服务的 key，不填 VLM 也能用，只是图片不会被转写。

也可以在 [Releases](https://github.com/Liunian06/study-buddy/releases/latest) 下载 `study-buddy-v<版本号>.zip`，解压后把 `study-buddy/` 目录放进你的 AI 编程助手的 skills 目录（如 `~/.claude/skills/`）。

### 更新

```bash
npx skills update study-buddy -g
```

脚本运行结束时会自动检查是否有新版本，有则提示更新命令。设置环境变量 `STUDY_BUDDY_NO_UPDATE_CHECK=1` 可关闭检查。

### 开始使用

安装完成后，在你的 AI 编程助手里直接说需求，例如：

- 老师发了一堆课件，帮我整理一下
- 这学期5门课，我怎么才能都过
- 带我学下一节
- 后天考微观，还没看过
- 下周三小组pre，我负责市场分析那段

首次使用建议先读[使用指南](.claude/skills/study-buddy/references/introduction.md)。

## 工作流程

skill 包含 7 个工作流，入口是 [SKILL.md](.claude/skills/study-buddy/SKILL.md)：

- **资料整理**：课件、教材、往年题转成讲义 `讲义/` 和目录 `讲义目录.md`
- **学期规划**：每门课的及格路径和每周清单
- **跟课学习**：按讲义逐节讲解、小测、错题本
- **期末速通**：速通计划、模拟卷、考前速记清单
- **Presentation**：pre 底稿、讲稿、Q&A 和 pptx
- **课程论文**：完整论文和提交文件
- **完成作业**：case study、task、报告、习题的完整作业和提交文件

工作流只编排步骤，详细规则在 skill 安装目录的 `references/` 下。

## 目录结构

### skill 目录

安装后位于 `~/<agent>/skills/study-buddy/`，`<agent>` 取决于你使用的 AI 编程助手（如 Claude Code 是 `~/.claude/skills/`）。

```
study-buddy/
├── SKILL.md        # skill 定义和工作流路由
├── CHANGELOG.md    # 版本记录
├── workflows/      # 7 个工作流
├── references/     # 各工作流引用的操作细则
├── assets/         # 模板、示例、VLM 转写指令
└── scripts/        # ocr_cli.py、build_pptx.py、convert_submission.py、check_update.py
```

### 数据目录

你的学习数据，在**工作目录**下运行后自动生成，不在 skill 目录里，也不会提交到仓库。

```
study-buddy-data/
├── config.json              # OCR 引擎和 VLM 的接口配置，含你自己填的 key
├── 学期规划.md              # 当前学期：课程清单、及格路径、每周清单
└── 0101_微观经济学/          # <学期号2位><课程序号2位>_<课程名>
    ├── 原始资料/            # 你放入的原始课件，只读不改
    ├── 课程资料/            # 整理后的副本（复制并改名）
    ├── 讲义目录.md          # 讲义目录、知识结构、考点权重、学习进度
    ├── 错题本.md            # 错题本
    ├── 讲义/                # 按知识点排好的讲义，第NN章_<知识点名>.md
    ├── 演示文稿/            # pre 底稿、pptx
    ├── 论文/                # 论文底稿和提交文件
    └── 作业/                # 作业底稿和提交文件
```

## 需要知道的限制

- 目标是及格，不追求高分或学透
- 考试或限时测验进行中不代答，考前复习和模拟照常帮忙
- 依赖 Windows 和 Office：转换 PPT 和 DOC、生成 pptx、导出 pdf 都通过 Office 的 COM 接口完成
- 课件会上传到 OCR 引擎（默认百度 Paddle，需要自己申请免费 key；没有 key 时退用 SoMark 的 OXR），图片会发给你配置的 VLM 服务商

遇到问题请在 [Issues](../../issues) 留言。

## 许可协议

本项目采用 [MIT](LICENSE) 协议。

Copyright © 2026 [Liunian06](https://github.com/Liunian06)
