# 考试提纲

按考试规则整理可打印的提纲：开卷的做全做可索引，只带一张A4的做小抄，只带教材的做教材查阅指南，闭卷的做背诵稿和填空版。产出是tex源文件和编译好的PDF。

## 适用条件

- 进入：用户说“开卷考试帮我整理提纲”“帮我做小抄”“闭卷怎么背”“帮我整理一份能打印的复习资料”，且这门课已有讲义
- 不进入：这门课还没有讲义，改去[organize-materials.md](organize-materials.md)；用户要的是限时冲刺的计划和模拟卷，改去[cram-for-exam.md](cram-for-exam.md)；考试正在进行中，不代答

## 输入

- 课程名
- 考试规则（必须问清，不替用户假设）：可以带什么、多大、几页、能不能手写批注
- 考试日期和范围（可选）
- 读取该课的`讲义目录.md`（考点权重）、`讲义/`、`错题本.md`，以及往年题草稿（如有）

## 步骤

1. 检查前提：这门课没有`讲义/`时，提示先走[organize-materials.md](organize-materials.md)。只剩1天以内、讲义还没写时，先让资料整理走精简模式
2. 问考试规则：一次问清，按[exam-notes.md](../references/exam-notes.md)的规则表对应成一种产出。规则不确定时让用户去查课程大纲或问老师，不猜。全闭卷只做背诵稿，不出任何能带进考场的东西
3. 确认点：说明将做什么（产出形式、范围、页数上限、取哪些章节、大概多长），用户确认后再写。教材查阅指南要先确认教材的版本，并确认页码用书上印的页码
4. 取内容：从讲义的“查阅”部分、`讲义目录.md`的考点权重、`错题本.md`和往年题草稿取材，按考点权重取舍。不凭记忆添加，术语、定义、公式、符号按课程资料原样
5. 写tex：复制对应的模板到课程文件夹的`考试提纲/<YYYY-MM-DD>_<形式>.tex`，只改标题信息和正文，写法按[exam-notes.md](../references/exam-notes.md)
6. 编译：运行`python scripts/compile_tex.py <文件>.tex`，参数见[setup.md](../references/setup.md)。小抄加`--max-pages`；闭卷再加`--fill`出填空版。编译失败就按返回的行号改tex，再编译，最多5轮；超页就删低权重内容；没装xelatex就按返回的安装建议告诉用户安装（Windows MiKTeX，macOS MacTeX，Linux TeX Live），tex保存，不做备选
7. 检查：按[exam-notes.md](../references/exam-notes.md)“编译后检查”核对，小抄看`last_page_fill`，空间多就按权重补内容。把PDF路径告诉用户，建议打印前自己翻一遍

## 产出

- `考试提纲/`下的tex源文件和编译好的PDF，给用户看；闭卷另有`_填空版.pdf`
- 命名：`<YYYY-MM-DD>_开卷提纲`、`_小抄`、`_教材查阅指南`、`_闭卷背诵稿`

## 分支

- 没装xelatex：tex已保存。按脚本返回的安装命令告诉用户（Windows `winget install MiKTeX.MiKTeX`，macOS `brew install --cask mactex-no-gui`，Linux `sudo apt install texlive-xetex texlive-lang-chinese`），装完重开终端再编译
- 规则改了（如课程组临时允许带更多）：新版本另存为`-2`，不覆盖已经交付的那一份
- 同一门课做了多种产出（如开卷提纲和小抄）：分别写分别编译，内容保持一致，公式以讲义为准
- 用户背完闭卷稿，想针对薄弱点再练：转[cram-for-exam.md](cram-for-exam.md)，错题写进`错题本.md`

## 后续（可选）

- 考前几天时间紧：进入[cram-for-exam.md](cram-for-exam.md)
- 发现讲义有错或缺：回到[study-with-notes.md](study-with-notes.md)修订那一节，再重新生成提纲
