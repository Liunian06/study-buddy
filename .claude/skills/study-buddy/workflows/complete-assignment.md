# 完成作业

完成case study、task、报告、习题等作业，个人或小组都适用，小组只做用户负责的部分。

## 适用条件

- 进入：用户说“这个case study怎么写”“帮我做这个task”“这份个人作业”
- 不进入：作业是pre，改去[prepare-presentation.md](prepare-presentation.md)；是课程论文，改去[write-course-paper.md](write-course-paper.md)；是正在进行的考试或限时测验，不代答

## 输入

- 作业要求（题目、案例材料、评分标准、字数或格式、交付形式），是文件时交给`ocr_cli.py`解析，类型选作业要求
- 个人还是小组；读取该课的`讲义/`

## 步骤

1. 拆要求：按[assignment.md](../references/assignment.md)套用[assignment.template.md](../assets/templates/assignment.template.md)逐条列出要回答的问题、要交的东西和评分点，判断作业类型；问用户这次提交用什么语言，见[deliverables.md](../references/deliverables.md)
2. 分模式：个人直接进入下一步。小组按[group-work.md](../references/group-work.md)出分工表，用户确认自己负责的部分之前，不往下做
3. 按类型完成：按[assignment.md](../references/assignment.md)中该类型的做法产出；用到的理论取自`讲义/`，需要图表时按[deliverables.md](../references/deliverables.md)由agent构建
4. 自查：套用[submission-check.template.md](../assets/templates/submission-check.template.md)，按评分点逐条核对，不通过先修正
5. 翻译（提交语言不是中文时）：按[deliverables.md](../references/deliverables.md)把定稿翻译成提交语言，重新核对字数
6. 转换提交格式：问用户要交什么格式，再运行[convert_submission.py](../scripts/convert_submission.py)从最终语言的Markdown转换

## 产出

- `作业/<YYYY-MM-DD>_<作业名>.md`（中文底稿），提交语言不是中文时另有`.<语言代码>.md`
- 用户确认格式后转出的提交文件、`_图表/`里的图表图片
- 小组模式只做用户负责的部分，另有分工表

## 分支

- 作业用到还没学过的知识点：可选跳到[study-with-notes.md](study-with-notes.md)
