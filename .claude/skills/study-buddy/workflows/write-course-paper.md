# 课程论文

从作业要求到可提交的完整论文：用户指导方向，agent起草和修改，直到定稿并转成提交格式。

## 适用条件

- 进入：用户说“要交一篇3000字的论文”“课程论文怎么写”
- 不进入：作业是pre，改去[prepare-presentation.md](prepare-presentation.md)；是case study、task、报告、习题，改去[complete-assignment.md](complete-assignment.md)

## 输入

- 作业要求（题目或范围、字数、格式规范、引用格式、截止时间），是文件时交给`ocr_cli.py`解析，类型选作业要求
- 用户的方向指导、相关资料；读取该课的`notes/`

## 步骤

1. 拆要求：按[course-paper.md](../references/course-paper.md)提取字数、结构、格式、引用格式和评分点，得到要求清单；问用户这次提交用什么语言，见[deliverables.md](../references/deliverables.md)
2. 分模式：个人直接进入下一步。小组按[group-work.md](../references/group-work.md)出分工表，用户确认自己负责的章节之前，不往下做
3. 确认点：给出2–3个选题或论点方案，用户选择或另给方向。确认之前不搭大纲
4. 搭大纲：套用[paper.template.md](../assets/templates/paper.template.md)出大纲，写清每节要点和字数分配，用户确认后才开始写
5. 找文献：用户提供的优先；不够时按[course-paper.md](../references/course-paper.md)检索并核实，没核实的明确标出，不编造
6. 写作：按大纲逐节写出完整正文，小组只写用户负责的章节；需要图表时按[deliverables.md](../references/deliverables.md)由agent构建
7. 修改与自查：用户提意见，agent改，每一版另存为`-2`、`-3`；每版定稿前套用[submission-check.template.md](../assets/templates/submission-check.template.md)自查，直到用户满意
8. 翻译（提交语言不是中文时）：按[deliverables.md](../references/deliverables.md)把定稿翻译成提交语言，参考文献改成目标语言格式，重新核对字数
9. 转换提交格式：问用户要交什么格式，再运行[convert_submission.py](../scripts/convert_submission.py)从最终语言的Markdown转换

## 产出

- `paper/<YYYY-MM-DD>_<题目>.md`及各修改版（中文底稿）
- 提交语言不是中文时另有翻译稿`.<语言代码>.md`
- 用户确认格式后转出的提交文件、`_assets/`里的图表图片
- 小组模式只写用户负责的章节，另有分工表
