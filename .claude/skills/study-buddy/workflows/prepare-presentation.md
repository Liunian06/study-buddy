# Presentation

从作业要求到可直接上台的pre：结构、PPT大纲、讲稿、Q&A和pptx。个人或小组都适用，小组只做用户负责的部分。

## 适用条件

- 进入：用户说“下周要做pre”“小组pre我负责第三部分”
- 不进入：作业是论文，改去[write-course-paper.md](write-course-paper.md)；是case study、task、报告、习题，改去[complete-assignment.md](complete-assignment.md)

## 输入

- 作业要求（主题、时长、评分标准、个人还是小组、要交什么），是文件时交给`ocr_cli.py`解析，类型选作业要求
- 相关资料；主题与课程内容相关时，读取该课的`讲义/`

## 步骤

1. 拆要求：按[presentation.md](../references/presentation.md)提取时长、页数、必须包含的内容和评分维度，得到要求清单；问用户这次提交用什么语言，见[deliverables.md](../references/deliverables.md)
2. 分模式：个人直接进入下一步。小组先按[group-work.md](../references/group-work.md)出分工表，用户确认自己负责的部分之前，不往下做
3. 搭结构：按[presentation.md](../references/presentation.md)定核心信息，排引言、主体、结论，估页数。小组只展开用户负责的部分，写好前后衔接
4. 写内容：套用[presentation.template.md](../assets/templates/presentation.template.md)逐页写大纲和口语化讲稿；需要图表时按[deliverables.md](../references/deliverables.md)由agent构建
5. 预演Q&A：预判5–8个可能被问到的问题，各写一段简答
6. 自查：套用[submission-check.template.md](../assets/templates/submission-check.template.md)，对照要求清单和评分维度逐项检查，不通过先修正
7. 翻译（提交语言不是中文时）：按[deliverables.md](../references/deliverables.md)把定稿翻译成提交语言，重新核对讲稿时长
8. 生成pptx：以最终语言的Markdown为底稿，运行`python scripts/build_pptx.py <底稿>`，检查页数、标题、备注栏和图片；脚本只排版，不套模板，用默认主题

## 产出

- `演示文稿/<YYYY-MM-DD>_<题目>.md`（中文底稿），提交语言不是中文时另有`.<语言代码>.md`
- 同名`.pptx`、`_图表/`里的图表图片
- 小组模式另有分工表

## 分支

- 用户要自己调整样式、套学校模板：脚本不处理，生成后由用户在PowerPoint里调整
