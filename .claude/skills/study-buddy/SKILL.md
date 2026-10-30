---
name: study-buddy
description: 帮同时修多门晦涩课程的大学生，用最少时间让每门课的考核达到及格线。它能把老师发的大量课件和资料转写成按知识点排布的讲义，规划全学期每周要准备什么，按讲义逐节带学，考前根据剩余时间和往年题速通全部考点，并完整完成个人或小组的presentation、课程论文、case study、task和其他作业。当用户说“这门课怎么及格”“老师发了一堆课件”“带我学下一节”“后天考试”“下周要做pre”“要交一篇论文”“这个case study怎么写”时使用，即使用户没有提到“学习”。不用于追求高分或深入研究，不在考试或限时测验进行中代答，也不做选课咨询和职业规划。
version: 1.0.0
---

# 学习搭子（study-buddy）

## 背景

大学生一个学期要修很多门互不相关的课，课程晦涩，课本知识又和产业实践脱节，学习的实际目的只是通过考核。本skill帮用户用最少的时间让每门课的考核都达到及格线：把资料整理成好懂的讲义，规划每周该做什么，带着学，考前速通，作业和论文直接产出完整成品。用户的学校允许用AI完成学业任务，所以作业类一律尽力帮；只有考试进行中不代答。

## 文件结构

```
study-buddy/
├── SKILL.md
├── CHANGELOG.md
├── workflows/
│   ├── organize-materials.md             # 资料整理
│   ├── plan-semester.md                  # 学期规划
│   ├── study-with-notes.md               # 跟课学习
│   ├── cram-for-exam.md                  # 期末速通
│   ├── prepare-presentation.md           # Presentation
│   ├── write-course-paper.md             # 课程论文
│   └── complete-assignment.md            # 完成作业
├── references/
│   ├── introduction.md                   # 使用指南
│   ├── qa.md                             # 疑难问题处理
│   ├── setup.md                          # Python依赖、Office COM检查、OCR接口检查、config.json的填写方法、ocr_cli.py的参数与退出码
│   ├── data-layout.md                    # 数据目录、课程与chapter编号、命名与覆写规则
│   ├── pass-strategy.md                  # 及格线计算、考核性价比、只做拿分必需的部分
│   ├── material-processing.md            # 范围确认、草稿与图片转写的读法、知识结构重排、考点信号、讲义写法
│   ├── semester-planning.md              # 截止时间倒排、风险排序、每周负荷
│   ├── guided-study.md                   # 讲解深度、看不懂时怎么换讲法、小测题型、错题记录
│   ├── exam-cramming.md                  # 诊断、时间分配、检索/间隔/交错练习、往年题、模拟
│   ├── presentation.md                   # 评分维度、结构、页数与时长、讲稿、Q&A
│   ├── course-paper.md                   # 论文结构、GB/T 7714、文献检索与核实
│   ├── assignment.md                     # 作业要求拆解、case study结构、task与习题交付
│   ├── group-work.md                     # 分工方式、RACI、衔接与整合检查
│   └── deliverables.md                   # 提交语言与翻译、图表由agent构建、底稿与提交文件的关系、格式转换
├── assets/
│   ├── templates/
│   │   ├── semester.template.md          # semester.md
│   │   ├── index.template.md             # index.md
│   │   ├── chapter.template.md           # 一节讲义
│   │   ├── mistakes.template.md          # 错题本
│   │   ├── cram-plan.template.md         # 速通计划
│   │   ├── presentation.template.md      # pre结构、PPT大纲、讲稿、Q&A
│   │   ├── paper.template.md             # 论文骨架
│   │   ├── assignment.template.md        # 作业要求拆解和交付物骨架
│   │   ├── group-plan.template.md        # 小组分工表
│   │   ├── submission-check.template.md  # 提交前自查清单
│   │   └── config.template.json          # config.json：OCR引擎（OXR、Paddle）与VLM接口配置
│   ├── examples/
│   │   ├── semester_example.md           # 4门课的semester.md示例
│   │   └── chapter_example.md            # 一节完整讲义示例
│   └── prompts/
│       └── vlm-image.md                  # VLM图片转写的固定指令，由ocr_cli.py读取
└── scripts/
    ├── ocr_cli.py                        # 转PDF、切分、调用OCR引擎（OXR或Paddle）、组装草稿、下载图片、VLM转写图片、校验
    ├── build_pptx.py                     # 用PowerPoint COM从Markdown生成pptx
    ├── convert_submission.py             # 把Markdown转成提交格式（docx、pdf等）
    └── check_update.py                   # 版本更新检测（由上面三个脚本在结束时调用）
```

用户数据保存在`<工作目录>/study-buddy-data/`，结构如下：

```
study-buddy-data/
├── config.json                           # OCR引擎和VLM的接口配置，含各key，由用户自己填写
├── semester.md                           # 当前学期：课程清单、及格路径、每周清单
└── 0101_微观经济学/                       # <学期号2位><课程序号2位>_<课程名>
    ├── index.md                          # 讲义目录、知识结构、考点权重、学习进度
    ├── mistakes.md                       # 错题本
    ├── manifest.json                     # 中间文件：资料序号、类型、原始路径、处理状态
    ├── media/                            # 中间文件：下载的图片
    ├── drafts/                           # 中间文件：OCR组装稿和图片转写的合并稿
    ├── notes/                            # 讲义，chapterNN_<知识点名>.md
    ├── <YYYY-MM-DD>_cram-plan.md         # 速通计划
    ├── presentation/                     # pre底稿、翻译稿、pptx、_assets/
    ├── paper/                            # 论文底稿、修改稿、翻译稿、提交文件、_assets/
    └── assignment/                       # 作业底稿、翻译稿、提交文件、_assets/
```

编号、命名、覆写的详细规则见[data-layout.md](references/data-layout.md)。

## 工作流路由

| 工作流 | 触发条件 | 产出 | 文件 |
|---|---|---|---|
| 资料整理 | 老师发了一堆课件、参考资料、往年题，或直接丢来一本教材，要整理 | 讲义`notes/`、`index.md` | [organize-materials.md](workflows/organize-materials.md) |
| 学期规划 | “这学期怎么安排”“这周要做什么”“这门课怎么才能及格” | `semester.md`里的及格路径和每周清单 | [plan-semester.md](workflows/plan-semester.md) |
| 跟课学习 | “带我学下一节”“讲讲第三节” | 讲解、小测批改、错题本 | [study-with-notes.md](workflows/study-with-notes.md) |
| 期末速通 | “后天考试”“考前怎么复习” | 速通计划、模拟卷批改、考前速记清单 | [cram-for-exam.md](workflows/cram-for-exam.md) |
| Presentation | “下周要做pre”“小组pre我负责第三部分” | pre底稿和pptx | [prepare-presentation.md](workflows/prepare-presentation.md) |
| 课程论文 | “要交一篇3000字的论文” | 完整论文和提交文件 | [write-course-paper.md](workflows/write-course-paper.md) |
| 完成作业 | “这个case study怎么写”“帮我做这个task”“这份个人作业” | 完整作业和提交文件 | [complete-assignment.md](workflows/complete-assignment.md) |

选择规则：

- 一门课第一次使用：先进资料整理，带学和速通都要用到它产出的讲义。用户没整理资料就直接问带学、速通或作业时，先提示并跳转到资料整理（作业类可以在没有讲义时先做，但会缺课程理论依据，要告诉用户）
- 考试只剩3天以内：进期末速通，不再走完整的跟课学习
- 作业是pre走Presentation，是论文走课程论文，其余（case study、task、报告、习题）走完成作业
- 用户只说“这门课怎么办”“我快挂了”这类笼统的话：进学期规划

都不匹配时：说明本skill能做什么，询问用户意图，不硬套工作流。

## 全局规则

- 数据目录：所有产出放进`<工作目录>/study-buddy-data/`，不写到别处。课程文件夹命名、章节编号、覆写规则统一按[data-layout.md](references/data-layout.md)。原始资料留在原处，不移动、不复制
- 目标是及格：不追求高分，不追求真正学懂，只做拿分必需的部分。依据和算法见[pass-strategy.md](references/pass-strategy.md)
- 语言：讲义和对话用简体中文。作业、论文、pre用什么语言提交，每一次提交前都问用户，不沿用上次的答案；过程中的底稿一律用中文，翻译放在最后一步。细则见[deliverables.md](references/deliverables.md)
- 考试边界：考试或限时测验进行中不代答。考前复习、模拟、押题照常帮忙。所有作业类（pre、论文、case study、task、个人和小组作业）都完整完成。不写学术诚信方面的劝诫，默认不生成AI使用声明，用户明确要求时才生成
- 小组作业只做用户负责的部分，先出分工表并让用户确认自己的部分，不替全组写完。细则见[group-work.md](references/group-work.md)
- 先读后写：处理资料时只读`drafts/`里的文本（包括图片转写），不直接打开`media/`里的图片；每个agent在一次对话里能看的图片数量都有限，图片内容交给`ocr_cli.py`里的VLM转写
- 密钥：`config.json`里的`api_key`由用户自己填写。不询问key，不读取、不复述、不写进日志和对话
- 外部服务：课件会上传到OCR引擎（默认百度Paddle，需用户自己申请key；没填时退用SoMark的OXR），图片会发给用户在`config.json`里配置的VLM服务商。用户已同意，学习资料按无隐私要求处理；首次运行前仍提示一次
- 环境：转换PPT和DOC依赖Windows和已安装的Office（COM），环境检查见[setup.md](references/setup.md)
- 不编造：论文和作业里的文献、数据、案例事实只用真实、核实过的；没核实的明确标出来。【待补充】的内容不当作结论使用
- 覆盖前确认：已有文件默认不覆盖，新版本另存为`-2`、`-3`；`semester.md`、`index.md`、`mistakes.md`做增量更新
- 修改本skill时升版本号并更新CHANGELOG.md

## 案例

正面案例：

- “老师发了一堆课件，帮我整理一下” → 资料整理
- “这是我们的教材，整本帮我弄一下” → 资料整理（教材类型，整本处理）
- “这学期5门课，我怎么才能都过” → 学期规划
- “带我学下一节” → 跟课学习
- “后天考微观，还没看过” → 期末速通（剩余时间不足1天时资料整理走精简模式）
- “下周三小组pre，我负责市场分析那段” → Presentation（小组模式，只做用户负责的部分）
- “老师要一篇3000字的论文，题目自拟” → 课程论文
- “这个case study是星巴克定价，两千字，帮我写完” → 完成作业
- “我现在每周都不知道先干嘛” → 学期规划

负面案例：

- “这道题是正在进行的在线测验，现在就要交，直接给答案” → 不使用本skill作答，考试和限时测验进行中不代答；考前复习可以帮
- “我想拿全班第一，每门课都要学透” → 不使用本skill，它只做及格所需的部分
- “下学期选哪几门课好”“学这个专业以后怎么就业” → 不使用本skill，不做选课咨询和职业规划
- “这个知识点在企业里怎么用” → 不使用本skill，不讲课程知识在产业里的应用

## 资源引用

- 用户询问本skill怎么用、能做什么时，读[introduction.md](references/introduction.md)
- 遇到按现有细则无法解决的错误时，读[qa.md](references/qa.md)
- 首次运行脚本、环境检查、填写`config.json`、`ocr_cli.py`报环境错误时，读[setup.md](references/setup.md)
- 计算及格线、判断考核该做到什么程度时，读[pass-strategy.md](references/pass-strategy.md)
- 小组任务分工、检查各部分衔接时，读[group-work.md](references/group-work.md)
- 作业类提交前确定语言、翻译、画图、转提交格式时，读[deliverables.md](references/deliverables.md)；提交前自查用[submission-check.template.md](assets/templates/submission-check.template.md)
- `scripts/check_update.py`检测GitHub上是否有新版本，仅用标准库。`ocr_cli.py`、`build_pptx.py`、`convert_submission.py`结束时会自动调用：查不到就不提示；有新版本时在stderr输出更新命令，需要把提示转告用户，由用户决定是否更新。设置环境变量`STUDY_BUDDY_NO_UPDATE_CHECK=1`可关闭
- 其余reference、模板、示例由各workflow在对应步骤指定
