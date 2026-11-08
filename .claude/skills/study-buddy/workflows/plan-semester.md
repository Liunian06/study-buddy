# 学期规划

综合各门课的考核，按每门课的目标（保过或高分）算出路径，排出全学期每周要准备的清单。

## 适用条件

- 进入：用户问“这学期怎么安排”“这周要做什么”“这门课怎么才能及格”“这门课想冲90”，或给出了新的评分构成、成绩、截止时间、目标
- 不进入：用户要整理资料，改去[organize-materials.md](organize-materials.md)；用户要开始学某一节，改去[study-with-notes.md](study-with-notes.md)

## 输入

- 各门课的评分构成、考核形式、截止时间、已有成绩，可以是口述，也可以是课程大纲文件（文件交给`ocr_cli.py`解析，类型选作业要求）
- `study-buddy-data/学期规划.md`（跨课总览），各课文件夹里的`课程信息.md`（单课考核信息），以及各课`讲义目录.md`里的知识点数量

## 步骤

1. 收集与更新：把用户给的信息按[course-info.template.md](../assets/templates/course-info.template.md)写进**对应那门课**的`课程信息.md`（文件不存在就新建），只读写这一门课自己的文件；同一门课以最新一次更新为准。评分构成是可选信息：学期末才最终评定，什么时候给都行，用户没给就不问、不催，也不因此停下；用户顺口提到评分、成绩或截止时间，直接记下，不追问
2. 算路径：只对填了评分构成的课，按[score-strategy.md](../references/score-strategy.md)算还差多少分（保过课对及格线，高分课对目标分），标出需要期末单独及格的课，结果写回该课的`课程信息.md`；没填的课这一步跳过，相关栏写“未填”，不猜比例。每项考核默认都按做满安排，保过课的“可以放弃”只在按最低投入也排不下、用户同意时才写，高分课不建议放弃。参照[course-info_example.md](../assets/examples/course-info_example.md)
3. 排周计划：按[semester-planning.md](../references/semester-planning.md)倒排截止时间（没有评分构成的课按截止时间和任务量排，每项照样按做满估小时），给每项任务分配最低小时数。先把所有课排到保过线，剩余时间再分给高分课，冲突时按风险排序，汇总每周总小时，某周偏高时把任务往前挪
4. 确认点：展示每门课的目标和路径、保过课可放弃的考核、高分课的额外投入小时，以及每周总小时。如果按最低投入也明显超出可承受范围，或高分课的目标分和可用时间对不上，如实说明差多少、缺多少小时，由用户决定怎么取舍（降目标、减别的课的投入、接受风险）。确认之前不把“放弃某项”写成定论
5. 写入并展示本周：把确认后的路径写进各课的`课程信息.md`；再按[semester.template.md](../assets/templates/semester.template.md)更新`学期规划.md`的总览，规则见[data-layout.md](../references/data-layout.md)的“两处信息的同步”；在对话里展示本周要做的事和每项的小时数。参照[semester_example.md](../assets/examples/semester_example.md)

## 产出

- 各课`课程信息.md`中的考核表和路径（保过路径或高分路径）
- `学期规划.md`中的课程清单、截止时间和每周准备清单
- 对话里本周的任务

## 分支

- 用户问“这周做什么”：只读`学期规划.md`的本周部分，输出，不重算，也不用去读各课的`课程信息.md`
- 有新成绩或新截止时间：回到第1步做增量更新，只读写受影响那门课的`课程信息.md`，再更新`学期规划.md`里这门课的行，只重算受影响的课和周
- 用户说“这门我想冲90”“这门保过就行”：只改那一门课的`课程信息.md`里的“目标”，重算这门课的路径，再更新`学期规划.md`里这门课的行和受影响的周；目标分要具体，用户只说“尽量高”就问一个分数，见[score-strategy.md](../references/score-strategy.md)
- 课程没有评分构成：不问，不算路径，直接按截止时间排每周清单；用户问“这门课怎么才能及格”时，说明没有评分构成算不出离目标差多少、随时补充即可，同时先给出每周清单。用户之后给了，回到第1步增量更新
- 换学期：先问用户，确认后把`学期规划.md`改写成新学期的内容。旧学期的课程文件夹连同其中的`课程信息.md`原样保留，上学期的评分构成、及格路径和已得成绩因此仍可回看

## 后续（可选）

- 本周任务对应某个工作流时，直接进入：整理资料进[organize-materials.md](organize-materials.md)，带学进[study-with-notes.md](study-with-notes.md)，考前进[cram-for-exam.md](cram-for-exam.md)，pre进[prepare-presentation.md](prepare-presentation.md)，论文进[write-course-paper.md](write-course-paper.md)，其余作业进[complete-assignment.md](complete-assignment.md)
