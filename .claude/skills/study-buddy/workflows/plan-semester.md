# 学期规划

综合各门课的考核，算出每门课的及格路径，排出全学期每周要准备的清单。

## 适用条件

- 进入：用户问“这学期怎么安排”“这周要做什么”“这门课怎么才能及格”，或给出了新的评分构成、成绩、截止时间
- 不进入：用户要整理资料，改去[organize-materials.md](organize-materials.md)；用户要开始学某一节，改去[study-with-notes.md](study-with-notes.md)

## 输入

- 各门课的评分构成、考核形式、截止时间、已有成绩，可以是口述，也可以是课程大纲文件（文件交给`ocr_cli.py`解析，类型选作业要求）
- `study-buddy-data/semester.md`，以及各课`index.md`里的知识点数量

## 步骤

1. 收集与更新：把用户给的信息按[semester.template.md](../assets/templates/semester.template.md)写进`semester.md`的课程清单和各课考核表；同一门课以最新一次更新为准。缺评分构成的课列出来问用户，没有评分构成就算不出及格路径
2. 算及格路径：按[pass-strategy.md](../references/pass-strategy.md)对每门课算还差多少分，选出必须做和可以放弃的考核，标出需要期末单独及格的课。参照[semester_example.md](../assets/examples/semester_example.md)
3. 排周计划：按[semester-planning.md](../references/semester-planning.md)倒排截止时间，给每项任务分配最低小时数，冲突时按风险排序，汇总每周总小时，某周偏高时把任务往前挪
4. 确认点：展示每门课的及格路径、可放弃的考核，以及每周总小时。如果按最低投入也明显超出可承受范围，如实说明，由用户决定怎么取舍。确认之前不把“放弃某项”写成定论
5. 写入并展示本周：把确认后的内容增量写进`semester.md`，在对话里展示本周要做的事和每项的小时数

## 产出

- `semester.md`中的及格路径和每周准备清单
- 对话里本周的任务

## 分支

- 用户问“这周做什么”：只读`semester.md`，输出本周部分，不重算
- 有新成绩或新截止时间：回到第1步做增量更新，只重算受影响的课和周
- 换学期：先问用户，确认后把`semester.md`改写成新学期的内容，旧学期的课程文件夹保留

## 后续（可选）

- 本周任务对应某个工作流时，直接进入：整理资料进[organize-materials.md](organize-materials.md)，带学进[study-with-notes.md](study-with-notes.md)，考前进[cram-for-exam.md](cram-for-exam.md)，pre进[prepare-presentation.md](prepare-presentation.md)，论文进[write-course-paper.md](write-course-paper.md)，其余作业进[complete-assignment.md](complete-assignment.md)
