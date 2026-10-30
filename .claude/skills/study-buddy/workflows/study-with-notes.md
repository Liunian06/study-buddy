# 跟课学习

按讲义带用户学一节，学到能答题为止，并及时发现没懂的地方。

## 适用条件

- 进入：用户说“带我学下一节”“讲讲第三节”，且这门课已有讲义
- 不进入：这门课还没有讲义，改去[organize-materials.md](organize-materials.md)；考试只剩3天以内，改去[cram-for-exam.md](cram-for-exam.md)

## 输入

- 课程名，以及指定的某一节或“下一节”
- 读取该课的`index.md`（学习进度）、对应的`notes/chapterNN_*.md`和`mistakes.md`

## 步骤

1. 定位：用户指定了节就用那一节，没指定就从`index.md`里取第一个“未学”或“需复习”的节。前置知识点还没学时提醒一句，不强制先学
2. 讲解：按[guided-study.md](../references/guided-study.md)讲，先白话建立直觉，再讲术语和例子，随时回答用户的提问。用户说看不懂时按该文件换讲法
3. 小测：出2–4道题，用户作答前不给答案；用户作答后批改并指出错因
4. 记录：答错的题按[mistakes.template.md](../assets/templates/mistakes.template.md)写进`mistakes.md`；按小测结果把`index.md`里这一节标为“已学”或“需复习”；用户反馈讲义哪里难懂时，直接修订那一节

## 产出

- 讲解和小测批改
- 更新后的`index.md`进度和`mistakes.md`

## 后续（可选）

- 用户要继续时，重复以上步骤学下一节
- 临近考试时，进入[cram-for-exam.md](cram-for-exam.md)
