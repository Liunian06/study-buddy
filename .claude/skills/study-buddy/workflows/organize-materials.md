# 资料整理

把老师发的课件、参考资料、往年题、作业要求转成按知识点排布的讲义和目录。

## 适用条件

- 进入：用户给出一个资料目录或一批文件，要整理成讲义；或某门课补充了新资料
- 不进入：用户要的是作业成品，改去[complete-assignment.md](complete-assignment.md)；资料已整理好、要开始学，改去[study-with-notes.md](study-with-notes.md)

## 输入

- 资料所在目录或文件，以及课程名；缺少时问用户
- 当前学期号：读`study-buddy-data/semester.md`，没有就问用户，见[data-layout.md](../references/data-layout.md)
- 首次使用时，用户需要先填写`config.json`里的VLM配置，见[setup.md](../references/setup.md)

## 步骤

1. 建课：读`semester.md`和已有课程文件夹，课程已存在就进入原文件夹，是新课就按[data-layout.md](../references/data-layout.md)取下一个课程序号，建课程文件夹，并在`semester.md`的课程清单登记（模板见[semester.template.md](../assets/templates/semester.template.md)）
2. 确认点：按[material-processing.md](../references/material-processing.md)列出候选文件（文件名、类型、大小），给出建议顺序，并标注每份的类型（课件/参考资料、教材、往年题、作业要求）。教材是几百页的整本书时，告诉用户页数和预计耗时，整本一次处理；OCR引擎默认用Paddle，需要用户先申请key；没填key时退用OXR（慢，且没有书页码），见[setup.md](../references/setup.md)。用户确认处理哪些、什么顺序之前，不运行CLI。首次运行前提示一次：课件会上传到OXR，图片会发给用户配置的VLM服务商
3. 调用CLI：把确认好的文件按顺序交给[ocr_cli.py](../scripts/ocr_cli.py)，参数见[setup.md](../references/setup.md)。CLI完成转PDF、切分、解析、下载图片、图片转写、合并和校验，agent只读最后输出的JSON摘要。资料多时放到后台运行
4. 处理结果：退出码0直接继续；退出码1时把失败的资料和页码告诉用户，让用户决定是否重新运行（CLI只重试失败的部分）；退出码2时按[setup.md](../references/setup.md)排查环境。摘要里有“未转写”的图时，说明数量，并提示填好`config.json`后重新运行即可补转写
5. 读草稿：按[material-processing.md](../references/material-processing.md)读`drafts/`，不打开图片；教材按该文件“教材”一节的方法定位小节，不整本通读
6. 建知识结构：合并重复内容，识别知识点和依赖关系，按先基础后应用重排，标考点信号和权重，得到知识结构和章节划分。有课件时课件定范围、教材补解释；只有教材时按[material-processing.md](../references/material-processing.md)判断怎么切知识点
7. 确认点：把知识结构、章节划分和考点权重给用户看，用户可以合并、拆分、调整顺序。确认前不写讲义，避免写完再返工
8. 转写讲义：套用[chapter.template.md](../assets/templates/chapter.template.md)，参照[chapter_example.md](../assets/examples/chapter_example.md)，一个知识点一节，写成`notes/chapterNN_<知识点名>.md`。资料多时分批写，每批报告进度
9. 建目录：套用[index.template.md](../assets/templates/index.template.md)生成`index.md`，学习进度初始全部为“未学”，并把资料缺失和未转写的位置写进去

## 产出

- `notes/`下的讲义和`index.md`，给用户看
- `manifest.json`、`media/`、`drafts/`是中间文件，只给模型用
- `semester.md`里登记了这门课

## 分支

- 补充资料：只把新文件交给CLI，已完成的资料按manifest自动跳过；新知识点插进已有结构，需要给已有讲义重新编号时，先列出受影响的文件，用户确认后再改
- 时间很紧：从期末速通跳来、且距考试1天以内时，做到第6步，只产出`index.md`和考点清单，不写完整讲义，见[material-processing.md](../references/material-processing.md)的精简模式

## 后续（可选）

- 用户想登记这门课的考核信息、规划每周任务时，进入[plan-semester.md](plan-semester.md)
- 用户想开始学时，进入[study-with-notes.md](study-with-notes.md)
