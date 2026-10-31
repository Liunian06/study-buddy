# 首次使用

第一次用本skill时，按顺序完成环境检查、数据目录和配置文件创建、学期总览建立，然后进入第一门课的资料整理。已经做过的步骤跳过，不重复做，更不覆盖用户已有的文件。

## 适用条件

- 进入：用户说“帮我初始化”“第一次用”“怎么开始”“帮我装好环境”，或工作目录下还没有`study-buddy-data/`
- 不进入：`study-buddy-data/`和`config.json`都已存在，用户只是要整理新资料，改去[organize-materials.md](organize-materials.md)

## 输入

- 工作目录（当前目录）
- 当前学期，如“大三上”，没有就问用户，学期号见[data-layout.md](../references/data-layout.md)

## 步骤

1. 检查环境：按[setup.md](../references/setup.md)的“验证”一节检查Windows、Office、Python版本和依赖。缺依赖时先问用户要装到哪个Python环境，用户同意后再装，不改系统执行策略。不满足条件的，如实告诉用户哪几步用不了，如没有Office就转不了PPT和DOC
2. 建数据目录和配置文件：运行`python scripts/ocr_cli.py --init-config`，在当前目录下建好`study-buddy-data/`并从模板生成`config.json`；已经存在的不覆盖。把`config.json`的路径告诉用户，请用户自己填：`paddle.api_key`（OCR引擎，不填会退用较慢的OXR）和`vlm`下的`base_url`、`model`、`api_key`（转写图片，不填则图片标为“未转写”）。模板里的`sk-123`等是示例值，不是真key，见[setup.md](../references/setup.md)。agent不询问、不读取、不复述key
3. 建学期总览：套用[semester.template.md](../assets/templates/semester.template.md)生成空的`学期规划.md`，填学期名、学期号和今天的日期，其余表格留空；文件已存在则不动
4. 建第一门课：问用户第一门课的名字，进入[organize-materials.md](organize-materials.md)的第1步建课，之后按那个工作流继续

## 产出

- `study-buddy-data/config.json`（用户自己填key）
- `study-buddy-data/学期规划.md`（空总览）
- 第一门课的文件夹，含`原始资料/`、`课程资料/`和空的`课程信息.md`

## 分支

- 用户只想先装环境、不想马上整理资料：做到第3步，告诉用户资料放在哪个目录、下次直接说“整理这门课的资料”即可
- 用户还没有课程资料：做到第3步，同上
- `config.json`已存在：不动，只提醒用户确认key已填好
- 用户暂时不想配置VLM：可以继续，告诉用户图片不会被转写，填好后重新运行`ocr_cli.py`即可补转写

## 后续

- 用户给了评分构成和截止时间：进入[plan-semester.md](plan-semester.md)
- 用户给了第一批资料：进入[organize-materials.md](organize-materials.md)
