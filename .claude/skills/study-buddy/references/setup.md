# 运行环境与安装

本skill支持Windows、macOS和Linux。涉及Office文件的三步（`ocr_cli.py`转换PPT和DOC、`build_pptx.py`生成pptx、`convert_submission.py`导出PDF）按平台选工具，由脚本自动判断：

| 平台 | 转PPT、DOC、docx为PDF | 生成pptx |
|---|---|---|
| Windows | 默认Office COM（PowerPoint、Word），排版最接近原文件；没装Office或COM出错时自动退到LibreOffice | 默认PowerPoint COM；没装PowerPoint或COM出错时自动退到`python-pptx` |
| macOS、Linux | LibreOffice（`soffice --headless`） | `python-pptx` |

Windows上两者都没有时，转换步骤报错并给出安装建议；也可以让用户手动另存为PDF后再交给脚本。macOS、Linux没有Office COM，所以不装`pywin32`。

## 环境要求

- Windows 10/11（本机验证），或macOS、Linux。Windows上建议安装Office（PowerPoint和Word，本机验证版本为16.0）；没有Office或在macOS、Linux上，需要安装LibreOffice：Windows `winget install TheDocumentFoundation.LibreOffice`，macOS `brew install --cask libreoffice`，Linux `sudo apt install libreoffice`
- Python 3.10以上（本机验证版本为3.14）
- OCR引擎有三家供应商：Paddle（百度，PaddleOCR-VL-1.6）、OXR（SoMark）、MinerU，key互不通用。默认用Paddle。没填某家的key时，该家只能用免费或免登录的接口（OXR免费接口、MinerU轻量接口），Paddle没有免费路径。
- 自动选择和失败退路用同一条优先级链：**paddle key → oxr key（商业版）→ mineru key（精准解析）→ oxr免费接口 → mineru免登录轻量接口**。没填key的档位跳过（oxr免费接口、mineru轻量接口不需要key，永远在链里）；运行从第一个可用的档位开始，某一档失败时依次退到后面的档位，退用的原因记在摘要的`warnings`里。用`--engine`切换起点，或在`config.json`里写`engine`：

| 引擎 | 要不要key | 单块上限 | 特点 |
|---|---|---|---|
| `oxr` | 免费接口不要（`api_key`留示例值或留空）；商业版要，填进`oxr.api_key` | 20页、10MB（用户给出的免费版上限，商业版沿用） | 备选。SoMark，同步。免费接口同一时间只处理一个任务，实测约5秒一页，不保证稳定，没有隐私声明；填了key就改走商业版地址`https://somark.ai/api/v1`，请求方式相同。没有书页码和页眉页脚信息 |
| `paddle` | 要，到 https://aistudio.baidu.com/paddleocr 申请 | 100页（免费版，用户给出）、本地上传50MB（官方文档） | 默认。百度PaddleOCR-VL-1.6，异步提交加轮询，整本443页教材实测约15分钟、0个失败块。免费额度每天2万页，0点刷新（用户给出，未核实） |

- Paddle适合整本教材这类大文件。额度用完或key无效时（HTTP 401、403），这份资料后面的块自动退到链里的下一档（见上），已完成的块保留；链上所有档位都失败才停止提交后面的块，次日重新运行即可补齐。退用意味着资料会上传到另一家供应商，页码精度和图片处理也可能不同，要在摘要里告诉用户
- 两个引擎可以混用：同一份资料先用一个、失败的块换另一个重跑，草稿格式一样
- Paddle读取的是它返回的块级JSON（`parsing_res_list`），不是Markdown。好处：页眉、页码、页脚不混进正文；页码块直接记成书页码，草稿标记里写成`p.061 | 书p.46`；公式编号并进公式（`\tag{4.8}`）；图和图注挨在一起。OXR没有这些信息
- VLM服务（可选）：任何支持图片输入的、兼容OpenAI Chat Completions格式的服务。没有配置时，图片会标为“未转写”，其余照常

## 安装

Windows：

```bash
python -m pip install pymupdf==2.0 pywin32==312 requests python-docx==1.2.0 python-pptx==1.0.2 markdown-it-py pillow
```

macOS、Linux（不装`pywin32`，它只有Windows能装）：

```bash
python3 -m pip install pymupdf==2.0 requests python-docx==1.2.0 python-pptx==1.0.2 markdown-it-py pillow
```

`python-pptx`是没有PowerPoint时生成pptx的后备方案，Windows上装着也无妨，没装PowerPoint时才会用到。

### LibreOffice（macOS、Linux必需；Windows没有Office时需要）

```bash
brew install --cask libreoffice          # macOS
sudo apt install libreoffice             # Linux（Debian、Ubuntu）
winget install TheDocumentFoundation.LibreOffice   # Windows
```

- 脚本会在PATH和各系统的默认安装位置（macOS是`/Applications/LibreOffice.app`）里找`soffice`，装完不用手动配置
- 转换时用独立的临时用户配置，不和用户正在开着的LibreOffice抢锁；每份文件首次启动要几秒到十几秒
- LibreOffice转出的PDF与Office版式有细微差别（特殊字体、艺术字）。OCR取字不受影响；提交用的PDF如果对版式有严格要求，建议在装了Word的电脑上导出

### TeX（只有整理考试提纲才需要）

考试提纲工作流用`compile_tex.py`把tex编译成PDF，需要`xelatex`：

```bash
winget install MiKTeX.MiKTeX                  # Windows，比TeX Live轻很多
brew install --cask mactex-no-gui             # macOS（BasicTeX更小，需自己用tlmgr补宏包）
sudo apt install texlive-xetex texlive-lang-chinese   # Linux
```

- 装完重开终端，让PATH生效；之后用`xelatex --version`确认能找到。已经装了TeX Live的也可以直接用。脚本在macOS上还会到`/Library/TeX/texbin`里找
- MiKTeX首次编译会自动下载缺少的宏包（中文需要ctex等），需要联网，可能慢几分钟；脚本在MiKTeX下会自动打开自动安装开关
- 没装也不影响其他功能。整理提纲时没装xelatex，tex源文件仍会保存，装好后再编译
- 中文字体：tex模板里写的是`fontset=windows`（宋体、黑体、楷体），`compile_tex.py`编译时按平台自动换成ctex的对应字体集：Windows用`windows`，macOS用`mac`（系统自带的宋体、黑体、楷体），Linux用`fandol`（随TeX Live附带的开源字体）。所以tex源文件不用改，输出JSON里的`fontset`记录了实际用的字体集。字体不同，各平台编译出的PDF版式和页数可能略有差别，小抄限页数时以当前平台编译结果为准

## 验证

Windows：

```bash
python -c "import pymupdf, win32com.client, requests, docx, pptx, markdown_it, PIL; print('ok')"
```

macOS、Linux：

```bash
python3 -c "import pymupdf, requests, docx, pptx, markdown_it, PIL; print('ok')" && soffice --version
```

预期输出：`ok`（macOS、Linux还有LibreOffice的版本号）。Windows上装了Office可以不装LibreOffice，没装Office则要装LibreOffice。

## config.json

`config.json`在数据目录（`study-buddy-data/`）根部，模板见[config.template.json](../assets/templates/config.template.json)。有两种方式生成，生成后都由用户自己填key：

- **推荐**：首次使用时运行`python scripts/ocr_cli.py --init-config`。只生成配置文件就退出，已存在则不覆盖，这样用户可以先填好key再整理资料，流程见[first-time-setup.md](../workflows/first-time-setup.md)
- 兜底：没做初始化就直接运行`ocr_cli.py`处理资料时，脚本会自动生成`config.json`，**不会中断**，而是按空配置继续处理：OCR从优先级链里第一个可用的档位开始（没有任何key时是OXR免费接口，失败再退用MinerU免登录轻量接口），图片标为“未转写”，摘要的`warnings`里提示。用户填好key后重新运行即可补转写

| 字段 | 含义 |
|---|---|
| `engine` | 默认引擎，`paddle`、`oxr`或`mineru`，模板里是`paddle`；命令行的`--engine`优先 |
| `paddle.api_key` | 用户自己申请的key。模板里是示例值`paddle-token-123`；不填或保留示例值时，跳过Paddle这一档，从优先级链的下一档开始 |
| `paddle.model` | 默认`PaddleOCR-VL-1.6`，能输出带公式和表格的Markdown。`PP-OCRv6`只返回带框的图，没有文字，不适用 |
| `paddle.max_pages_per_chunk`、`paddle.max_mb_per_chunk` | 单块页数和大小上限，默认100页、45MB |
| `oxr.base_url` | OXR接口地址，默认免费接口；改用正式版时换成正式地址 |
| `oxr.api_key`、`oxr.paid_base_url` | 商业版key，模板里的示例值`OXR_API_KEY`（旧版的`dummy`同样）视为没填，走免费接口；填了真实key就用`paid_base_url`（默认`https://somark.ai/api/v1`） |
| `oxr.max_pages_per_chunk`、`oxr.max_mb_per_chunk` | 单块页数和大小上限，默认20页、10MB |
| `mineru.api_key` | MinerU的token，在MinerU官网的API管理页创建。模板里是示例值`MinerU_API_KEY`，保留示例值时视为没填。**填了**走精准解析接口，用`vlm`模型，有分页和图片文件；**没填**走免登录的轻量接口（pipeline模型，按IP限频）：每块最多20页、10MB，只返回一份Markdown，没有分页，所以草稿里页码只能标到块级（“第N–M页的合并输出”），也没有表格和公式识别 |
| `mineru.base_url`、`mineru.model_version` | 接口地址默认`https://mineru.net`；模型默认`vlm`（文档推荐），也可用`pipeline` |
| `mineru.max_pages_per_chunk`、`mineru.max_mb_per_chunk` | 单块页数和大小上限，默认100页、100MB。接口本身的上限是200页、200MB，留出余量；块小一些，中断后重做的部分也少 |
| `mineru.poll_seconds`、`mineru.job_timeout_seconds`、`mineru.agent_job_timeout_seconds` | 轮询间隔默认3秒；精准解析单个任务最长等待默认1800秒（实测多个任务同时排队时，等待可达数分钟）；轻量接口没有队列优先级，排队时间不可控，等待上限单独设为120秒，超过视为失败（它是链上最后一档，没有后续档位可退，这一块标记失败）。实测它在一台机器上提交后30多分钟仍在排队，所以不要设太长；网络和接口正常时几十秒内应有结果 |
| `vlm.base_url` | 服务地址，模板里是示例`https://api.openai.com/v1`，脚本会请求`<base_url>/chat/completions`；用其他兼容服务商时换成它的地址 |
| `vlm.model` | 模型名，必须支持图片输入；模板里是示例`gpt-4o` |
| `vlm.api_key` | 用户自己填写的key。模板里是示例值`sk-123`，保留示例值时视为没填，VLM不启用 |
| `vlm.min_image_pixels` | 图片最短边小于这个值时视为图标，跳过转写，默认64 |
| `vlm.max_concurrency` | 同时在途的VLM请求数上限，默认10，用线程池实现（不是分批）；服务商限流（HTTP 429）时调小，最小为1 |
| `vlm.max_attempts` | 单张图片请求失败后的重试上限，按总尝试次数计（含第一次），默认3，最小为1（不重试）；两次尝试之间依次等2秒、4秒 |
| `download.max_attempts` | OCR引擎返回的图片链接下载失败后的重试上限，按总尝试次数计，默认3；403、404、410等说明链接已失效的状态不重试，直接补OCR |
| `download.max_concurrency` | 同时下载的图片数，默认4。下载和后面块的OCR同时进行 |
| `download.max_reocr_failures` | 连续这么多页单页补OCR后图片仍下载失败，就停止补OCR，默认3。这种情况多半是网络问题而不是链接过期，继续补只会白白消耗额度 |

- 图片转写用的固定指令在[vlm-image.md](../assets/prompts/vlm-image.md)，由脚本直接读取；想调整转写风格时改这个文件
- agent不会询问key，也不读取或复述它（包括`vlm.api_key`、`paddle.api_key`和`mineru.api_key`）；脚本的日志和输出里不出现key
- `vlm.api_key`、`paddle.api_key`和`mineru.api_key`以明文保存。工作目录如果处在网盘同步的文件夹里，key会随文件一起同步；用户要避免时，把数据目录放到不同步的位置，或用`--data-dir`指向别处
- 课件会上传到OXR，图片会发给用户在`vlm`里配置的服务商。用户已同意，学习资料按无隐私要求处理

## 脚本用法

| 脚本 | 用途 | 命令 |
|---|---|---|
| `scripts/ocr_cli.py` | 把一个或多个文件转成带页码和图片转写的Markdown草稿 | `python scripts/ocr_cli.py <文件...> --course-dir <课程文件夹> --types material` |
| `scripts/build_pptx.py` | 从pre底稿生成pptx | `python scripts/build_pptx.py <底稿.md>` |
| `scripts/convert_submission.py` | 把Markdown转成docx或pdf | `python scripts/convert_submission.py <稿.md> --format docx` |
| `scripts/compile_tex.py` | 把考试提纲的tex编译成PDF | `python scripts/compile_tex.py <文件>.tex` |

### ocr_cli.py的参数

| 参数 | 含义 |
|---|---|
| `files` | 一个或多个文件（ppt、pptx、doc、docx、pdf、ipynb），按传入顺序编号 |
| `--course-dir` | 课程文件夹，输出写在这里 |
| `--data-dir` | 放`config.json`的数据目录，默认是课程文件夹的上一级；与`--init-config`一起用时默认是当前目录下的`study-buddy-data/` |
| `--init-config` | 只生成`config.json`（已存在不覆盖）然后退出，不需要传资料文件和`--course-dir` |
| `--types` | `material`课件/参考资料、`textbook`教材、`past-exam`往年题、`assignment`作业要求；逗号分隔与文件一一对应，或只写一个应用到全部 |
| `--pages` | 只处理指定页（1起，如`10,17,20-22`），用于测试 |
| `--engine` | `paddle`、`oxr`或`mineru`，默认取`config.json`的`engine`，没写则用`paddle`。优先级链是paddle key → oxr key → mineru key → oxr免费接口 → mineru轻量接口：默认引擎是paddle但没填key时，从链上第一个可用的档位开始并在摘要的`warnings`里提示；`--engine oxr`或`mineru`从该家在链里最靠前的可用档位开始（有key用key，没有用免费或免登录接口）；运行中某一档失败，依次退到后面的档位；显式`--engine paddle`又没有key则报错退出，不悄悄换引擎 |
| `--keep-temp` | 保留临时PDF和切块 |
| `--dry-run` | 只预估不处理，输出每份资料的页数、将用的引擎和OCR耗时范围，然后退出（退出码0）。不联网、不用key、不改`manifest.json`、不生成`config.json` |

- `--dry-run`的速度假设：Paddle按30–50页/分钟（用户给出），OXR按约12页/分钟（约5秒一页，单次实测）。只估OCR解析；Office转PDF、图片转写（VLM）、下载图片的时间不计入，所以实际总耗时会比预估长，图片多时尤其明显
- PPT和Word没有页数，`--dry-run`要先转成PDF才能数页，每份要几十秒。转出的PDF缓存在`<课程文件夹>/.临时/预转换/`（键含路径、修改时间和大小，文件改过会重转），正式运行直接复用，处理成功后自动删除
- 摘要里已完成的资料标“跳过（已完成）”，中断过的标“续跑”，只算没完成的块；ipynb不走OCR，只报单元格数

- 退出码：0全部成功；1部分失败（某块解析失败、图片下载失败或图片转写失败），重新运行只会重试失败的部分；2参数或环境错误
- 图片下载：OCR引擎返回的图片是临时链接，会过期。每解析完一块就马上下载该块的图片，和后面块的OCR同时进行，不等整份资料解析完。下载失败先重试（`download.max_attempts`）；仍失败，或链接已失效（403、404等），就对这一页单独再请求一次OCR拿新链接，再下载。补OCR在主流程里串行做，因为免费OXR同一时间只处理一个任务；每页最多补一次。上次中断留下的“已解析但缺图”的块，重新运行时会先下载，不重新OCR
- 摘要里的`image_download`：`retries`是重试成功的次数，`reocr_pages`是补过OCR的页码，`reocr_stopped`为真表示因连续失败停止了补OCR，`failed`是最终仍没下载下来的图片数。`reocr_stopped`为真或`failed`不为0时，先检查网络、代理和防火墙，再重新运行；补OCR会消耗OCR额度，所以不要反复重跑
- 脚本异常退出（出现Python traceback，或退出码不是上面列出的值）不在这些约定内：按[qa.md](qa.md)整理问题描述，并按[community.md](community.md)给出反馈群二维码（过期后是作者的微信二维码）和Issues链接
- 已完成的资料会自动跳过；配置好VLM后重新运行，只补转写没转写的图
- 断点续传：每个块解析完（成功或失败）就保存进度，所以脚本被中断（报错退出，或被强杀、断电、后台任务被取消）后重新运行，已解析完的块不会重新OCR，已下载的图片也不会重新下载，只丢弃中断时正在进行、还没返回结果的那一块。唯一的例外是Paddle已提交但还没出结果的任务：任务编号不保存，重跑会重新提交，这一块的额度会再用一次
- 输出：stdout是JSON摘要，每份资料的状态、页数、图片数、已转写和未转写数、草稿路径、失败页码；stderr是进度日志
- 进度实时输出到stderr，进度行统一带`[进度]`前缀，内容有：
  - 每份资料开始和结束，结束行含状态、页数、失败页、图片情况和用时
  - 每个块完成时的累计：已处理页数、百分比、已用时间；第一块完成后还有预计剩余（按已完成的块外推）
  - OCR等待期间每60秒一行“仍在处理，已等待X”，用来判断没有卡住
  - 图片下载和图片转写的n/total（转写最多打20行左右，图片多时不刷屏）
  - 全部结束时一行总结
- 资料多、耗时长时放到后台运行；助手运行期间读这几行并转告用户，见[organize-materials.md](../workflows/organize-materials.md)的第3步

### build_pptx.py和convert_submission.py

- `build_pptx.py`读取pre底稿里“### 第N页 标题”下的“要点”“图表”“讲稿”“备注”；讲稿和备注写进备注栏；用默认主题，不套模板；输出已存在时另存为`-2`
- `convert_submission.py`保留标题层级、粗斜体、列表、表格、代码块、引用块和图片；公式（`$$...$$`）按文本原样保留；不添加封面、页眉页脚、目录。pdf先生成docx，再导出（Windows优先Word COM，没有Word或在macOS、Linux上用LibreOffice）

### compile_tex.py的参数

| 参数 | 含义 |
|---|---|
| `tex` | tex源文件，UTF-8编码 |
| `-o`、`--output` | 输出PDF路径，默认与tex同名；`--fill`时文件名加“_填空版”。同名PDF会被覆盖 |
| `--fill` | 编译填空版：模板里的`\kw{答案}`变成下划线空白；不加时是加粗 |
| `--max-pages` | 页数上限（小抄填1）。超页时依次缩小字号、行距和页边距（9pt到6pt）重新编译，缩到最小仍超则返回`overflow` |
| `--keep` | 保留临时编译目录，排查用 |

- 退出码：0成功；1编译失败或超页，按输出里`errors`的行号改tex后重新编译；2参数或环境错误（找不到xelatex等）
- 输出：stdout是JSON，含`status`（ok、error、overflow、no_engine）、`pages`、`pdf`、`scale`（缩放后的字号等）、`last_page_fill`（最后一页内容占页面高度的比例，小抄用来判断还能不能补内容）、`warnings`（缺字、行宽溢出）、`errors`和`hints`；stderr是进度日志
- 脚本在临时的ASCII目录里编译，所以tex路径含中文或空格没有问题；每次最多编译三遍，让目录、页码引用和longtable稳定

### 版本更新检测

`ocr_cli.py`、`build_pptx.py`、`convert_submission.py`、`compile_tex.py`运行结束时，会调用`check_update.py`查询GitHub上最新的Release；有新版本时在stderr提示更新命令，并注明“自动更新：是/否”，同时附上本地版本之后到最新版本之间所有版本的CHANGELOG段落（按最新版本的tag从GitHub取，新的在前，最多约6000字，超出截断；取不到时只给Release链接）：同一主版本升级时助手直接运行更新命令，不用问用户；主版本升级（可能有不兼容的变更）时不自动更新，先告诉用户更新内容，用户同意后再更新。每次对话第一次用到本skill时，助手也会先运行一次`python scripts/check_update.py`。脚本按自己所在位置判断安装方式：装在用户主目录下（全局安装）提示`npx skills update study-buddy -g`；装在其他文件夹里（项目内安装）提示不带`-g`的`npx skills update study-buddy`，并写明要在哪个文件夹里运行。请求超时2秒，结果缓存（成功24小时、失败6小时）；断网、被限流等任何查询失败都不提示，不影响脚本结果。设置环境变量`STUDY_BUDDY_NO_UPDATE_CHECK=1`可同时关闭检测和自动更新。

## 网络与权限

- `ocr_cli.py`访问所选的OCR接口和（配置后）VLM服务；引擎返回的图片链接会被下载到本地
- `check_update.py`访问`api.github.com`查询最新版本；发现有新版本时再访问`raw.githubusercontent.com`取该版本的`CHANGELOG.md`。都是只读，不上传任何数据
- 转换Office文件时：Windows会启动PowerPoint或Word的后台实例，完成后退出，已打开同名文件时可能冲突，先关闭；LibreOffice以无界面方式启动，用临时配置目录，结束后清理
- 脚本只写课程文件夹和数据目录，不修改`原始资料/`里的原件；它读取的是`课程资料/`里的副本

## 常见安装问题

| 现象 | 处理 |
|---|---|
| `缺少pywin32`（Windows） | 运行上方安装命令；装完仍报错时确认用的是同一个Python。macOS、Linux不需要pywin32，也不会看到这条 |
| 转换PPT或DOC时COM报错 | Windows上脚本会自动退到LibreOffice；两者都没有时按提示安装其一。关闭已打开的同名文件；不能用WPS代替Office |
| `没有找到LibreOffice` | macOS、Linux或没有Office的Windows上没装LibreOffice。按上方“LibreOffice”一节安装；也可以让用户手动另存为PDF再交给脚本 |
| LibreOffice转换超时或没生成PDF | 首次启动慢，或文件损坏、有密码。再试一次；仍失败就用Office或手动另存为PDF |
| OXR返回415 | 文件类型不被接受。脚本已先把Office文件转成PDF；仍出现时检查文件是否损坏 |
| OXR超时、5xx或429 | 免费接口不稳定，同一时间只处理一个任务。脚本遇到429会等30秒再试；仍失败时稍后重新运行，只会重试失败的块 |
| Paddle报401或403 | key无效，或今日免费额度用完。检查`paddle.api_key`；额度用完等次日0点后重新运行 |
| 用`--engine paddle`提示需要key | 到 https://aistudio.baidu.com/paddleocr 申请，填进`config.json`的`paddle.api_key` |
| 摘要里图片全是“未转写” | 还没配置VLM，或配置的服务不支持图片输入；填好`config.json`后重新运行 |
| 图片转写部分失败 | 重新运行即可，只重试失败的图；持续失败检查`vlm.model`是否支持图片输入 |
| 路径里有空格或中文 | 脚本支持；讲义里的图片路径中资料名的空格会换成下划线 |
