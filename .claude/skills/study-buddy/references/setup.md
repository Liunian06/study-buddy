# 运行环境与安装

本skill的脚本依赖Windows和已安装的Office：`ocr_cli.py`转换PPT和DOC、`build_pptx.py`生成pptx、`convert_submission.py`导出PDF，都通过Office的COM接口完成。没有Office的电脑上，这几步无法运行。

## 环境要求

- Windows，已安装PowerPoint和Word（本机验证版本为16.0）
- Python 3.10以上（本机验证版本为3.14）
- OCR引擎默认用Paddle（PaddleOCR-VL-1.6），OXR作为备选。用`--engine`切换，或在`config.json`里写`engine`：

| 引擎 | 要不要key | 单块上限 | 特点 |
|---|---|---|---|
| `oxr` | 不要（填`dummy`） | 20页、10MB（用户给出的免费版上限） | 备选。SoMark免费接口，同步，同一时间只处理一个任务，实测约5秒一页，不保证稳定，没有隐私声明；没有书页码和页眉页脚信息 |
| `paddle` | 要，到 https://aistudio.baidu.com/paddleocr 申请 | 100页（免费版，用户给出）、本地上传50MB（官方文档） | 默认。百度PaddleOCR-VL-1.6，异步提交加轮询，整本443页教材实测约15分钟、0个失败块。免费额度每天2万页，0点刷新（用户给出，未核实） |

- Paddle适合整本教材这类大文件。额度用完时（HTTP 403），脚本停止提交后面的块，已完成的块保留，次日重新运行即可补齐
- 两个引擎可以混用：同一份资料先用一个、失败的块换另一个重跑，草稿格式一样
- Paddle读取的是它返回的块级JSON（`parsing_res_list`），不是Markdown。好处：页眉、页码、页脚不混进正文；页码块直接记成书页码，草稿标记里写成`p.061 | 书p.46`；公式编号并进公式（`\tag{4.8}`）；图和图注挨在一起。OXR没有这些信息
- VLM服务（可选）：任何支持图片输入的、兼容OpenAI Chat Completions格式的服务。没有配置时，图片会标为“未转写”，其余照常

## 安装

```bash
python -m pip install pymupdf==2.0 pywin32==312 requests python-docx==1.2.0 markdown-it-py pillow
```

`python-pptx`不是必需依赖，只在检查pptx内容时用到。

## 验证

```bash
python -c "import pymupdf, win32com.client, requests, docx, markdown_it, PIL; print('ok')"
```

预期输出：`ok`

## config.json

`config.json`在数据目录（`study-buddy-data/`）根部，模板见[config.template.json](../assets/templates/config.template.json)。有两种方式生成，生成后都由用户自己填key：

- **推荐**：首次使用时运行`python scripts/ocr_cli.py --init-config`。只生成配置文件就退出，已存在则不覆盖，这样用户可以先填好key再整理资料，流程见[first-time-setup.md](../workflows/first-time-setup.md)
- 兜底：没做初始化就直接运行`ocr_cli.py`处理资料时，脚本会自动生成`config.json`，**不会中断**，而是按空配置继续处理：OCR退用OXR，图片标为“未转写”，摘要的`warnings`里提示。用户填好key后重新运行即可补转写

| 字段 | 含义 |
|---|---|
| `engine` | 默认引擎，`paddle`或`oxr`，模板里是`paddle`；命令行的`--engine`优先 |
| `paddle.api_key` | 用户自己申请的key。模板里是示例值`paddle-token-123`；不填或保留示例值时，默认引擎退用OXR |
| `paddle.model` | 默认`PaddleOCR-VL-1.6`，能输出带公式和表格的Markdown。`PP-OCRv6`只返回带框的图，没有文字，不适用 |
| `paddle.max_pages_per_chunk`、`paddle.max_mb_per_chunk` | 单块页数和大小上限，默认100页、45MB |
| `oxr.base_url` | OXR接口地址，默认免费接口；改用正式版时换成正式地址 |
| `oxr.api_key` | 免费接口填`dummy`；正式版填自己的key |
| `oxr.max_pages_per_chunk`、`oxr.max_mb_per_chunk` | 单块页数和大小上限，默认20页、10MB |
| `vlm.base_url` | 服务地址，模板里是示例`https://api.openai.com/v1`，脚本会请求`<base_url>/chat/completions`；用其他兼容服务商时换成它的地址 |
| `vlm.model` | 模型名，必须支持图片输入；模板里是示例`gpt-4o` |
| `vlm.api_key` | 用户自己填写的key。模板里是示例值`sk-123`，保留示例值时视为没填，VLM不启用 |
| `vlm.min_image_pixels` | 图片最短边小于这个值时视为图标，跳过转写，默认64 |

- 图片转写用的固定指令在[vlm-image.md](../assets/prompts/vlm-image.md)，由脚本直接读取；想调整转写风格时改这个文件
- agent不会询问key，也不读取或复述它（包括`vlm.api_key`和`paddle.api_key`）；脚本的日志和输出里不出现key
- `vlm.api_key`和`paddle.api_key`以明文保存。工作目录如果处在网盘同步的文件夹里，key会随文件一起同步；用户要避免时，把数据目录放到不同步的位置，或用`--data-dir`指向别处
- 课件会上传到OXR，图片会发给用户在`vlm`里配置的服务商。用户已同意，学习资料按无隐私要求处理

## 脚本用法

| 脚本 | 用途 | 命令 |
|---|---|---|
| `scripts/ocr_cli.py` | 把一个或多个文件转成带页码和图片转写的Markdown草稿 | `python scripts/ocr_cli.py <文件...> --course-dir <课程文件夹> --types material` |
| `scripts/build_pptx.py` | 从pre底稿生成pptx | `python scripts/build_pptx.py <底稿.md>` |
| `scripts/convert_submission.py` | 把Markdown转成docx或pdf | `python scripts/convert_submission.py <稿.md> --format docx` |

### ocr_cli.py的参数

| 参数 | 含义 |
|---|---|
| `files` | 一个或多个文件（ppt、pptx、doc、docx、pdf、ipynb），按传入顺序编号 |
| `--course-dir` | 课程文件夹，输出写在这里 |
| `--data-dir` | 放`config.json`的数据目录，默认是课程文件夹的上一级；与`--init-config`一起用时默认是当前目录下的`study-buddy-data/` |
| `--init-config` | 只生成`config.json`（已存在不覆盖）然后退出，不需要传资料文件和`--course-dir` |
| `--types` | `material`课件/参考资料、`textbook`教材、`past-exam`往年题、`assignment`作业要求；逗号分隔与文件一一对应，或只写一个应用到全部 |
| `--pages` | 只处理指定页（1起，如`10,17,20-22`），用于测试 |
| `--engine` | `paddle`或`oxr`，默认取`config.json`的`engine`，没写则用`paddle`。默认引擎是paddle但没填key时，自动退用oxr并在摘要的`warnings`里提示；显式`--engine paddle`又没有key则报错退出 |
| `--keep-temp` | 保留临时PDF和切块 |

- 退出码：0全部成功；1部分失败（某块解析失败、图片下载失败或图片转写失败），重新运行只会重试失败的部分；2参数或环境错误
- 已完成的资料会自动跳过；配置好VLM后重新运行，只补转写没转写的图
- 输出：stdout是JSON摘要，每份资料的状态、页数、图片数、已转写和未转写数、草稿路径、失败页码；stderr是进度日志
- 资料多、耗时长时放到后台运行

### build_pptx.py和convert_submission.py

- `build_pptx.py`读取pre底稿里“### 第N页 标题”下的“要点”“图表”“讲稿”“备注”；讲稿和备注写进备注栏；用默认主题，不套模板；输出已存在时另存为`-2`
- `convert_submission.py`保留标题层级、粗斜体、列表、表格、代码块、引用块和图片；公式（`$$...$$`）按文本原样保留；不添加封面、页眉页脚、目录。pdf先生成docx，再由Word导出

### 版本更新检测

`ocr_cli.py`、`build_pptx.py`、`convert_submission.py`运行结束时，会调用`check_update.py`查询GitHub上最新的Release；有新版本时在stderr提示更新命令`npx skills update study-buddy -g`。请求超时2秒，结果缓存（成功24小时、失败6小时）；断网、被限流等任何查询失败都不提示，不影响脚本结果。设置环境变量`STUDY_BUDDY_NO_UPDATE_CHECK=1`可关闭。

## 网络与权限

- `ocr_cli.py`访问所选的OCR接口和（配置后）VLM服务；引擎返回的图片链接会被下载到本地
- `check_update.py`访问`api.github.com`查询最新版本，只读，不上传任何数据
- 转换Office文件时会启动PowerPoint或Word的后台实例，完成后退出；已打开同名文件时可能冲突，先关闭
- 脚本只写课程文件夹和数据目录，不修改`原始资料/`里的原件；它读取的是`课程资料/`里的副本

## 常见安装问题

| 现象 | 处理 |
|---|---|
| `缺少pywin32` | 运行上方安装命令；装完仍报错时确认用的是同一个Python |
| 转换PPT或DOC时COM报错 | 确认已安装Office；关闭已打开的同名文件；不能用WPS代替 |
| OXR返回415 | 文件类型不被接受。脚本已先把Office文件转成PDF；仍出现时检查文件是否损坏 |
| OXR超时、5xx或429 | 免费接口不稳定，同一时间只处理一个任务。脚本遇到429会等30秒再试；仍失败时稍后重新运行，只会重试失败的块 |
| Paddle报401或403 | key无效，或今日免费额度用完。检查`paddle.api_key`；额度用完等次日0点后重新运行 |
| 用`--engine paddle`提示需要key | 到 https://aistudio.baidu.com/paddleocr 申请，填进`config.json`的`paddle.api_key` |
| 摘要里图片全是“未转写” | 还没配置VLM，或配置的服务不支持图片输入；填好`config.json`后重新运行 |
| 图片转写部分失败 | 重新运行即可，只重试失败的图；持续失败检查`vlm.model`是否支持图片输入 |
| 路径里有空格或中文 | 脚本支持；讲义里的图片路径中资料名的空格会换成下划线 |
