# 疑难问题处理

用于遇到按现有细则无法解决的错误。目标是把问题整理清楚，先尝试自行解决，解决不了时给用户一份可以直接转交开发者的问题描述。环境和依赖问题先看[setup.md](setup.md)的常见安装问题。

## 处理顺序

1. 梳理现象：发生了什么、期望是什么、在哪个工作流的哪个环节出现、相关文件路径、完整报错文本
2. 复现：写出最少的复现条件，确认问题能稳定出现；无法复现时记录出现的条件
3. 尝试方案：按可能性从高到低列出原因和对应方案，逐个尝试并记录结果
4. 仍未解决：停止尝试，按下方格式整理问题描述交给用户，并按[community.md](community.md)给出反馈群二维码（过期后是作者的微信二维码）和Issues链接，请用户带着这份描述联系开发者

不要为了绕过错误而修改用户的数据，也不要读取或复述`config.json`里的key。

## 常见问题

| 现象 | 可能原因 | 处理 |
|---|---|---|
| 讲义里某页写着“资料缺失” | OCR解析该页失败 | 重新运行`ocr_cli.py`，只会重试失败的块；仍失败时告诉用户该页内容缺失，请用户手动补充 |
| 讲义里图下写着“未转写”，且图是重点 | VLM未配置或调用失败 | 填好`config.json`后重新运行；不要凭图片周围的文字编造图意 |
| 处理到一半被中断（关了窗口、断电、后台任务被取消） | 每个块解析完都会保存进度 | 重新运行同样的命令即可，只会重做中断时正在进行的那一块，已完成的块和已下载的图片不重做，见[setup.md](setup.md) |
| 摘要里`images_broken`或`image_download.failed`不为0，图片没下载下来 | OCR引擎返回的图片链接是临时的，过期了或网络不通 | 脚本已经自动重试并对出问题的页补过OCR；仍失败时检查网络、代理、防火墙，再重新运行。`image_download.reocr_stopped`为真说明连续多页补OCR也没用，多半是网络问题，不要反复重跑，会消耗OCR额度，见[setup.md](setup.md) |
| 课程信息和总览对不上 | 总览是汇总，单课信息可能后来改过 | 以该课的`课程信息.md`为准，重新生成`学期规划.md`里这门课的行，见[data-layout.md](data-layout.md)的“两处信息的同步” |
| 课程文件夹编号不对 | `学期规划.md`丢失或学期号错 | 问用户当前学期，按[data-layout.md](data-layout.md)修正，已有文件夹不改名 |
| 摘要警告“默认引擎是paddle，但key还没填”，处理比较慢，或页码只有范围、没有表格公式 | 没填Paddle key，从优先级链的下一档开始（有oxr key用oxr商业版，再是mineru token，都没有就用OXR免费接口）；免费OXR同一时间只处理一个任务，约5秒一页；最后一档MinerU轻量接口没有分页和表格公式识别 | 申请key填进`config.json`的`paddle.api_key`，或填`oxr.api_key`、`mineru.api_key`，见[setup.md](setup.md) |
| 摘要警告“xx失败……改用xx”，处理很慢或OXR频繁报“已有任务在进行” | 链上某一档key无效、额度用尽或排队超时，退到了下一档，资料因此上传到了另一家供应商；免费OXR同一时间只处理一个任务，约5秒一页 | 检查摘要里写的那一档的key是否有效，或稍后重跑；也可以填Paddle key，见[setup.md](setup.md) |
| 没有Office的电脑上转不了PPT | Windows上脚本会自动退到LibreOffice，两者都没有才失败；macOS、Linux直接用LibreOffice | 按`setup.md`装LibreOffice；仍不行就请用户手动另存为PDF后再交给脚本 |
| `compile_tex.py`返回`no_engine` | 没装xelatex | 按返回的`message`装：Windows `winget install MiKTeX.MiKTeX`，macOS `brew install --cask mactex-no-gui`，Linux `sudo apt install texlive-xetex texlive-lang-chinese`；装完重开终端，tex源文件已保存，装好后重新编译，见[setup.md](setup.md) |
| MiKTeX首次编译很久没结果，或超时 | 在联网下载缺少的宏包 | 确认网络可用，再编译一次；仍超时就用`--keep`看临时目录里的日志 |
| 提纲PDF里有的字显示成空白 | 字体里没有这个字符，`warnings`里会列出 | 换成常见字符；数学环境里不要放汉字，用`\text{}` |
| 开卷提纲速查表页码显示“??” | 页码引用需要多编译一遍才稳定 | 再运行一次`compile_tex.py` |
| tex一直编译不过 | 特殊字符没转义、括号不配对、用了模板外的宏 | 按`errors`的行号改，先只改第一个错误；5轮仍不行，把最后的错误交给用户，见[exam-notes.md](exam-notes.md) |
| 生成的pptx没有样式 | 脚本不套模板，只用默认主题 | 属预期；用户在PowerPoint里套学校模板 |
| 转出的docx不符合学校格式 | 脚本不添加封面、页眉页脚、目录 | 属预期；按作业要求由用户在Word里补，或告知具体要求后调整稿件 |
| 翻译后字数不符 | 中英文计数方式不同 | 按目标语言重新计数，调整后再核对 |
| 引用格式拿不准 | 老师有指定版本 | 以老师给的模板为准，见[course-paper.md](course-paper.md) |

## 问题描述格式

```markdown
## 问题
一句话描述现象

## 环境
- 客户端与版本：
- 操作系统：
- 操作系统（Windows、macOS或Linux）、Office或LibreOffice版本：
- study-buddy版本：

## 复现步骤
1.
2.

## 期望结果 / 实际结果

## 完整报错
（ocr_cli.py的JSON摘要和stderr日志；不要包含config.json的内容）

## 已尝试的方案与结果
```
