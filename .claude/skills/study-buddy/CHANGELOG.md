# Changelog

study-buddy的版本记录。版本号规则：主版本.次版本.修订号。

## [2.0.0] - 2026-10-31

### 变更（不兼容）

- 数据目录的文件名和文件夹名改为中文，旧版本生成的课程数据需要手动改名后才能被识别。`config.json`和`manifest.json`不变
  - `notes/`改为`讲义/`，讲义文件名`chapterNN_`改为`第NN章_`
  - `index.md`改为`讲义目录.md`，`mistakes.md`改为`错题本.md`，`semester.md`改为`学期规划.md`
  - `drafts/`改为`草稿/`，`media/`改为`媒体/`，`.cache/`改为`.缓存/`，`.work/`改为`.临时/`
  - `presentation/`改为`演示文稿/`，`paper/`改为`论文/`，`assignment/`改为`作业/`
  - `_cram-plan.md`改为`_速通计划.md`，`_assets/`改为`_图表/`，`.group-plan.md`改为`.分工表.md`
- `ocr_cli.py`同步使用新的中文目录名，并读取`课程资料/`里的副本

### 新增

- 每门课新增`原始资料/`和`课程资料/`两个目录：原件放`原始资料/`，只读不改；资料整理第1步建课时建好目录，第2、3步把确认后的文件复制到`课程资料/`并改成`<类型>_<名称>.<扩展名>`

## [1.0.1] - 2026-10-30

### 变更

- `config.template.json`补充示例值：`vlm`的`base_url`为`https://api.openai.com/v1`、`model`为`gpt-4o`，`vlm.api_key`为`sk-123`，`paddle.api_key`为`paddle-token-123`
- `ocr_cli.py`把示例key视为没填：VLM不启用，默认引擎仍会在没有真实paddle key时退用OXR

## [1.0.0] - 2026-10-29

### 新增

- 初始版本，包含工作流：资料整理、学期规划、跟课学习、期末速通、Presentation、课程论文、完成作业
- 版本更新检测：`ocr_cli.py`、`build_pptx.py`、`convert_submission.py`运行结束时自动检查是否有新版本，有则在stderr提示更新命令；查询失败时不提示。设置环境变量`STUDY_BUDDY_NO_UPDATE_CHECK=1`可关闭
