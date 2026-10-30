# Changelog

study-buddy的版本记录。版本号规则：主版本.次版本.修订号。

## [1.0.0] - 2026-10-29

### 新增

- 初始版本，包含工作流：资料整理、学期规划、跟课学习、期末速通、Presentation、课程论文、完成作业
- 版本更新检测：`ocr_cli.py`、`build_pptx.py`、`convert_submission.py`运行结束时自动检查是否有新版本，有则在stderr提示更新命令；查询失败时不提示。设置环境变量`STUDY_BUDDY_NO_UPDATE_CHECK=1`可关闭
