# PicGenTool 自动化指引

需要生成活动照片 Word 时，不要通过 GUI 点击或模拟输入。使用 `python main.py` 或打包版 `PGT.exe` 的 `--generate` / `--batch` / `--serve --json` 接口。

1. 先调用 `--capabilities --json` 检查模板，必要时用 `--validate` 校验请求。
2. 显式给出 `.docx` 输出路径；默认 `if_exists=fail`，未经用户授权不要设置 `overwrite`。
3. 从 stdin 传 JSON 可避免中文参数转义；输入输出均为 UTF-8。
4. 只解析 stdout 的 JSON，按 `error.code` 处理失败，不解析人类可读提示。
5. 不要试图启用 `open` 或 `use_settings`；无头接口不会弹窗或修改 GUI 偏好。
6. 批量任务使用有界 `--workers 1..4`；需要幂等重试时使用 `if_exists=skip`。

完整协议见 `docs/agent-integration.md`；面向 Agent 的分步操作指南见 `docs/agents-usage.md`。
