# Agent 自动化 P0/P3 设计

## 目标

为 PicGenTool 提供不依赖 GUI 的稳定自动化入口：单任务、请求校验、能力发现和 JSON Schema 属于 P0；JSONL 批处理、常驻 stdin/stdout 服务、有界队列与失败策略属于 P3。GUI 和现有 Explorer 菜单行为保持兼容。

## 范围与约束

- 不包含 IPC v2、MCP Server（P1/P2）。
- 所有自动化入口共用一份版本化请求、结果、错误码和 Schema。
- 无头路径不导入 PyQt，不创建窗口，不显示对话框，不读写 QSettings，不自动打开 Word。
- 输出必须显式指定且为 `.docx`；默认拒绝覆盖，`if_exists=overwrite` 才允许替换，`skip` 用于幂等重试。
- 缺失、不支持或不可读输入逐文件报告；SHA-256 内容去重，重复只记数量。
- stdout 仅输出最终 JSON（批处理/常驻模式为逐行 JSON）；日志写 stderr。
- 复用现有原子保存、协作取消和 docxtpl 单线程约束；临时图片总是清理。
- 默认单任务执行。`--workers 1..4` 控制批处理并发；图片并发仍由现有 `MAX_IMAGE_WORKERS` 控制。

## 请求契约

```json
{
  "schema_version": 1,
  "title": "春季运动会",
  "location": "学校操场",
  "activity_date": "2026-11-27",
  "images": ["C:/photos/01.jpg"],
  "output": "D:/out/result.docx",
  "options": {"if_exists": "fail", "on_image_error": "fail"}
}
```

`options.open`、`options.use_settings` 即使提供也不生效，无头模式不会启动 GUI 或 QSettings。请求只接受明确字段，未知字段、错误类型、无效日期和空字符串返回 `INVALID_REQUEST`。`--schema` 输出 Draft 2020-12 JSON Schema。

成功结果包含 `ok/job_id/output/image_count/duplicates_skipped/warnings/elapsed_ms/app_version`。失败包含稳定错误码、消息和可选详情。错误码：`INVALID_REQUEST`、`IMAGE_NOT_FOUND`、`UNSUPPORTED_FORMAT`、`IMAGE_READ_FAILED`、`TOO_MANY_IMAGES`、`TEMPLATE_MISSING`、`OUTPUT_EXISTS`、`OUTPUT_NOT_WRITABLE`、`CANCELLED`、`INTERNAL_ERROR`。退出码依次为通用参数错误 2、资源错误 3、输出错误 4、取消 5、内部错误 1。

## 组件

- `models/agent_contract.py`：冻结的请求、结果、错误与 Schema 常量；JSON 可序列化。
- `services/generation_service.py`：从 `GenerationWorker` 抽出流水线编排，统一清理，不负责打开文件。
- `services/agent_api.py`：输入校验、输出策略、去重、模板检查、任务 ID、异常映射。
- `services/agent_runner.py`：批量读取、有界 `ThreadPoolExecutor`、`continue/stop`、逐任务超时和 Ctrl-C 取消。
- `agent_cli.py`：`--generate/--validate/--batch/--capabilities/--schema/--diagnose --json`，以及 `--serve` 的 NDJSON stdin/stdout 循环。
- `main.py`：在 PyQt 导入前分派自动化命令；其余参数继续走原 GUI 启动。
- `ui/main_window.py`：worker 变为 `GenerationService` 的信号适配器，保留现有弹窗与偏好行为。

`--validate` 检查请求、输入文件、去重、模板和输出冲突，但不创建输出。`--capabilities` 报告格式、模板缺失和入口能力。`--generate -` 从 stdin 读取一个 JSON 请求。`--batch jobs.jsonl` 每行一个请求，允许空行和 `#` 注释，逐行输出结果，带任务索引。`--serve` 每行接受 `{ "id": 任意JSON值, "request": {...} }`，同步返回带同 id 的结果；Ctrl-C/EOF 安全结束。批量和常驻入口不产生最终聚合文档，也不删除成功文档；`stop` 在首个失败后停止提交新任务，已运行任务正常完成。

## 备选方案

- 只增加 CLI 参数、直接调用现有服务：改动少，但校验、错误码和批处理会散落，GUI 逻辑仍不可复用，未采用。
- 先做 HTTP 服务：可跨进程调用，但增加端口、认证和依赖，超出 P0/P3 本地自动化需求，未采用。
- 采用共用服务与标准流 JSON：无额外依赖、可测试、冻结版 console EXE 可直接使用，采用。

## 验证

覆盖成功生成、逐文件错误、模板缺失、输出冲突、去重、无 PyQt 导入、真实子进程 stdout、JSONL 失败继续/停止、并发上限、常驻多请求与中断。完整运行 pytest、compileall、模板深检、真实 CLI 和 PyInstaller 打包烟测（环境允许时）。
