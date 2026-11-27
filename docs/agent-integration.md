# Agent 自动化接口

PicGenTool 支持不启动 GUI 的本地自动化。请求、图片和文档不会上传到网络；调用方必须有权读取指定图片和写入指定目录。

## 请求

```json
{
  "schema_version": 1,
  "title": "春季运动会",
  "location": "学校操场",
  "activity_date": "2026-11-27",
  "images": ["C:/photos/01.jpg", "C:/photos/02.heic"],
  "output": "D:/out/result.docx",
  "options": {"if_exists": "fail"}
}
```

字段均必填（`options` 除外），未知字段会报错。图片最多 30 张，按内容 SHA-256 去重，顺序保持不变。输出必须为已存在目录下的 `.docx`；程序不会自动创建目录。

`options.if_exists`：

- `fail`：默认，存在即失败，不覆盖。
- `skip`：保留现有文件，返回成功及 `skipped: true`，用于幂等重试。
- `overwrite`：通过同目录临时文件原子替换；要求调用方明确授权覆盖。

`open=true`、`use_settings=true` 和 `on_image_error` 的非 `fail` 值不受支持，会返回参数错误。无头模式不读写 QSettings，不打开 Word，不显示对话框。

## 单任务

开发环境：

```bash
python main.py --generate request.json --json
python main.py --generate - --json < request.json
python main.py --validate request.json --json
python main.py --capabilities --json
python main.py --schema --json
python main.py --diagnose --json
```

Windows 打包版将 `python main.py` 换成 `PGT.exe`。`--schema` 返回 Draft 2020-12 JSON Schema，可供 Agent 构造工具参数。`--capabilities` 返回当前格式和模板可用性。

成功返回 `ok/job_id/output/image_count/duplicates_skipped/warnings/elapsed_ms/app_version`。失败返回稳定 `error.code/message/details`。stdout 仅为 JSON，日志为 stderr；当前无头入口不写文件日志。所有标准流强制使用 UTF-8。

| 错误码 | 退出码 |
| --- | ---: |
| `INVALID_REQUEST`, `IMAGE_NOT_FOUND`, `UNSUPPORTED_FORMAT`, `IMAGE_READ_FAILED`, `TOO_MANY_IMAGES` | 2 |
| `TEMPLATE_MISSING` | 3 |
| `OUTPUT_EXISTS`, `OUTPUT_NOT_WRITABLE` | 4 |
| `CANCELLED` | 5 |
| `INTERNAL_ERROR` | 1 |

## 批处理（P3）

`jobs.jsonl` 每行一个完整请求，空行和以 `#` 开头的行被忽略：

```bash
python main.py --batch jobs.jsonl --workers 2 --on-error continue --timeout 60 --json
```

- `--workers` 为 1–4，默认 1；结果按输入顺序输出，并带原始 `line`。
- `--on-error continue` 默认继续；`stop` 在发现失败后停止提交新任务，已运行的并发任务正常结束。
- `--timeout` 是单个任务从启动开始的秒数，采用协作取消，不能强制中断正在进行的第三方库调用。
- 每行独立成功或失败，不生成聚合 Word，不删除成功文档。进程退出码取所有已执行结果中最严重的错误码。
- 输入文件和每行均限制为 1 MiB；Ctrl-C 会取消所有正在运行的任务。

## 常驻本地服务（P3）

```bash
python main.py --serve --json
```

向 stdin 逐行写 NDJSON，每个请求最多 1 MiB：

```json
{"id":"job-1","request":{"schema_version":1,"title":"春游","location":"公园","activity_date":"2026-11-27","images":["C:/photos/a.jpg"],"output":"D:/out/a.docx"}}
```

stdout 每行返回 `{"id":"job-1","result":{...}}`。服务按输入顺序同步执行，避免重入；坏请求返回错误后可继续处理下一行，EOF 正常退出。此接口不提供远程任务查询或异步 job ID 轮询；需要异步编排时由 Agent 管理子进程。

`--result-file result.json` 可将单任务、能力、Schema 或诊断结果同时写入指定文件，适用于无法捕获 stdout 的启动器。结果文件使用独占创建，已存在即拒绝，且不得与 DOCX 输出路径相同；批处理和常驻模式不支持该选项。

## 安全与重试

- 仅监听本地标准流，不开放 TCP/HTTP 端口；不要把服务暴露到不受信任的输入。
- Agent 必须先校验路径和输出冲突，再执行写操作；默认不覆盖用户文档。
- 失败时可按错误码修正请求；重试优先使用 `if_exists=skip`。
- 当前不包含 GUI IPC v2 或 MCP Server。
