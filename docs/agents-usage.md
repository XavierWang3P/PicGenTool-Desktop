# PicGenTool Agents 使用说明

面向需要自动生成活动照片 Word 文档的 AI Agent、脚本和自动化工作流。PicGenTool 将标题、地点、日期和 1～30 张照片写入现有模板，生成 `.docx` 文件。

## 1. 最小使用流程

```text
检查环境 → 校验请求 → 生成文档 → 检查 JSON → 验证输出 → 返回文件路径
```

1. 安装并启动依赖完整的 PicGenTool，或进入项目目录并激活 Python 虚拟环境。
2. 调用 `--capabilities --json`，确认模板齐全。
3. 收集用户明确提供的标题、地点、日期、图片和输出路径。
4. 必要时先 `--validate`，再 `--generate`。
5. 仅在 `ok=true` 时报告成功；可将文档路径返回给用户。

不要模拟鼠标点击 GUI、打开文件选择框或解析窗口文字。

## 2. 运行方式

以下示例中的 `$PGT` 是程序入口：

- PowerShell / CMD 打包版：`$PGT = "D:\Apps\PGT\PGT.exe"`
- PowerShell 开发版：`$PGT = "D:\Projects\PicGenTool\.venv\Scripts\python.exe"`，命令后加 `main.py`
- Bash 开发版：`python main.py ...`

所有入口都使用 UTF-8 JSON。PowerShell 中建议显式设置：

```powershell
$env:PYTHONUTF8 = "1"
$OutputEncoding = [System.Text.UTF8Encoding]::new()
```

打包版自带运行环境；开发版需 Python 3.11+ 和项目 `requirements.txt`。

## 3. 环境检查

```powershell
& $PGT --capabilities --json
```

成功响应示例：

```json
{
  "ok": true,
  "schema_version": 1,
  "app_version": "3.1.1",
  "max_images": 30,
  "formats": [".bmp", ".gif", ".heic", ".heif", ".jpeg", ".jpg", ".png"],
  "templates": {"available": 30, "missing": []}
}
```

若 `templates.missing` 非空，停止生成并修复模板。获取诊断信息：

```powershell
& $PGT --diagnose --json
```

## 4. 单任务：推荐从 stdin 调用

请求示例：

```json
{
  "schema_version": 1,
  "title": "春季运动会",
  "location": "学校操场",
  "activity_date": "2026-11-27",
  "images": [
    "C:/photos/01.jpg",
    "C:/photos/02.heic"
  ],
  "output": "D:/out/2026.11.27-春季运动会照片.docx",
  "options": {
    "if_exists": "fail"
  }
}
```

Windows 打包版 PowerShell：

```powershell
$json = Get-Content .\request.json -Raw -Encoding UTF8
$json | & $PGT --generate - --json
```

开发版 PowerShell：

```powershell
$json | & $PGT main.py --generate - --json
```

Bash：

```bash
python main.py --generate - --json < request.json
```

也可直接传请求文件：

```powershell
& $PGT --generate .\request.json --json
```

### 成功响应

```json
{
  "ok": true,
  "job_id": "ef01b3e310bc4124a2ad31a72bb8a86f",
  "output": "D:\\out\\2026.11.27-春季运动会照片.docx",
  "image_count": 2,
  "duplicates_skipped": 0,
  "warnings": [],
  "elapsed_ms": 847.936,
  "app_version": "3.1.1"
}
```

`job_id` 仅用于关联本次结果；不是可远程轮询的持久任务 ID。

### 先校验再生成

```powershell
$json | & $PGT --validate - --json
$json | & $PGT --generate - --json
```

`--validate` 检查字段、文件、格式、去重、模板和输出策略，不创建文档。校验通过后仍可能因文件占用、磁盘错误、并发状态变化而生成失败，必须检查最终响应。

## 5. 字段约束

| 字段 | 要求 |
| --- | --- |
| `schema_version` | 必须为整数 `1` |
| `title` | 非空字符串，作为活动标题 |
| `location` | 非空字符串，作为活动地点 |
| `activity_date` | 有效 `YYYY-MM-DD` 日期 |
| `images` | 1～30 个现有图片路径；按内容 SHA-256 去重，保留首次出现顺序 |
| `output` | 已存在目录下的 `.docx` 路径；程序不会自动创建目录 |
| `options` | 可选；不接受未知字段 |

支持 JPG/JPEG、PNG、GIF、BMP、HEIC、HEIF。图片按模板比例居中裁剪并校正 EXIF 方向；不是“仅插入原图”。不要在未经确认的情况下选择用户图片或覆盖文档。

### 输出策略

| `options.if_exists` | 用途 |
| --- | --- |
| `fail`（默认） | 输出已存在即失败；同一进程内同路径并发也只会成功一个 |
| `skip` | 保留现有文件并返回 `ok=true, skipped=true`，适合幂等重试 |
| `overwrite` | 原子替换已有文档；仅在用户明确授权时使用 |

`skip` 不验证现有文件是否由本次相同请求生成。若内容变化，应使用新的输出路径，或获得明确授权后覆盖。

无头模式不支持 `open=true`、`use_settings=true`；`on_image_error` 只能为 `fail`。

## 6. 批处理

创建 `jobs.jsonl`，每行一个完整请求：

```jsonl
# 每个活动使用不同的输出路径
{"schema_version":1,"title":"春游","location":"公园","activity_date":"2026-11-27","images":["C:/photos/a.jpg"],"output":"D:/out/spring.docx"}
{"schema_version":1,"title":"读书会","location":"图书馆","activity_date":"2026-11-28","images":["C:/photos/b.jpg"],"output":"D:/out/reading.docx"}
```

```powershell
& $PGT --batch .\jobs.jsonl --workers 2 --on-error continue --timeout 60 --json
```

每行输出一个带 `line` 的结果：

```json
{"line":2,"ok":true,"job_id":"...","output":"D:\\out\\spring.docx","image_count":1}
```

- `--workers 1..4`：默认 1；结果按输入顺序输出。
- `--on-error continue`：默认继续后续任务。
- `--on-error stop`：发现失败后停止提交新任务；已运行的并发任务仍会完成并输出结果。
- `--timeout`：单任务秒数，从该任务实际启动时计算；使用协作取消，不是强制终止第三方库。
- 单行最大 1 MiB；忽略空行和以 `#` 开头的行。
- Ctrl-C 取消正在运行的任务；等待安全的原子步骤结束。
- 每个任务独立生成一个文档，不生成合并报告，不删除已成功文档。

## 7. 常驻本地服务

多个连续任务可复用一个进程，避免重复启动：

```powershell
& $PGT --serve --json
```

stdin 每行一个 NDJSON 请求，外层只含 `id` 和 `request`：

```json
{"id":"job-1","request":{"schema_version":1,"title":"春游","location":"公园","activity_date":"2026-11-27","images":["C:/photos/a.jpg"],"output":"D:/out/a.docx"}}
```

响应：

```json
{"id":"job-1","result":{"ok":true,"job_id":"...","output":"D:\\out\\a.docx"}}
```

服务按输入顺序同步执行；坏行返回错误后可继续下一行；关闭 stdin（EOF）正常退出。不要并发写入同一 stdin 流，由 Agent 串行管理请求与响应。该模式没有异步任务查询、历史任务存储、网络端口或 MCP。

## 8. Python Agent 调用示例

```python
import json
import subprocess
import sys
from pathlib import Path
from zipfile import ZipFile

ROOT = Path(r"D:\Projects\PicGenTool")
request = {
    "schema_version": 1,
    "title": "春季运动会",
    "location": "学校操场",
    "activity_date": "2026-11-27",
    "images": [r"C:\photos\01.jpg"],
    "output": r"D:\out\result.docx",
    "options": {"if_exists": "fail"},
}

completed = subprocess.run(
    [sys.executable, str(ROOT / "main.py"), "--generate", "-", "--json"],
    input=json.dumps(request, ensure_ascii=False),
    text=True,
    encoding="utf-8",
    capture_output=True,
    cwd=ROOT,
    timeout=90,
    check=False,
)

if not completed.stdout.strip():
    raise RuntimeError(f"工具未返回 JSON，退出码={completed.returncode}")

result = json.loads(completed.stdout)
if not result.get("ok"):
    error = result["error"]
    raise RuntimeError(f"{error['code']}: {error['message']}")

output = Path(result["output"])
with ZipFile(output) as archive:
    assert "word/document.xml" in archive.namelist()
print(output)
```

使用 `--result-file result.json` 可在无法捕获 stdout 的启动器中保存单任务、能力、Schema 或诊断结果。结果文件必须不存在，且不得等于 DOCX 输出路径；批处理和常驻服务不支持此选项。

## 9. 错误处理

始终检查 `ok`，不能只依赖退出码；用 `error.code` 分支，不解析中文提示。

| 错误码 | 退出码 | 建议动作 |
| --- | ---: | --- |
| `INVALID_REQUEST` | 2 | 修正字段、日期、路径或选项 |
| `IMAGE_NOT_FOUND` | 2 | 重新确认图片路径 |
| `UNSUPPORTED_FORMAT` | 2 | 移除不支持文件或转换格式 |
| `IMAGE_READ_FAILED` | 2 | 检查损坏、权限或解码问题，不重复盲目重试 |
| `TOO_MANY_IMAGES` | 2 | 将图片减少到 30 张以内 |
| `TEMPLATE_MISSING` | 3 | 修复模板或运行 `--diagnose --json` |
| `OUTPUT_EXISTS` | 4 | 换输出名，或明确选择 `skip` / `overwrite` |
| `OUTPUT_NOT_WRITABLE` | 4 | 检查目录、权限、磁盘和 Word 文件占用 |
| `CANCELLED` | 5 | 视为用户取消或超时，不宣称成功 |
| `INTERNAL_ERROR` | 1 | 保留结构化错误，运行诊断并检查日志 |

批处理退出码为已执行任务中的最大错误码；`skipped=true` 仍属于成功响应。stdout 无内容或 JSON 无法解析时，报告启动器/环境问题，不得推断文档已生成。

## 10. 安全与验收清单

- [ ] 先执行 `--capabilities --json`，模板无缺失。
- [ ] 图片路径、输出路径和用户授权均明确。
- [ ] 默认 `if_exists=fail`；不自动覆盖、不自动打开 Word。
- [ ] 使用 UTF-8 stdin，解析 stdout JSON。
- [ ] 成功后确认 `output` 存在，必要时检查 DOCX 的 `word/document.xml`。
- [ ] 不将 stdin 服务转发为未受保护的 HTTP/TCP 接口。
- [ ] 无头失败不污染 GUI 偏好；错误响应中的路径属于用户请求，不应上传到第三方。
- [ ] 重试前检查目标文件；幂等重试可用 `skip`，内容变化不要误用 `skip`。

协议详情见 `docs/agent-integration.md`；仓库内的简短自动化约束见根目录 `AGENTS.md`。
