# PicGenTool Quality and Release Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 为 PicGenTool 增加 HEIC/HEIF、精确去重、可靠性和诊断能力，改善桌面交互，并建立测试、体积预算和可选代码签名组成的 Windows 发布链路。

**Architecture:** 保留现有 `ui → services → models` 分层，将图片兼容与去重集中在 `ImageService`，将模板、日志、异常、设置和资源管理做成职责单一的服务。生成线程同时启动图片准备 future 和同线程 Word 模板准备，随后仍在同一 Worker 线程渲染保存；CI 按测试、打包、体积、可选签名、发布顺序执行。

**Tech Stack:** Python 3.11/3.12、PyQt6 6.11、Pillow 12.2、pi-heif 1.4.0、docxtpl 0.20.2、python-docx 1.2、pytest、pytest-qt、PyInstaller 6.x、GitHub Actions、Windows SignTool

**Spec:** `docs/superpowers/specs/2026-11-26-quality-release-optimization-design.md`

## Global Constraints

- 保持 Python 3.11 及以上；CI 同时验证 3.11 和 3.12。
- 图片上限 30 张；输出保持 6.4 cm × 4.8 cm、240 DPI、JPEG 质量 60。
- `pi-heif==1.4.0`（解码专用发行版），通过 `register_heif_opener()` 接入现有 Pillow 流程；不升级项目固定的 Pillow 12.2，不引入 HEIF 编码依赖。
- 重复图片只按完整文件内容 SHA-256 判定，保留第一次出现的顺序。
- 不保存照片列表、输出路径、证书、密码或用户表单内容到诊断日志。
- 日志轮转为 5 MiB、3 个备份。
- Word/docxtpl 对象只在创建它的 GenerationWorker 线程准备、渲染和保存。
- 模板要求 `Template-1.docx` 至 `Template-30.docx`，占位符恰为公共文本字段和 `image1…N`。
- 保持 PyInstaller one-dir 和 Qt 软件 OpenGL；打包警戒预算为 130 MiB。
- 代码签名凭据只能由环境/CI secret 注入；签名后必须验签，临时 PFX 和导入证书必须在 `finally` 中删除。
- 当前目录不是 Git 仓库；每个任务末尾用“文件清单 + 定向测试 + progress.md 记录”代替 commit，不运行任何初始化仓库或模拟提交操作。

## Review Focus

- 扩展名正确但内容损坏的 HEIC/HEIF 必须给安全错误且不泄露父目录；Task 1 覆盖。
- 同一路径、硬拷贝、同大小不同内容和已有列表重复必须保持稳定顺序且在上限计算前去重；Task 1 覆盖。
- 模板准备或图片准备任一失败/取消都必须等待另一侧安全退出并清理临时图片，且 Word 对象不跨线程；Task 3 覆盖。
- 全局异常处理器自身失败、后台线程异常和 Qt 关闭期间异常不能递归或弹出含私人路径的消息；Task 4 覆盖。
- CI 没有证书 secret 时应成功跳过签名，有证书时任何签名/时间戳/验签失败都必须阻止发布且清理凭据；Task 9 覆盖。

---

## File Map

- `config.py`：应用版本、HEIC/HEIF 扩展名、日志轮转和打包预算常量。
- `requirements.txt`：固定 pi-heif 运行依赖。
- `services/image_service.py`：HEIF 插件注册、内容哈希与去重。
- `models/image_selection.py`：去重结果数据对象。
- `services/template_service.py`：轻量和深度模板完整性检查。
- `models/template_validation.py`：模板检查结果数据对象。
- `services/word_exporter.py`：同线程 Word 准备对象与准备后导出。
- `services/document_service.py`：构建 Word payload、准备和导出阶段接口。
- `ui/main_window.py`：流水线重叠、设置恢复、拖拽反馈、启动检查和诊断入口。
- `utils/logging_utils.py`：轮转 handler 和日志路径。
- `utils/exception_handler.py`：主线程/后台线程异常钩子和 Qt 安全提示。
- `services/explorer_integration.py`：Windows 右键菜单注册、状态和路径自愈。
- `tools/register_explorer_menu.py`：调用共享服务的命令行包装。
- `services/diagnostics.py`：隐私安全的诊断快照和文本格式化。
- `ui/diagnostics_dialog.py`：关于/诊断对话框。
- `tools/check_templates.py`：CI 可调用的深度模板检查命令。
- `tools/package_report.py`：分发体积报告和预算退出码。
- `tools/sign_release.ps1`：导入证书、签名、验签和清理。
- `pgt.spec`：HEIF 二进制收集和经验证的排除项。
- `.github/workflows/compile.yml`：测试矩阵、构建、体积、可选签名、发布依赖。
- `tests/`：新增对应单元、Qt、脚本和集成测试。
- `README.md`、`CHANGELOG.md`：用户行为、诊断、CI、签名和体积说明。

### Task 1: HEIC/HEIF 解码与精确图片去重

**Files:**
- Create: `models/image_selection.py`
- Modify: `config.py`
- Modify: `requirements.txt`
- Modify: `services/image_service.py`
- Modify: `main.py`
- Modify: `ui/main_window.py`
- Test: `tests/test_image_service.py`
- Test: `tests/test_cli_and_ipc.py`

**Interfaces:**
- Produces: `DeduplicationFailure(filename: str, reason: str)` 和 `DeduplicationResult(paths: list[Path], duplicate_count: int, failures: tuple[DeduplicationFailure, ...])`。
- Produces: `ImageService.deduplicate(files: list[str | Path], existing: list[Path] | None = None) -> DeduplicationResult`。
- Produces: `ImageService._content_digest(path: Path) -> str`，以 1 MiB 块流式 SHA-256。
- Produces: `ImageImportWorker(files: list[str], existing: list[Path])` 的 `completed = pyqtSignal(object)` 和 `finished = pyqtSignal()`；`MainWindow._queue_image_import(files: list[str]) -> None` 串行处理批次。
- Preserves: `ImageService.prepare_images(...) -> PreparedImages` 和 `extract_image_paths(arguments) -> list[str]`。

- [x] **Step 1: 写 HEIF 和去重失败测试**

在 `tests/test_image_service.py` 添加：

- `test_prepare_images_decodes_heic_and_heif_to_rgb_jpeg`
- `test_deduplicate_removes_same_path_and_identical_copy_preserving_first_order`
- `test_deduplicate_keeps_same_size_different_content`
- `test_deduplicate_compares_new_files_with_existing_selection`
- `test_deduplicate_unreadable_file_reports_safe_failure_and_keeps_readable_files`
- `test_image_import_runs_off_ui_thread_and_queues_second_batch`

在 `tests/test_cli_and_ipc.py` 扩展 CLI 支持 `.heic/.heif` 的断言。HEIF 样本由测试通过 Pillow plugin 写入，不依赖外部 fixture。

- [x] **Step 2: 运行定向测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_image_service.py tests/test_cli_and_ipc.py -q`

Expected: FAIL，原因是 pillow-heif 未安装、扩展名和 `deduplicate` 尚不存在。

- [x] **Step 3: 固定依赖并安装开发环境**

在 `requirements.txt` 增加兼容 Pillow 12.2 的 HEIF 解码依赖，并安装开发与打包环境：

`.\.venv\Scripts\python.exe -m pip install -r requirements.txt`

Expected: Windows CPython 3.11 wheel 安装成功。HEIF 依赖最终固定为 `pi-heif==1.4.0`；`python -c "import pi_heif; print(pi_heif.__version__)"` 输出 `1.4.0`。

- [x] **Step 4: 实现插件注册和去重接口**

在 `ImageService` 模块内只注册一次 opener。`deduplicate` 先扩展名过滤、再规范路径、按大小分组并计算 SHA-256；读取失败加入只包含 `path.name` 的 failure 并继续。窗口把批次放入导入 QThread，完成后才在 UI 线程更新路径；新批次排队并使用上一批更新后的 existing 快照。最大数量判断在结果返回后执行，并把添加、跳过和失败数量写入完整状态文案。

- [x] **Step 5: 运行定向和全量测试**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_image_service.py tests/test_cli_and_ipc.py tests/test_main_window.py -q`

Expected: 全部 PASS；状态测试断言“已添加 N 张，跳过 M 张重复图片”和“未添加：M 张图片已存在”。

- [x] **Step 6: 记录检查点**

把修改文件、依赖版本和测试结果写入 `progress.md`；确认没有证书、用户图片或临时 HEIF 样本进入项目目录。

### Task 2: 模板完整性自检与 CI 命令

**Files:**
- Create: `models/template_validation.py`
- Create: `tools/check_templates.py`
- Modify: `services/template_service.py`
- Modify: `ui/main_window.py`
- Test: `tests/test_template_service.py`
- Test: `tests/test_document_integration.py`

**Interfaces:**
- Produces: `TemplateIssue(template_number: int, code: str, message: str)`。
- Produces: `TemplateValidationResult(valid: bool, checked_count: int, issues: tuple[TemplateIssue, ...])`。
- Produces: `TemplateService.validate_all(deep: bool = False) -> TemplateValidationResult`。
- Produces: `python tools/check_templates.py [--template-dir PATH]`，成功退出 0，失败打印摘要并退出 1。

- [x] **Step 1: 写模板检查失败测试**

覆盖真实 1～30 模板通过，以及缺失文件、损坏 ZIP、缺少 `word/document.xml`、缺公共占位符、缺/多图片占位符。测试生成最小 DOCX ZIP，不修改真实模板。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_template_service.py -q`

Expected: FAIL，因验证模型和 `validate_all` 不存在。

- [x] **Step 3: 实现轻量与深度校验**

轻量模式只检查文件集合；深度模式读取 ZIP 中的 `word/document.xml`，合并 XML 文本后提取 `{{...}}`，与精确期望集合比较。问题按模板编号、问题代码排序，确保诊断稳定。

- [x] **Step 4: 接入启动状态和生成门禁**

窗口创建完成后用 `QTimer.singleShot(0, ...)` 运行轻量检查。失败时状态显示“模板不完整，请打开关于 / 诊断”，禁用生成但保持图片选择、清空和诊断可用。

- [x] **Step 5: 实现并验证命令行检查**

Run: `.\.venv\Scripts\python.exe tools\check_templates.py`

Expected: 输出 `30/30 templates valid`（或等价稳定文案），退出 0。

- [x] **Step 6: 运行集成测试并记录检查点**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_template_service.py tests/test_document_integration.py tests/test_main_window.py -q`

Expected: PASS。记录真实模板检查结果到 `progress.md`。

### Task 3: Word 模板准备与图片阶段安全重叠

**Files:**
- Create: `models/word_export_session.py`
- Modify: `services/layout_service.py`
- Modify: `services/document_service.py`
- Modify: `services/word_exporter.py`
- Modify: `ui/main_window.py`
- Test: `tests/test_generation_worker.py`
- Test: `tests/test_word_exporter.py`
- Test: `tests/test_document_integration.py`

**Interfaces:**
- Produces: `PreparedWordExport`，记录创建线程 ID、docxtpl 对象、文本上下文和输出路径；不对 UI 暴露 doc 对象。
- Produces: `DocumentService.prepare(task: GenerationTask, cancellation: CancellationToken | None = None) -> PreparedWordExport`。
- Produces: `DocumentService.export_prepared(prepared: PreparedWordExport, image_paths: list[Path], progress: ProgressCallback | None = None, cancellation: CancellationToken | None = None) -> Path`。
- Preserves: `DocumentService.generate(...) -> Path` 作为串行兼容包装。

- [x] **Step 1: 写重叠、线程归属和失败清理测试**

新增：

- `test_worker_opens_template_while_images_are_still_preparing`
- `test_prepared_word_export_is_consumed_on_creator_thread`
- `test_template_prepare_failure_cancels_and_cleans_image_future`
- `test_image_failure_does_not_render_or_leave_temp_files`
- `test_cancel_during_overlap_emits_cancelled_once`

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_generation_worker.py tests/test_word_exporter.py -q`

Expected: 新增接口和时序断言 FAIL。

- [x] **Step 3: 拆分 Word 准备和导出**

`WordExporter.prepare(...)` 在当前线程打开 `DocxTemplate`；`export_prepared(...)` 首先断言/检查创建线程 ID，再构造 `InlineImage`、render、原子 save。串行 `export(...)` 调用这两个接口保持兼容。

- [x] **Step 4: 在 GenerationWorker 中编排重叠**

用单 worker 的 `ThreadPoolExecutor` 提交完整 `ImageService.prepare_images`；GenerationWorker 当前线程执行 `DocumentService.prepare`，等待图片 future 后调用 `export_prepared`。异常路径设置 cancellation、等待 future、清理 `PreparedImages`；不强制终止线程。

- [x] **Step 5: 运行定向、集成和重复测试**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_generation_worker.py tests/test_word_exporter.py tests/test_document_integration.py -q`

Run: `1..5 | ForEach-Object { .\.venv\Scripts\python.exe -m pytest tests/test_generation_worker.py -q }`

Expected: 所有运行 PASS，无线程泄漏或 Qt 销毁警告。

- [x] **Step 6: 运行 1/10/30 基准并记录检查点**

Run: `.\.venv\Scripts\python.exe benchmarks\image_pipeline.py --count 30 --runs 3 --output benchmark-overlap-30.json`

记录图片阶段未回归；Word 重叠收益通过阶段日志比较，不设置易波动的 CI 绝对耗时断言。

### Task 4: 日志轮转与全局异常兜底

**Files:**
- Create: `utils/exception_handler.py`
- Modify: `config.py`
- Modify: `utils/logging_utils.py`
- Modify: `main.py`
- Test: `tests/test_logging_utils.py`
- Create: `tests/test_exception_handler.py`

**Interfaces:**
- Produces: `current_log_path() -> Path | None`。
- Produces: `GlobalExceptionHandler(QObject)` 的 `set_sensitive_values(values: Iterable[str | Path]) -> None`、`clear_sensitive_values() -> None`、`capture(exc_type, exc_value, traceback) -> None` 和 `capture_thread(args: threading.ExceptHookArgs) -> None`。
- Produces: `install_exception_handlers(app: QApplication, logger: logging.Logger) -> GlobalExceptionHandler`。

- [x] **Step 1: 写轮转与异常钩子失败测试**

测试 RotatingFileHandler 的 `maxBytes=5 MiB`、`backupCount=3`，小阈值下注入消息后产生有限备份；测试主线程和线程异常记录 `critical`，UI 文案不含异常中的私人路径；测试处理器内部失败回退原 hook 且不递归。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_logging_utils.py tests/test_exception_handler.py -q`

Expected: FAIL，现有 handler 不轮转且异常模块不存在。

- [x] **Step 3: 实现轮转 handler 和路径查询**

保持原目录候选顺序；只在 handler 成功创建后设置当前日志路径。重复调用 `setup_logging` 不重复添加 handlers，测试显式关闭临时 handlers。

- [x] **Step 4: 实现并安装异常处理器**

处理器先格式化 traceback，再用用户目录和当前任务登记值做最长优先字符串替换，最后记录脱敏后的完整堆栈；通过 Qt signal 将通用提示排队到 UI 线程；跳过 `KeyboardInterrupt/SystemExit`。GenerationWorker 开始前由窗口登记 title/location/image/save，finished 后清除。`main()` 在 QApplication 创建后安装并把对象保存在 app 属性中；启动前异常记录并返回非零，不再把无处理异常重新抛给无控制台的打包程序。

- [x] **Step 5: 运行定向和全量测试并记录检查点**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_logging_utils.py tests/test_exception_handler.py tests/test_generation_worker.py -q`

Expected: PASS，日志和弹窗文本不泄露测试私人目录。

### Task 5: Windows 右键菜单状态与路径自愈

**Files:**
- Create: `services/explorer_integration.py`
- Modify: `tools/register_explorer_menu.py`
- Modify: `main.py`
- Test: `tests/test_explorer_integration.py`
- Modify: `tests/test_cli_and_ipc.py`

**Interfaces:**
- Produces: `ExplorerRegistrationStatus` enum：`UNSUPPORTED`、`NOT_REGISTERED`、`CURRENT`、`STALE`、`PARTIAL`。
- Produces: `build_command(executable: Path) -> str`。
- Produces: `registration_status(executable: Path) -> ExplorerRegistrationStatus`。
- Produces: `register(executable: Path) -> None`、`unregister() -> None`、`repair_if_registered(executable: Path, packaged: bool) -> bool`。

- [x] **Step 1: 写纯函数和假注册表失败测试**

用 fake `winreg`/后端覆盖命令引号、支持扩展名含 HEIC/HEIF、全部当前、部分缺失、全部旧路径、未注册、非 Windows、开发态不修复和打包态修复。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_explorer_integration.py tests/test_cli_and_ipc.py -q`

Expected: FAIL，服务模块不存在。

- [x] **Step 3: 提取共享服务并改造 CLI**

`winreg` 延迟导入；每用户 HKCU 路径和菜单文案保持兼容。工具脚本只负责参数解析、文件存在检查、调用服务和退出码。

- [x] **Step 4: 启动时执行保守自愈**

在应用已经能记录日志后调用 `repair_if_registered(Path(sys.executable), packaged=getattr(sys, "frozen", False))`。只修复既有注册；失败记录 warning，不阻止启动。

- [x] **Step 5: 运行测试并记录检查点**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_explorer_integration.py tests/test_cli_and_ipc.py -q`

Expected: PASS；不实际写入开发机注册表。

### Task 6: 记住上次输入与拖拽高亮反馈

**Files:**
- Create: `services/settings_service.py`
- Modify: `ui/main_window.py`
- Test: `tests/test_settings_service.py`
- Modify: `tests/test_main_window.py`

**Interfaces:**
- Produces: `UserPreferences(title: str, location: str, activity_date: date, open_file: bool, always_on_top: bool, recent_image_dir: Path | None)`。
- Produces: `SettingsService.load() -> UserPreferences`、`save(preferences: UserPreferences) -> None`、`clear_title() -> None`。
- Produces: `MainWindow._set_drag_active(active: bool) -> None`。

- [x] **Step 1: 写设置边界和拖拽状态失败测试**

使用临时 INI `QSettings` 覆盖跨实例恢复、坏日期回退今天、失效目录回退 None、不存在 `image_paths` 键、成功后清标题但保留地点/日期、明确清空后标题/地点/日期持久化为默认值。构造 `QMimeData/QDragEnterEvent/QDragLeaveEvent/QDropEvent` 覆盖有效和无效拖拽，断言动态属性、样式刷新和文字恢复。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_settings_service.py tests/test_main_window.py -q`

Expected: FAIL，设置服务和拖拽属性不存在。

- [x] **Step 3: 实现设置服务并注入窗口**

`MainWindow(settings_service: SettingsService | None = None)` 默认创建真实 QSettings 服务，测试注入临时实例。创建 widgets 后加载；开始生成前和正常关闭时保存；成功后清标题/照片但保留地点和日期；明确清空时恢复并保存标题/地点/日期默认值。最近目录只在文件选择成功时更新。

- [x] **Step 4: 实现稳定拖拽反馈**

有效 drag enter 设置 `dragActive=true` 和“松开即可添加图片”；无效输入清理状态。`dragLeaveEvent` 和 `dropEvent` 使用 `finally` 清理；QSS 用属性选择器同时改变边框和背景，不改变几何尺寸。

- [x] **Step 5: 运行 UI 测试并记录检查点**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_settings_service.py tests/test_main_window.py tests/test_cli_and_ipc.py -q`

Expected: PASS，测试后没有写入真实用户 QSettings。

### Task 7: 关于 / 诊断入口

**Files:**
- Create: `models/diagnostics.py`
- Create: `services/diagnostics.py`
- Create: `ui/diagnostics_dialog.py`
- Modify: `config.py`
- Modify: `ui/main_window.py`
- Test: `tests/test_diagnostics.py`
- Modify: `tests/test_main_window.py`

**Interfaces:**
- Produces: `DiagnosticSnapshot`，字段包含 app version、Python、platform、packaged、HEIF、template result、log availability 和 Explorer status。
- Produces: `DiagnosticsService.collect(deep_templates: bool = True) -> DiagnosticSnapshot`。
- Produces: `DiagnosticsService.format_for_copy(snapshot: DiagnosticSnapshot) -> str`，对用户目录和可执行路径归一化。
- Produces: `DiagnosticsDialog(snapshot, log_directory, parent=None)`。

- [x] **Step 1: 写诊断内容与隐私失败测试**

断言稳定文本包含版本、Python、HEIF、模板、日志、打包和右键状态；不包含标题、地点、图片路径、输出路径或测试用户目录。测试按钮可访问名称、复制到 clipboard 和日志目录打开调用。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_diagnostics.py tests/test_main_window.py -q`

Expected: FAIL，诊断模块和按钮不存在。

- [x] **Step 3: 实现快照服务和对话框**

快照服务只消费已有服务状态，不扫描用户文件。对话框使用 `QPlainTextEdit(readOnly=True)`；三个文字按钮均有 accessible name；复制后状态写“诊断信息已复制”，不弹额外模态框。

- [x] **Step 4: 接入主窗口布局和 Tab 顺序**

增加视觉次要的“关于 / 诊断”按钮，不改变“生成文档”的主 CTA。生成中仍允许打开诊断，但禁用会修改输入的控件；模板错误时诊断入口保持可用。

- [x] **Step 5: 运行 UI 测试和人工无头快照检查**

Run: `$env:QT_QPA_PLATFORM='offscreen'; .\.venv\Scripts\python.exe -m pytest tests/test_diagnostics.py tests/test_main_window.py -q`

Expected: PASS；主窗口最小尺寸下按钮不重叠，Tab 顺序与视觉顺序一致。

- [x] **Step 6: 记录 UI/UX 检查点**

在 `progress.md` 记录：拖拽有文字替代、颜色不是唯一状态、选择按钮保留、所有新按钮有可访问名称和可见焦点。

### Task 8: CI 测试矩阵与模板门禁

**Files:**
- Modify: `.github/workflows/compile.yml`
- Modify: `requirements-dev.txt`
- Test: `tests/test_workflow_contract.py`

**Interfaces:**
- Produces: `test` job，`windows-latest × Python {3.11, 3.12}`。
- Modifies: `build.needs == "test"`。

- [x] **Step 1: 写工作流契约失败测试**

测试以文本/安全 YAML 解析断言：存在 test job、矩阵 3.11/3.12、安装两个 requirements、运行 compileall/pytest/check_templates、build 依赖 test。避免引入会把 GitHub `${{ }}` 当模板执行的解析方式。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_workflow_contract.py -q`

Expected: FAIL，现有 workflow 没有 test job。

- [x] **Step 3: 重构 workflow job 顺序**

保留现有 main push 和 workflow_dispatch 触发语义；test job 使用最小 `contents: read`，build/release 才使用 `contents: write`。build 固定安装 requirements、requirements-dev 和已验证 PyInstaller 范围。

- [x] **Step 4: 本地运行与 CI 等价命令**

Run: `.\.venv\Scripts\python.exe -m compileall main.py app.py config.py ui services models utils tools tests`

Run: `.\.venv\Scripts\python.exe -m pytest -q`

Run: `.\.venv\Scripts\python.exe tools\check_templates.py`

Expected: 全部成功。

- [x] **Step 5: 记录检查点**

将 workflow 文件、命令和本地结果记录到 `progress.md`；由于没有 Git 元数据，明确 CI 远端实际运行尚需推送后验证。

### Task 9: 打包体积报告、PyInstaller 收集与代码签名

**Files:**
- Create: `tools/package_report.py`
- Create: `tools/sign_release.ps1`
- Modify: `pgt.spec`
- Modify: `.github/workflows/compile.yml`
- Create: `tests/test_package_report.py`
- Create: `tests/test_signing_contract.py`
- Modify: `tests/test_workflow_contract.py`

**Interfaces:**
- Produces: `scan_distribution(root: Path) -> PackageReport`。
- Produces: `python tools/package_report.py DIST_DIR [--max-mib 130] [--json PATH]`。
- Produces: `tools/sign_release.ps1 -TargetPath PATH -CertificateBase64 ENV -CertificatePassword ENV [-TimestampUrl URL]`；只接受应用主 EXE，成功签名并验签退出 0，失败退出非零，始终清理。

- [x] **Step 1: 写体积报告和签名契约失败测试**

体积测试覆盖文件数、总字节、分组、最大文件、预算边界和 JSON。签名契约静态断言：目标固定为单个 `PGT.exe`、SecureString、CurrentUser\My、thumbprint、`/fd SHA256`、`/tr`、`/td SHA256`、`verify /pa /tw`、默认 `http://timestamp.digicert.com`、finally 删除证书/PFX，且脚本没有打印 secret。

- [x] **Step 2: 运行测试确认失败**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_package_report.py tests/test_signing_contract.py tests/test_workflow_contract.py -q`

Expected: FAIL，脚本和 workflow 步骤不存在。

- [x] **Step 3: 实现体积工具**

报告按相对顶层目录分组、稳定排序，默认打印总 MiB 和前 20 大文件。超过 `--max-mib` 返回 1，但仍写完整 JSON 供 CI artifact/日志诊断。

- [x] **Step 4: 调整 PyInstaller 配置并构建一次**

从 datas 移除重复 `config.py`；显式收集 pi-heif 二进制/metadata；仅收集编号模板；排除未使用的 `PyQt6.QtPdf`、`PyQt6.QtPdfWidgets` 和 `PIL.AvifImagePlugin` 及其 QtPdf DLL。保留 `opengl32sw.dll`、QtNetwork、one-dir、`upx=False`。使用控制台 bootloader 以支持可捕获的 `--diagnose` 输出。

Run: `.\.venv\Scripts\pyinstaller.exe pgt.spec --clean --noconfirm`

Expected: `dist/PGT/PGT.exe` 存在，HEIF 二进制和 30 个模板均在产物中。

- [x] **Step 5: 运行包内烟测与体积预算**

为打包入口提供不显示窗口的 `--diagnose` 或独立导入烟测，验证 HEIF opener、模板自检和核心 imports。

Run: `.\.venv\Scripts\python.exe tools\package_report.py dist\PGT --max-mib 130 --json package-report.json`

Expected: 退出 0；记录总 MiB、文件数、相比 119.01 MiB 基线的变化和最大项。

- [x] **Step 6: 实现安全签名脚本和可选 workflow 步骤**

workflow 把 secrets 映射为 job env，再用 `if: env.WINDOWS_CERTIFICATE_BASE64 != '' && env.WINDOWS_CERTIFICATE_PASSWORD != ''`。`WINDOWS_TIMESTAMP_URL` 允许变量覆盖，否则脚本采用 `http://timestamp.digicert.com`。签名发生在压缩前且仅针对 `dist/PGT/PGT.exe`；脚本导入证书、按 thumbprint 签名和验签，在 `finally` 删除证书和 PFX。无 secret 时输出明确 skip 步骤，不能假称已签名。

- [x] **Step 7: 运行脚本契约和无凭据路径测试**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_package_report.py tests/test_signing_contract.py tests/test_workflow_contract.py -q`

Run: `pwsh -NoProfile -File tools\sign_release.ps1 -TargetPath dist\PGT\PGT.exe -WhatIf`

Expected: 契约 PASS；`-WhatIf` 只验证 SignTool 定位、目标和清理计划，不要求真实证书且不修改 EXE。

- [x] **Step 8: 记录检查点**

在 `progress.md` 记录构建命令、体积前后、烟测、排除项理由和“未执行真实签名（无证书）”或真实验签结果。

### Task 10: 文档、全量验证与交付

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Modify: `task_plan.md`
- Modify: `findings.md`
- Modify: `progress.md`

**Interfaces:**
- Consumes: Tasks 1～9 的测试、构建、体积和签名结果。
- Produces: 用户操作说明、维护者发布说明和可审计验收记录。

- [x] **Step 1: 更新用户与维护文档**

README 说明 HEIC/HEIF、精确去重、设置恢复边界、拖拽反馈、诊断入口、右键菜单自愈（新位置启动一次）、日志轮转、模板检查、CI、体积和签名 secret 名称。CHANGELOG 只声明已验证结果。

- [x] **Step 2: 运行静态与全量自动化验证**

Run: `.\.venv\Scripts\python.exe -m compileall main.py app.py config.py ui services models utils tools benchmarks tests`

Run: `.\.venv\Scripts\python.exe -m pytest -q`

Run: `.\.venv\Scripts\python.exe tools\check_templates.py`

Expected: 全部成功。

- [x] **Step 3: 运行真实模板和性能验收**

Run: `.\.venv\Scripts\python.exe -m pytest tests/test_document_integration.py -q`

Run: `.\.venv\Scripts\python.exe benchmarks\image_pipeline.py --count 1 --runs 3 --output benchmark-final-1.json`

对 10、30 重复。Expected: 1/10/30 DOCX 兼容测试 PASS；图片阶段不比同机既有 after 基线显著回归。

- [x] **Step 4: 重新构建并验证最终产物**

Run: `.\.venv\Scripts\pyinstaller.exe pgt.spec --clean --noconfirm`

Run: `.\.venv\Scripts\python.exe tools\package_report.py dist\PGT --max-mib 130 --json package-report-final.json`

Expected: 构建和预算 PASS；模板、HEIF、诊断、单实例 imports 均通过包内烟测。

- [x] **Step 5: UI/UX 最终检查**

核对：焦点可见、Tab 顺序、拖拽不是唯一入口、状态不只靠颜色、错误文案含恢复路径、生成中按钮语义、模板失败时诊断可用、最小窗口尺寸不重叠。

- [x] **Step 6: 需求覆盖自检**

逐项映射并核验：HEIC/HEIF、去重、日志轮转、右键自愈、CI、全局异常、模板自检、体积、Word/图片重叠、记住输入、拖拽反馈、关于/诊断、代码签名。任何未验证项保持 pending，不以代码存在代替验收。

- [x] **Step 7: 完成持久化计划和交付记录**

把阶段状态更新为 complete，记录所有改动文件、测试数量、构建体积、未执行的外部动作（例如没有真实证书时的 Authenticode 实签）和后续远端 CI 验证事项。

## Completion Gate

- HEIC/HEIF 真实样本通过完整图片准备流程。
- 去重四类边界测试通过，最大数量在去重后计算。
- 1～30 模板深度检查通过；缺失和损坏案例被识别。
- Word/图片阶段时序和线程归属测试通过，取消/失败无临时文件。
- 日志轮转和主/后台全局异常测试通过，无私人路径泄漏。
- 右键菜单只修复既有打包态注册，开发态测试不写真实注册表。
- 设置不保存照片列表；拖拽和诊断 UI 满足键盘、焦点和非颜色反馈要求。
- CI 测试矩阵、build 依赖和模板门禁契约通过。
- PyInstaller 最终产物通过烟测且不超过 130 MiB，报告已保存。
- 无证书路径明确跳过；若提供证书，则签名、时间戳和 `/pa /tw` 验签全部成功并清理凭据。
- `compileall`、全量 pytest、1/10/30 DOCX 集成和最终构建全部通过。
