# PicGenTool Python Performance Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在不改变现有 UI、Word 模板和输出格式的前提下，通过可测量的图片并发与编解码优化降低 PicGenTool 启动及批量生成耗时。

**Architecture:** 保持现有 `ui → services → models` 分层，在服务边界加入轻量性能计时和协作式取消。图片处理提取为无共享状态的单图任务，通过受控线程池并行执行；Word 导出继续使用 docxtpl，但改为分阶段计时和临时文件原子保存。

**Tech Stack:** Python 3.11、PyQt6 6.11、Pillow 12.2、docxtpl 0.20.2、python-docx 1.2、pytest、pytest-qt

**Spec:** `docs/superpowers/specs/2026-11-25-python-performance-optimization-design.md`

## Global Constraints

- 保持 Python 3.11 及以上，不引入新的运行时依赖。
- 图片上限 30 张；输出保持 6.4 cm × 4.8 cm、240 DPI、JPEG 质量 60。
- 默认并发数为 `min(4, os.cpu_count() or 1, image_count)`。
- JPEG 默认 `optimize=False`，通过 `config.py` 配置。
- 不修改 30 个 Word 模板，不改变 UI 布局和文案以外的交互。
- 性能比较使用同机、同样本、至少 3 次运行的中位数。
- 日志不得记录图片完整路径、活动标题、地点等私人内容。

## Review Focus

- EXIF 方向为 6/8 的手机照片必须按视觉方向裁剪；由 Task 3 的方向测试覆盖。
- 并发任务乱序完成时返回列表仍须与输入顺序一致；由 Task 4 的乱序测试覆盖。
- 某张图片损坏或编码失败时必须取消剩余工作并清理临时目录；由 Task 4 的失败测试覆盖。
- Word 保存失败或取消时不得覆盖已有目标文件或留下半成品；由 Task 6 的原子保存测试覆盖。
- 用户连续点击生成/取消、窗口关闭期间取消不能访问已释放的 Qt 对象；由 Task 7 的 UI 状态测试和人工回归覆盖。

---

## File Map

- `config.py`：新增图片并发数和 JPEG 优化开关。
- `utils/performance.py`：统一的阶段计时与结构化日志。
- `services/cancellation.py`：取消异常和取消令牌协议。
- `services/image_service.py`：单图流水线、受控并发、进度与取消。
- `services/document_service.py`：布局和导出阶段计时、取消检查。
- `services/word_exporter.py`：细分计时和 DOCX 原子保存。
- `ui/main_window.py`：传递取消事件、取消按钮状态和启动计时。
- `benchmarks/image_pipeline.py`：生成固定样本并输出 JSON 基准结果。
- `tests/`：单元测试与服务集成测试，不依赖真实 Word/Qt 弹窗。
- `requirements-dev.txt`：开发测试依赖 `pytest` 和 `pytest-qt`。

### Task 1: 建立测试基线和可重复的性能基准

**Files:**
- Create: `requirements-dev.txt`
- Create: `tests/conftest.py`
- Create: `tests/test_image_service.py`
- Create: `benchmarks/image_pipeline.py`
- Modify: `README.md`

**Interfaces:**
- Consumes: 现有 `ImageService.prepare_images(image_paths, progress=None) -> PreparedImages`。
- Produces: `benchmarks/image_pipeline.py --count {1,10,30} --runs N --output PATH`，输出含 `count`、`runs_ms`、`median_ms` 的 JSON。

- [ ] **Step 1: 写现有行为测试**

在 `tests/test_image_service.py` 创建 `test_prepare_images_preserves_count_size_mode_and_order`：使用 3 张不同颜色的临时图片，断言输出数量为 3、文件名为 `image_01.jpg` 至 `image_03.jpg`、尺寸为根据配置计算的目标像素、模式为 RGB，并调用 `cleanup()`。

- [ ] **Step 2: 添加测试依赖并运行基线测试**

`requirements-dev.txt` 固定 `pytest>=8.0,<9.0` 和 `pytest-qt>=4.4,<5.0`。运行：`python -m pytest tests/test_image_service.py -v`。预期：PASS。

- [ ] **Step 3: 编写基准脚本**

脚本签名：`run_benchmark(count: int, runs: int, source_dir: Path | None = None) -> dict[str, object]`。未提供源目录时生成 4032×3024、方向和颜色固定的 JPEG；每轮调用现有 `ImageService` 并清理结果；JSON 不包含源文件路径。

- [ ] **Step 4: 记录优化前基线**

运行：`python benchmarks/image_pipeline.py --count 1 --runs 3 --output benchmark-before-1.json`，并对 10、30 重复。预期：三个 JSON 均包含非空 `runs_ms` 和正数 `median_ms`。

- [ ] **Step 5: 更新 README 性能测试章节**

记录基准命令、同机同样本规则及“自动化测试不使用绝对耗时断言”。

- [ ] **Step 6: 提交**

```bash
git add requirements-dev.txt tests benchmarks README.md
git commit -m "test: add image pipeline performance baseline"
```

### Task 2: 增加统一的性能计时

**Files:**
- Create: `utils/performance.py`
- Create: `tests/test_performance.py`
- Modify: `main.py`
- Modify: `services/document_service.py`
- Modify: `services/word_exporter.py`

**Interfaces:**
- Consumes: 标准库 `logging` 和 `time.perf_counter`。
- Produces: `measure(logger: logging.Logger, stage: str, **fields: int | str) -> ContextManager[None]`；`log_elapsed(logger: logging.Logger, stage: str, started_at: float, **fields: int | str) -> None`。

- [ ] **Step 1: 写计时上下文失败测试**

测试 `test_measure_logs_stage_elapsed_and_safe_fields`，替换 perf counter 为固定值，断言日志含 `stage=image_total`、`elapsed_ms=125.0`、`image_count=10`，且没有路径字段。

- [ ] **Step 2: 运行测试确认失败**

运行：`python -m pytest tests/test_performance.py -v`。预期：因 `utils.performance` 不存在而 FAIL。

- [ ] **Step 3: 实现 `measure`**

只允许字段值为整数或短字符串；调用方不得传用户路径、标题或地点。异常离开上下文时仍记录耗时并重新抛出异常。`log_elapsed` 供跨越 Qt 事件循环边界的启动计时使用。

- [ ] **Step 4: 为现有阶段接入计时**

`main.py` 在模块加载时保存 `perf_counter()`，窗口 `show()` 后通过 `QTimer.singleShot(0, ...)` 调用 `log_elapsed(..., stage="startup_ready")`，以首个 Qt 事件循环回调作为启动完成点；`DocumentService.generate()` 记录 `layout`；`WordExporter.export()` 分别记录 `word_open`、`word_render`、`word_save`。不改变公开返回值和异常文案。

- [ ] **Step 5: 运行测试**

运行：`python -m pytest tests -v`。预期：全部 PASS。

- [ ] **Step 6: 提交**

```bash
git add utils/performance.py tests/test_performance.py main.py services/document_service.py services/word_exporter.py
git commit -m "perf: add stage timing instrumentation"
```

### Task 3: 优化单张图片处理流水线

**Files:**
- Modify: `config.py`
- Modify: `services/image_service.py`
- Modify: `tests/test_image_service.py`

**Interfaces:**
- Consumes: `IMAGE_QUALITY`、`DPI`、目标尺寸配置。
- Produces: `ImageService._prepare_one(image_path: Path, output_path: Path, target_size: tuple[int, int], target_ratio: float) -> Path`。

- [ ] **Step 1: 写 EXIF、模式和 JPEG 参数测试**

增加 `test_prepare_one_applies_exif_orientation_before_crop`、`test_prepare_one_converts_rgba_and_grayscale_to_rgb`、`test_prepare_one_uses_configured_quality_without_optimize`。最后一个测试 mock `Image.Image.save` 并断言 `quality=60`、`optimize=False`、`dpi=(240, 240)`。

- [ ] **Step 2: 运行新增测试确认失败**

运行：`python -m pytest tests/test_image_service.py -v`。预期：因 `_prepare_one` 和新配置不存在而 FAIL。

- [ ] **Step 3: 添加配置**

在 `config.py` 新增 `IMAGE_OPTIMIZE = False` 和 `MAX_IMAGE_WORKERS = 4`。

- [ ] **Step 4: 实现单图流水线**

在 `_prepare_one` 中对 JPEG 调用 `draft("RGB", target_size)`，随后调用 `ImageOps.exif_transpose()`；转 RGB、居中裁剪、LANCZOS 缩放并保存。每个中间 Pillow 对象由当前调用独占并可靠关闭。

- [ ] **Step 5: 运行图片测试和基准烟测**

运行：`python -m pytest tests/test_image_service.py -v`；再运行 `python benchmarks/image_pipeline.py --count 10 --runs 1 --output benchmark-single-optimized.json`。预期：测试 PASS，基准 JSON 有效。

- [ ] **Step 6: 提交**

```bash
git add config.py services/image_service.py tests/test_image_service.py
git commit -m "perf: optimize single image processing pipeline"
```

### Task 4: 并发处理图片并保持确定性顺序

**Files:**
- Modify: `services/image_service.py`
- Modify: `tests/test_image_service.py`

**Interfaces:**
- Consumes: Task 3 的 `_prepare_one(...) -> Path` 和 `MAX_IMAGE_WORKERS`。
- Produces: `ImageService._worker_count(image_count: int) -> int`；保持 `prepare_images(...) -> PreparedImages` 外部接口不变。

- [ ] **Step 1: 写并发数、乱序、进度和失败清理测试**

增加以下测试：`test_worker_count_is_bounded_by_cpu_config_and_image_count`；`test_parallel_completion_preserves_input_order` 用不同延迟模拟乱序完成；`test_progress_is_monotonic_and_reaches_60`；`test_worker_failure_cleans_temp_directory_and_names_source_file`。

- [ ] **Step 2: 运行测试确认失败**

运行：`python -m pytest tests/test_image_service.py -v`。预期：新增测试 FAIL。

- [ ] **Step 3: 实现受控线程池**

使用 `ThreadPoolExecutor(max_workers=self._worker_count(total), thread_name_prefix="pgt-image")`。future 映射到原始索引；完成时更新进度，结果写入预分配列表；发生异常时取消未开始任务并等待执行中的任务结束，再清理临时目录。

- [ ] **Step 4: 运行全部测试**

运行：`python -m pytest tests -v`。预期：全部 PASS，重复运行 5 次无偶发失败。

- [ ] **Step 5: 采集优化后基准并计算提升比例**

分别运行 1、10、30 张、每组 3 次，输出 `benchmark-after-*.json`。计算 `(before_median-after_median)/before_median`；10 和 30 张目标均不低于 40%。若未达到，保留数据并在进入 Task 5 前检查 CPU 利用率和 `optimize` 参数是否生效，不增加超过 4 个线程。

- [ ] **Step 6: 提交**

```bash
git add services/image_service.py tests/test_image_service.py benchmark-after-*.json
git commit -m "perf: process images with bounded concurrency"
```

### Task 5: 增加协作式取消模型

**Files:**
- Create: `services/cancellation.py`
- Create: `tests/test_cancellation.py`
- Modify: `services/image_service.py`
- Modify: `services/document_service.py`

**Interfaces:**
- Produces: `GenerationCancelled(Exception)`；`CancellationToken(event: threading.Event)`，提供 `cancel() -> None`、`is_cancelled() -> bool`、`raise_if_cancelled() -> None`。
- 修改: `ImageService.prepare_images(..., cancellation: CancellationToken | None = None) -> PreparedImages`。
- 修改: `DocumentService.generate(..., cancellation: CancellationToken | None = None) -> Path`。

- [ ] **Step 1: 写令牌和服务取消测试**

测试令牌初始未取消、`cancel()` 幂等、`raise_if_cancelled()` 抛出 `GenerationCancelled`；测试图片任务在提交前取消不创建输出；处理中取消会停止提交后续结果并清理临时目录；布局和导出前会检查取消。

- [ ] **Step 2: 运行测试确认失败**

运行：`python -m pytest tests/test_cancellation.py -v`。预期：模块不存在而 FAIL。

- [ ] **Step 3: 实现取消对象和检查点**

取消仅设置 `threading.Event`，不强杀线程。在每个单图任务开始、完成后、布局前和导出前调用 `raise_if_cancelled()`。

- [ ] **Step 4: 运行全部测试**

运行：`python -m pytest tests -v`。预期：全部 PASS，且测试进程无遗留 `pgt-image` 线程。

- [ ] **Step 5: 提交**

```bash
git add services/cancellation.py services/image_service.py services/document_service.py tests
git commit -m "feat: add cooperative generation cancellation"
```

### Task 6: Word 文档原子保存与取消安全

**Files:**
- Modify: `services/word_exporter.py`
- Create: `tests/test_word_exporter.py`

**Interfaces:**
- Consumes: `CancellationToken | None`。
- 修改: `WordExporter.export(payload: WordExportPayload, cancellation: CancellationToken | None = None) -> Path`。

- [ ] **Step 1: 写原子保存测试**

mock `DocxTemplate`，测试成功时先保存到目标目录临时文件再 `replace()`；保存失败时已有目标文件内容不变且临时文件被删除；在 render 后取消时不调用正式替换；成功后返回正式路径。

- [ ] **Step 2: 运行测试确认失败**

运行：`python -m pytest tests/test_word_exporter.py -v`。预期：现有实现直接保存目标文件导致 FAIL。

- [ ] **Step 3: 实现同目录临时文件保存**

使用 `tempfile.NamedTemporaryFile(dir=payload.output_path.parent, suffix=".docx", delete=False)` 获得路径并立即关闭句柄；渲染后检查取消，保存临时路径，成功后 `replace(payload.output_path)`；所有异常和取消路径删除临时文件。

- [ ] **Step 4: 运行 Word 和全量测试**

运行：`python -m pytest tests/test_word_exporter.py tests -v`。预期：全部 PASS。

- [ ] **Step 5: 提交**

```bash
git add services/word_exporter.py tests/test_word_exporter.py
git commit -m "fix: save generated documents atomically"
```

### Task 7: 把取消能力接入 PyQt UI

**Files:**
- Modify: `ui/main_window.py`
- Create: `tests/test_generation_worker.py`
- Create: `tests/test_main_window.py`

**Interfaces:**
- Consumes: `CancellationToken` 和 `GenerationCancelled`。
- Produces: `GenerationWorker.cancel() -> None`；新增 `cancelled = pyqtSignal()`；`MainWindow._cancel_generation() -> None`。

- [ ] **Step 1: 写 Worker 取消结果测试**

测试取消异常触发 `cancelled` 而不是 `error`；`finished` 始终触发一次；连续两次 `cancel()` 不抛异常。

- [ ] **Step 2: 写窗口状态测试**

使用 Qt 的事件循环测试开始任务后生成按钮保持禁用、取消按钮启用；点击取消后状态显示“正在取消”；收到 `cancelled` 后状态显示“已取消”，所有常规按钮恢复。

- [ ] **Step 3: 运行测试确认失败**

运行：`python -m pytest tests/test_generation_worker.py tests/test_main_window.py -v`。预期：取消接口和信号不存在而 FAIL。

- [ ] **Step 4: 接入 UI 和窗口关闭处理**

增加“取消任务”按钮或在忙碌状态将生成按钮切换为取消操作，保持窗口整体布局尺寸不变。`closeEvent` 在任务运行时请求取消并忽略第一次关闭，待线程结束后允许关闭，避免强制终止 `QThread`。

- [ ] **Step 5: 运行 UI 与全量测试**

运行：`python -m pytest tests -v`。预期：全部 PASS，无 `QThread: Destroyed while thread is still running` 警告。

- [ ] **Step 6: 提交**

```bash
git add ui/main_window.py tests/test_generation_worker.py tests/test_main_window.py
git commit -m "feat: support safe generation cancellation"
```

### Task 8: 启动、打包和端到端验收

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Modify: `pgt.spec` only if measurements identify a regression
- Create: `docs/performance/2026-11-25-results.md`

**Interfaces:**
- Consumes: Task 1～7 的日志、基准 JSON 和打包配置。
- Produces: 性能结果报告和可复现的 Windows 验收记录。

- [ ] **Step 1: 执行静态和自动化验证**

运行：`python -m compileall main.py app.py config.py ui services models utils benchmarks tests`；运行：`python -m pytest tests -v`。预期：无语法错误，全部测试 PASS。

- [ ] **Step 2: 测量开发版和打包版冷启动**

分别记录从进程启动到 `MainWindow.show()` 后首个事件循环可处理事件的毫秒数，每种运行 3 次取中位数。目标较基线降低 30%；未达到时在报告中拆分 Python 导入、QApplication 和 MainWindow 创建占比，不盲目修改 `pgt.spec`。

- [ ] **Step 3: 执行 1、10、30 张端到端人工回归**

每组验证：图片顺序、方向、尺寸观感、标题/地点/日期、模板版式、生成后打开、取消、损坏图片提示、目标文件被占用提示。

- [ ] **Step 4: 打包并验证 Windows one-dir 版本**

运行：`pyinstaller pgt.spec`。预期：保持 `upx=False`、`noarchive=True`、one-dir；打包产物能完成 Task 8 Step 3 的关键路径。

- [ ] **Step 5: 写性能结果和回退条件**

`docs/performance/2026-11-25-results.md` 记录设备、样本尺寸、前后中位数、提升比例、峰值内存、未达到目标的原因。若 1 张图片回归超过 10% 或 30 张峰值内存超过串行基线 2 倍，将默认 worker 从 4 降为 2 后重新测量。

- [ ] **Step 6: 更新用户文档和变更日志**

README 增加取消行为和性能基准方法；CHANGELOG 记录并发图片处理、EXIF 修正、原子保存和实测结果，不宣称未测量的性能收益。

- [ ] **Step 7: 最终提交**

```bash
git add README.md CHANGELOG.md docs/performance pgt.spec
git commit -m "docs: record Python performance optimization results"
```

## Completion Gate

以下条件全部满足才视为完成：

- 自动化测试与 `compileall` 全部通过。
- 1、10、30 张图片的 Word 输出完成兼容性检查。
- 10 和 30 张图片处理时间相对基线中位数至少改善 40%，或有数据充分的阻塞原因和下一阶段提案。
- 取消、损坏图片、保存失败均不遗留临时文件或损坏目标文件。
- Windows one-dir 包完成至少一次完整人工回归。
- 若 `word_render + word_save` 仍占总耗时超过 50%，另建 DOCX 输出专项设计，不在本计划中临时扩大范围。
