# SDD ledger — plan: D:\Projects\PicGenTool\docs\superpowers\plans\2026-11-26-quality-release-optimization.md

Setup: baseline 42 tests passed before implementation.

Ruling: Execute in the existing workspace without a Git worktree — the project is not a Git repository, and initializing one would be an out-of-scope state change — cost if wrong: edits are not isolated by Git and cannot be reviewed as commit ranges; mitigation is exact file lists, RED→GREEN tests, persistent progress, and no unrelated rewrites.

Ruling: Replace unavailable `test-driven-development` and `verification-before-completion` skills with the plan's explicit failing-test-first sequence and a recorded completion checklist — the referenced skills are not installed in this harness — cost if wrong: process guidance may be less comprehensive; mitigation is to watch each named test fail, make the minimum change, rerun task tests, then the full suite.

Pre-flight: Task 1 `SUPPORTED_FORMATS` and deduplicated image list feed Tasks 5–7 Explorer, settings/UI, and diagnostics; use `config.SUPPORTED_FORMATS` as the single extension source and `DeduplicationResult` as the only import result.

Pre-flight: Task 2 `TemplateValidationResult` feeds Task 7 diagnostics and Task 8 CI; keep a platform-neutral dataclass and stable issue codes/messages.

Pre-flight: Task 3 `PreparedWordExport` consumes existing cancellation and atomic-save semantics; keep the serial `DocumentService.generate` wrapper so existing integrations remain valid.

Pre-flight: Task 4 `current_log_path` and `GlobalExceptionHandler` feed Task 7 diagnostics and Task 3 worker privacy; expose read-only status and register/clear current task sensitive values around worker lifetime.

Pre-flight: Task 5 Explorer status feeds Task 7 diagnostics and Task 8/9 packaged behavior; keep registry imports lazy and status side-effect-free.

Pre-flight: Task 6 settings and drag state share `MainWindow` with Tasks 1, 2, 3, and 7; preserve constructor defaults and inject services only through optional parameters so existing tests remain valid.

Pre-flight: Task 8 workflow is extended by Task 9 signing and package gates; Task 9 must preserve the test job and `build.needs: test`.

Task 1 RED: `.venv\\Scripts\\python.exe -m pytest tests\\test_image_service.py tests\\test_cli_and_ipc.py tests\\test_main_window.py -q` → 7 failed, 22 passed. Failures matched missing pillow-heif, deduplication API, HEIC extensions, and async import queue.

Task 1: dependency install finding — `pillow-heif==1.9.0` requires `Pillow>=13.0.0`, conflicting with the project's pinned `Pillow==12.2.0`; investigate the latest compatible pillow-heif instead of widening the core image dependency without evidence.

Task 1: Ruling: pin `pillow-heif==1.8.1` instead of planned 1.9.0 — PyPI metadata and pip dry-run confirm 1.8.1 supports Pillow>=11.1 and Windows CPython 3.11/3.12, while 1.9.0 requires Pillow>=13 — cost if wrong: misses fixes introduced only in 1.9.0; mitigation is exact HEIC/HEIF decode tests and a future Pillow 13 upgrade can move the pin.

Task 1: Ruling: change the pre-existing two-image UI test fixture to use different pixel colors — its two PNGs were byte-identical and are correctly deduplicated by the new spec — cost if wrong: the test no longer exercises adding identical-looking files; dedicated same-content dedup tests cover that behavior explicitly.

Task 1: complete (no git commits: non-repository workspace; tests: `.venv\\Scripts\\python.exe -m pytest -q` → 49 passed in 2.01s; compileall → success; pillow-heif → 1.8.1).

Task 2 RED: template validation tests → 6 failed because models/API/startup gate were absent.

Task 2 investigation: startup fields were accidentally inserted into `GenerationWorker`; moving them to `MainWindow` made the focused Qt tests pass. Real `Template-23.docx` contained `{{%image1}}`; adding image count 23 to the integration test reproduced `jinja2.TemplateSyntaxError: unexpected '%'`.

Task 2: Ruling: repair `Template-23.docx` instead of relaxing validator syntax — docxtpl cannot render the malformed tag, while all other templates use `{{image1}}` — cost if wrong: binary package metadata/compression changes; mitigation: preserve original in approved temp storage, alter only one text run, validate package against original, and render all 23 images.

Task 2: complete (no git commits; template CLI → 30/30 valid; DOCX validator → all passed, paragraphs 73→73; targeted → 13 passed; full suite → 56 passed in 3.12s).

Task 3 RED: exporter tests failed collection because `PreparedWordExport` and split prepare/export APIs did not exist.

Task 3 debugging: one overlap cancellation test emitted an AttributeError because the future passed the cancellation token positionally while the boundary test consumed the documented keyword; changed submission to `cancellation=...`.

Task 3: complete (no git commits; targeted → 17 passed; worker repeated → 6 passed × 5; cancellation → 7 passed; full suite → 62 passed in 2.99s; 30-image median → 190.063 ms vs previous 188.833 ms, ~0.65% variance).

Task 4 RED: exception tests failed collection because `utils.exception_handler` did not exist.

Task 4: complete (no git commits; rotating handler 5 MiB/3 backups; registered task values and user home sanitized before critical logging; sys/thread hooks installed after QApplication; handler failure falls back once; targeted → 13 passed; full suite → 68 passed in 3.24s).

Task 5 RED: Explorer integration tests failed collection because the shared service did not exist.

Task 5: complete (no git commits; lazy winreg service covers 7 supported extensions including HEIC/HEIF; packaged startup repairs only STALE/PARTIAL registrations; fake-registry targeted → 13 passed; full suite → 73 passed in 3.11s; development machine registry remained read-only and reported `partial`).

Task 6 RED: settings service initially missing; after service unit implementation, 12 window tests failed because injection and drag state were not implemented.

Task 6: Ruling: compare dropped paths as `Path` values in tests — Qt normalizes Windows local URLs with `/` while `str(Path)` uses `\\`; separator spelling is not part of the UI contract — cost if wrong: could hide malformed path text; image service existence/read tests cover actual usability.

Task 6: complete (no git commits; only title/location/date/checkboxes/recent directory persisted; success preserves location/date, explicit clear persists defaults; dynamic drag state restores safely; targeted → 20 passed; full suite → 81 passed in 3.32s).

Task 7 RED: diagnostics tests failed collection because snapshot/service/dialog modules did not exist.

Task 7: ui-ux-pro-max guidance influenced the implementation: modal focus starts inside the dialog, native focus handling remains visible, controls are text-labelled with accessible names, and Tab order follows visual order. Headless screenshots at 440×520 and 520×400 showed no overlap; the offscreen Qt image used missing-font boxes, a test-environment font limitation rather than layout clipping.

Task 7: complete (no git commits; diagnostic copy omits user paths and all task inputs; targeted → 11 passed; full suite → 84 passed in 3.42s).

Task 8 RED: workflow contract → 2 failed because no test job existed.

Task 8: complete locally (no git commits; windows-latest Python 3.11/3.12 matrix; compileall/pytest/template gates; build needs test; PyInstaller range pinned; contract → 2 passed; local equivalent → 86 passed and 30/30 templates valid). Remote Actions execution requires a future push because this workspace has no Git metadata.

Task 9 RED: package/signing/workflow tests initially failed collection because the report/signing modules were absent.

Task 9: Ruling: use `pi-heif==1.4.0`, the upstream decoder-only distribution — full pillow-heif's libheif PE import table directly imports libx265, so deleting that DLL would break loading, while this application only decodes HEIF to JPEG — cost if wrong: HEIF encoding is unavailable and decoder libheif is 1.23.0 rather than full wheel 1.23.6; mitigation: fixed hashed HEIC/HEIF fixtures and frozen-package decode smoke test.

Task 9: complete (no git commits; PyInstaller one-dir console bootloader supports captured `--diagnose`; QtPdf binaries and PIL Avif excluded; exactly 30 templates; frozen HEIC sample and core imports pass; 944 files / 113.09 MiB vs 119.01 MiB baseline; full suite → 94 passed). Real local `-WhatIf` correctly failed for missing SDK; a controlled SignTool stub validated the WhatIf plan. No Authenticode signing claimed.

Final review: independent reviewer found 6 Important release blockers and 4 Minor hardening issues. All were verified against code and reproduced where applicable. Fixed import/clear/generation/drop state gating and final feedback, safe Word errors and cleanup logs, unconditional worker finished, monotonic progress, immutable tags, explicit SDK SignTool discovery, safe template issue details, frozen GUI/Word smoke, and asynchronous fragmented IPC with one-shot completion.

Final review tests: focused blocker/workflow suite → 35 passed; full suite → 105 passed in 3.86s; compileall and 30/30 templates pass; rebuilt frozen HEIC + one-image Word + offscreen GUI diagnostics pass; package 944 files / 113.10 MiB; pip check clean. IPC probe confirmed the test initially waited for a second readyRead after waitUntil consumed the first, not a production acknowledgement loss. Real credentialed signing and remote CI still require external execution.

Final re-review: save-dialog event reentrancy could bypass the initial import gate. Added `_generation_starting` before `_build_task`, gates shared by dialog/drag/IPC, stable image snapshot, and post-dialog import recheck. Deferred clear now always confirms once, disables input/import, discards late results, and restores idle status. Diagnostic temp cleanup cannot mask success. Focused UI/IPC/diagnostics → 28 passed; final full suite → 108 passed in 4.02s; rebuilt frozen diagnostics pass. Independent final reviewer reported no Critical/Important issues. Accepted-save path is tested; busy IPC returns `busy` rather than a false `ok`.

Final: all 13 requested optimizations and review fixes complete locally. Evidence: 108 tests, compileall, pip check, 30/30 templates, real 1/10/23/30 DOCX integration, frozen HEIC + one-image Word + offscreen GUI smoke, package 944 files / 113.10 MiB. No Git commit/push; remote CI matrix and credentialed Authenticode signing remain external verification steps.
