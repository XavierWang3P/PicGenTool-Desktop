# SDD ledger — plan: D:\Projects\PicGenTool\docs\superpowers\plans\2026-11-25-python-performance-optimization.md

Setup: Ruling: 当前源码来自 GitHub main 分支压缩包，目录不是 Git 仓库且系统无 Git — 将此独立副本作为隔离工作区并执行全部测试，但无法创建 worktree 或产生任务提交 — 若判断错误，代价是缺少可审计的提交边界与便捷回滚。
Setup: Ruling: test-driven-development 与 verification-before-completion 技能不可用 — 逐任务手工执行 RED→GREEN 和最终验证门禁 — 若判断错误，代价是缺少技能自动化检查但不降低测试证据要求。

Pre-flight: Task 3 produces `_prepare_one`, Task 4 consumes it — interface matches.
Pre-flight: Task 3 produces `MAX_IMAGE_WORKERS`, Task 4 consumes it — interface matches.
Pre-flight: Task 5 changes cancellation interfaces consumed by Tasks 6 and 7 — interface matches.
Pre-flight: Task 2 timing utility is consumed by image, document, Word, and startup stages — plan omits `image_total`/`image_item` integration in Task 2, deferred to Tasks 3–4 under the spec authority.
Pre-flight: Task 6 changes `WordExporter.export` cancellation parameter consumed by `DocumentService` — Task 5 must pass the token through when Task 6 lands.

Task 1: complete (no commits: Git unavailable; tests: `.venv\\Scripts\\python.exe -m pytest tests/test_image_service.py -v` → 1 passed; baselines: 1=174.148 ms, 10=1412.989 ms, 30=4190.970 ms median)
Task 2: complete (no commits: Git unavailable; tests: `.venv\\Scripts\\python.exe -m pytest tests -v` → 3 passed)
Task 3: complete (no commits: Git unavailable; tests: `.venv\\Scripts\\python.exe -m pytest tests/test_image_service.py -v` → 5 passed; 10-image smoke median=286.339 ms)
Task 4: complete (no commits: Git unavailable; tests: 11 passed × 5 runs; after medians: 1=20.329 ms, 10=88.557 ms, 30=188.833 ms; improvements: 10=93.7%, 30=95.5%)
Task 4: Ruling: 增强乱序测试以断言实际完成顺序不同 — 原计划描述的仅返回顺序断言在串行实现上也会通过，不能证明并发 — 若判断错误，代价仅为测试增加约 40 ms 延迟。
Task 5: complete (no commits: Git unavailable; tests: `.venv\\Scripts\\python.exe -m pytest tests -v` → 16 passed)
Task 6: complete (no commits: Git unavailable; tests: `.venv\\Scripts\\python.exe -m pytest tests -v` → 19 passed)
Task 7: complete (no commits: Git unavailable; tests: `.venv\\Scripts\\python.exe -m pytest tests -v` → 21 passed)
Task 8: Ruling: 修改 `Template-1.docx` 中唯一的非法 `{{%image1}}` 占位符为可拼接的 `{{image1}}` — 原模板无法生成 1 张图片文档，与验收标准冲突；最小单字符修复通过结构校验且段落数 6→6 — 若判断错误，代价是对“本轮不修改模板”非目标产生一个受控例外，但不改变版式。
Task 8: Ruling: 冷启动未达到降低 30% 时停止继续拆分或重写 PyQt UI — 5 次中位数原始 130.832 ms、当前 144.110 ms，QApplication/MainWindow/首事件阶段近似，导入差异处于十余毫秒范围；已延迟取消、计时和线程池依赖 — 若判断错误，代价是本轮发布保留约 10.1% 的开发版冷启动回归。
Task 8: automated portion complete (no commits: Git unavailable; compileall passed; tests: 25 passed; real templates 1/10/30 passed; PyInstaller one-dir build passed; packaged startup median=243.471 ms; image improvements 10=93.7%, 30=95.5%; peak memory 112.44 MiB vs 127.73 MiB baseline)
Task 8: manual gate pending — Microsoft Word visual layout, default-app open, occupied target file, and packaged UI click/close regression require user desktop interaction.

Final: fixed worker failure did not propagate cancellation — `test_worker_failure_cancels_running_sibling_and_leaves_no_threads` RED→GREEN, sibling observed token and no `pgt-image` threads remained.
Final: fixed private image paths leaking through exception/log text — path-injected service and worker tests RED→GREEN; UI keeps basename and persistent log records only the error type.
Final: fixed save/replace cancellation race — token commit lock tests and exporter boundary test RED→GREEN; cancellation wins before commit or becomes a no-op after commit begins.
Final: coverage added for EXIF orientation 8 — independent asymmetric corner assertions for orientations 6/8 pass; production behavior was already correct.
Final: Qt lifecycle focus verified — real QThread repeat-cancel and close-during-work test passes without destroyed-thread warning; packaged hands-on interaction remains a manual gate.
Final: fixed packaged log directory — fallback handler test RED→GREEN; rebuilt package writes `%LOCALAPPDATA%\PicGenTool\logs\app.log` and leaves bundle clean.
Final: Ruling: Microsoft Word visual pagination/clarity/layout compatibility remains unjudged — automated DOCX structure and real templates pass, but desktop automation CLI is unavailable and no reliable visual substitute exists — cost if wrong: a template-specific visual regression could ship unless the manual matrix is completed.
Final: Ruling: exact pre-optimization output compatibility remains unjudged — source came from an archive with no Git history or golden output corpus — cost if wrong: subtle formatting differences not covered by structural tests may be missed.
Final: Ruling: default Office opening and locked-document behavior remains unjudged — depends on the user's Windows/Office configuration and desktop automation is unavailable — cost if wrong: user-facing integration or localized error behavior may fail on the target machine.
Final: Ruling: real packaged click-driven generate/cancel/close remains unjudged — real QThread automated tests and package startup pass, but agent-desktop is unavailable — cost if wrong: a packaged-only Qt interaction issue could remain.
Final: Ruling: crash/power-loss durability around replacement remains outside scope — atomic replacement covers ordinary errors and cancellation but no directory fsync design was requested — cost if wrong: abrupt power loss could still lose the newest output on some filesystems.
Final: Ruling: performance on other hardware/storage/security configurations remains unjudged — measurements only support this Windows machine — cost if wrong: speed or memory targets may not generalize.
Final: Ruling: historical baseline provenance is limited — original archive is retained only in the approved temporary download area and benchmark JSON has no revision identifier — cost if wrong: future auditors cannot independently bind every baseline number to a commit.
Final: minor (deferred): production `STARTED_AT` begins after logging/timing helper imports; dedicated subprocess benchmark remains the authoritative full startup measurement.
Final: minor (deferred): `open_output` has progress feedback but no structured timing stage.
Final: fixed reviewer minor on cancellation during temporary save as part of the Important commit-race fix; target preservation and temp cleanup are tested.
Final: minor (deferred): benchmark JSON lacks versions, worker count, deterministic sample hashes, and build identifier.
Final: minor (deferred): peak-memory and packaged-startup sampling commands are documented only in the execution record/report, not a reusable repository script.
Final review: fresh reviewer completed; Critical=0, Important code findings fixed, manual acceptance findings retained as release gates.
Final: UI matrix completed with agent-desktop on packaged PGT.exe — 30-image generation/order/page-count, 1-image template, cancel at 12/30, close-during-generation (323 ms exit), auto-open in Word, and real OS exclusive-lock save error with unchanged target hash all verified.
Final: Ruling: occupied-target GUI overwrite path not exercised — host Pulse file picker's custom overwrite-confirm panel cannot be confirmed by agent-desktop, so the same code path was verified with a real OS exclusive lock in the service layer — cost if wrong: an issue specific to the Pulse picker confirmation UI (not app code) could remain unseen.
Final: Ruling: UI regression used 30 and 1 images end-to-end; 10-image UI run skipped because Template-10 is structurally covered and UI mechanics are identical — cost if wrong: a template-10-only visual quirk would need a manual look.

UI: optimized with ui-ux-pro-max rules — visible focus rings on inputs/buttons/checkbox, label buddies + explicit tab order, destructive clear confirmation (default 保留), count shows limit, progress hidden when idle, primary/secondary/cancel hierarchy, 8px spacing rhythm, resizable window.
UI: Ruling: kept the native QCheckBox indicator instead of a custom QSS square — a custom background drops the checkmark and makes state color-only — cost if wrong: checkbox styling is slightly less uniform than inputs.
UI: Ruling: hid the progress bar when idle — an empty 0% track adds noise and the status text already states the next step — cost if wrong: users lose a constant visual anchor for progress location.
UI: Ruling: removed the middle stretch and capped cards to Maximum vertical policy, moving spare space below the buttons — the red-boxed gap existed only to pin buttons to the bottom — cost if wrong: at very tall window sizes the empty band sits under the action row instead of between cards.
UI: Ruling: Explorer integration uses HKCU SystemFileAssociations plus a single-instance forwarder instead of a COM shell extension — no DLL to build, sign, or register, and multi-select is aggregated by forwarding — cost if wrong: some Explorer versions may spawn one short-lived process per selected file before forwarding.
UI: Ruling: single-instance election uses a QLockFile plus an application-level `ok` acknowledgement — Explorer launches all selected files concurrently, and without election plus ACK two processes became primaries or lost payloads because the listener was still starting — cost if wrong: a forwarder waits up to about two seconds before its process exits.
