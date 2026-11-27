# Agent Automation P0/P3 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Provide deterministic headless generation, JSON contracts, batch processing, and a persistent standard-stream automation service without changing the GUI workflow.

**Architecture:** Extract the current worker pipeline into a Qt-free `GenerationService`. Put validation and result mapping in `AgentAPI`, then share it between one-shot, JSONL, and persistent NDJSON CLI adapters. Preserve the GUI as a signal-based adapter.

**Tech Stack:** Python 3.11+, standard-library argparse/json/concurrent.futures/uuid, existing Pillow/docxtpl services, pytest.

**Spec:** `docs/superpowers/specs/2026-11-27-agent-automation-design.md`

## Global Constraints

- Headless paths must not import PyQt, read QSettings, display dialogs, or open Word.
- stdout contains final JSON only; logs and progress go to stderr.
- Existing GUI cancellation, atomic save, template checks, and Explorer commands remain compatible.
- Default existing-output behavior is `fail`; no undocumented field or input is silently ignored.

## Review Focus

- Corrupt images must produce `IMAGE_READ_FAILED`, not a traceback or a partial DOCX; test API and real CLI generation with invalid image bytes.
- A concurrent batch must never use docxtpl objects across threads; test service thread affinity and real two-job generation.
- A timeout/Ctrl-C must clean prepared images and cancel running work; test runner cancellation without waiting for unbounded work.
- A malformed persistent request must return an error and leave the stream usable; test malformed JSON followed by a valid request.
- Existing GUI mocks must still exercise the shared service and preserve signal behavior; run the entire existing generation worker suite.

---

### Task 1: Shared generation service

**Files:**
- Create: `services/generation_service.py`
- Modify: `ui/main_window.py`
- Test: `tests/test_generation_worker.py`

**Interfaces:** `GenerationService.run(task: GenerationTask, progress=None, cancellation=None) -> Path`. `GenerationWorker` delegates pipeline work and retains `open_file` and UI signals.

- [ ] Add service tests proving overlap, cleanup on success/failure/cancel, and template failure waiting for image completion.
- [ ] Run the focused tests and confirm missing `GenerationService` fails.
- [ ] Move the existing `GenerationWorker.run()` pipeline into the service without changing ordering; make the worker a thin adapter.
- [ ] Run `python -m pytest tests/test_generation_worker.py tests/test_main_window.py -q`.
- [ ] Commit `refactor: share headless generation pipeline`.

### Task 2: Versioned request and result contract

**Files:**
- Create: `models/agent_contract.py`
- Create: `services/agent_api.py`
- Test: `tests/test_agent_api.py`

**Interfaces:** `AgentAPI.validate(payload: object) -> dict`, `AgentAPI.generate(payload: object, progress=None, cancellation=None) -> dict`, `AgentAPI.capabilities() -> dict`, `request_schema() -> dict`, `AgentError(code, message, details=None)`.

- [ ] Write tests for valid generation, all invalid input classes, missing templates, all output policies, duplicates, unknown fields, schema shape, and stable error codes.
- [ ] Confirm tests fail before implementation.
- [ ] Implement explicit validation, deduplication, path/output policy, template checks, error mapping, UUID job IDs and elapsed time. Never persist settings.
- [ ] Run `python -m pytest tests/test_agent_api.py -q`.
- [ ] Commit `feat: add structured agent generation API`.

### Task 3: Headless one-shot CLI

**Files:**
- Create: `agent_cli.py`
- Modify: `main.py`
- Modify: `pgt.spec` if required by frozen imports
- Test: `tests/test_agent_cli.py`

**Interfaces:** `run_agent_cli(argv: list[str], stdin=None, stdout=None, stderr=None) -> int`, called before GUI/PyQt imports. Supports `--generate`, `--validate`, `--capabilities`, `--schema`, `--diagnose` and `--json`.

- [ ] Write subprocess tests asserting exact JSON envelope, exit codes, DOCX validity, no Qt imports in child generation, and stdin support.
- [ ] Confirm tests fail before implementation.
- [ ] Implement argparse dispatch, UTF-8 streams, bounded input, stable exception handling, and early main dispatch; preserve `--diagnose` default behavior.
- [ ] Run `python -m pytest tests/test_agent_cli.py tests/test_cli_and_ipc.py -q` and direct smoke commands.
- [ ] Commit `feat: add headless JSON command line interface`.

### Task 4: Batch and persistent runner

**Files:**
- Create: `services/agent_runner.py`
- Modify: `agent_cli.py`
- Test: `tests/test_agent_runner.py`
- Test: `tests/test_agent_cli.py`

**Interfaces:** `AgentRunner.run_batch(payloads, workers=1, on_error='continue', timeout=None) -> Iterable[dict]`; CLI `--batch jobs.jsonl --workers 1..4 --on-error continue|stop --timeout seconds`; `--serve` consumes/responds to NDJSON requests synchronously.

- [ ] Test JSONL parse/index isolation, continue/stop, worker bounds, real independent jobs, per-job timeout/cancellation, persistent correlation IDs, malformed request recovery, and EOF.
- [ ] Confirm tests fail before implementation.
- [ ] Implement bounded in-order submission/completion with `ThreadPoolExecutor`; stop scheduling after failure in stop mode, cancel via shared tokens on interruption, always emit one result per submitted item.
- [ ] Run focused runner/CLI tests plus real batch and persistent smoke tests.
- [ ] Commit `feat: add batch and persistent agent automation`.

### Task 5: Documentation, contracts, and full verification

**Files:**
- Create: `docs/agent-integration.md`
- Create: `AGENTS.md`
- Modify: `README.md`, `CHANGELOG.md`, `.github/workflows/compile.yml`
- Modify: `tests/test_workflow_contract.py`

- [ ] Add contract assertions for headless compilation and CLI smoke tests in CI; do not change the release version or existing tag semantics.
- [ ] Document exact request examples, output errors, batch/persistent use and security boundaries.
- [ ] Run `python -m compileall main.py app.py agent_cli.py config.py ui services models utils tools tests`, full pytest, `python tools/check_templates.py`, and frozen diagnostics/build where available.
- [ ] Review `git diff --check` and tracked files; commit `docs: document agent automation contracts`.
- [ ] Push ordinary commits to `origin/main`, verify remote SHA and report CI/release status separately.
