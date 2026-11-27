from __future__ import annotations

import argparse
import json
import os
import statistics
import subprocess
import sys
from pathlib import Path


HARNESS = r"""
import json
import time
started_at = time.perf_counter()
from PyQt6.QtCore import QTimer
from ui.main_window import MainWindow, create_application
imports_ready = time.perf_counter()
app = create_application()
app_ready = time.perf_counter()
window = MainWindow()
window_ready = time.perf_counter()
window.show()
def ready():
    event_ready = time.perf_counter()
    result = {
        "imports": (imports_ready - started_at) * 1000,
        "qapplication": (app_ready - imports_ready) * 1000,
        "main_window": (window_ready - app_ready) * 1000,
        "first_event": (event_ready - window_ready) * 1000,
        "total": (event_ready - started_at) * 1000,
    }
    print("STARTUP_RESULT=" + json.dumps(result), flush=True)
    app.quit()
QTimer.singleShot(0, ready)
app.exec()
"""


def run_benchmark(python: Path, project_root: Path, runs: int) -> dict[str, object]:
    if runs < 1:
        raise ValueError("runs 必须大于 0")

    environment = os.environ.copy()
    environment.setdefault("QT_QPA_PLATFORM", "offscreen")
    stage_runs: dict[str, list[float]] = {
        "imports": [],
        "qapplication": [],
        "main_window": [],
        "first_event": [],
        "total": [],
    }
    for _ in range(runs):
        result = subprocess.run(
            [str(python), "-c", HARNESS],
            cwd=project_root,
            env=environment,
            capture_output=True,
            text=True,
            timeout=30,
            check=True,
        )
        line = next(
            (
                line
                for line in result.stdout.splitlines()
                if line.startswith("STARTUP_RESULT=")
            ),
            None,
        )
        if line is None:
            raise RuntimeError("启动进程未返回计时结果")
        stages = json.loads(line.partition("=")[2])
        for stage in stage_runs:
            stage_runs[stage].append(round(float(stages[stage]), 3))

    return {
        "runs_ms": stage_runs["total"],
        "median_ms": round(statistics.median(stage_runs["total"]), 3),
        "stages_median_ms": {
            stage: round(statistics.median(values), 3)
            for stage, values in stage_runs.items()
            if stage != "total"
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="PicGenTool 开发版冷启动基准")
    parser.add_argument("--python", type=Path, default=Path(sys.executable))
    parser.add_argument("--project-root", type=Path, default=Path.cwd())
    parser.add_argument("--runs", type=int, default=3)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()

    result = run_benchmark(args.python, args.project_root, args.runs)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
