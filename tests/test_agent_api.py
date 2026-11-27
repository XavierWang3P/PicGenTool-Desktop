from __future__ import annotations

import threading
import time
from pathlib import Path
from unittest.mock import Mock

from services.agent_api import AgentAPI, AgentError
from services.agent_runner import read_jsonl, run_batch


def test_batch_stop_does_not_submit_following_jobs() -> None:
    api = Mock(spec=AgentAPI)
    api.generate.side_effect = [
        {"ok": False, "error": {"code": "INVALID_REQUEST"}},
        {"ok": True},
    ]
    results = list(run_batch([(1, {}), (2, {}), (3, {})], api, on_error="stop"))

    assert [result["line"] for result in results] == [1]
    assert api.generate.call_count == 1


def test_batch_worker_limit_and_thread_affinity() -> None:
    api = Mock(spec=AgentAPI)
    active = 0
    maximum = 0
    lock = threading.Lock()
    threads: set[int] = set()

    def generate(payload, cancellation=None):
        nonlocal active, maximum
        with lock:
            active += 1
            maximum = max(maximum, active)
            threads.add(threading.get_ident())
        time.sleep(0.03)
        with lock:
            active -= 1
        return {"ok": True}

    api.generate.side_effect = generate
    results = list(run_batch([(index, {}) for index in range(6)], api, workers=2))

    assert len(results) == 6
    assert maximum == 2
    assert len(threads) == 2


def test_batch_timeout_uses_shared_cancellation_token() -> None:
    api = Mock(spec=AgentAPI)
    seen: list[object] = []

    def generate(payload, cancellation=None):
        seen.append(cancellation)
        time.sleep(0.04)
        return {"ok": False, "error": {"code": "CANCELLED"}}

    api.generate.side_effect = generate
    result = list(run_batch([(1, {})], api, timeout=0.001))[0]

    assert result["error"]["code"] == "CANCELLED"
    assert seen[0].is_cancelled()


def test_invalid_utf8_line_is_isolated(tmp_path: Path) -> None:
    batch = tmp_path / "jobs.jsonl"
    batch.write_bytes(b"\xff\n{}\n")
    items = list(read_jsonl(batch))

    assert len(items) == 2
    assert items[0][0] == 1
    assert items[1] == (2, {})


def test_output_existing_directory_is_not_an_overwrite_target(tmp_path: Path) -> None:
    image = tmp_path / "a.jpg"
    image.write_bytes(b"x")
    api = AgentAPI()
    result = api.validate({
        "schema_version": 1, "title": "t", "location": "l",
        "activity_date": "2026-11-27", "images": [str(image)],
        "output": str(tmp_path / "folder.docx"),
        "options": {"if_exists": "overwrite"},
    })
    (tmp_path / "folder.docx").mkdir()
    result = api.validate({
        "schema_version": 1, "title": "t", "location": "l",
        "activity_date": "2026-11-27", "images": [str(image)],
        "output": str(tmp_path / "folder.docx"),
        "options": {"if_exists": "overwrite"},
    })

    assert result["error"]["code"] == "OUTPUT_NOT_WRITABLE"


def test_stable_public_error_codes() -> None:
    api = AgentAPI()
    base = {
        "schema_version": 1, "title": "t", "location": "l",
        "activity_date": "2026-11-27", "images": ["missing.jpg"],
        "output": "result.docx",
    }
    for field, value, code in [
        ("schema_version", 2, "INVALID_REQUEST"),
        ("activity_date", "not-date", "INVALID_REQUEST"),
        ("images", [], "INVALID_REQUEST"),
        ("images", ["x.jpg"] * 31, "TOO_MANY_IMAGES"),
        ("options", {"open": True}, "INVALID_REQUEST"),
    ]:
        payload = {**base, field: value}
        result = api.validate(payload)
        assert result["ok"] is False
        assert result["error"]["code"] == code


def test_corrupt_image_failure_is_not_hidden_by_cooperative_cancel(tmp_path: Path) -> None:
    image = tmp_path / "corrupt.jpg"
    image.write_bytes(b"not an image")
    api = AgentAPI()
    result = api.generate({
        "schema_version": 1, "title": "t", "location": "l",
        "activity_date": "2026-11-27", "images": [str(image)],
        "output": str(tmp_path / "out.docx"),
    })

    assert result["error"]["code"] == "IMAGE_READ_FAILED"
    assert not (tmp_path / "out.docx").exists()
