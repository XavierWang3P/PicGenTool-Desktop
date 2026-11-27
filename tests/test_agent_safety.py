from __future__ import annotations

import threading
import time
from datetime import date
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from models.generation_task import GenerationTask
from services.agent_api import AgentAPI
from services.generation_service import GenerationService


def task(tmp_path: Path) -> GenerationTask:
    return GenerationTask("title", "location", date.today(), [], tmp_path / "out.docx", False)


def test_template_failure_cleans_successful_prepared_images(tmp_path: Path) -> None:
    prepared = SimpleNamespace(image_paths=[], cleanup=Mock())
    service = GenerationService()
    service.image_service.prepare_images = Mock(return_value=prepared)
    service.document_service.prepare = Mock(side_effect=OSError("template failed"))

    with pytest.raises(OSError):
        service.run(task(tmp_path))

    prepared.cleanup.assert_called_once_with()


def test_image_failure_during_template_error_remains_original(tmp_path: Path) -> None:
    service = GenerationService()
    service.image_service.prepare_images = Mock(side_effect=ValueError("image failed"))
    service.document_service.prepare = Mock(side_effect=OSError("template failed"))

    with pytest.raises(ValueError, match="image failed"):
        service.run(task(tmp_path))


def test_concurrent_fail_policy_never_overwrites_same_output(tmp_path: Path) -> None:
    from PIL import Image
    from zipfile import ZipFile

    image = tmp_path / "image.jpg"
    Image.new("RGB", (32, 24), (40, 80, 120)).save(image)
    output = tmp_path / "same.docx"
    payload = {
        "schema_version": 1, "title": "t", "location": "l",
        "activity_date": "2026-11-27", "images": [str(image)],
        "output": str(output),
    }
    api = AgentAPI()
    with ThreadPoolExecutorForTest() as pool:
        results = list(pool.map(lambda _: api.generate(payload), range(2)))

    assert sum(result["ok"] for result in results) == 1
    assert any(
        result.get("error", {}).get("code") == "OUTPUT_EXISTS" for result in results
    )
    with ZipFile(output) as archive:
        assert "word/document.xml" in archive.namelist()


class ThreadPoolExecutorForTest:
    def __enter__(self):
        from concurrent.futures import ThreadPoolExecutor
        self.pool = ThreadPoolExecutor(max_workers=2)
        return self.pool

    def __exit__(self, *args):
        self.pool.shutdown(wait=True)


def test_corrupt_template_has_template_error_and_cleans_output(tmp_path: Path) -> None:
    from PIL import Image
    from services.template_service import TemplateService

    image = tmp_path / "photo.jpg"
    Image.new("RGB", (32, 24)).save(image)
    template = tmp_path / "Template-1.docx"
    template.write_bytes(b"broken")
    api = AgentAPI(template_service=TemplateService(tmp_path))
    result = api.generate({
        "schema_version": 1, "title": "t", "location": "l",
        "activity_date": "2026-11-27", "images": [str(image)],
        "output": str(tmp_path / "out.docx"),
    })

    assert result["error"]["code"] == "TEMPLATE_MISSING"
    assert not (tmp_path / "out.docx").exists()


def test_interrupt_cancels_all_active_tokens_outside_future_wait() -> None:
    from services.agent_runner import run_batch

    api = Mock(spec=AgentAPI)
    worker_started = threading.Event()
    tokens: list[object] = []

    def generate(payload, cancellation=None):
        tokens.append(cancellation)
        worker_started.set()
        while not cancellation.is_cancelled():
            time.sleep(0.002)
        return {"ok": False, "error": {"code": "CANCELLED"}}

    api.generate.side_effect = generate

    def inputs():
        yield (1, {})
        assert worker_started.wait(timeout=2)
        raise KeyboardInterrupt

    with pytest.raises(KeyboardInterrupt):
        list(run_batch(inputs(), api, workers=2))

    assert tokens
    assert all(token.is_cancelled() for token in tokens)


def test_container_options_are_invalid_requests(tmp_path: Path) -> None:
    for value in ([], {}, 1, True):
        result = AgentAPI().validate({
            "schema_version": 1, "title": "t", "location": "l",
            "activity_date": "2026-11-27", "images": ["missing.jpg"],
            "output": "out.docx", "options": {"if_exists": value},
        })
        assert result["error"]["code"] == "INVALID_REQUEST"
