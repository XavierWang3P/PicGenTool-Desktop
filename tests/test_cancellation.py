from __future__ import annotations

import threading
import time
from datetime import date
from pathlib import Path
from unittest.mock import Mock

import pytest

from models.document_layout import WordExportPayload
from models.generation_task import GenerationTask
from services.cancellation import CancellationToken, GenerationCancelled
from services.document_service import DocumentService
from services.image_service import ImageService


def test_cancellation_token_is_idempotent_and_raises() -> None:
    token = CancellationToken(threading.Event())
    assert not token.is_cancelled()

    token.cancel()
    token.cancel()

    assert token.is_cancelled()
    with pytest.raises(GenerationCancelled, match="已取消"):
        token.raise_if_cancelled()


def test_prepare_images_cancelled_before_start_does_not_process(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    token = CancellationToken()
    token.cancel()
    service = ImageService()
    prepare_one = Mock()
    monkeypatch.setattr(service, "_prepare_one", prepare_one)

    with pytest.raises(GenerationCancelled):
        service.prepare_images([Path("photo.jpg")], cancellation=token)

    prepare_one.assert_not_called()


def test_prepare_images_cancelled_during_processing_cleans_temp_directory(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    token = CancellationToken()
    service = ImageService()
    sources = [tmp_path / f"source-{index}.jpg" for index in range(3)]
    output_directories: list[Path] = []

    monkeypatch.setattr(service, "_worker_count", lambda _count: 1)

    def prepare(source, output, _target_size, _target_ratio):
        output_directories.append(output.parent)
        output.write_text(source.name, encoding="utf-8")
        token.cancel()
        return output

    monkeypatch.setattr(service, "_prepare_one", prepare)

    with pytest.raises(GenerationCancelled):
        service.prepare_images(sources, cancellation=token)

    assert output_directories
    assert all(not directory.exists() for directory in output_directories)


def test_worker_failure_cancels_running_sibling_and_leaves_no_threads(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    token = CancellationToken()
    service = ImageService()
    sources = [tmp_path / "broken.jpg", tmp_path / "sibling.jpg"]
    sibling_started = threading.Event()
    sibling_observed_cancellation = threading.Event()

    monkeypatch.setattr(service, "_worker_count", lambda _count: 2)

    def prepare(source, output, _target_size, _target_ratio):
        if source.name == "broken.jpg":
            assert sibling_started.wait(timeout=1)
            raise OSError("decode failed")

        sibling_started.set()
        deadline = time.monotonic() + 0.5
        while time.monotonic() < deadline:
            if token.is_cancelled():
                sibling_observed_cancellation.set()
                break
            time.sleep(0.005)
        output.write_bytes(b"done")
        return output

    monkeypatch.setattr(service, "_prepare_one", prepare)

    with pytest.raises(ValueError, match="broken.jpg"):
        service.prepare_images(sources, cancellation=token)

    assert token.is_cancelled()
    assert sibling_observed_cancellation.is_set()
    assert not any(thread.name.startswith("pgt-image") for thread in threading.enumerate())


def test_commit_linearizes_late_cancellation() -> None:
    token = CancellationToken()
    commit_started = threading.Event()
    allow_commit = threading.Event()

    def action() -> str:
        commit_started.set()
        assert allow_commit.wait(timeout=1)
        return "committed"

    result: list[str] = []
    commit_thread = threading.Thread(target=lambda: result.append(token.commit(action)))
    commit_thread.start()
    assert commit_started.wait(timeout=1)

    cancel_thread = threading.Thread(target=token.cancel)
    cancel_thread.start()
    time.sleep(0.02)
    cancel_was_blocked = cancel_thread.is_alive()
    allow_commit.set()
    commit_thread.join(timeout=1)
    cancel_thread.join(timeout=1)

    assert cancel_was_blocked
    assert result == ["committed"]
    assert not token.is_cancelled()


def _task(tmp_path: Path) -> GenerationTask:
    return GenerationTask(
        title="title",
        location="location",
        activity_date=date(2026, 11, 25),
        image_paths=[tmp_path / "source.jpg"],
        save_path=tmp_path / "result.docx",
        open_file=False,
    )


def test_document_service_checks_cancellation_before_layout(tmp_path: Path) -> None:
    token = CancellationToken()
    token.cancel()
    layout_service = Mock()
    word_exporter = Mock()
    service = DocumentService(layout_service=layout_service, word_exporter=word_exporter)

    with pytest.raises(GenerationCancelled):
        service.generate(_task(tmp_path), [tmp_path / "prepared.jpg"], cancellation=token)

    layout_service.build_word_payload.assert_not_called()
    word_exporter.export.assert_not_called()


def test_document_service_checks_cancellation_before_export(tmp_path: Path) -> None:
    token = CancellationToken()
    payload = WordExportPayload(
        template_path=tmp_path / "template.docx",
        text_context={},
        image_paths=[tmp_path / "prepared.jpg"],
        output_path=tmp_path / "result.docx",
    )
    layout_service = Mock()

    def build_payload(*_args):
        token.cancel()
        return payload

    layout_service.build_word_payload.side_effect = build_payload
    word_exporter = Mock()
    service = DocumentService(layout_service=layout_service, word_exporter=word_exporter)

    with pytest.raises(GenerationCancelled):
        service.generate(_task(tmp_path), payload.image_paths, cancellation=token)

    word_exporter.export.assert_not_called()
