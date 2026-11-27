from __future__ import annotations

from datetime import date
from pathlib import Path
from types import SimpleNamespace
import threading
import time
from unittest.mock import Mock

from PyQt6.QtTest import QSignalSpy
import pytest

from models.generation_task import GenerationTask
from services.cancellation import GenerationCancelled
from ui.main_window import GenerationWorker


def test_worker_cancel_emits_cancelled_not_error_and_finishes_once(tmp_path: Path) -> None:
    task = GenerationTask(
        title="title",
        location="location",
        activity_date=date(2026, 11, 25),
        image_paths=[tmp_path / "source.jpg"],
        save_path=tmp_path / "result.docx",
        open_file=False,
    )
    worker = GenerationWorker(task)
    cancelled = QSignalSpy(worker.cancelled)
    errors = QSignalSpy(worker.error)
    finished = QSignalSpy(worker.finished)

    worker.cancel()
    worker.cancel()
    worker.run()

    assert len(cancelled) == 1
    assert len(errors) == 0
    assert len(finished) == 1


def test_worker_does_not_log_or_emit_private_image_directory(
    tmp_path: Path,
    caplog: pytest.LogCaptureFixture,
) -> None:
    private_directory = tmp_path / "private-customer-album"
    private_directory.mkdir()
    corrupt_path = private_directory / "corrupt.jpg"
    corrupt_path.write_bytes(b"invalid")
    task = GenerationTask(
        title="title",
        location="location",
        activity_date=date(2026, 11, 25),
        image_paths=[corrupt_path],
        save_path=tmp_path / "result.docx",
        open_file=False,
    )
    worker = GenerationWorker(task)
    worker.image_service.prepare_images = Mock(
        side_effect=ValueError(f"decoder failed for {corrupt_path}")
    )
    errors = QSignalSpy(worker.error)

    with caplog.at_level("ERROR", logger="pgt.worker"):
        worker.run()

    assert len(errors) == 1
    assert str(private_directory) not in str(errors[0][0])
    assert str(private_directory) not in caplog.text


def _task(tmp_path: Path) -> GenerationTask:
    return GenerationTask(
        title="title",
        location="location",
        activity_date=date(2026, 11, 25),
        image_paths=[tmp_path / "source.jpg"],
        save_path=tmp_path / "result.docx",
        open_file=False,
    )


def test_worker_opens_template_while_images_are_still_preparing(tmp_path: Path) -> None:
    worker = GenerationWorker(_task(tmp_path))
    image_started = threading.Event()
    template_opened = threading.Event()
    image_completed = threading.Event()
    prepared_images = SimpleNamespace(
        image_paths=[tmp_path / "prepared.jpg"],
        cleanup=Mock(),
    )

    def prepare_images(*_args, **_kwargs):
        image_started.set()
        assert template_opened.wait(timeout=1)
        image_completed.set()
        return prepared_images

    def prepare_word(*_args, **_kwargs):
        assert image_started.wait(timeout=1)
        assert not image_completed.is_set()
        template_opened.set()
        return SimpleNamespace()

    worker.image_service.prepare_images = prepare_images
    worker.document_service.prepare = Mock(side_effect=prepare_word)
    worker.document_service.export_prepared = Mock(return_value=worker.task.save_path)

    worker.run()

    worker.document_service.export_prepared.assert_called_once()
    prepared_images.cleanup.assert_called_once_with()


def test_template_prepare_failure_cancels_and_waits_for_image_future(
    tmp_path: Path,
) -> None:
    worker = GenerationWorker(_task(tmp_path))
    image_exited = threading.Event()

    def prepare_images(*_args, cancellation=None, **_kwargs):
        try:
            while not cancellation.is_cancelled():
                time.sleep(0.005)
            raise GenerationCancelled()
        finally:
            image_exited.set()

    worker.image_service.prepare_images = prepare_images
    worker.document_service.prepare = Mock(side_effect=OSError("template broken"))
    worker.document_service.export_prepared = Mock()
    errors = QSignalSpy(worker.error)
    finished = QSignalSpy(worker.finished)

    worker.run()

    assert image_exited.is_set()
    assert len(errors) == 1
    assert len(finished) == 1
    worker.document_service.export_prepared.assert_not_called()


def test_image_failure_does_not_render_or_leave_prepared_images(tmp_path: Path) -> None:
    worker = GenerationWorker(_task(tmp_path))
    worker.image_service.prepare_images = Mock(side_effect=OSError("image broken"))
    worker.document_service.prepare = Mock(return_value=SimpleNamespace())
    worker.document_service.export_prepared = Mock()
    errors = QSignalSpy(worker.error)

    worker.run()

    assert len(errors) == 1
    worker.document_service.export_prepared.assert_not_called()


def test_cancel_during_overlap_emits_cancelled_once(tmp_path: Path) -> None:
    worker = GenerationWorker(_task(tmp_path))
    image_started = threading.Event()

    def prepare_images(*_args, cancellation=None, **_kwargs):
        image_started.set()
        while not cancellation.is_cancelled():
            time.sleep(0.005)
        raise GenerationCancelled()

    def prepare_word(*_args, **_kwargs):
        assert image_started.wait(timeout=1)
        worker.cancel()
        return SimpleNamespace()

    worker.image_service.prepare_images = prepare_images
    worker.document_service.prepare = Mock(side_effect=prepare_word)
    worker.document_service.export_prepared = Mock()
    cancelled = QSignalSpy(worker.cancelled)
    errors = QSignalSpy(worker.error)
    finished = QSignalSpy(worker.finished)

    worker.run()

    assert len(cancelled) == 1
    assert len(errors) == 0
    assert len(finished) == 1
    worker.document_service.export_prepared.assert_not_called()


def test_cleanup_failure_still_finishes_and_closes_worker_thread(
    tmp_path: Path,
) -> None:
    worker = GenerationWorker(_task(tmp_path))
    prepared = SimpleNamespace(
        image_paths=[tmp_path / "prepared.jpg"],
        cleanup=Mock(side_effect=PermissionError(r"C:\private\locked")),
    )
    worker.image_service.prepare_images = Mock(return_value=prepared)
    worker.document_service.prepare = Mock(return_value=SimpleNamespace())
    worker.document_service.export_prepared = Mock(return_value=worker.task.save_path)
    finished = QSignalSpy(worker.finished)
    success = QSignalSpy(worker.success)

    worker.run()

    assert len(success) == 1
    assert len(finished) == 1
    prepared.cleanup.assert_called_once_with()


def test_worker_progress_is_monotonic_across_image_and_word_phases(
    tmp_path: Path,
) -> None:
    worker = GenerationWorker(_task(tmp_path))
    values = QSignalSpy(worker.progress)

    worker._emit_progress("正在准备 Word 模板", 60)
    worker._emit_progress("已完成图片 1/2", 30)
    worker._emit_progress("正在导出 Word 文档", 82)

    assert [entry[1] for entry in values] == [60, 60, 82]
