from __future__ import annotations

from unittest.mock import Mock
from types import SimpleNamespace
import threading
import time

from PyQt6.QtCore import QMimeData, QPoint, QPointF, Qt, QUrl
from PyQt6.QtGui import QDragEnterEvent, QDragLeaveEvent, QDropEvent

from ui.main_window import MainWindow
from ui.main_window import GenerationWorker
from models.generation_task import GenerationTask
from services.image_service import ImageService
from services.template_service import TemplateService
from datetime import date
from pathlib import Path


def test_cancel_button_updates_busy_and_cancelled_states(qtbot, window_settings) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.show()
    window.worker = Mock()

    window._set_busy(True)

    assert not window.generate_button.isVisible()
    assert window.cancel_button.isVisible()
    assert window.cancel_button.isEnabled()
    assert not window.image_button.isEnabled()
    assert not window.clear_button.isEnabled()

    qtbot.mouseClick(window.cancel_button, Qt.MouseButton.LeftButton)

    window.worker.cancel.assert_called_once_with()
    assert window.status_label.text() == "正在取消"
    assert not window.cancel_button.isEnabled()

    window._handle_cancelled()
    window._cleanup_worker()

    assert window.status_label.text() == "已取消"
    assert window.generate_button.isVisible()
    assert window.generate_button.isEnabled()
    assert not window.cancel_button.isVisible()
    assert window.image_button.isEnabled()
    assert window.clear_button.isEnabled()


def test_real_thread_supports_repeat_cancel_and_close_without_qthread_warning(
    qtbot,
    qtlog,
    tmp_path: Path,
    monkeypatch,
    window_settings,
) -> None:
    task = GenerationTask(
        title="title",
        location="location",
        activity_date=date(2026, 11, 25),
        image_paths=[tmp_path / "source.jpg"],
        save_path=tmp_path / "result.docx",
        open_file=False,
    )

    def controlled_run(worker: GenerationWorker) -> None:
        while not worker.cancellation.is_cancelled():
            time.sleep(0.005)
        worker.cancelled.emit()
        worker.finished.emit()

    monkeypatch.setattr(GenerationWorker, "run", controlled_run)
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.show()
    monkeypatch.setattr(window, "_build_task", lambda: task)

    window._start_generation()
    qtbot.waitUntil(lambda: window.worker_thread is not None and window.worker_thread.isRunning())
    window._cancel_generation()
    qtbot.waitUntil(lambda: window.worker_thread is None, timeout=2000)
    assert window.status_label.text() == "已取消"

    window._start_generation()
    qtbot.waitUntil(lambda: window.worker_thread is not None and window.worker_thread.isRunning())
    window.close()
    assert window._close_requested
    qtbot.waitUntil(lambda: window.worker_thread is None, timeout=2000)
    qtbot.waitUntil(lambda: not window.isVisible(), timeout=2000)

    messages = "\n".join(record.message for record in qtlog.records)
    assert "QThread: Destroyed while thread is still running" not in messages


def test_image_import_runs_off_ui_thread_and_queues_second_batch(
    qtbot,
    tmp_path: Path,
    monkeypatch,
    window_settings,
) -> None:
    first = tmp_path / "first.jpg"
    second = tmp_path / "second.jpg"
    first.write_bytes(b"first")
    second.write_bytes(b"second")
    started = threading.Event()
    release = threading.Event()
    calls: list[tuple[list[Path], list[Path], int]] = []

    def deduplicate(_service, files, existing=None):
        paths = [Path(path) for path in files]
        calls.append((paths, list(existing or []), threading.get_ident()))
        if len(calls) == 1:
            started.set()
            assert release.wait(timeout=2)
        return SimpleNamespace(paths=paths, duplicate_count=0, failures=())

    monkeypatch.setattr(ImageService, "deduplicate", deduplicate)
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    main_thread_id = threading.get_ident()

    window._queue_image_import([str(first)])
    assert started.wait(timeout=1)
    window._queue_image_import([str(second)])
    release.set()

    qtbot.waitUntil(lambda: len(calls) == 2, timeout=3000)
    qtbot.waitUntil(lambda: window.image_paths == [first, second], timeout=3000)
    qtbot.waitUntil(lambda: window.image_import_thread is None, timeout=3000)

    assert calls[0][2] != main_thread_id
    assert calls[1][1] == [first]


def test_startup_template_failure_disables_generation_but_keeps_diagnostics_path_open(
    qtbot,
    monkeypatch,
    window_settings,
) -> None:
    invalid_result = SimpleNamespace(
        valid=False,
        checked_count=29,
        issues=(SimpleNamespace(template_number=7, code="missing_file"),),
    )
    monkeypatch.setattr(
        TemplateService,
        "validate_all",
        lambda _service, deep=False: invalid_result,
        raising=False,
    )

    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)

    qtbot.waitUntil(lambda: "模板不完整" in window.status_label.text(), timeout=1000)
    assert not window.generate_button.isEnabled()
    assert window.image_button.isEnabled()
    assert window.clear_button.isEnabled()


def test_success_clears_title_and_images_but_preserves_location_and_date(
    qtbot,
    window_settings,
) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.title_input.setText("Activity")
    window.location_input.setText("Hall")
    window.date_input.setDate(window.date_input.date().addDays(-2))
    expected_date = window.date_input.date()
    window.image_paths = [Path("photo.jpg")]
    window._show_message = Mock()

    window._handle_success({"docx_path": "result.docx"})

    assert window.title_input.text() == ""
    assert window.image_paths == []
    assert window.location_input.text() == "Hall"
    assert window.date_input.date() == expected_date
    loaded = window_settings.load()
    assert loaded.title == ""
    assert loaded.location == "Hall"


def test_explicit_clear_persists_default_title_location_and_date(
    qtbot,
    window_settings,
) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.title_input.setText("Activity")
    window.location_input.setText("Hall")
    window._confirm_clear = Mock(return_value=True)

    window._reset_form()

    loaded = window_settings.load()
    assert loaded.title == ""
    assert loaded.location == ""
    assert loaded.activity_date == date.today()


def test_drag_feedback_activates_for_supported_file_and_restores_on_leave(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(image))])
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    original_status = window.status_label.text()
    enter = QDragEnterEvent(
        QPoint(1, 1),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )

    window.dragEnterEvent(enter)

    assert enter.isAccepted()
    assert window.property("dragActive") is True
    assert window.status_label.text() == "松开即可添加图片"

    window.dragLeaveEvent(QDragLeaveEvent())

    assert window.property("dragActive") is False
    assert window.status_label.text() == original_status


def test_drop_always_clears_drag_feedback(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(image))])
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window._add_images = Mock()
    window._set_drag_active(True)
    event = QDropEvent(
        QPointF(1, 1),
        Qt.DropAction.CopyAction,
        mime,
        Qt.MouseButton.LeftButton,
        Qt.KeyboardModifier.NoModifier,
    )

    window.dropEvent(event)

    assert window.property("dragActive") is False
    dropped_files = window._add_images.call_args.args[0]
    assert [Path(path) for path in dropped_files] == [image]


def test_diagnostics_entry_remains_available_when_busy_or_templates_fail(
    qtbot,
    window_settings,
) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)

    assert window.diagnostics_button.accessibleName() == "关于和诊断"
    window._templates_valid = False
    window._set_busy(True)

    assert window.diagnostics_button.isEnabled()


def _blocking_import(window: MainWindow):
    started = threading.Event()
    release = threading.Event()
    original = window.image_import_worker.image_service.deduplicate

    def blocked(*args, **kwargs):
        started.set()
        assert release.wait(timeout=2)
        return original(*args, **kwargs)

    window.image_import_worker.image_service.deduplicate = blocked
    return started, release


def test_clear_waits_for_import_and_discards_late_result(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window._confirm_clear = Mock(return_value=True)
    window._queue_image_import([str(image)])
    started, release = _blocking_import(window)
    assert started.wait(timeout=1)

    window._reset_form()
    release.set()
    qtbot.waitUntil(lambda: window.image_import_thread is None, timeout=3000)

    assert window.image_paths == []


def test_clear_rejection_does_not_schedule_deferred_clear(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window._confirm_clear = Mock(return_value=False)
    window._queue_image_import([str(image)])
    started, release = _blocking_import(window)
    assert started.wait(timeout=1)

    window._reset_form()

    assert not window._clear_after_import
    assert not window.clear_button.isEnabled()
    release.set()
    qtbot.waitUntil(lambda: window.image_paths == [image], timeout=3000)


def test_drop_is_rejected_while_deferred_clear_is_pending(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window._confirm_clear = Mock(return_value=True)
    window._queue_image_import([str(image)])
    started, release = _blocking_import(window)
    assert started.wait(timeout=1)
    window._reset_form()
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(image))])
    event = QDropEvent(
        QPointF(1, 1), Qt.DropAction.CopyAction, mime,
        Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier,
    )

    window.dropEvent(event)

    assert not event.isAccepted()
    assert not window._image_import_queue
    release.set()
    qtbot.waitUntil(lambda: window.image_import_thread is None, timeout=3000)
    assert window.status_label.text() == "请选择照片"
    assert window.image_button.isEnabled()


def test_ipc_during_save_dialog_cannot_start_import_or_generation_race(
    qtbot,
    tmp_path: Path,
    window_settings,
    monkeypatch,
) -> None:
    source = tmp_path / "source.jpg"
    source.write_bytes(b"source")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.title_input.setText("Activity")
    window.location_input.setText("Hall")
    window.image_paths = [source]
    window._show_message = Mock()
    import_attempts: list[list[str]] = []
    original_queue = window._queue_image_import
    window._queue_image_import = lambda files: import_attempts.append(files) or original_queue(files)

    def save_dialog(*_args, **_kwargs):
        window.add_image_files([str(tmp_path / "late.heic")])
        assert window._generation_starting
        assert not window.add_image_files([str(tmp_path / "late.heic")])
        return str(tmp_path / "result.docx"), "Word 文档 (*.docx)"

    monkeypatch.setattr("ui.main_window.QFileDialog.getSaveFileName", save_dialog)
    from ui.main_window import GenerationWorker

    def immediate_run(worker):
        worker.finished.emit()

    monkeypatch.setattr(GenerationWorker, "run", immediate_run)

    window._start_generation()

    assert import_attempts == []
    assert window.image_paths == [source]
    assert not window._generation_starting
    assert window.image_import_thread is None
    assert window.worker_thread is not None
    assert window._generation_starting is False
    qtbot.waitUntil(lambda: window.worker_thread is None, timeout=3000)


def test_generation_rejected_while_import_is_active(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window._queue_image_import([str(image)])
    started, release = _blocking_import(window)
    assert started.wait(timeout=1)
    window._build_task = Mock()

    window._start_generation()

    assert window.worker_thread is None
    assert window.worker is None
    assert "请先完成当前操作" in window.status_label.text()
    release.set()
    qtbot.waitUntil(lambda: window.image_import_thread is None, timeout=3000)


def test_drop_ignored_while_generation_is_busy(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    mime = QMimeData()
    mime.setUrls([QUrl.fromLocalFile(str(image))])
    event = QDropEvent(
        QPointF(1, 1), Qt.DropAction.CopyAction, mime,
        Qt.MouseButton.LeftButton, Qt.KeyboardModifier.NoModifier,
    )
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.worker_thread = Mock()
    window.worker_thread.isRunning.return_value = True

    window.dropEvent(event)

    assert window.image_import_thread is None
    assert window.image_paths == []


def test_final_import_feedback_remains_visible_after_thread_finishes(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    image = tmp_path / "photo.jpg"
    image.write_bytes(b"image")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)

    window._queue_image_import([str(image), str(image)])

    qtbot.waitUntil(lambda: window.image_import_thread is None, timeout=3000)
    assert "重复" in window.status_label.text()

