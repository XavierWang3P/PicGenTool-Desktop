from __future__ import annotations

import sys
from pathlib import Path

import pytest
from PIL import Image
from PyQt6.QtCore import Qt

from main import extract_image_paths
from ui.main_window import MainWindow


def test_extract_image_paths_keeps_supported_files_only(tmp_path: Path) -> None:
    image = tmp_path / "photo.jpg"
    Image.new("RGB", (20, 20), (10, 20, 30)).save(image)
    text = tmp_path / "notes.txt"
    text.write_text("x", encoding="utf-8")
    missing = tmp_path / "missing.png"

    result = extract_image_paths(
        ["--add-images", str(missing), "flagless", str(text), str(image)]
    )

    assert result == [str(image)]


def test_extract_image_paths_accepts_existing_heic_and_heif_files(tmp_path: Path) -> None:
    heic = tmp_path / "photo.heic"
    heif = tmp_path / "photo.heif"
    heic.write_bytes(b"heic-placeholder")
    heif.write_bytes(b"heif-placeholder")

    assert extract_image_paths([str(heic), str(heif)]) == [str(heic), str(heif)]


def test_window_is_always_on_top_by_default(qtbot, window_settings) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    assert window.windowFlags() & Qt.WindowType.WindowStaysOnTopHint
    assert window.always_on_top_checkbox.isChecked()


def test_always_on_top_checkbox_removes_flag(qtbot, window_settings) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window._apply_always_on_top(False)
    assert not (window.windowFlags() & Qt.WindowType.WindowStaysOnTopHint)


def test_add_image_files_updates_count_and_status(qtbot, tmp_path: Path, window_settings) -> None:
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    first = tmp_path / "a.png"
    second = tmp_path / "b.png"
    for path, color in ((first, (30, 40, 50)), (second, (60, 70, 80))):
        Image.new("RGB", (20, 20), color).save(path)

    window.add_image_files([str(first), str(second)])

    qtbot.waitUntil(lambda: len(window.image_paths) == 2, timeout=3000)
    qtbot.waitUntil(lambda: window.image_import_thread is None, timeout=3000)
    assert window.image_count_label.text() == "已选择 2 / 30 张图片"
    assert window.status_label.text() in {"已添加 2 张", "图片已就绪"}


def test_add_image_files_ignored_while_generating(qtbot, window_settings) -> None:
    from unittest.mock import Mock

    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.worker_thread = Mock()
    window.worker_thread.isRunning.return_value = True

    assert not window.add_image_files(["C:/nonexistent.jpg"])

    assert window.image_paths == []


def test_register_script_builds_verb_path() -> None:
    pytest.importorskip("winreg")
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
    from tools.register_explorer_menu import _verb_key

    assert _verb_key(".jpg").endswith(r"\SystemFileAssociations\.jpg\shell\PicGenToolAddImages")


def test_election_lock_serializes_primary_start(tmp_path: Path) -> None:
    from PyQt6.QtCore import QLockFile

    lock_path = tmp_path / "instance.lock"
    first = QLockFile(str(lock_path))
    second = QLockFile(str(lock_path))

    assert first.tryLock(200)
    assert not second.tryLock(100)

    first.unlock()
    assert second.tryLock(200)
    second.unlock()


def test_ipc_accepts_fragmented_json_without_blocking_ui(
    qtbot,
    tmp_path: Path,
    window_settings,
) -> None:
    import json
    from unittest.mock import Mock
    from PyQt6.QtCore import QTimer
    from PyQt6.QtNetwork import QLocalServer, QLocalSocket
    from main import attach_instance_server

    QLocalServer.removeServer("PicGenToolDesktop")
    window = MainWindow(settings_service=window_settings)
    qtbot.addWidget(window)
    window.add_image_files = Mock()
    assert attach_instance_server(window)
    client = QLocalSocket()
    client.connectToServer("PicGenToolDesktop")
    assert client.waitForConnected(500)
    payload = json.dumps({"paths": ["photo.heic"]}).encode("utf-8")
    acknowledgements: list[bytes] = []
    client.readyRead.connect(lambda: acknowledgements.append(bytes(client.readAll())))
    midpoint = len(payload) // 2
    client.write(payload[:midpoint])
    client.flush()
    QTimer.singleShot(20, lambda: client.write(payload[midpoint:]))
    QTimer.singleShot(40, client.flush)

    qtbot.waitUntil(lambda: window.add_image_files.call_count == 1, timeout=2000)
    window.add_image_files.assert_called_once_with(["photo.heic"])
    qtbot.waitUntil(lambda: acknowledgements == [b"ok"], timeout=500)
    client.disconnectFromServer()
    window.ipc_server.close()
