from __future__ import annotations

import json
import sys
import time
from pathlib import Path

from config import SUPPORTED_FORMATS
from utils.logging_utils import setup_logging
from utils.path_utils import user_data_root
from utils.performance import log_elapsed


STARTED_AT = time.perf_counter()
SERVER_NAME = "PicGenToolDesktop"
ELECTION_LOCK_NAME = "instance.lock"


def extract_image_paths(arguments: list[str]) -> list[str]:
    paths: list[str] = []
    for argument in arguments:
        if argument.startswith("-"):
            continue
        path = Path(argument)
        if path.suffix.lower() in SUPPORTED_FORMATS and path.is_file():
            paths.append(str(path))
    return paths


def notify_running_instance(paths: list[str], wait_ms: int = 350) -> bool:
    from PyQt6.QtNetwork import QLocalSocket

    socket = QLocalSocket()
    socket.connectToServer(SERVER_NAME)
    if not socket.waitForConnected(wait_ms):
        return False
    socket.write(json.dumps({"paths": paths}, ensure_ascii=False).encode("utf-8"))
    socket.flush()
    if not socket.waitForBytesWritten(500):
        return False
    acknowledged = socket.waitForReadyRead(2000) and bytes(socket.readAll()) == b"ok"
    socket.disconnectFromServer()
    return acknowledged


def forward_with_retry(paths: list[str], attempts: int = 2) -> bool:
    for attempt in range(attempts):
        if notify_running_instance(paths, 350 if attempt == 0 else 700):
            return True
    return False


def instance_lock_path() -> Path:
    directory = user_data_root()
    directory.mkdir(parents=True, exist_ok=True)
    return directory / ELECTION_LOCK_NAME


def attach_instance_server(window) -> bool:
    from PyQt6.QtNetwork import QLocalServer

    server = QLocalServer(window)
    if not server.listen(SERVER_NAME):
        QLocalServer.removeServer(SERVER_NAME)
        if not server.listen(SERVER_NAME):
            return False
    window.ipc_server = server

    def handle_connection() -> None:
        connection = server.nextPendingConnection()
        if connection is None:
            return
        receive_instance_paths(connection, window)

    server.newConnection.connect(handle_connection)
    return True


def receive_instance_paths(connection, window) -> None:
    """Receive one bounded JSON request without blocking the Qt UI thread."""
    payload = bytearray()
    completed = False

    def reject() -> None:
        if completed:
            return
        connection.disconnectFromServer()
        connection.deleteLater()

    def acknowledge(paths: list[str]) -> None:
        accepted = True
        if paths:
            accepted = bool(window.add_image_files(paths))
        else:
            window.present()
        connection.write(b"ok" if accepted else b"busy")
        connection.flush()
        if connection.bytesToWrite() > 0:
            connection.waitForBytesWritten(300)
        connection.disconnectFromServer()
        connection.deleteLater()

    def read_available() -> None:
        nonlocal completed
        if completed:
            return
        payload.extend(bytes(connection.readAll()))
        if len(payload) > 1024 * 1024:
            reject()
            return
        try:
            data = json.loads(payload.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            return
        if not isinstance(data, dict):
            reject()
            return
        values = data.get("paths", [])
        if not isinstance(values, list) or not all(isinstance(item, str) for item in values):
            reject()
            return
        connection.readyRead.disconnect(read_available)
        completed = True
        acknowledge(values)

    connection.readyRead.connect(read_available)
    connection.disconnected.connect(connection.deleteLater)
    read_available()


def run_diagnostics() -> int:
    """Exercise packaged decoders, Word output, and Qt startup without user data."""
    try:
        from datetime import date
        import os
        from tempfile import TemporaryDirectory

        from PIL import Image
        from pi_heif import register_heif_opener
        from services.diagnostics import DiagnosticsService

        register_heif_opener()
        for module in ("docxtpl", "docx", "PyQt6.QtNetwork"):
            __import__(module)
        if "HEIF" not in Image.OPEN:
            print("HEIF opener registration failed", file=sys.stderr)
            return 1
        if getattr(sys, "frozen", False):
            fixture = Path(sys._MEIPASS) / "diagnostics" / "sample.heic"
            with Image.open(fixture) as sample:
                sample.load()
                if sample.size != (320, 240) or sample.mode != "RGB":
                    print("HEIF sample decode returned unexpected pixels", file=sys.stderr)
                    return 1
            print("HEIF sample decode: passed")
        snapshot = DiagnosticsService().collect(deep_templates=True)
        print(DiagnosticsService.format_for_copy(snapshot))
        if not snapshot.template_result.valid:
            return 1

        temporary_directory = TemporaryDirectory(prefix="pgt-diagnostics-")
        temporary = Path(temporary_directory.name)
        try:
            prepared = None
            try:
                from models.generation_task import GenerationTask
                from services.document_service import DocumentService
                from services.image_service import ImageService
                from services.settings_service import SettingsService
                from ui.main_window import MainWindow, create_application
                from PyQt6.QtCore import QSettings
                from zipfile import ZipFile

                image_service = ImageService()
                if getattr(sys, "frozen", False):
                    source_paths = [fixture]
                else:
                    source = temporary / "source.jpg"
                    Image.new("RGB", (320, 240), (40, 80, 120)).save(source)
                    source_paths = [source]
                prepared = image_service.prepare_images(source_paths)
                output = temporary / "diagnostic.docx"
                task = GenerationTask(
                    title="诊断",
                    location="自检",
                    activity_date=date.today(),
                    image_paths=source_paths,
                    save_path=output,
                    open_file=False,
                )
                DocumentService().generate(task, prepared.image_paths)
                with ZipFile(output) as archive:
                    if "word/document.xml" not in archive.namelist():
                        raise ValueError("Word output is incomplete")
                print("Word one-image generation: passed")

                if getattr(sys, "frozen", False):
                    os.environ["QT_QPA_PLATFORM"] = "offscreen"
                app = create_application()
                window = MainWindow(
                    settings_service=SettingsService(
                        QSettings(
                            str(temporary / "settings.ini"),
                            QSettings.Format.IniFormat,
                        )
                    )
                )
                window.show()
                app.processEvents()
                window.close()
                app.processEvents()
                print("Offscreen GUI startup: passed")
            finally:
                if prepared is not None:
                    try:
                        prepared.cleanup()
                    except OSError:
                        pass
        finally:
            try:
                temporary_directory.cleanup()
            except OSError as exc:
                print(
                    f"Diagnostic temporary cleanup warning: {type(exc).__name__}",
                    file=sys.stderr,
                )

        print("Packaged diagnostics passed")
        return 0
    except Exception as exc:
        print(f"Packaged diagnostics failed: {type(exc).__name__}", file=sys.stderr)
        return 1


def main() -> int:
    from agent_cli import AGENT_OPTIONS, run_agent_cli

    if any(argument in AGENT_OPTIONS for argument in sys.argv[1:]):
        return run_agent_cli(sys.argv[1:])
    logger = setup_logging()
    if "--diagnose" in sys.argv[1:]:
        return run_diagnostics()
    exception_handler = None
    try:
        from services.explorer_integration import repair_if_registered

        try:
            repaired = repair_if_registered(
                Path(sys.executable),
                packaged=bool(getattr(sys, "frozen", False)),
            )
            if repaired:
                logger.info("已修复 Explorer 右键菜单命令路径")
        except OSError as exc:
            logger.warning(
                "Explorer 右键菜单自愈失败: error_type=%s",
                type(exc).__name__,
            )

        from PyQt6.QtCore import QLockFile, QTimer

        from ui.main_window import MainWindow, create_application

        app = create_application()
        from utils.exception_handler import install_exception_handlers

        exception_handler = install_exception_handlers(app, logger)
        paths = extract_image_paths(sys.argv[1:])
        if forward_with_retry(paths, attempts=1):
            return 0

        lock = QLockFile(str(instance_lock_path()))
        lock.setStaleLockTime(5000)
        locked = lock.tryLock(6000)
        try:
            if forward_with_retry(paths):
                return 0
            window = MainWindow()
            attach_instance_server(window)
        finally:
            if locked:
                lock.unlock()

        window.show()
        window.showNormal()
        if paths:
            QTimer.singleShot(0, lambda: window.add_image_files(paths))
        QTimer.singleShot(0, lambda: log_elapsed(logger, "startup_ready", STARTED_AT))
        return app.exec()
    except Exception as exc:
        if exception_handler is not None:
            exception_handler.capture(type(exc), exc, exc.__traceback__)
        else:
            logger.critical("应用启动失败: error_type=%s", type(exc).__name__)
        return 1


if __name__ == "__main__":
    sys.exit(main())
