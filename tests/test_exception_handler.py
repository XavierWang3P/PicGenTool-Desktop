from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

from PyQt6.QtTest import QSignalSpy

from utils.exception_handler import GlobalExceptionHandler


def _captured_exception(message: str):
    try:
        raise RuntimeError(message)
    except RuntimeError:
        return sys.exc_info()


def test_main_exception_logs_sanitized_traceback_and_emits_generic_message(
    tmp_path: Path,
) -> None:
    logger = Mock()
    fallback = Mock()
    handler = GlobalExceptionHandler(logger, original_sys_hook=fallback)
    private_path = tmp_path / "private-customer" / "photo.jpg"
    handler.set_sensitive_values([private_path, "secret activity title"])
    notification = QSignalSpy(handler.notification)
    exc_type, exc_value, traceback = _captured_exception(
        f"failed for {private_path}: secret activity title"
    )

    handler.capture(exc_type, exc_value, traceback)

    logged = " ".join(str(value) for value in logger.critical.call_args.args)
    assert str(private_path) not in logged
    assert "secret activity title" not in logged
    assert "<redacted>" in logged
    assert len(notification) == 1
    assert str(private_path) not in str(notification[0][0])
    fallback.assert_not_called()


def test_thread_exception_uses_same_sanitization(tmp_path: Path) -> None:
    logger = Mock()
    fallback = Mock()
    handler = GlobalExceptionHandler(logger, original_thread_hook=fallback)
    private_path = tmp_path / "private" / "photo.jpg"
    handler.set_sensitive_values([private_path])
    exc_type, exc_value, traceback = _captured_exception(str(private_path))
    args = SimpleNamespace(
        exc_type=exc_type,
        exc_value=exc_value,
        exc_traceback=traceback,
        thread=SimpleNamespace(name="worker"),
    )

    handler.capture_thread(args)

    logged = " ".join(str(value) for value in logger.critical.call_args.args)
    assert str(private_path) not in logged
    assert "<redacted>" in logged
    fallback.assert_not_called()


def test_handler_failure_falls_back_once_without_recursion() -> None:
    logger = Mock()
    logger.critical.side_effect = RuntimeError("logger failed")
    fallback = Mock()
    handler = GlobalExceptionHandler(logger, original_sys_hook=fallback)
    exc_type, exc_value, traceback = _captured_exception("boom")

    handler.capture(exc_type, exc_value, traceback)

    fallback.assert_called_once_with(exc_type, exc_value, traceback)


def test_clear_sensitive_values_stops_replacing_prior_task_value() -> None:
    logger = Mock()
    handler = GlobalExceptionHandler(logger)
    handler.set_sensitive_values(["old title"])
    handler.clear_sensitive_values()
    exc_type, exc_value, traceback = _captured_exception("old title")

    handler.capture(exc_type, exc_value, traceback)

    logged = " ".join(str(value) for value in logger.critical.call_args.args)
    assert "old title" in logged
