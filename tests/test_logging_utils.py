from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path

import pytest

from utils import logging_utils


def test_file_handler_falls_back_when_user_log_directory_is_unwritable(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    primary_root = tmp_path / "primary"
    fallback_root = tmp_path / "fallback"
    real_file_handler = RotatingFileHandler

    monkeypatch.setattr(logging_utils, "user_data_root", lambda: primary_root)
    monkeypatch.setattr(logging_utils.tempfile, "gettempdir", lambda: str(fallback_root))

    def file_handler(path, *args, **kwargs):
        if primary_root in Path(path).parents:
            raise PermissionError("read only install")
        return real_file_handler(path, *args, **kwargs)

    monkeypatch.setattr(logging_utils, "RotatingFileHandler", file_handler)
    handler = logging_utils._create_file_handler()
    try:
        assert handler is not None
        assert Path(handler.baseFilename).is_relative_to(fallback_root)
    finally:
        if handler is not None:
            handler.close()


def test_file_handler_uses_bounded_rotation_configuration(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(logging_utils, "user_data_root", lambda: tmp_path)

    handler = logging_utils._create_file_handler()
    try:
        assert isinstance(handler, RotatingFileHandler)
        assert handler.maxBytes == 5 * 1024 * 1024
        assert handler.backupCount == 3
    finally:
        if handler is not None:
            handler.close()


def test_rotating_handler_creates_no_more_than_configured_backups(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr(logging_utils, "user_data_root", lambda: tmp_path)
    monkeypatch.setattr(logging_utils, "LOG_MAX_BYTES", 80)
    monkeypatch.setattr(logging_utils, "LOG_BACKUP_COUNT", 3)
    handler = logging_utils._create_file_handler()
    assert handler is not None
    handler.setFormatter(logging.Formatter("%(message)s"))
    logger = logging.getLogger("test.rotation")
    logger.handlers = [handler]
    logger.propagate = False
    logger.setLevel(logging.INFO)
    try:
        for index in range(30):
            logger.info("message-%02d-%s", index, "x" * 40)
        handler.flush()
        backups = sorted((tmp_path / "logs").glob("app.log.*"))
        assert 1 <= len(backups) <= 3
        assert logging_utils.current_log_path() == tmp_path / "logs" / "app.log"
    finally:
        handler.close()
        logger.handlers.clear()
