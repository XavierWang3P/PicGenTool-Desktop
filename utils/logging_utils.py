from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
import tempfile
from pathlib import Path

from config import LOG_BACKUP_COUNT, LOG_FILE_NAME, LOG_LEVEL, LOG_MAX_BYTES
from utils.path_utils import ensure_directory, user_data_root


_CURRENT_LOG_PATH: Path | None = None


def _create_file_handler() -> RotatingFileHandler | None:
    global _CURRENT_LOG_PATH

    candidates = [
        user_data_root() / "logs",
        Path(tempfile.gettempdir()) / "PicGenTool" / "logs",
    ]
    for directory in candidates:
        try:
            log_dir = ensure_directory(directory)
            log_path = log_dir / LOG_FILE_NAME
            handler = RotatingFileHandler(
                log_path,
                maxBytes=LOG_MAX_BYTES,
                backupCount=LOG_BACKUP_COUNT,
                encoding="utf-8",
            )
            _CURRENT_LOG_PATH = log_path
            return handler
        except OSError:
            continue
    return None


def current_log_path() -> Path | None:
    return _CURRENT_LOG_PATH


def setup_logging() -> logging.Logger:
    root_logger = logging.getLogger()
    if any(getattr(handler, "_pgt_handler", False) for handler in root_logger.handlers):
        return logging.getLogger("pgt")

    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )
    handlers: list[logging.Handler] = [logging.StreamHandler()]
    file_handler = _create_file_handler()
    if file_handler is not None:
        handlers.insert(0, file_handler)

    for handler in handlers:
        handler.setFormatter(formatter)
        handler._pgt_handler = True  # type: ignore[attr-defined]
        root_logger.addHandler(handler)
    root_logger.setLevel(LOG_LEVEL)
    return logging.getLogger("pgt")
