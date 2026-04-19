from __future__ import annotations

import logging

from config import LOG_FILE_NAME, LOG_LEVEL
from utils.path_utils import ensure_directory, project_root


def setup_logging() -> logging.Logger:
    log_dir = ensure_directory(project_root() / "logs")
    log_file = log_dir / LOG_FILE_NAME

    logging.basicConfig(
        level=LOG_LEVEL,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(),
        ],
    )
    return logging.getLogger("pgt")

