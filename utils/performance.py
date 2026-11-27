from __future__ import annotations

import logging
import time
from contextlib import contextmanager
from typing import Iterator


_MAX_FIELD_LENGTH = 80
_SENSITIVE_FIELD_PARTS = ("path", "title", "location")


def _safe_fields(fields: dict[str, object]) -> str:
    parts: list[str] = []
    for name, value in fields.items():
        if any(part in name.lower() for part in _SENSITIVE_FIELD_PARTS):
            continue
        if isinstance(value, int):
            parts.append(f"{name}={value}")
        elif isinstance(value, str) and len(value) <= _MAX_FIELD_LENGTH:
            parts.append(f"{name}={value}")
    return " ".join(parts)


def log_elapsed(
    logger: logging.Logger,
    stage: str,
    started_at: float,
    **fields: int | str,
) -> None:
    elapsed_ms = (time.perf_counter() - started_at) * 1000
    suffix = _safe_fields(fields)
    message = f"stage={stage} elapsed_ms={elapsed_ms:.1f}"
    if suffix:
        message = f"{message} {suffix}"
    logger.info(message)


@contextmanager
def measure(
    logger: logging.Logger,
    stage: str,
    **fields: int | str,
) -> Iterator[None]:
    started_at = time.perf_counter()
    try:
        yield
    finally:
        log_elapsed(logger, stage, started_at, **fields)
