from __future__ import annotations

import logging
from pathlib import Path

import pytest


def test_measure_logs_stage_elapsed_and_safe_fields(
    caplog: pytest.LogCaptureFixture,
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    from utils import performance

    readings = iter([10.0, 10.125])
    monkeypatch.setattr(performance.time, "perf_counter", lambda: next(readings))
    logger = logging.getLogger("test.performance")

    with caplog.at_level(logging.INFO, logger=logger.name):
        with performance.measure(
            logger,
            "image_total",
            image_count=10,
            image_path=Path("C:/private/photo.jpg"),
        ):
            pass

    assert "stage=image_total" in caplog.text
    assert "elapsed_ms=125.0" in caplog.text
    assert "image_count=10" in caplog.text
    assert "image_path" not in caplog.text
    assert "private" not in caplog.text


def test_measure_logs_and_reraises_exceptions(
    caplog: pytest.LogCaptureFixture,
) -> None:
    from utils.performance import measure

    logger = logging.getLogger("test.performance.exception")
    with caplog.at_level(logging.INFO, logger=logger.name):
        with pytest.raises(RuntimeError, match="boom"):
            with measure(logger, "word_render"):
                raise RuntimeError("boom")

    assert "stage=word_render" in caplog.text
