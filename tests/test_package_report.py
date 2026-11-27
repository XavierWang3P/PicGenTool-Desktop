from __future__ import annotations

import json
from pathlib import Path

import pytest

from tools.package_report import build_report, scan_distribution, write_report


def test_scan_distribution_counts_groups_and_largest_files(tmp_path: Path) -> None:
    (tmp_path / "_internal" / "Qt").mkdir(parents=True)
    (tmp_path / "templates").mkdir()
    (tmp_path / "PGT.exe").write_bytes(b"e" * 30)
    (tmp_path / "_internal" / "Qt" / "core.dll").write_bytes(b"c" * 20)
    (tmp_path / "templates" / "Template-1.docx").write_bytes(b"t" * 10)

    report = scan_distribution(tmp_path)

    assert report.root == tmp_path
    assert report.file_count == 3
    assert report.total_bytes == 60
    assert report.largest_files[0] == ("PGT.exe", 30)
    assert report.groups["_internal"] == 20
    assert report.groups["templates"] == 10


def test_build_report_enforces_budget_without_losing_measurements(tmp_path: Path) -> None:
    (tmp_path / "PGT.exe").write_bytes(b"x" * 10)

    within = build_report(tmp_path, max_bytes=10)
    over = build_report(tmp_path, max_bytes=9)

    assert within.within_budget
    assert not over.within_budget
    assert over.total_bytes == 10
    assert over.max_bytes == 9


def test_write_report_emits_stable_json_even_when_over_budget(tmp_path: Path) -> None:
    dist = tmp_path / "dist"
    dist.mkdir()
    (dist / "PGT.exe").write_bytes(b"x" * 10)
    destination = tmp_path / "report.json"

    write_report(build_report(dist, max_bytes=9), destination)
    data = json.loads(destination.read_text(encoding="utf-8"))

    assert data["file_count"] == 1
    assert data["total_bytes"] == 10
    assert data["total_mib"] == pytest.approx(10 / 1024 / 1024)
    assert data["within_budget"] is False
    assert data["largest_files"] == [{"path": "PGT.exe", "bytes": 10}]
