from __future__ import annotations

from pathlib import Path
import re


WORKFLOW = Path(__file__).resolve().parents[1] / ".github" / "workflows" / "compile.yml"


def test_workflow_runs_test_matrix_before_build() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")

    assert re.search(r"(?m)^  test:\s*$", text)
    assert "matrix:" in text
    assert re.search(r"python-version:\s*\[\s*[\"']?3\.11[\"']?\s*,\s*[\"']?3\.12", text)
    assert "pip install -r requirements.txt" in text
    assert "pip install -r requirements-dev.txt" in text
    assert "python -m compileall main.py app.py config.py ui services models utils tools tests" in text
    assert "python -m pytest -q" in text
    assert "python tools/check_templates.py" in text
    assert re.search(r"(?ms)^  build:\s*\n\s+needs:\s*test\s*$", text)


def test_test_job_has_read_only_contents_permission() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")
    test_job = text.split("  test:", 1)[1].split("\n  build:", 1)[0]

    assert re.search(r"(?ms)permissions:\s*\n\s+contents:\s*read", test_job)


def test_workflow_checks_package_size_and_uploads_report() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")
    build_job = text.split("  build:", 1)[1]

    assert "python tools/package_report.py dist/PGT --max-mib 130 --json package-report.json" in text
    assert "./dist/PGT/PGT.exe --diagnose" in build_job
    assert "actions/upload-artifact@" in build_job
    assert "path: package-report.json" in build_job


def test_workflow_signs_only_main_exe_before_archive_and_skips_without_secrets() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")
    build_job = text.split("  build:", 1)[1]
    sign_position = build_job.index("tools/sign_release.ps1")
    archive_position = build_job.index("Compress-Archive")

    assert "-TargetPath dist/PGT/PGT.exe" in build_job
    assert sign_position < archive_position
    assert "if: env.WINDOWS_CERTIFICATE_BASE64 != '' && env.WINDOWS_CERTIFICATE_PASSWORD != ''" in build_job
    assert "WINDOWS_CERTIFICATE_BASE64: ${{ secrets.WINDOWS_CERTIFICATE_BASE64 }}" in build_job
    assert "WINDOWS_CERTIFICATE_PASSWORD: ${{ secrets.WINDOWS_CERTIFICATE_PASSWORD }}" in build_job
    assert "WINDOWS_TIMESTAMP_URL: ${{ vars.WINDOWS_TIMESTAMP_URL }}" in build_job
    assert "Code signing skipped: signing credentials are not configured" in build_job


def test_workflow_never_rewrites_existing_release_tags_and_locates_sdk_signtool() -> None:
    text = WORKFLOW.read_text(encoding="utf-8")

    assert "git tag -d" not in text
    assert 'git push origin ":refs/tags/' not in text
    assert "signtool.exe" in text
    assert "Windows Kits" in text
    assert "GITHUB_PATH" in text
