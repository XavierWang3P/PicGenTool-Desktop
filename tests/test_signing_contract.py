from __future__ import annotations

from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_signing_script_uses_certificate_store_and_verifies_timestamped_signature() -> None:
    text = (PROJECT_ROOT / "tools" / "sign_release.ps1").read_text(encoding="utf-8")

    for required in (
        "[Parameter(Mandatory = $true)]",
        "$TargetPath",
        "$CertificateBase64",
        "$CertificatePassword",
        "SecureString",
        "CurrentUser",
        "My",
        "Thumbprint",
        "/fd SHA256",
        "/tr $TimestampUrl",
        "/td SHA256",
        "verify /pa /tw",
        "http://timestamp.digicert.com",
        "finally",
        "Remove-Item -LiteralPath $pfxPath",
        "Remove-Item -LiteralPath \"Cert:\\CurrentUser\\My\\$thumbprint\"",
    ):
        assert required in text

    assert "Write-Host $CertificatePassword" not in text
    assert "Write-Output $CertificatePassword" not in text


def test_signing_script_whatif_does_not_require_certificate_or_modify_exe() -> None:
    text = (PROJECT_ROOT / "tools" / "sign_release.ps1").read_text(encoding="utf-8")

    assert "[CmdletBinding(SupportsShouldProcess)]" in text
    assert "ShouldProcess" in text
    assert "Get-Command signtool.exe" in text
    assert "Join-Path (Split-Path $TargetPath) PGT.exe" not in text


def test_spec_collects_heif_and_removes_only_planned_unused_modules() -> None:
    text = (PROJECT_ROOT / "pgt.spec").read_text(encoding="utf-8")

    assert "collect_all('pi_heif')" in text
    assert "('config.py', '.')" not in text
    assert "templates/Template-[0-9]*.docx" in text
    assert "a.binaries" in text
    assert "a.datas" in text
    assert "console=True" in text
    for excluded in ("PyQt6.QtPdf", "PyQt6.QtPdfWidgets", "PIL.AvifImagePlugin"):
        assert f"'{excluded}'" in text
    assert "noarchive=True" in text
    assert "upx=False" in text
    assert "PyQt6.QtNetwork" in text
    assert "opengl32sw.dll" not in text.split("excludes=[", 1)[1].split("]", 1)[0]
