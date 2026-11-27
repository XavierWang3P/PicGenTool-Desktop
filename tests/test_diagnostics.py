from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from unittest.mock import Mock

from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import QApplication

from models.diagnostics import DiagnosticSnapshot
from models.template_validation import TemplateIssue, TemplateValidationResult
from services.diagnostics import DiagnosticsService
from services.explorer_integration import ExplorerRegistrationStatus
from ui.diagnostics_dialog import DiagnosticsDialog


def _snapshot(private_log: Path) -> DiagnosticSnapshot:
    return DiagnosticSnapshot(
        app_version="3.1.1",
        python_version="3.11.9",
        platform="Windows 11 AMD64",
        packaged=True,
        heif_available=True,
        heif_version="1.4.0",
        template_result=TemplateValidationResult(True, 30, ()),
        log_path=private_log,
        explorer_status=ExplorerRegistrationStatus.CURRENT,
    )


def test_diagnostics_copy_text_is_stable_and_omits_private_values(tmp_path: Path) -> None:
    private_directory = tmp_path / "private-user" / "logs"
    snapshot = _snapshot(private_directory / "app.log")

    text = DiagnosticsService.format_for_copy(snapshot)

    for expected in (
        "PicGenTool: 3.1.1",
        "Python: 3.11.9",
        "HEIF: available (1.4.0)",
        "Templates: valid (30/30)",
        "Log: available (app.log)",
        "Packaged: yes",
        "Explorer menu: current",
    ):
        assert expected in text
    assert str(private_directory) not in text
    assert "secret activity" not in text
    assert "photo.jpg" not in text
    assert "result.docx" not in text


def test_diagnostics_copy_text_includes_safe_template_issue_details(tmp_path: Path) -> None:
    snapshot = _snapshot(tmp_path / "app.log")
    snapshot = replace(
        snapshot,
        template_result=TemplateValidationResult(
            False,
            29,
            (
                TemplateIssue(
                    7,
                    "missing_placeholders",
                    "缺少占位符：image2、location",
                ),
            ),
        ),
    )

    text = DiagnosticsService.format_for_copy(snapshot)

    assert "Template-7: missing_placeholders: 缺少占位符：image2、location" in text


def test_dialog_buttons_are_accessible_and_copy_without_persisting(
    qtbot,
    tmp_path: Path,
    monkeypatch,
) -> None:
    log_directory = tmp_path / "logs"
    log_directory.mkdir()
    opener = Mock()
    monkeypatch.setattr("ui.diagnostics_dialog.open_with_default_app", opener)
    dialog = DiagnosticsDialog(_snapshot(log_directory / "app.log"), log_directory)
    qtbot.addWidget(dialog)

    assert dialog.copy_button.accessibleName() == "复制诊断信息"
    assert dialog.open_log_button.accessibleName() == "打开日志目录"
    assert dialog.close_button.accessibleName() == "关闭诊断窗口"
    assert dialog.details_edit.isReadOnly()

    qtbot.mouseClick(dialog.copy_button, Qt.MouseButton.LeftButton)
    assert QApplication.clipboard().text() == dialog.details_edit.toPlainText()
    assert dialog.status_label.text() == "诊断信息已复制"

    qtbot.mouseClick(dialog.open_log_button, Qt.MouseButton.LeftButton)
    opener.assert_called_once_with(log_directory)
