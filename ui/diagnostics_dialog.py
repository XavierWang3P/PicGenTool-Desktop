from __future__ import annotations

from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
    QHBoxLayout,
    QLabel,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from models.diagnostics import DiagnosticSnapshot
from services.diagnostics import DiagnosticsService
from utils.path_utils import open_with_default_app


class DiagnosticsDialog(QDialog):
    def __init__(
        self,
        snapshot: DiagnosticSnapshot,
        log_directory: Path | None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.log_directory = log_directory
        self.setWindowTitle("关于 / 诊断")
        self.setMinimumSize(520, 400)

        heading = QLabel(f"PicGenTool {snapshot.app_version}")
        heading.setStyleSheet("font-size: 18px; font-weight: 700;")
        self.details_edit = QPlainTextEdit(
            DiagnosticsService.format_for_copy(snapshot)
        )
        self.details_edit.setReadOnly(True)
        self.details_edit.setAccessibleName("诊断信息")

        self.status_label = QLabel("")
        self.status_label.setAccessibleName("诊断操作状态")

        self.copy_button = QPushButton("复制诊断信息")
        self.copy_button.setAccessibleName("复制诊断信息")
        self.open_log_button = QPushButton("打开日志目录")
        self.open_log_button.setAccessibleName("打开日志目录")
        self.open_log_button.setEnabled(
            log_directory is not None and log_directory.is_dir()
        )
        self.close_button = QPushButton("关闭")
        self.close_button.setAccessibleName("关闭诊断窗口")
        for button in (self.copy_button, self.open_log_button, self.close_button):
            button.setMinimumHeight(44)

        buttons = QHBoxLayout()
        buttons.setSpacing(8)
        buttons.addWidget(self.copy_button)
        buttons.addWidget(self.open_log_button)
        buttons.addStretch(1)
        buttons.addWidget(self.close_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)
        layout.addWidget(heading)
        layout.addWidget(self.details_edit, 1)
        layout.addWidget(self.status_label)
        layout.addLayout(buttons)

        self.copy_button.clicked.connect(self._copy_details)
        self.open_log_button.clicked.connect(self._open_log_directory)
        self.close_button.clicked.connect(self.accept)
        self.setTabOrder(self.copy_button, self.open_log_button)
        self.setTabOrder(self.open_log_button, self.close_button)
        self.copy_button.setFocus()

    def _copy_details(self) -> None:
        QApplication.clipboard().setText(self.details_edit.toPlainText())
        self.status_label.setText("诊断信息已复制")

    def _open_log_directory(self) -> None:
        if self.log_directory is not None and self.log_directory.is_dir():
            open_with_default_app(self.log_directory)
