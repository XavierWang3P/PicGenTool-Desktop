from __future__ import annotations

import logging
from pathlib import Path

from PyQt6.QtCore import QDate, QObject, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from config import (
    APP_STYLES,
    APP_TITLE,
    DEFAULT_ACTIVITY_LOCATION,
    MAX_IMAGES,
    MARGINS,
    SPACING,
    SUPPORTED_FORMATS,
)
from models.generation_task import GenerationTask
from services.document_service import DocumentService
from services.image_service import ImageService
from utils.path_utils import build_default_output_path, open_with_default_app


class GenerationWorker(QObject):
    progress = pyqtSignal(str, int)
    success = pyqtSignal(dict)
    error = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, task: GenerationTask) -> None:
        super().__init__()
        self.task = task
        self.logger = logging.getLogger("pgt.worker")
        self.image_service = ImageService()
        self.document_service = DocumentService()

    def run(self) -> None:
        prepared = None
        try:
            prepared = self.image_service.prepare_images(self.task.image_paths, self._emit_progress)
            docx_path = self.document_service.generate(
                self.task, prepared.image_paths, self._emit_progress
            )

            if self.task.open_file:
                self._emit_progress("正在打开生成文件", 94)
                open_with_default_app(docx_path)

            self._emit_progress("生成完成", 100)
            self.success.emit(
                {
                    "docx_path": str(docx_path),
                }
            )
        except Exception as exc:
            self.logger.exception("生成文档失败")
            self.error.emit(str(exc))
        finally:
            if prepared:
                prepared.cleanup()
            self.finished.emit()

    def _emit_progress(self, message: str, value: int) -> None:
        self.progress.emit(message, max(0, min(value, 100)))


class MainWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.logger = logging.getLogger("pgt.ui")
        self.image_service = ImageService()
        self.image_paths: list[Path] = []
        self.worker_thread: QThread | None = None
        self.worker: GenerationWorker | None = None

        self.setWindowTitle(APP_TITLE)
        self.setGeometry(300, 200, 420, 440)
        self.setMinimumSize(420, 440)
        self.setMaximumSize(420, 440)
        self.setWindowIcon(QIcon(str(Path(__file__).resolve().parent.parent / "favicon.png")))
        self.setAcceptDrops(True)

        self._create_widgets()
        self._setup_layout()
        self._connect_signals()
        self._setup_style()

    def _create_widgets(self) -> None:
        self.title_label = QLabel("活动标题")
        self.title_input = QLineEdit(placeholderText="请输入活动标题")
        self.location_label = QLabel("活动地点")
        self.location_input = QLineEdit(placeholderText="请输入活动地点")
        self.location_input.setText(DEFAULT_ACTIVITY_LOCATION)

        self.date_label = QLabel("活动日期")
        self.date_input = QDateEdit()
        self.date_input.setDate(QDate.currentDate())
        self.date_input.setDisplayFormat("yyyy年M月d日")
        self.date_input.setCalendarPopup(True)
        self.date_input.setMinimumDate(QDate(2025, 1, 1))
        self.date_input.setMaximumDate(QDate(2099, 12, 31))
        self.date_input.setFixedHeight(42)

        self.open_file_checkbox = QCheckBox("生成后打开文件")
        self.open_file_checkbox.setChecked(True)

        self.image_count_label = QLabel("已选择 0 张图片")
        self.image_count_label.setObjectName("countLabel")
        self.status_label = QLabel("准备就绪")
        self.status_label.setObjectName("statusLabel")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(14)

        self.form_card = QFrame()
        self.form_card.setObjectName("card")
        self.status_card = QFrame()
        self.status_card.setObjectName("card")

        self.image_button = QPushButton("选择照片")
        self.clear_button = QPushButton("清空内容")
        self.generate_button = QPushButton("生成文档")
        for button in [self.image_button, self.clear_button, self.generate_button]:
            button.setFixedHeight(46)

    def _setup_layout(self) -> None:
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(*MARGINS)
        main_layout.setSpacing(12)

        form_layout = QVBoxLayout()
        form_layout.setContentsMargins(14, 14, 14, 14)
        form_layout.setSpacing(12)
        form_layout.addLayout(self._create_row(self.title_label, self.title_input))
        form_layout.addLayout(self._create_row(self.location_label, self.location_input))
        form_layout.addLayout(self._create_row(self.date_label, self.date_input))
        form_layout.addSpacing(4)

        checkbox_layout = QHBoxLayout()
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        checkbox_layout.setSpacing(22)
        checkbox_layout.addWidget(self.open_file_checkbox)
        checkbox_layout.addStretch()
        form_layout.addLayout(checkbox_layout)
        self.form_card.setLayout(form_layout)

        status_layout = QVBoxLayout()
        status_layout.setContentsMargins(14, 12, 14, 14)
        status_layout.setSpacing(10)
        status_layout.addWidget(self.image_count_label, alignment=Qt.AlignmentFlag.AlignCenter)
        status_layout.addWidget(self.status_label)
        status_layout.addWidget(self.progress_bar)
        self.status_card.setLayout(status_layout)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        button_layout.addWidget(self.image_button, 1)
        button_layout.addWidget(self.clear_button, 1)
        button_layout.addWidget(self.generate_button, 1)

        main_layout.addWidget(self.form_card)
        main_layout.addWidget(self.status_card)
        main_layout.addStretch(1)
        main_layout.addLayout(button_layout)
        self.setLayout(main_layout)

    def _create_row(self, label: QLabel, field: QWidget) -> QHBoxLayout:
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(12)
        label.setFixedWidth(72)
        layout.addWidget(label, 0, Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(field, 1)
        return layout

    def _connect_signals(self) -> None:
        self.image_button.clicked.connect(self._handle_image_selection)
        self.clear_button.clicked.connect(self._reset_form)
        self.generate_button.clicked.connect(self._start_generation)

    def _setup_style(self) -> None:
        style_sheet = f"""
            QWidget {{
                background: #F4F6F8;
                font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
                color: #2F3437;
            }}
            QLabel {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['label'].items())} }}
            QLabel {{
                background: transparent;
            }}
            QLineEdit {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit'].items())} }}
            QLineEdit:focus {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit_focus'].items())} }}
            QDateEdit {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit'].items())} }}
            QDateEdit:focus {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit_focus'].items())} }}
            QPushButton {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['button'].items())} }}
            QPushButton:hover {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['button_hover'].items())} }}
            QPushButton#generateBtn {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['generate_button'].items())} }}
            QPushButton#generateBtn:hover {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['generate_button_hover'].items())} }}
            QPushButton#clearBtn {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['clear_button'].items())} }}
            QCheckBox {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['checkbox'].items())} }}
            QCheckBox {{
                background: transparent;
            }}
            QFrame#card {{
                background: white;
                border: 1px solid #E3E8EF;
                border-radius: 16px;
            }}
            QLabel#countLabel {{
                font-size: 15px;
                font-weight: 700;
                color: #1F2933;
                padding: 2px 0 0 0;
            }}
            QLabel#statusLabel {{
                font-size: 14px;
                font-weight: 600;
                color: #4D5966;
            }}
            QProgressBar {{
                border: none;
                border-radius: 7px;
                background: #E8EDF3;
                text-align: center;
                color: #4D5966;
                font-size: 12px;
                font-weight: 600;
            }}
            QProgressBar::chunk {{
                background-color: #0F6FFF;
                border-radius: 7px;
            }}
            QDateEdit::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 28px;
                border: none;
                margin-right: 6px;
                background: transparent;
            }}
            QDateEdit::down-arrow {{
                width: 12px;
                height: 12px;
            }}
        """
        self.setStyleSheet(style_sheet)
        self.generate_button.setObjectName("generateBtn")
        self.clear_button.setObjectName("clearBtn")

    def _handle_image_selection(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择图片文件",
            "",
            f"图片文件 ({' '.join(f'*{fmt}' for fmt in SUPPORTED_FORMATS)})",
        )
        if files:
            self._add_images(files)

    def _add_images(self, files: list[str]) -> None:
        valid_images = self.image_service.filter_supported(files)
        if not valid_images:
            self._show_message("提示", "没有找到有效的图片文件", QMessageBox.Icon.Warning)
            return

        if len(self.image_paths) + len(valid_images) > MAX_IMAGES:
            self._show_message("提示", f"最多只能选择 {MAX_IMAGES} 张图片", QMessageBox.Icon.Warning)
            return

        self.image_paths.extend(valid_images)
        self._update_image_count()
        self.status_label.setText("图片已就绪")

    def _update_image_count(self) -> None:
        self.image_count_label.setText(f"已选择 {len(self.image_paths)} 张图片")

    def _reset_form(self) -> None:
        if self.worker_thread and self.worker_thread.isRunning():
            return
        self._clear_form_fields()
        self.status_label.setText("准备就绪")
        self.progress_bar.setValue(0)

    def _clear_form_fields(self) -> None:
        self.title_input.clear()
        self.location_input.setText(DEFAULT_ACTIVITY_LOCATION)
        self.date_input.setDate(QDate.currentDate())
        self.image_paths.clear()
        self._update_image_count()

    def _start_generation(self) -> None:
        try:
            task = self._build_task()
        except Exception as exc:
            self._show_message("输入错误", str(exc), QMessageBox.Icon.Warning)
            return

        self._set_busy(True)
        self.status_label.setText("准备开始生成")
        self.progress_bar.setValue(0)

        self.worker_thread = QThread(self)
        self.worker = GenerationWorker(task)
        self.worker.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker.run)
        self.worker.progress.connect(self._handle_progress)
        self.worker.success.connect(self._handle_success)
        self.worker.error.connect(self._handle_error)
        self.worker.finished.connect(self._cleanup_worker)
        self.worker.finished.connect(self.worker_thread.quit)
        self.worker_thread.finished.connect(self.worker_thread.deleteLater)
        self.worker_thread.start()

    def _build_task(self) -> GenerationTask:
        title = self.title_input.text().strip()
        location = self.location_input.text().strip()

        if not title:
            raise ValueError("请输入活动标题")
        if not location:
            raise ValueError("请输入活动地点")
        if not self.image_paths:
            raise ValueError("请选择至少一张图片")
        if len(self.image_paths) > MAX_IMAGES:
            raise ValueError(f"最多只能选择 {MAX_IMAGES} 张图片")

        selected_date = self.date_input.date().toPyDate()
        default_path = build_default_output_path(self.image_paths[0], title, selected_date)
        save_path_str, _ = QFileDialog.getSaveFileName(
            self,
            "保存文档",
            str(default_path),
            "Word 文档 (*.docx)",
        )
        if not save_path_str:
            raise ValueError("已取消保存")

        return GenerationTask(
            title=title,
            location=location,
            activity_date=selected_date,
            image_paths=[Path(path) for path in self.image_paths],
            save_path=Path(save_path_str),
            open_file=self.open_file_checkbox.isChecked(),
        )

    def _handle_progress(self, message: str, value: int) -> None:
        self.status_label.setText(message)
        self.progress_bar.setValue(value)

    def _handle_success(self, result: dict) -> None:
        docx_path = result["docx_path"]
        message = f"Word 文档已生成：\n{docx_path}"
        self._show_message("成功", message, QMessageBox.Icon.Information)
        self._clear_form_fields()

    def _handle_error(self, message: str) -> None:
        self._show_message("生成失败", message, QMessageBox.Icon.Critical)

    def _cleanup_worker(self) -> None:
        self._set_busy(False)
        self.status_label.setText("准备就绪")
        if self.worker:
            self.worker.deleteLater()
            self.worker = None
        self.worker_thread = None

    def _set_busy(self, busy: bool) -> None:
        self.generate_button.setDisabled(busy)
        self.image_button.setDisabled(busy)
        self.clear_button.setDisabled(busy)
        self.open_file_checkbox.setDisabled(busy)

    def _show_message(self, title: str, message: str, icon: QMessageBox.Icon) -> None:
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(icon)
        box.exec()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        if any(
            url.isLocalFile() and self.image_service.is_supported_image(url.toLocalFile())
            for url in event.mimeData().urls()
        ):
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent) -> None:
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        self._add_images(files)


def create_application() -> QApplication:
    return QApplication.instance() or QApplication([])
