from __future__ import annotations

import logging
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path

from PyQt6.QtCore import QDate, QObject, Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QCloseEvent, QDragEnterEvent, QDragLeaveEvent, QDropEvent, QIcon
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
    QSizePolicy,
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
from services.generation_service import GenerationService
from services.image_service import ImageService
from services.settings_service import SettingsService, UserPreferences
from services.template_service import TemplateService
from utils.path_utils import build_default_output_path, open_with_default_app


class GenerationWorker(QObject):
    progress = pyqtSignal(str, int)
    success = pyqtSignal(dict)
    error = pyqtSignal(str)
    cancelled = pyqtSignal()
    finished = pyqtSignal()

    def __init__(self, task: GenerationTask) -> None:
        super().__init__()
        from services.cancellation import CancellationToken

        self.task = task
        self.logger = logging.getLogger("pgt.worker")
        self.service = GenerationService()
        self.image_service = self.service.image_service
        self.document_service = self.service.document_service
        self.cancellation = CancellationToken()

    def cancel(self) -> None:
        self.cancellation.cancel()

    def run(self) -> None:
        from services.cancellation import GenerationCancelled

        try:
            docx_path = self.service.run(
                self.task,
                self._emit_progress,
                self.cancellation,
            )
            if self.task.open_file:
                self.cancellation.raise_if_cancelled()
                self._emit_progress("正在打开生成文件", 94)
                open_with_default_app(docx_path)
            self._emit_progress("生成完成", 100)
            self.success.emit({"docx_path": str(docx_path)})
        except GenerationCancelled:
            self.cancelled.emit()
        except Exception as exc:
            self.logger.error("生成文档失败: error_type=%s", type(exc).__name__)
            self.error.emit(self._safe_error_message(exc))
        finally:
            self.finished.emit()

    def _emit_progress(self, message: str, value: int) -> None:
        bounded = max(0, min(value, 100))
        previous = getattr(self, "_last_progress", 0)
        bounded = max(previous, bounded)
        self._last_progress = bounded
        self.progress.emit(message, bounded)

    def _safe_error_message(self, error: Exception) -> str:
        message = str(error)
        private_paths = [*self.task.image_paths, self.task.save_path]
        for path in private_paths:
            message = message.replace(str(path), path.name)
            try:
                message = message.replace(str(path.resolve()), path.name)
            except OSError:
                pass
        return message


class ImageImportWorker(QObject):
    completed = pyqtSignal(object)
    error = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, files: list[str], existing: list[Path]) -> None:
        super().__init__()
        self.files = files
        self.existing = existing
        self.image_service = ImageService()

    def run(self) -> None:
        try:
            self.completed.emit(
                self.image_service.deduplicate(self.files, existing=self.existing)
            )
        except Exception as exc:
            logging.getLogger("pgt.image_import").error(
                "导入图片失败: error_type=%s",
                type(exc).__name__,
            )
            self.error.emit("检查图片时发生异常，请重试")
        finally:
            self.finished.emit()


class MainWindow(QWidget):
    def __init__(self, settings_service: SettingsService | None = None) -> None:
        super().__init__()
        self.logger = logging.getLogger("pgt.ui")
        self.image_service = ImageService()
        self.template_service = TemplateService()
        self.settings_service = settings_service or SettingsService()
        self.template_validation = None
        self._templates_valid = True
        self.recent_image_dir: Path | None = None
        self._drag_previous_status: str | None = None
        self._clear_after_import = False
        self._generation_starting = False
        self.image_paths: list[Path] = []
        self.worker_thread: QThread | None = None
        self.worker: GenerationWorker | None = None
        self.image_import_thread: QThread | None = None
        self.image_import_worker: ImageImportWorker | None = None
        self._image_import_queue: list[list[str]] = []
        self._generation_cancelled = False
        self._close_requested = False

        self.setWindowTitle(APP_TITLE)
        self.resize(460, 540)
        self.setMinimumSize(440, 520)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setWindowIcon(QIcon(str(Path(__file__).resolve().parent.parent / "favicon.png")))
        self.setAcceptDrops(True)

        self._create_widgets()
        self._setup_layout()
        self._load_preferences()
        self._connect_signals()
        self._setup_style()
        QTimer.singleShot(0, self._run_startup_checks)

    def _create_widgets(self) -> None:
        self.title_label = QLabel("活动标题")
        self.title_input = QLineEdit(placeholderText="请输入活动标题")
        self.title_input.setAccessibleName("活动标题")
        self.title_label.setBuddy(self.title_input)

        self.location_label = QLabel("活动地点")
        self.location_input = QLineEdit(placeholderText="请输入活动地点")
        self.location_input.setAccessibleName("活动地点")
        self.location_label.setBuddy(self.location_input)
        self.location_input.setText(DEFAULT_ACTIVITY_LOCATION)

        self.date_label = QLabel("活动日期")
        self.date_input = QDateEdit()
        self.date_input.setAccessibleName("活动日期")
        self.date_label.setBuddy(self.date_input)
        self.date_input.setDate(QDate.currentDate())
        self.date_input.setDisplayFormat("yyyy年M月d日")
        self.date_input.setCalendarPopup(True)
        self.date_input.setMinimumDate(QDate(2025, 1, 1))
        self.date_input.setMaximumDate(QDate(2099, 12, 31))

        self.open_file_checkbox = QCheckBox("生成后打开文件")
        self.open_file_checkbox.setAccessibleName("生成后打开文件")
        self.open_file_checkbox.setChecked(True)

        self.always_on_top_checkbox = QCheckBox("窗口置顶")
        self.always_on_top_checkbox.setAccessibleName("窗口置顶")
        self.always_on_top_checkbox.setToolTip("让窗口保持在其他窗口前面")
        self.always_on_top_checkbox.setChecked(True)

        self.image_count_label = QLabel("已选择 0 / 30 张图片")
        self.image_count_label.setObjectName("countLabel")
        self.status_label = QLabel("请选择照片")
        self.status_label.setObjectName("statusLabel")
        self.diagnostics_button = QPushButton("关于 / 诊断")
        self.diagnostics_button.setAccessibleName("关于和诊断")
        self.diagnostics_button.setToolTip("查看版本、模板、日志和右键菜单状态")
        self.diagnostics_button.setObjectName("diagnosticsBtn")
        self.diagnostics_button.setMinimumHeight(40)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setFixedHeight(16)
        self.progress_bar.hide()

        self.form_card = QFrame()
        self.form_card.setObjectName("card")
        self.status_card = QFrame()
        self.status_card.setObjectName("card")
        for card in (self.form_card, self.status_card):
            card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Maximum)

        self.image_button = QPushButton("选择照片")
        self.image_button.setAccessibleName("选择照片")
        self.image_button.setToolTip(
            f"可多选 JPG、PNG、GIF、BMP、HEIC、HEIF 格式，最多 {MAX_IMAGES} 张"
        )
        self.clear_button = QPushButton("清空内容")
        self.clear_button.setAccessibleName("清空内容")
        self.clear_button.setToolTip("清空已填写的信息和已选照片")
        self.generate_button = QPushButton("生成文档")
        self.generate_button.setAccessibleName("生成文档")
        self.generate_button.setToolTip("按所选图片数量选择模板并生成 Word 文档")
        self.cancel_button = QPushButton("取消任务")
        self.cancel_button.setAccessibleName("取消任务")
        self.cancel_button.setToolTip("停止当前生成，正在处理的步骤会安全结束")
        self.cancel_button.hide()
        for button in [
            self.image_button,
            self.clear_button,
            self.generate_button,
            self.cancel_button,
        ]:
            button.setFixedHeight(46)

    def _setup_layout(self) -> None:
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(*MARGINS)
        main_layout.setSpacing(12)

        form_layout = QVBoxLayout()
        form_layout.setContentsMargins(16, 16, 16, 16)
        form_layout.setSpacing(12)
        form_layout.addLayout(self._create_row(self.title_label, self.title_input))
        form_layout.addLayout(self._create_row(self.location_label, self.location_input))
        form_layout.addLayout(self._create_row(self.date_label, self.date_input))
        form_layout.addSpacing(4)

        checkbox_layout = QHBoxLayout()
        checkbox_layout.setContentsMargins(0, 0, 0, 0)
        checkbox_layout.setSpacing(16)
        checkbox_layout.addWidget(self.open_file_checkbox)
        checkbox_layout.addWidget(self.always_on_top_checkbox)
        checkbox_layout.addStretch()
        form_layout.addLayout(checkbox_layout)
        self.form_card.setLayout(form_layout)

        status_layout = QVBoxLayout()
        status_layout.setContentsMargins(16, 14, 16, 16)
        status_layout.setSpacing(10)
        status_row = QHBoxLayout()
        status_row.setSpacing(10)
        status_row.addWidget(self.image_count_label)
        status_row.addStretch(1)
        status_row.addWidget(self.status_label)
        status_row.addWidget(self.diagnostics_button)
        status_layout.addLayout(status_row)
        status_layout.addWidget(self.progress_bar)
        self.status_card.setLayout(status_layout)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)
        button_layout.addWidget(self.image_button, 1)
        button_layout.addWidget(self.clear_button, 1)
        button_layout.addWidget(self.generate_button, 1)
        button_layout.addWidget(self.cancel_button, 1)

        main_layout.addWidget(self.form_card)
        main_layout.addWidget(self.status_card)
        main_layout.addLayout(button_layout)
        main_layout.addStretch(1)
        self.setLayout(main_layout)

        self.setTabOrder(self.title_input, self.location_input)
        self.setTabOrder(self.location_input, self.date_input)
        self.setTabOrder(self.date_input, self.open_file_checkbox)
        self.setTabOrder(self.open_file_checkbox, self.always_on_top_checkbox)
        self.setTabOrder(self.always_on_top_checkbox, self.diagnostics_button)
        self.setTabOrder(self.diagnostics_button, self.image_button)
        self.setTabOrder(self.image_button, self.clear_button)
        self.setTabOrder(self.clear_button, self.generate_button)
        self.setTabOrder(self.generate_button, self.cancel_button)
        self.title_input.setFocus()

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
        self.cancel_button.clicked.connect(self._cancel_generation)
        self.diagnostics_button.clicked.connect(self._show_diagnostics)
        self.always_on_top_checkbox.toggled.connect(self._apply_always_on_top)

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
            QLineEdit:disabled {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit_disabled'].items())} }}
            QDateEdit {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit'].items())} }}
            QDateEdit:focus {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit_focus'].items())} }}
            QDateEdit:disabled {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['line_edit_disabled'].items())} }}
            QPushButton {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['button'].items())} }}
            QPushButton:hover {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['button_hover'].items())} }}
            QPushButton:focus {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['button_focus'].items())} }}
            QPushButton:disabled {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['button_disabled'].items())} }}
            QPushButton#generateBtn {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['generate_button'].items())} }}
            QPushButton#generateBtn:hover {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['generate_button_hover'].items())} }}
            QPushButton#generateBtn:focus {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['generate_button_focus'].items())} }}
            QPushButton#generateBtn:disabled {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['generate_button_disabled'].items())} }}
            QPushButton#clearBtn {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['clear_button'].items())} }}
            QPushButton#cancelBtn {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['cancel_button'].items())} }}
            QPushButton#diagnosticsBtn {{
                min-width: 88px;
                padding: 0 12px;
                background: #F6F9FC;
                color: #425466;
            }}
            QCheckBox {{ {'; '.join(f'{k}: {v}' for k, v in APP_STYLES['checkbox'].items())} }}
            QCheckBox {{
                background: transparent;
                padding: 4px 2px;
                border: 1px solid transparent;
                border-radius: 8px;
            }}
            QCheckBox:focus {{
                color: #0B4FBF;
                border-color: #0F6FFF;
            }}
            QFrame#card {{
                background: white;
                border: 1px solid #E3E8EF;
                border-radius: 16px;
            }}
            QFrame#card[dragActive="true"] {{
                background: #EEF5FF;
                border: 2px solid #0F6FFF;
            }}
            QLabel#countLabel {{
                font-size: 15px;
                font-weight: 700;
                color: #1F2933;
            }}
            QLabel#statusLabel {{
                font-size: 13px;
                font-weight: 600;
                color: #5A6674;
            }}
            QProgressBar {{
                border: none;
                border-radius: 8px;
                background: #E8EDF3;
                text-align: center;
                color: #33404F;
                font-size: 12px;
                font-weight: 600;
            }}
            QProgressBar::chunk {{
                background-color: #0F6FFF;
                border-radius: 8px;
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
            QToolTip {{
                background: #1F2933;
                color: white;
                border: none;
                padding: 6px 8px;
                border-radius: 6px;
                font-size: 12px;
            }}
        """
        self.setStyleSheet(style_sheet)
        self.generate_button.setObjectName("generateBtn")
        self.clear_button.setObjectName("clearBtn")
        self.cancel_button.setObjectName("cancelBtn")

    def _handle_image_selection(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "选择图片文件",
            str(self.recent_image_dir or ""),
            f"图片文件 ({' '.join(f'*{fmt}' for fmt in SUPPORTED_FORMATS)})",
        )
        if files:
            self.recent_image_dir = Path(files[0]).resolve().parent
            self._save_preferences()
            self._add_images(files)

    def _add_images(self, files: list[str]) -> None:
        if self._generation_starting or self._clear_after_import:
            self.status_label.setText("请先完成当前操作")
            return
        valid_images = self.image_service.filter_supported(files)
        if not valid_images:
            self._show_message("提示", "所选文件不是支持的图片格式", QMessageBox.Icon.Warning)
            return

        self._queue_image_import([str(path) for path in valid_images])

    def _queue_image_import(self, files: list[str]) -> None:
        if not files:
            return
        self._image_import_queue.append(list(files))
        self.clear_button.setDisabled(True)
        if self.image_import_thread is None:
            self._start_next_image_import()

    def _start_next_image_import(self) -> None:
        if not self._image_import_queue:
            self._show_idle_status()
            return

        files = self._image_import_queue.pop(0)
        self.status_label.setText("正在检查图片")
        thread = QThread(self)
        worker = ImageImportWorker(files, list(self.image_paths))
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.completed.connect(self._handle_image_import_result)
        worker.error.connect(self._handle_image_import_error)
        worker.finished.connect(worker.deleteLater)
        worker.finished.connect(thread.quit)
        thread.finished.connect(self._finish_image_import)
        thread.finished.connect(thread.deleteLater)
        self.image_import_thread = thread
        self.image_import_worker = worker
        thread.start()

    def _handle_image_import_result(self, result) -> None:
        if self._clear_after_import:
            return
        if len(self.image_paths) + len(result.paths) > MAX_IMAGES:
            self._show_message("提示", f"最多可选择 {MAX_IMAGES} 张图片", QMessageBox.Icon.Warning)
            return

        self.image_paths.extend(result.paths)
        self._update_image_count()

        parts: list[str] = []
        if result.paths:
            parts.append(f"已添加 {len(result.paths)} 张")
        if result.duplicate_count:
            parts.append(f"跳过 {result.duplicate_count} 张重复图片")
        if result.failures:
            failed_names = "、".join(failure.filename for failure in result.failures[:3])
            if len(result.failures) > 3:
                failed_names += "等"
            parts.append(f"{len(result.failures)} 张无法读取（{failed_names}）")
        if not parts:
            parts.append("未添加图片")
        elif not result.paths and result.duplicate_count:
            parts[0] = f"未添加：{result.duplicate_count} 张图片已存在"
            if len(parts) > 1 and parts[1].startswith("跳过"):
                parts.pop(1)
        self.status_label.setText("，".join(parts))

    def _handle_image_import_error(self, message: str) -> None:
        if not self._clear_after_import:
            self.status_label.setText(message)

    def _finish_image_import(self) -> None:
        self.image_import_worker = None
        self.image_import_thread = None
        self.clear_button.setEnabled(True)
        if self._close_requested:
            self._image_import_queue.clear()
            if not (self.worker_thread and self.worker_thread.isRunning()):
                self.close()
            return
        if self._clear_after_import:
            self._clear_after_import = False
            self._image_import_queue.clear()
            self._clear_form_fields()
            self._save_preferences()
            self.title_input.setEnabled(True)
            self.location_input.setEnabled(True)
            self.date_input.setEnabled(True)
            self.image_button.setEnabled(True)
            self._show_idle_status()
            self.progress_bar.setValue(0)
            self.progress_bar.hide()
        if self._image_import_queue:
            self._start_next_image_import()

    def _update_image_count(self) -> None:
        self.image_count_label.setText(
            f"已选择 {len(self.image_paths)} / {MAX_IMAGES} 张图片"
        )

    def _reset_form(self) -> None:
        if self.worker_thread and self.worker_thread.isRunning():
            return
        has_content = bool(
            self.title_input.text().strip()
            or self.location_input.text().strip()
            or self.image_paths
        )
        import_active = self.image_import_thread is not None or bool(self._image_import_queue)
        if has_content or import_active:
            if not self._confirm_clear():
                return
        if import_active:
            self._clear_after_import = True
            self._image_import_queue.clear()
            self.status_label.setText("正在清空")
            self.clear_button.setDisabled(True)
            self.image_button.setDisabled(True)
            self.title_input.setDisabled(True)
            self.location_input.setDisabled(True)
            self.date_input.setDisabled(True)
            return
        self._clear_form_fields()
        self._save_preferences()
        self._show_idle_status()
        self.progress_bar.setValue(0)
        self.progress_bar.hide()

    def _confirm_clear(self) -> bool:
        box = QMessageBox(self)
        box.setWindowTitle("清空内容")
        box.setText("清空后，已填写的信息和已选照片会被移除。")
        clear_button = box.addButton("清空", QMessageBox.ButtonRole.DestructiveRole)
        keep_button = box.addButton("保留", QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(keep_button)
        box.exec()
        return box.clickedButton() is clear_button

    def _apply_always_on_top(self, checked: bool) -> None:
        geometry = self.geometry()
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, checked)
        self.show()
        self.setGeometry(geometry)

    def add_image_files(self, files: list[str]) -> bool:
        if (
            not files
            or self._generation_starting
            or self._clear_after_import
            or (self.worker_thread and self.worker_thread.isRunning())
            or self._close_requested
        ):
            return False
        self._add_images(files)
        self.present()
        return True

    def present(self) -> None:
        if self.isMinimized():
            self.setWindowState(self.windowState() & ~Qt.WindowState.WindowMinimized)
        self.show()
        self.raise_()
        self.activateWindow()

    def _show_idle_status(self) -> None:
        if not self._templates_valid:
            self.status_label.setText("模板不完整，请打开关于 / 诊断")
            return
        self.status_label.setText("图片已就绪" if self.image_paths else "请选择照片")

    def _run_startup_checks(self) -> None:
        self.template_validation = self.template_service.validate_all(deep=False)
        self._templates_valid = self.template_validation.valid
        self.generate_button.setEnabled(self._templates_valid)
        self._show_idle_status()

    def _clear_form_fields(self) -> None:
        self.title_input.clear()
        self.location_input.setText(DEFAULT_ACTIVITY_LOCATION)
        self.date_input.setDate(QDate.currentDate())
        self.image_paths.clear()
        self._update_image_count()

    def _load_preferences(self) -> None:
        preferences = self.settings_service.load()
        self.title_input.setText(preferences.title)
        self.location_input.setText(preferences.location)
        self.date_input.setDate(
            QDate(
                preferences.activity_date.year,
                preferences.activity_date.month,
                preferences.activity_date.day,
            )
        )
        self.open_file_checkbox.setChecked(preferences.open_file)
        self.always_on_top_checkbox.setChecked(preferences.always_on_top)
        self.recent_image_dir = preferences.recent_image_dir
        self.setWindowFlag(
            Qt.WindowType.WindowStaysOnTopHint,
            preferences.always_on_top,
        )

    def _save_preferences(self) -> None:
        self.settings_service.save(
            UserPreferences(
                title=self.title_input.text().strip(),
                location=self.location_input.text().strip(),
                activity_date=self.date_input.date().toPyDate(),
                open_file=self.open_file_checkbox.isChecked(),
                always_on_top=self.always_on_top_checkbox.isChecked(),
                recent_image_dir=self.recent_image_dir,
            )
        )

    def _start_generation(self) -> None:
        if (
            self._generation_starting
            or self._clear_after_import
            or self.image_import_thread is not None
            or self._image_import_queue
        ):
            self.status_label.setText("请先完成当前操作")
            return
        self._generation_starting = True
        try:
            try:
                task = self._build_task()
            except Exception as exc:
                self._show_message("输入错误", str(exc), QMessageBox.Icon.Warning)
                return
            if self.image_import_thread is not None or self._image_import_queue:
                self.status_label.setText("正在检查图片，请完成后重试生成")
                return

            self._save_preferences()
            self._set_busy(True)
            self._generation_cancelled = False
            self.status_label.setText("正在准备生成")
            self.progress_bar.setValue(0)

            self.worker_thread = QThread(self)
            self.worker = GenerationWorker(task)
            exception_handler = getattr(
                QApplication.instance(),
                "_pgt_exception_handler",
                None,
            )
            if exception_handler is not None:
                exception_handler.set_sensitive_values(
                    [task.title, task.location, *task.image_paths, task.save_path]
                )
            self.worker.moveToThread(self.worker_thread)
            self.worker_thread.started.connect(self.worker.run)
            self.worker.progress.connect(self._handle_progress)
            self.worker.success.connect(self._handle_success)
            self.worker.error.connect(self._handle_error)
            self.worker.cancelled.connect(self._handle_cancelled)
            self.worker.finished.connect(self.worker.deleteLater)
            self.worker.finished.connect(self.worker_thread.quit)
            self.worker_thread.finished.connect(self._cleanup_worker)
            self.worker_thread.finished.connect(self.worker_thread.deleteLater)
            self.worker_thread.start()
        finally:
            self._generation_starting = False

    def _cancel_generation(self) -> None:
        if self.worker is None:
            return
        self.worker.cancel()
        self.status_label.setText("正在取消")
        self.cancel_button.setDisabled(True)

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
            raise ValueError(f"最多可选择 {MAX_IMAGES} 张图片")

        selected_date = self.date_input.date().toPyDate()
        image_paths = list(self.image_paths)
        default_path = build_default_output_path(image_paths[0], title, selected_date)
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
            image_paths=[Path(path) for path in image_paths],
            save_path=Path(save_path_str),
            open_file=self.open_file_checkbox.isChecked(),
        )

    def _handle_progress(self, message: str, value: int) -> None:
        self.status_label.setText(message)
        self.progress_bar.setValue(value)
        self.progress_bar.show()

    def _handle_success(self, result: dict) -> None:
        docx_path = result["docx_path"]
        message = f"Word 文档已生成：\n{docx_path}"
        self._show_message("成功", message, QMessageBox.Icon.Information)
        self.title_input.clear()
        self.image_paths.clear()
        self._update_image_count()
        self.settings_service.clear_title()
        self._save_preferences()

    def _handle_error(self, message: str) -> None:
        self._show_message("生成失败", message, QMessageBox.Icon.Critical)

    def _handle_cancelled(self) -> None:
        self._generation_cancelled = True
        self.status_label.setText("已取消")

    def _cleanup_worker(self) -> None:
        exception_handler = getattr(
            QApplication.instance(),
            "_pgt_exception_handler",
            None,
        )
        if exception_handler is not None:
            exception_handler.clear_sensitive_values()
        self._set_busy(False)
        if self._generation_cancelled:
            self.status_label.setText("已取消")
        else:
            self._show_idle_status()
        self.progress_bar.hide()
        self.worker = None
        self.worker_thread = None
        if self._close_requested:
            self.close()

    def _set_busy(self, busy: bool) -> None:
        self.generate_button.setDisabled(busy or not self._templates_valid)
        self.generate_button.setVisible(not busy)
        self.cancel_button.setVisible(busy)
        self.cancel_button.setEnabled(busy)
        self.image_button.setDisabled(busy)
        self.clear_button.setDisabled(
            busy or self.image_import_thread is not None or bool(self._image_import_queue)
        )
        self.open_file_checkbox.setDisabled(busy)
        if busy:
            self.progress_bar.show()
        else:
            self.progress_bar.hide()

    def _show_message(self, title: str, message: str, icon: QMessageBox.Icon) -> None:
        box = QMessageBox(self)
        box.setWindowTitle(title)
        box.setText(message)
        box.setIcon(icon)
        box.exec()

    def _show_diagnostics(self) -> None:
        from services.diagnostics import DiagnosticsService
        from ui.diagnostics_dialog import DiagnosticsDialog
        from utils.logging_utils import current_log_path

        snapshot = DiagnosticsService(self.template_service).collect(deep_templates=True)
        log_path = current_log_path()
        dialog = DiagnosticsDialog(
            snapshot,
            log_path.parent if log_path is not None else None,
            self,
        )
        dialog.exec()

    def dragEnterEvent(self, event: QDragEnterEvent) -> None:
        valid = any(
            url.isLocalFile() and self.image_service.is_supported_image(url.toLocalFile())
            for url in event.mimeData().urls()
        )
        if valid:
            self._set_drag_active(True)
            event.acceptProposedAction()
        else:
            self._set_drag_active(False)
            event.ignore()

    def dragLeaveEvent(self, event: QDragLeaveEvent) -> None:
        self._set_drag_active(False)
        event.accept()

    def dropEvent(self, event: QDropEvent) -> None:
        try:
            if (
                self._generation_starting
                or self._clear_after_import
                or self._close_requested
                or (self.worker_thread and self.worker_thread.isRunning())
            ):
                event.ignore()
                return
            files = [
                url.toLocalFile()
                for url in event.mimeData().urls()
                if url.isLocalFile()
                and self.image_service.is_supported_image(url.toLocalFile())
            ]
            if files:
                self._add_images(files)
                event.acceptProposedAction()
            else:
                event.ignore()
        finally:
            self._set_drag_active(False)

    def _set_drag_active(self, active: bool) -> None:
        if active and not bool(self.property("dragActive")):
            self._drag_previous_status = self.status_label.text()
        self.setProperty("dragActive", active)
        self.status_card.setProperty("dragActive", active)
        for widget in (self, self.status_card):
            widget.style().unpolish(widget)
            widget.style().polish(widget)
        if active:
            self.status_label.setText("松开即可添加图片")
        elif self.status_label.text() == "松开即可添加图片":
            self.status_label.setText(self._drag_previous_status or "请选择照片")
        if not active:
            self._drag_previous_status = None

    def closeEvent(self, event: QCloseEvent) -> None:
        self._save_preferences()
        if self.image_import_thread is not None and self.image_import_thread.isRunning():
            self._close_requested = True
            self._image_import_queue.clear()
            event.ignore()
            return
        if self.worker_thread is not None and self.worker_thread.isRunning():
            self._close_requested = True
            self._cancel_generation()
            event.ignore()
            return
        event.accept()


def create_application() -> QApplication:
    return QApplication.instance() or QApplication([])
