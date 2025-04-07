import sys
import os
import logging
from typing import List, Optional

from PyQt6.QtWidgets import (QWidget, QPushButton, QLineEdit, QLabel,
                             QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox,
                             QApplication, QCalendarWidget)
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon
from PyQt6.QtCore import Qt, QSize, QDate
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from PIL import Image
from docx2pdf import convert

# 常量定义
MAX_IMAGES = 30
IMAGE_PREVIEW_SIZE = QSize(40, 40)
PREVIEW_COLUMNS = 1
MARGINS = (12, 12, 12, 12)
SPACING = 6

class ImageProcessor:
    """处理图片相关的逻辑"""
    def __init__(pgt):
        pgt.images: List[str] = []
        pgt._supported_formats = {'.png', '.jpg', '.jpeg', '.gif', '.bmp'}
        pgt.logger = logging.getLogger(__name__)
        pgt.compressed_images: List[str] = []

    def add_images(pgt, files: List[str]) -> bool:
        """添加图片并返回是否成功"""
        valid_images = [f for f in files if pgt._is_image(f)]
        if not valid_images:
            pgt.logger.warning("没有找到有效的图片文件")
            return False

        if len(pgt.images) + len(valid_images) > MAX_IMAGES:
            pgt.logger.warning(f"图片总数超过最大限制 {MAX_IMAGES}")
            return False

        pgt.images.extend(valid_images)
        pgt.logger.info(f"成功添加 {len(valid_images)} 张图片")
        return True

    def clear_images(pgt):
        """清空图片"""
        pgt.images.clear()
        pgt.compressed_images.clear()
        pgt.logger.info("已清空所有图片")

    def _is_image(pgt, path: str) -> bool:
        """检查是否为支持的图片格式"""
        ext = os.path.splitext(path)[1].lower()
        return ext in pgt._supported_formats

    def compress_images(pgt) -> List[str]:
        """压缩所有图片并返回压缩后的图片路径列表"""
        if not pgt.images:
            return []

        pgt.compressed_images.clear()
        for img_path in pgt.images:
            try:
                # 打开图片
                with Image.open(img_path) as img:
                    # 创建临时文件路径，使用系统临时目录
                    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.temp')
                    os.makedirs(temp_dir, exist_ok=True)
                    compressed_path = os.path.join(
                        temp_dir,
                        f'compressed_{os.path.basename(img_path)}'
                    )

                    # 调整图片尺寸
                    max_size = 1200
                    width, height = img.size
                    if width > max_size or height > max_size:
                        # 计算缩放比例
                        ratio = max_size / max(width, height)
                        new_size = (int(width * ratio), int(height * ratio))
                        img = img.resize(new_size, Image.Resampling.LANCZOS)

                    # 保存压缩后的图片
                    img.save(
                        compressed_path,
                        quality=60,  # 压缩质量
                        optimize=True  # 优化文件大小
                    )
                    pgt.compressed_images.append(compressed_path)
                    pgt.logger.info(f"成功压缩图片: {img_path}")
            except Exception as e:
                pgt.logger.error(f"压缩图片失败: {img_path}, 错误: {str(e)}")
                # 如果压缩失败，使用原图
                pgt.compressed_images.append(img_path)

        return pgt.compressed_images

class DocumentGenerator:
    """处理文档生成的逻辑"""
    def __init__(pgt, image_processor: ImageProcessor):
        pgt.image_processor = image_processor

    def generate_document(pgt, title: str, location: str, date: QDate) -> Optional[str]:
        """生成文档并返回保存路径"""
        if not pgt._validate_inputs(title, location):
            return None

        template_path = pgt._get_template_path()
        if not template_path:
            raise FileNotFoundError("模板文件未找到或路径不正确")

        doc = DocxTemplate(template_path)
        context = pgt._prepare_context(doc, title, location, date)
        doc.render(context)

        save_path = pgt._get_save_path(title, date)
        if save_path:
            doc.save(save_path)
            # 生成PDF文件
            try:
                from docx2pdf import convert
                pdf_path = save_path.replace('.docx', '.pdf')
                convert(save_path, pdf_path)
            except Exception as e:
                pgt.logger.error(f"PDF生成失败: {str(e)}")
            return save_path
        return None

    def _validate_inputs(pgt, title: str, location: str) -> bool:
        """验证用户输入"""
        return bool(title and location and pgt.image_processor.images)

    def _prepare_context(pgt, doc: DocxTemplate, title: str, location: str, date: QDate) -> dict:
        """准备模板上下文数据"""
        context = {
            'title': title,
            'location': location,
            'dateForDoc': f'{date.year()}年{date.month()}月{date.day()}日'
        }

        # 使用压缩后的图片
        compressed_images = pgt.image_processor.compress_images()
        for i, path in enumerate(compressed_images, 1):
            context[f'image{i}'] = InlineImage(doc, path, width=Mm(72))
        return context

    def _get_template_path(pgt) -> Optional[str]:
        """获取模板文件路径"""
        path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'templates', f'Template-{len(pgt.image_processor.images)}.docx')
        return path if os.path.exists(path) else None

    def _get_save_path(pgt, title: str, date: QDate) -> Optional[str]:
        """获取保存路径"""
        date_str = f"{str(date.year())[2:]}.{date.month():02}.{date.day():02}"
        file_name = f'{date_str}-{title}照片.docx'
        # 确保文件名中不包含非法字符
        file_name = ''.join(c for c in file_name if c not in '<>:"/\\|?*')
        return QFileDialog.getSaveFileName(
            None,
            '保存文档',
            file_name,
            'Word 文档 (*.docx)'
        )[0]

class DocGeneratorUI(QWidget):
    """负责UI布局和用户交互"""
    def __init__(pgt, image_processor: ImageProcessor, document_generator: DocumentGenerator, parent=None):
        super().__init__(parent)
        pgt.image_processor = image_processor
        pgt.document_generator = document_generator
        pgt._init_ui()

    def _init_ui(pgt):
        """初始化UI"""
        pgt.setAcceptDrops(True)
        pgt._create_widgets()
        pgt._setup_layout()
        pgt._connect_signals()
        pgt._setup_style()

    def _create_widgets(pgt):
        """创建界面控件"""
        pgt.title_input = QLineEdit(placeholderText="请输入活动标题")
        pgt.location_input = QLineEdit(placeholderText="请输入活动地点")
        pgt.calendar = QCalendarWidget()
        pgt._configure_calendar()
        pgt.image_count_label = QLabel("已选择 0 张图片")
        pgt.image_button = QPushButton("选择文件")
        pgt.clear_button = QPushButton("清空内容")
        pgt.generate_button = QPushButton("生成文档")

    def _configure_calendar(pgt):
        """配置日历组件样式"""
        pgt.calendar.setGridVisible(True)
        pgt.calendar.setMinimumDate(QDate(2025, 1, 1))
        pgt.calendar.setMaximumDate(QDate(2026, 12, 31))
        pgt.calendar.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        pgt.calendar.setFixedHeight(220)
        pgt.calendar.setMaximumWidth(380)
        pgt.calendar.setStyleSheet("""
            QCalendarWidget QWidget { background-color: white; }
            QCalendarWidget QToolButton { color: #333; padding: 6px; }
            QCalendarWidget QMenu { width: 150px; left: 20px; color: #333; }
            QCalendarWidget QSpinBox { width: 60px; font-size: 14px; }
            QCalendarWidget QTableView { selection-background-color: #06F; }
        """)

    def _setup_layout(pgt):
        """设置界面布局"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(*MARGINS)
        main_layout.setSpacing(SPACING)
    
        # 输入区域布局
        main_layout.addLayout(pgt._create_input_layout())
        # 图片计数布局
        main_layout.addLayout(pgt._create_image_count_layout())
        main_layout.addStretch()
        # 按钮区域布局
        main_layout.addLayout(pgt._create_button_layout())
    
        pgt.setLayout(main_layout)

    def _create_input_layout(pgt):
        """创建输入区域布局"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        for widget in [pgt.title_input, pgt.location_input, pgt.calendar]:
            layout.addWidget(widget)
        return layout

    def _create_image_count_layout(pgt):
        """创建图片计数布局"""
        layout = QHBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(0, 2, 0, 2)
        layout.addStretch()
        layout.addWidget(pgt.image_count_label)
        layout.addStretch()
        return layout

    def _create_button_layout(pgt):
        """创建按钮区域布局"""
        layout = QHBoxLayout()
        layout.setSpacing(8)
        for button in [pgt.image_button, pgt.clear_button, pgt.generate_button]:
            layout.addWidget(button)
        return layout

    def _connect_signals(pgt):
        """连接信号与槽"""
        pgt.image_button.clicked.connect(pgt._handle_image_selection)
        pgt.clear_button.clicked.connect(pgt._reset_ui)
        pgt.generate_button.clicked.connect(pgt._generate_document)

    def _setup_style(pgt):
        """设置组件样式"""
        pgt.setStyleSheet("""
            QLabel { font-size: 13px; color: #333; }
            QLineEdit {
                padding: 6px;
                border: 1px solid #DDD;
                border-radius: 4px;
                background: white;
                font-size: 13px;
            }
            QLineEdit:focus { border-color: #06F; }
            QPushButton {
                padding: 6px 12px;
                border-radius: 4px;
                font-size: 13px;
                border: 1px solid #DDD;
                background: white;
            }
            QPushButton:hover {
                background: #F5F5F5;
                border-color: #CCC;
            }
            QPushButton#generateBtn {
                background: #06F; 
                color: white; 
                min-height: 32px;
                padding: 0 16px;
            }
            QPushButton#generateBtn:hover { background: #05C; }
            QPushButton#clearBtn { background: #F5F5F5; color: #666; }
        """)
        pgt.generate_button.setObjectName("generateBtn")
        pgt.clear_button.setObjectName("clearBtn")

    def _handle_image_selection(pgt):
        """处理图片选择"""
        files, _ = QFileDialog.getOpenFileNames(
            parent=pgt,
            caption="选择图片文件",
            filter="图片文件 (*.png *.jpg *.jpeg *.gif *.bmp)"
        )
        if files and pgt.image_processor.add_images(files):
            pgt._update_image_count()
        else:
            pgt._show_message("提示", "最多只能选择30张图片！", QMessageBox.warning)

    def _update_image_count(pgt):
        """更新图片数量显示"""
        pgt.image_count_label.setText(f"已选择 {len(pgt.image_processor.images)} 张图片")

    def _generate_document(pgt):
        """生成文档处理逻辑"""
        title = pgt.title_input.text()
        location = pgt.location_input.text()
        date = pgt.calendar.selectedDate()

        try:
            save_path = pgt.document_generator.generate_document(title, location, date)
            if save_path:
                pgt._show_message("成功", "文档生成成功！", QMessageBox.Icon.Information)
                # 清理临时文件夹
                pgt._cleanup_temp_files()
                pgt._reset_ui()
        except Exception as e:
            pgt._show_message("错误", str(e), QMessageBox.Icon.Critical)

    def _cleanup_temp_files(pgt):
        """清理临时文件夹"""
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.temp')
        if os.path.exists(temp_dir):
            for file in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, file)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"清理临时文件失败: {e}")

    def _reset_ui(pgt):
        """重置界面状态"""
        pgt.title_input.clear()
        pgt.location_input.clear()
        pgt.calendar.setSelectedDate(QDate.currentDate())
        pgt.image_processor.clear_images()
        pgt._update_image_count()

    def _show_message(pgt, title: str, message: str, icon: QMessageBox.Icon):
        """显示消息框"""
        msg_box = QMessageBox()
        msg_box.setIcon(icon)
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.exec()

    # 拖放事件处理
    def dragEnterEvent(pgt, event: QDragEnterEvent):
        if any(url.isLocalFile() and pgt.image_processor._is_image(url.toLocalFile())
               for url in event.mimeData().urls()):
            event.acceptProposedAction()

    def dropEvent(pgt, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files and pgt.image_processor.add_images(files):
            pgt._update_image_count()
        else:
            pgt._show_message("提示", "最多只能选择30张图片！", QMessageBox.warning)

class MainWindow(QWidget):
    def __init__(pgt):
        super().__init__()
        pgt.setWindowTitle("活动照片文档生成")
        pgt.setGeometry(300, 200, 380, 420)
        pgt.setMinimumSize(330, 400)
        pgt.setMaximumSize(330, 400)
        # 设置程序图标
        icon = QIcon("favicon.png")
        pgt.setWindowIcon(icon)
        pgt._setup_ui()

    def closeEvent(pgt, event):
        """程序退出时的处理"""
        # 清理临时文件夹
        temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.temp')
        if os.path.exists(temp_dir):
            for file in os.listdir(temp_dir):
                file_path = os.path.join(temp_dir, file)
                try:
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                except Exception as e:
                    print(f"清理临时文件失败: {e}")
        event.accept()

    def _setup_ui(pgt):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        image_processor = ImageProcessor()
        document_generator = DocumentGenerator(image_processor)
        layout.addWidget(DocGeneratorUI(image_processor, document_generator))
        pgt.setLayout(layout)
        pgt.setStyleSheet("""
            QWidget { 
                background: #FAFAFA;
                font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
            }
        """)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())