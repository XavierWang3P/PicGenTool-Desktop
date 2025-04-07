import sys
import os
import logging
from typing import List, Optional
import subprocess

from PyQt6.QtWidgets import (QWidget, QPushButton, QLineEdit, QLabel,
                             QHBoxLayout, QVBoxLayout, QFileDialog, QMessageBox,
                             QApplication, QCalendarWidget, QCheckBox)
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon
from PyQt6.QtCore import Qt, QSize, QDate, QTemporaryFile
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
from PIL import Image
from docx2pdf import convert

# 导入配置文件
from config import *

class ImageProcessor:
    """处理图片相关的逻辑
    
    负责图片的添加、压缩和清理等操作。支持拖拽和文件选择两种方式添加图片。
    图片会被压缩和调整大小以适应文档模板的要求。
    
    Attributes:
        images: 存储原始图片路径的列表
        compressed_images: 存储压缩后图片路径的列表
        logger: 日志记录器实例
    """
    def __init__(pgt):
        pgt.images: List[str] = []
        pgt.logger = logging.getLogger(__name__)
        pgt.compressed_images: List[str] = []

    def add_images(pgt, files: List[str]) -> bool:
        """添加图片并返回是否成功
        
        验证并添加用户选择的图片文件到处理队列中。会检查文件格式是否支持，
        以及是否超出最大图片数量限制。
        
        Args:
            files: 待添加的图片文件路径列表
            
        Returns:
            bool: 添加是否成功
        """
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
        """检查是否为支持的图片格式
        
        检查文件扩展名是否在支持的图片格式列表中。
        支持的格式定义在配置文件的SUPPORTED_FORMATS中。
        
        Args:
            path: 待检查的文件路径
            
        Returns:
            bool: 是否为支持的图片格式
        """
        ext = os.path.splitext(path)[1].lower()
        return ext in SUPPORTED_FORMATS

    def compress_images(pgt) -> List[str]:
        """压缩所有图片并返回压缩后的图片路径列表
        
        将所有添加的图片按照配置的尺寸和质量要求进行压缩处理。
        包括调整图片尺寸、裁剪以保持比例、压缩质量等操作。
        处理后的图片保存为临时文件。
        
        Returns:
            List[str]: 压缩后的图片文件路径列表
        
        Raises:
            ValueError: 图片格式不支持或处理过程出错
            IOError: 文件操作失败
        """
        if not pgt.images:
            return []

        pgt.compressed_images.clear()
        for img_path in pgt.images:
            try:
                # 打开图片
                with Image.open(img_path) as img:
                    # 验证图片格式
                    if img.format not in ['JPEG', 'PNG', 'GIF', 'BMP']:
                        raise ValueError(f"不支持的图片格式: {img.format}")

                    # 创建带扩展名的临时文件
                    temp_file = QTemporaryFile()
                    temp_file.setAutoRemove(True)
                    if not temp_file.open():
                        raise IOError("无法创建临时文件，请检查系统临时目录权限")
                    # 确保临时文件有.jpg扩展名
                    compressed_path = temp_file.fileName() + ".jpg"
                    temp_file.close()  # 关闭临时文件以便后续写入

                    # 根据配置计算目标尺寸（转换为像素）
                    pixels_per_cm = DPI / 2.54  # 将DPI转换为每厘米像素数
                    target_width = int(TARGET_WIDTH_CM * pixels_per_cm)
                    target_height = int(TARGET_HEIGHT_CM * pixels_per_cm)
                    target_ratio = target_width / target_height

                    # 获取原始图片尺寸
                    try:
                        width, height = img.size
                        if width == 0 or height == 0:
                            raise ValueError("图片尺寸无效")
                    except Exception as e:
                        raise ValueError(f"无法获取图片尺寸: {str(e)}")

                    current_ratio = width / height

                    # 根据目标比例调整图片
                    try:
                        if current_ratio > target_ratio:
                            # 图片太宽，需要裁剪宽度
                            new_width = int(height * target_ratio)
                            left = (width - new_width) // 2
                            img = img.crop((left, 0, left + new_width, height))
                        elif current_ratio < target_ratio:
                            # 图片太高，需要裁剪高度
                            new_height = int(width / target_ratio)
                            top = (height - new_height) // 2
                            img = img.crop((0, top, width, top + new_height))
                    except Exception as e:
                        raise ValueError(f"图片裁剪失败: {str(e)}")

                    # 调整到目标尺寸
                    try:
                        img = img.resize((target_width, target_height), Image.Resampling.LANCZOS)
                    except Exception as e:
                        raise ValueError(f"图片缩放失败: {str(e)}")

                    # 如果图片模式是RGBA，转换为RGB
                    if img.mode == 'RGBA':
                        img = img.convert('RGB')
                    
                    # 保存压缩后的图片
                    try:
                        img.save(
                            compressed_path,
                            format='JPEG',
                            quality=60,  # 压缩质量
                            optimize=True  # 优化文件大小
                        )
                    except Exception as e:
                        raise IOError(f"图片保存失败: {str(e)}")

                    pgt.compressed_images.append(compressed_path)
                    pgt.logger.info(f"成功压缩图片: {img_path}")

            except (IOError, OSError) as e:
                error_msg = f"图片文件访问错误: {str(e)}"
                pgt.logger.error(f"{error_msg}, 图片路径: {img_path}")
                raise ValueError(error_msg)
            except ValueError as e:
                error_msg = str(e)
                pgt.logger.error(f"{error_msg}, 图片路径: {img_path}")
                raise ValueError(error_msg)
            except Exception as e:
                error_msg = f"图片处理过程中发生未知错误: {str(e)}"
                pgt.logger.error(f"{error_msg}, 图片路径: {img_path}")
                raise ValueError(error_msg)

        return pgt.compressed_images

class DocumentGenerator:
    """处理文档生成的逻辑"""
    def __init__(pgt, image_processor: ImageProcessor, pdf_checkbox=None, open_file_checkbox=None):
        pgt.image_processor = image_processor
        pgt.pdf_checkbox = pdf_checkbox
        pgt.open_file_checkbox = open_file_checkbox
        pgt.logger = logging.getLogger(__name__)

    def generate_document(pgt, title: str, location: str, date: QDate) -> Optional[str]:
        """生成文档并返回保存路径"""
        try:
            if not pgt._validate_inputs(title, location):
                if not title:
                    raise ValueError("请输入活动标题")
                if not location:
                    raise ValueError("请输入活动地点")
                if not pgt.image_processor.images:
                    raise ValueError("请选择至少一张图片")
                return None

            template_path = pgt._get_template_path()
            if not template_path:
                raise FileNotFoundError(f"未找到对应{len(pgt.image_processor.images)}张图片的模板文件，请检查templates目录")

            try:
                doc = DocxTemplate(template_path)
            except Exception as e:
                raise IOError(f"模板文件打开失败: {str(e)}")

            try:
                context = pgt._prepare_context(doc, title, location, date)
                doc.render(context)
            except Exception as e:
                raise ValueError(f"文档渲染失败: {str(e)}")

            save_path = pgt._get_save_path(title, date)
            if not save_path:
                return None

            try:
                doc.save(save_path)
            except PermissionError:
                raise IOError("无法保存文档，请检查文件权限或是否被其他程序占用")
            except Exception as e:
                raise IOError(f"文档保存失败: {str(e)}")

            # 生成PDF文件
            if pgt.pdf_checkbox and pgt.pdf_checkbox.isChecked():
                try:
                    pdf_path = save_path.replace('.docx', '.pdf')
                    convert(save_path, pdf_path)
                except Exception as e:
                    pgt.logger.error(f"PDF生成失败: {str(e)}")
                    raise ValueError(f"PDF文档生成失败: {str(e)}")
            
            # 如果勾选了自动打开文档，则打开生成的文档
            if pgt.open_file_checkbox and pgt.open_file_checkbox.isChecked():
                try:
                    if sys.platform == 'darwin':  # macOS
                        subprocess.run(['open', save_path])
                    elif sys.platform == 'win32':  # Windows
                        os.startfile(save_path)
                    else:  # Linux
                        subprocess.run(['xdg-open', save_path])
                except Exception as e:
                    error_msg = f"无法自动打开文档: {str(e)}"
                    pgt.logger.error(error_msg)
                    raise IOError(error_msg)
            
            return save_path

        except (ValueError, FileNotFoundError, IOError) as e:
            pgt.logger.error(str(e))
            raise
        except Exception as e:
            error_msg = f"文档生成过程中发生未知错误: {str(e)}"
            pgt.logger.error(error_msg)
            raise ValueError(error_msg)

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
        
        # 获取第一张照片所在的文件夹作为默认保存位置
        default_dir = os.path.dirname(pgt.image_processor.images[0]) if pgt.image_processor.images else ''
        
        return QFileDialog.getSaveFileName(
            None,
            '保存文档',
            os.path.join(default_dir, file_name),
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
        # 创建标签和输入框
        pgt.title_label = QLabel("活动标题")
        pgt.title_input = QLineEdit(placeholderText="请输入活动标题")
        pgt.location_label = QLabel("活动地点")
        pgt.location_input = QLineEdit(placeholderText="请输入活动地点")
        
        # 创建日历组件
        pgt.calendar = QCalendarWidget()
        pgt._configure_calendar()
        
        # 创建复选框
        pgt.pdf_checkbox = QCheckBox("同时生成 PDF 文档")
        pgt.open_file_checkbox = QCheckBox("生成后打开文档")
        pgt.open_file_checkbox.setChecked(True)
        
        # 创建图片计数标签
        pgt.image_count_label = QLabel("已选择 0 张图片")
        
        # 创建按钮
        pgt.image_button = QPushButton("选择照片")
        pgt.clear_button = QPushButton("清空内容")
        pgt.generate_button = QPushButton("生成文档")
        
        # 设置按钮大小策略
        for button in [pgt.image_button, pgt.clear_button, pgt.generate_button]:
            button.setFixedHeight(42)
            button.setFixedWidth(100)
        
    def _configure_calendar(pgt):
        """配置日历组件样式"""
        pgt.calendar.setGridVisible(True)
        pgt.calendar.setMinimumDate(QDate(2025, 1, 1))
        pgt.calendar.setMaximumDate(QDate(2099, 12, 31))
        pgt.calendar.setFirstDayOfWeek(Qt.DayOfWeek.Monday)
        pgt.calendar.setFixedHeight(230)
        pgt.calendar.setFixedWidth(305)
        pgt.calendar.setHorizontalHeaderFormat(QCalendarWidget.HorizontalHeaderFormat.ShortDayNames)
        pgt.calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        pgt.calendar.setStyleSheet("""
            QCalendarWidget QWidget { background-color: white; }
            QCalendarWidget QToolButton { color: #333; padding: 4px; }
            QCalendarWidget QMenu { width: 150px; left: 20px; color: #333; }
            QCalendarWidget QSpinBox { width: 60px; font-size: 13px; }
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
            pgt._show_message("提示", "最多只能选择30张图片！", QMessageBox.Icon.Warning)

    def dropEvent(pgt, event: QDropEvent):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files and pgt.image_processor.add_images(files):
            pgt._update_image_count()
        else:
            pgt._show_message("提示", "最多只能选择30张图片！", QMessageBox.Icon.Warning)

    # 在_reset_ui中也需要清空预览
    def _reset_ui(pgt):
        """重置界面状态"""
        pgt.title_input.clear()
        pgt.location_input.clear()
        pgt.calendar.setSelectedDate(QDate.currentDate())
        pgt.image_processor.clear_images()
        pgt._update_image_count()

    def _create_input_layout(pgt):
        """创建输入区域布局"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        
        # 标题区域
        title_layout = QHBoxLayout()
        title_layout.addWidget(pgt.title_label)
        title_layout.addWidget(pgt.title_input)
        layout.addLayout(title_layout)
        
        # 地点区域
        location_layout = QHBoxLayout()
        location_layout.addWidget(pgt.location_label)
        location_layout.addWidget(pgt.location_input)
        layout.addLayout(location_layout)
        
        # 设置标签宽度
        pgt.title_label.setFixedWidth(60)
        pgt.location_label.setFixedWidth(60)
        
        # 日历组件
        layout.addWidget(pgt.calendar)
        return layout

    def _create_image_count_layout(pgt):
        """创建图片计数布局"""
        layout = QVBoxLayout()
        layout.setSpacing(8)
        layout.setContentsMargins(0, 2, 0, 2)
        
        # 复选框区域
        checkbox_layout = QHBoxLayout()
        checkbox_layout.addWidget(pgt.pdf_checkbox)
        checkbox_layout.addWidget(pgt.open_file_checkbox)
        layout.addLayout(checkbox_layout)
        
        # 图片计数标签居中显示
        count_layout = QHBoxLayout()
        count_layout.addStretch()
        count_layout.addWidget(pgt.image_count_label)
        count_layout.addStretch()
        layout.addLayout(count_layout)
        
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
        """设置组件样式
        
        应用UI_STYLES中定义的样式到各个UI组件。包括标签、输入框、按钮等的样式设置。
        样式包括字体、颜色、边框、背景等属性。
        
        Returns:
            None
        """
        style_sheet = f"""
            QLabel {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['label'].items())} }}
            QLineEdit {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['line_edit'].items())} }}
            QLineEdit:focus {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['line_edit_focus'].items())} }}
            QPushButton {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['button'].items())} }}
            QPushButton:hover {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['button_hover'].items())} }}
            QPushButton#generateBtn {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['generate_button'].items())} }}
            QPushButton#generateBtn:hover {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['generate_button_hover'].items())} }}
            QPushButton#clearBtn {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['clear_button'].items())} }}
            QCheckBox {{ {'; '.join(f'{k}: {v}' for k, v in UI_STYLES['checkbox'].items())} }}
        """
        pgt.setStyleSheet(style_sheet)
        pgt.generate_button.setObjectName("generateBtn")
        pgt.clear_button.setObjectName("clearBtn")

    def _handle_image_selection(pgt):
        """处理图片选择
        
        打开文件选择对话框，允许用户选择图片文件。支持多选功能。
        选择完成后会更新图片计数显示，如果超出最大限制会显示警告信息。
        
        Returns:
            None
        """
        files, _ = QFileDialog.getOpenFileNames(
            parent=pgt,
            caption="选择图片文件",
            filter=f"图片文件 ({' '.join(f'*{fmt}' for fmt in SUPPORTED_FORMATS)})"
        )
        if files and pgt.image_processor.add_images(files):
            pgt._update_image_count()
        else:
            pgt._show_message("提示", f"最多只能选择{MAX_IMAGES}张图片！", QMessageBox.Icon.Warning)

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
                pgt._reset_ui()
        except ValueError as e:
            # 用户输入或图片处理相关的错误
            pgt._show_message("输入错误", str(e), QMessageBox.Icon.Warning)
        except FileNotFoundError as e:
            # 模板文件相关的错误
            pgt._show_message("模板错误", str(e), QMessageBox.Icon.Critical)
        except IOError as e:
            # 文件操作相关的错误
            pgt._show_message("文件操作错误", str(e), QMessageBox.Icon.Critical)
        except Exception as e:
            # 未预期的错误
            pgt._show_message("系统错误", f"发生未知错误: {str(e)}\n请联系技术支持", QMessageBox.Icon.Critical)

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
        pgt.setGeometry(300, 200, 380, 450)
        pgt.setMinimumSize(330, 450)
        pgt.setMaximumSize(330, 450)
        # 设置程序图标
        icon = QIcon("favicon.png")
        pgt.setWindowIcon(icon)
        pgt._setup_ui()

    def closeEvent(pgt, event):
        """程序退出时的处理"""
        event.accept()

    def _setup_ui(pgt):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        image_processor = ImageProcessor()
        
        # 创建UI组件
        ui = DocGeneratorUI(image_processor, None)
        
        # 创建文档生成器并传递checkbox
        document_generator = DocumentGenerator(image_processor, ui.pdf_checkbox, ui.open_file_checkbox)
        ui.document_generator = document_generator
        
        layout.addWidget(ui)
        pgt.setLayout(layout)
        pgt.setStyleSheet("""
            QWidget { 
                background: #FAFAFA;
                font-family: "Microsoft YaHei", "PingFang SC", sans-serif;
            }
        """)

# 在文件开头设置日志配置
import logging
import os

# 设置日志
def setup_logging():
    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
    os.makedirs(log_dir, exist_ok=True)
    
    log_file = os.path.join(log_dir, 'app.log')
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    
    return logging.getLogger(__name__)

# 在主函数中调用
if __name__ == "__main__":
    logger = setup_logging()
    try:
        app = QApplication(sys.argv)
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logger.critical(f"应用崩溃: {str(e)}", exc_info=True)