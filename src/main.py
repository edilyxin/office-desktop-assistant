import sys
import os
import requests

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QLabel, QProgressBar,
    QMessageBox, QAction, QStatusBar, QGroupBox, QCheckBox,
    QSplitter, QScrollArea, QTabWidget
)
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor
from PyQt5.QtCore import Qt, QRect, QPoint, QThread, pyqtSignal

from .ocr_api import PaddleOCRVL
from .screenshot import Screenshot
from .file_utils import FileUtils
from .clipboard_manager import ClipboardManager
from .markdown_manager import MarkdownManager
# from .config_manager import load_config
from .log_manager import logger
# 在 main.py 文件顶部添加 requests 导入


class ScreenshotWindow(QWidget):
    """截图窗口，用于选择截图区域"""
    screenshot_taken = pyqtSignal(bytes)

    def __init__(self, screenshot):
        super().__init__()
        self.screenshot = screenshot
        self.start_pos = QPoint()
        self.end_pos = QPoint()
        self.is_dragging = False

        self.init_ui()

    def init_ui(self):
        """初始化截图窗口"""
        # 全屏显示
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.setWindowState(Qt.WindowFullScreen)

        # 截取全屏
        self.fullscreen_bytes = self.screenshot.capture_fullscreen()
        self.pixmap = QPixmap()
        self.pixmap.loadFromData(self.fullscreen_bytes)

    def paintEvent(self, event):
        """绘制截图区域"""
        painter = QPainter(self)

        # 绘制半透明背景
        painter.fillRect(self.rect(), QColor(0, 0, 0, 128))

        # 绘制原图
        painter.drawPixmap(0, 0, self.pixmap)

        # 绘制选择区域
        if self.is_dragging:
            rect = self.get_rect()
            # 绘制选择区域边框
            pen = QPen(Qt.red, 2, Qt.SolidLine)
            painter.setPen(pen)
            painter.drawRect(rect)

            # 绘制选择区域内部（高亮）
            painter.fillRect(rect, QColor(255, 255, 255, 128))

    def mousePressEvent(self, event):
        """鼠标按下事件"""
        if event.button() == Qt.LeftButton:
            self.start_pos = event.pos()
            self.end_pos = event.pos()
            self.is_dragging = True

    def mouseMoveEvent(self, event):
        """鼠标移动事件"""
        if self.is_dragging:
            self.end_pos = event.pos()
            self.update()

    def mouseReleaseEvent(self, event):
        """鼠标释放事件"""
        if event.button() == Qt.LeftButton and self.is_dragging:
            self.is_dragging = False
            rect = self.get_rect()

            # 截取选择区域
            if rect.width() > 10 and rect.height() > 10:  # 最小选择区域
                # 转换为(left, top, width, height)
                region = (rect.left(), rect.top(), rect.width(), rect.height())
                img_bytes = self.screenshot.capture_region(region)
                self.screenshot_taken.emit(img_bytes)

            self.close()

    def keyPressEvent(self, event):
        """键盘事件"""
        if event.key() == Qt.Key_Escape:
            # 取消截图
            self.close()

    def get_rect(self):
        """获取选择区域的矩形"""
        left = min(self.start_pos.x(), self.end_pos.x())
        top = min(self.start_pos.y(), self.end_pos.y())
        width = abs(self.start_pos.x() - self.end_pos.x())
        height = abs(self.start_pos.y() - self.end_pos.y())
        return QRect(left, top, width, height)


class OCRThread(QThread):
    """OCR识别线程"""
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, ocr_client, file_path=None, image_bytes=None, options=None):
        super().__init__()
        self.ocr_client = ocr_client
        self.file_path = file_path
        self.image_bytes = image_bytes
        self.options = options or {}

    def run(self):
        try:
            if self.file_path:
                # 从文件识别
                result = self.ocr_client.recognize(
                    self.file_path,
                    use_doc_orientation_classify=self.options.get(
                        'use_doc_orientation_classify', False),
                    use_doc_unwarping=self.options.get(
                        'use_doc_unwarping', False),
                    use_chart_recognition=self.options.get(
                        'use_chart_recognition', False),
                    prettifyMarkdown=True
                )
            elif self.image_bytes:
                # 从字节流识别
                result = self.ocr_client.recognize_image_from_bytes(
                    self.image_bytes,
                    use_doc_orientation_classify=self.options.get(
                        'use_doc_orientation_classify', False),
                    use_doc_unwarping=self.options.get(
                        'use_doc_unwarping', False),
                    use_chart_recognition=self.options.get(
                        'use_chart_recognition', False),
                    prettifyMarkdown=True
                )
            else:
                raise ValueError("必须提供file_path或image_bytes")

            self.finished.emit(result)
        except Exception as e:
            self.error.emit(str(e))


class PaddleOCRVLAssistant(QMainWindow):
    """PaddleOCR-vl桌面助手主窗口"""

    def __init__(self):
        super().__init__()
        logger.info("PaddleOCR-vl桌面助手启动")
        self.ocr_client = PaddleOCRVL()
        self.screenshot = Screenshot()
        self.clipboard_manager = ClipboardManager()
        self.markdown_manager = MarkdownManager()
        self.current_ocr_result = None
        self.current_image_bytes = None

        self.init_ui()

    def init_ui(self):
        """初始化UI界面"""
        self.setWindowTitle("办公桌面助手")
        self.setGeometry(100, 100, 1200, 600)

        # 创建菜单
        self.create_menu()

        # 创建中心部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 主布局
        main_layout = QVBoxLayout(central_widget)

        # 创建Tab控件
        self.tab_widget = QTabWidget()

        # 创建第一个Tab：文件识别
        self.create_file_recognition_tab()

        # 创建第二个Tab：文件样式
        self.create_file_style_tab()

        main_layout.addWidget(self.tab_widget)

        # 状态栏
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("就绪")

    def create_file_recognition_tab(self):
        """创建文件识别Tab页"""
        # 创建Tab页面
        tab_widget = QWidget()
        tab_layout = QVBoxLayout(tab_widget)

        # 功能按钮区域
        button_group = QGroupBox("功能选择")
        button_layout = QHBoxLayout()

        # 截图按钮
        self.screenshot_btn = QPushButton("截图识别")
        self.screenshot_btn.clicked.connect(self.capture_screenshot)
        button_layout.addWidget(self.screenshot_btn)

        # 文件上传按钮
        self.upload_btn = QPushButton("上传文件")
        self.upload_btn.clicked.connect(self.upload_file)
        button_layout.addWidget(self.upload_btn)

        button_group.setLayout(button_layout)
        tab_layout.addWidget(button_group)

        # 选项设置区域
        options_group = QGroupBox("识别选项")
        options_layout = QHBoxLayout()

        # 文档方向分类
        self.orientation_check = QCheckBox("文档方向分类")
        options_layout.addWidget(self.orientation_check)

        # 文档矫正
        self.unwarping_check = QCheckBox("文档矫正")
        options_layout.addWidget(self.unwarping_check)

        # 图表识别
        self.chart_check = QCheckBox("图表识别")
        options_layout.addWidget(self.chart_check)

        options_group.setLayout(options_layout)
        tab_layout.addWidget(options_group)

        # 水平分割布局，左侧文本结果，右侧图片
        splitter = QSplitter(Qt.Horizontal)

        # 左侧：结果显示区域
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        # 结果显示（使用WebEngineView渲染Markdown）
        self.result_view = QWebEngineView()
        self.result_view.setMinimumSize(600, 400)
        left_layout.addWidget(self.result_view)

        # 操作按钮区域
        action_layout = QHBoxLayout()

        # 复制到剪贴板
        self.copy_btn = QPushButton("复制到剪贴板")
        self.copy_btn.clicked.connect(self.copy_to_clipboard)
        self.copy_btn.setEnabled(False)
        action_layout.addWidget(self.copy_btn)

        # 保存为markdown
        self.save_btn = QPushButton("保存为Markdown")
        self.save_btn.clicked.connect(self.save_as_markdown)
        self.save_btn.setEnabled(False)
        action_layout.addWidget(self.save_btn)

        left_layout.addLayout(action_layout)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        left_layout.addWidget(self.progress_bar)

        splitter.addWidget(left_widget)

        # 右侧：图片显示区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        # 图片标签
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(400, 400)

        # 滚动区域
        scroll_area = QScrollArea()
        scroll_area.setWidget(self.image_label)
        scroll_area.setWidgetResizable(True)

        right_layout.addWidget(scroll_area)
        splitter.addWidget(right_widget)

        # 设置分割比例
        splitter.setSizes([600, 600])

        tab_layout.addWidget(splitter)

        # 添加Tab页
        self.tab_widget.addTab(tab_widget, "文件识别")

    def create_file_style_tab(self):
        """创建文件样式Tab页"""
        # 创建Tab页面
        tab_widget = QWidget()
        tab_layout = QVBoxLayout(tab_widget)

        # 样式模板选择区域
        template_group = QGroupBox("样式模板")
        template_layout = QHBoxLayout()

        # 模板文件路径显示
        self.template_label = QLabel("未选择模板文件")
        template_layout.addWidget(self.template_label)

        # 选择模板按钮
        self.select_template_btn = QPushButton("选择模板")
        self.select_template_btn.clicked.connect(self.select_template)
        template_layout.addWidget(self.select_template_btn)

        template_group.setLayout(template_layout)
        tab_layout.addWidget(template_group)

        # 目标文件上传区域
        target_group = QGroupBox("目标文件")
        target_layout = QHBoxLayout()

        # 目标文件路径显示
        self.target_label = QLabel("未选择目标文件")
        target_layout.addWidget(self.target_label)

        # 上传目标文件按钮
        self.upload_target_btn = QPushButton("上传文件")
        self.upload_target_btn.clicked.connect(self.upload_target_file)
        target_layout.addWidget(self.upload_target_btn)

        target_group.setLayout(target_layout)
        tab_layout.addWidget(target_group)

        # 处理按钮区域
        process_layout = QHBoxLayout()

        # 处理文件按钮
        self.process_btn = QPushButton("处理文件")
        self.process_btn.clicked.connect(self.process_file)
        self.process_btn.setEnabled(False)
        process_layout.addWidget(self.process_btn)

        # 保存结果按钮
        self.save_result_btn = QPushButton("保存结果")
        self.save_result_btn.clicked.connect(self.save_result)
        self.save_result_btn.setEnabled(False)
        process_layout.addWidget(self.save_result_btn)

        tab_layout.addLayout(process_layout)

        # 结果显示区域
        self.style_result_label = QLabel("处理结果将显示在这里")
        self.style_result_label.setAlignment(Qt.AlignCenter)
        self.style_result_label.setMinimumHeight(200)
        tab_layout.addWidget(self.style_result_label)

        # 添加Tab页
        self.tab_widget.addTab(tab_widget, "文件样式")

    def create_menu(self):
        """创建菜单"""
        menubar = self.menuBar()

        # 文件菜单
        file_menu = menubar.addMenu("文件")

        # 退出动作
        exit_action = QAction("退出", self)
        exit_action.setShortcut("Ctrl+Q")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # 帮助菜单
        help_menu = menubar.addMenu("帮助")

        # 关于动作
        about_action = QAction("关于", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def capture_screenshot(self):
        """截图识别"""
        try:
            logger.info("开始截图识别")
            self.status_bar.showMessage("正在准备截图...")

            # 创建截图窗口
            self.screenshot_window = ScreenshotWindow(self.screenshot)
            self.screenshot_window.screenshot_taken.connect(
                self.on_screenshot_taken)
            self.screenshot_window.show()
        except Exception as e:
            logger.error(f"截图失败: {str(e)}")
            QMessageBox.critical(self, "错误", f"截图失败: {str(e)}")
            self.status_bar.showMessage("截图失败")

    def on_screenshot_taken(self, img_bytes):
        """截图完成回调"""
        logger.info("截图完成，开始OCR识别")
        self.status_bar.showMessage("截图完成，正在识别...")
        self.start_ocr_recognition(image_bytes=img_bytes)

    def upload_file(self):
        """上传文件识别"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择文件", "", "支持的文件 (*.png *.jpg *.jpeg *.bmp *.gif *.tiff *.pdf)"
        )

        if file_path:
            logger.info(f"上传文件: {file_path}")
            # 检查文件格式
            if not FileUtils.is_supported_file(file_path):
                logger.warning(f"不支持的文件格式: {file_path}")
                QMessageBox.warning(self, "警告", "不支持的文件格式")
                return

            # 开始OCR识别
            self.start_ocr_recognition(file_path=file_path)

    def start_ocr_recognition(self, file_path=None, image_bytes=None):
        """开始OCR识别"""
        # 获取识别选项
        options = {
            'use_doc_orientation_classify': self.orientation_check.isChecked(),
            'use_doc_unwarping': self.unwarping_check.isChecked(),
            'use_chart_recognition': self.chart_check.isChecked()
        }

        logger.debug(f"OCR识别选项: {options}")

        # 保存当前图片字节
        if image_bytes:
            self.current_image_bytes = image_bytes
            logger.info("使用截图进行OCR识别")
        elif file_path:
            # 从文件读取图片字节
            self.current_image_bytes = FileUtils.read_file_bytes(file_path)
            logger.info(f"使用文件进行OCR识别: {file_path}")

        # 显示图片（等比例缩放）
        if self.current_image_bytes:
            pixmap = QPixmap()
            pixmap.loadFromData(self.current_image_bytes)

            # 等比例缩放图片，保持原始比例
            scaled_pixmap = pixmap.scaled(
                self.image_label.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            )

            self.image_label.setPixmap(scaled_pixmap)
            logger.debug(
                f"图片显示完成，原始尺寸: {pixmap.size()}, 缩放后尺寸: {scaled_pixmap.size()}")

        # 禁用按钮
        self.screenshot_btn.setEnabled(False)
        self.upload_btn.setEnabled(False)
        self.copy_btn.setEnabled(False)
        self.save_btn.setEnabled(False)

        # 显示进度条
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # 无限循环
        self.status_bar.showMessage("正在识别...")

        # 清空结果
        self.result_view.setHtml("<html><body><h3>正在识别...</h3></body></html>")

        # 创建并启动OCR线程
        self.ocr_thread = OCRThread(
            self.ocr_client, file_path, image_bytes, options)
        self.ocr_thread.finished.connect(self.on_ocr_finished)
        self.ocr_thread.error.connect(self.on_ocr_error)
        self.ocr_thread.start()
        logger.info("OCR识别线程已启动")

    def on_ocr_finished(self, result):
        """OCR识别完成回调"""
        logger.info("OCR识别完成")
        # 保存结果
        self.current_ocr_result = result

        # 提取markdown文本
        markdown_text = self.markdown_manager.extract_markdown_text(result)

        processed_markdown = self.process_images(result, markdown_text)

        # 记录完整的识别结果到日志
        logger.debug(f"识别结果: {processed_markdown}")
        logger.info(f"识别结果长度: {len(processed_markdown)} 字符")

        # 将Markdown转换为HTML并显示
        html_content = self.markdown_to_html(processed_markdown)
        self.result_view.setHtml(html_content)

        # 启用按钮
        self.screenshot_btn.setEnabled(True)
        self.upload_btn.setEnabled(True)
        self.copy_btn.setEnabled(True)
        self.save_btn.setEnabled(True)

        # 隐藏进度条
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage("识别完成")

    def process_images(self, ocr_result, markdown_text):
        """处理识别结果中的图片，保存到本地imgs文件夹"""
        # 创建imgs目录
        imgs_dir = "imgs"
        os.makedirs(imgs_dir, exist_ok=True)

        # 处理每个页面的结果
        for i, res in enumerate(ocr_result["layoutParsingResults"]):
            # 处理markdown中的图片
            for img_path, img_url in res["markdown"]["images"].items():
                # 构建本地图片路径
                local_img_path = os.path.join(
                    imgs_dir, os.path.basename(img_path))

                # 下载图片
                try:
                    img_response = requests.get(img_url)
                    if img_response.status_code == 200:
                        with open(local_img_path, "wb") as img_file:
                            img_file.write(img_response.content)
                        logger.info(f"图片已保存到本地: {local_img_path}")

                        # 替换markdown中的图片路径为本地路径
                        markdown_text = markdown_text.replace(
                            img_path, local_img_path)
                except Exception as e:
                    logger.error(f"下载图片失败: {img_url}, 错误: {str(e)}")

        return markdown_text

    def markdown_to_html(self, markdown_text):
        """将Markdown文本转换为HTML"""
        # 简单的HTML模板，不使用f-string，避免CSS中的{}导致语法错误
        html_template = '''<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>OCR识别结果</title>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 100%;
            padding: 20px;
            margin: 0;
            background-color: #f9f9f9;
        }
        h1, h2, h3, h4, h5, h6 {
            color: #2c3e50;
            margin-top: 1.5em;
            margin-bottom: 0.5em;
        }
        p {
            margin: 0.8em 0;
        }
        code {
            background-color: #f0f0f0;
            padding: 2px 6px;
            border-radius: 3px;
            font-family: Consolas, Monaco, 'Andale Mono', monospace;
        }
        pre {
            background-color: #f0f0f0;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            margin: 1em 0;
        }
        pre code {
            background-color: transparent;
            padding: 0;
        }
        ul, ol {
            margin: 1em 0;
            padding-left: 2em;
        }
        li {
            margin: 0.5em 0;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            margin: 1em 0;
        }
        th, td {
            border: 1px solid #ddd;
            padding: 8px;
            text-align: left;
        }
        th {
            background-color: #f2f2f2;
        }
        img {
            max-width: 100%;
            height: auto;
            margin: 1em 0;
        }
        blockquote {
            border-left: 4px solid #ddd;
            padding-left: 15px;
            margin: 1em 0;
            color: #666;
            font-style: italic;
        }
    </style>
</head>
<body>
'''

        # 简单的Markdown到HTML转换（仅处理基本格式）
        html_content = html_template + markdown_text + '''
</body>
</html>'''

        return html_content

    def on_ocr_error(self, error_msg):
        """OCR识别错误回调"""
        logger.error(f"OCR识别失败: {error_msg}")
        # 显示错误信息
        QMessageBox.critical(self, "识别错误", f"识别失败: {error_msg}")

        # 启用按钮
        self.screenshot_btn.setEnabled(True)
        self.upload_btn.setEnabled(True)

        # 隐藏进度条
        self.progress_bar.setVisible(False)
        self.status_bar.showMessage("识别失败")

    def copy_to_clipboard(self):
        """复制结果到剪贴板"""
        if self.current_ocr_result:
            markdown_text = self.markdown_manager.extract_markdown_text(
                self.current_ocr_result)
            self.clipboard_manager.copy_to_clipboard(markdown_text)
            logger.info("识别结果已复制到剪贴板")
            self.status_bar.showMessage("已复制到剪贴板")

    def save_as_markdown(self):
        """保存结果为markdown文件"""
        if self.current_ocr_result:
            # 选择保存目录
            output_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
            if output_dir:
                try:
                    logger.info(f"开始保存Markdown文件到: {output_dir}")
                    # 保存文件
                    saved_files = self.markdown_manager.save_markdown(
                        self.current_ocr_result, output_dir)
                    logger.info(f"已保存 {len(saved_files)} 个文件到: {output_dir}")
                    self.status_bar.showMessage(f"已保存 {len(saved_files)} 个文件")
                    QMessageBox.information(
                        self, "保存成功", f"已保存到: {output_dir}")
                except Exception as e:
                    logger.error(f"保存Markdown文件失败: {str(e)}")
                    QMessageBox.critical(self, "保存错误", f"保存失败: {str(e)}")

    def select_template(self):
        """选择样式模板文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择模板文件", "", "Word文件 (*.docx)"
        )

        if file_path:
            # 不复制模板文件，直接使用原始路径
            # 更新显示
            template_name = os.path.basename(file_path)
            self.template_label.setText(f"已选择模板: {template_name}")
            self.template_path = file_path

            # 检查是否可以启用处理按钮
            self.check_process_enabled()

            logger.info(f"已选择模板文件: {file_path}")

    def upload_target_file(self):
        """上传目标文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择目标文件", "", "Word文件 (*.docx)"
        )

        if file_path:
            # 更新显示
            self.target_label.setText(
                f"已选择目标文件: {os.path.basename(file_path)}")
            self.target_path = file_path

            # 检查是否可以启用处理按钮
            self.check_process_enabled()

            logger.info(f"已上传目标文件: {file_path}")

    def check_process_enabled(self):
        """检查是否可以启用处理按钮"""
        if hasattr(self, 'template_path') and hasattr(self, 'target_path'):
            self.process_btn.setEnabled(True)
        else:
            self.process_btn.setEnabled(False)

    def process_file(self):
        """处理文件，应用样式模板"""
        try:
            self.status_bar.showMessage("正在处理文件...")
            logger.info(f"开始处理文件: {self.target_path}")

            # 导入样式处理器
            from src.style_processor import WordStyleProcessor

            # 保存处理后的文件
            processed_name = f"processed_{os.path.basename(self.target_path)}"
            self.processed_path = os.path.join("output", processed_name)
            os.makedirs("output", exist_ok=True)

            # 创建样式处理器实例
            style_processor = WordStyleProcessor(
                template_path=self.template_path,
                target_path=self.target_path,
                output_path=self.processed_path
            )

            # 执行样式处理流程
            if style_processor.process():
                # 获取处理统计信息
                stats = style_processor.get_processing_stats()
                
                # 更新界面显示
                self.style_result_label.setText(
                    f"文件处理完成!\n"\
                    f"结果文件: {processed_name}\n"\
                    f"共处理段落: {stats['total_paragraphs']}，修改: {stats['processed_paragraphs']}\n"\
                    f"共处理表格: {stats['total_tables']}，修改: {stats['processed_tables']}\n"\
                    f"成功应用: {stats['successful_applications']}，失败: {stats['failed_applications']}")
                self.save_result_btn.setEnabled(True)

                self.status_bar.showMessage("文件处理完成")
                logger.info(f"文件处理完成: {self.processed_path}")
            else:
                raise Exception("样式处理流程执行失败")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"处理文件失败: {str(e)}")
            self.status_bar.showMessage("处理文件失败")
            logger.error(f"处理文件失败: {str(e)}", exc_info=True)

    def save_result(self):
        """保存处理结果"""
        if not hasattr(self, 'processed_path'):
            QMessageBox.warning(self, "警告", "没有可保存的处理结果")
            return

        # 选择保存目录
        save_dir = QFileDialog.getExistingDirectory(self, "选择保存目录")
        if save_dir:
            try:
                import shutil

                # 获取处理后的文件名
                processed_name = os.path.basename(self.processed_path)
                save_path = os.path.join(save_dir, processed_name)

                # 复制文件
                shutil.copy2(self.processed_path, save_path)

                QMessageBox.information(self, "成功", f"结果已保存到: {save_path}")
                logger.info(f"结果已保存到: {save_path}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存结果失败: {str(e)}")
                logger.error(f"保存结果失败: {str(e)}")

    def show_about(self):
        """显示关于对话框"""
        QMessageBox.about(self, "关于办公桌面助手",
                          "办公桌面助手 v0.5\n\n"
                          "基于PaddleOCR-vl API开发的桌面OCR工具，"
                          "支持截图识别和文件上传识别，"
                          "识别结果可复制到剪贴板或保存为Markdown文件。")


def main():
    """主函数"""
    app = QApplication(sys.argv)
    window = PaddleOCRVLAssistant()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
