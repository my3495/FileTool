"""
主窗口模块
程序的主界面
"""
import os
import sys
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QPushButton, QLabel, QFileDialog, QLineEdit, QMessageBox,
    QProgressBar, QTextEdit, QCheckBox, QGroupBox, QFormLayout,
    QComboBox, QSplitter, QApplication, QFrame, QMenu
)
from PySide6.QtCore import Qt, Signal
from PySide6.QtGui import QIcon, QFont, QPixmap, QDragEnterEvent, QDropEvent, QAction

from src.gui.word_extractor_tab import WordExtractorTab
from src.gui.word_filler_tab import WordFillerTab
from src.utils.logger import setup_logger
from src.gui.styles import AppStyles
from src.utils.config_manager import ConfigManager
from src.gui.theme_manager import ThemeManager


class MainWindow(QMainWindow):
    """主窗口类"""
    
    def __init__(self):
        """初始化主窗口"""
        super().__init__()
        
        # 初始化配置管理器
        self.config_manager = ConfigManager()
        
        # 初始化主题管理器
        self.theme_manager = ThemeManager(self.config_manager)
        
        # 设置日志
        self.logger = setup_logger()
        
        # 设置窗口
        self.setWindowTitle("Word数据提取与填充工具")
        self.setMinimumSize(1000, 788)
        
        # 启用拖放
        self.setAcceptDrops(True)
        
        # 应用主题
        theme_id = self.config_manager.get("theme", "default")
        self.theme_manager.apply_theme(theme_id)
        
        # 应用样式
        self.setStyleSheet(AppStyles.get_main_style())
        
        # 创建UI
        self._create_ui()
        
        self.logger.info("主窗口已初始化")
    
    def _create_ui(self):
        """创建用户界面"""
        # 主布局
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # 创建左侧导航栏
        self._create_sidebar(main_layout)
        
        # 创建右侧内容区
        self._create_content_area(main_layout)
    
    def _create_sidebar(self, main_layout):
        """
        创建左侧导航栏
        
        Args:
            main_layout: 主布局
        """
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setMaximumWidth(220)
        sidebar.setMinimumWidth(180)
        sidebar_layout = QVBoxLayout(sidebar)
        sidebar_layout.setContentsMargins(15, 20, 15, 20)
        sidebar_layout.setSpacing(10)
        
        # 应用侧边栏样式
        sidebar.setStyleSheet(AppStyles.get_sidebar_style())
        
        # 标题
        title_label = QLabel("Word工具")
        title_label.setObjectName("title_label")
        title_label.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(title_label)
        
        # 添加分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: rgba(255, 255, 255, 0.2);")
        separator.setMaximumHeight(1)
        sidebar_layout.addWidget(separator)
        sidebar_layout.addSpacing(20)
        
        # 添加导航按钮
        self.extract_btn = QPushButton("  数据提取")
        self.fill_btn = QPushButton("  数据填充")
        self.settings_btn = QPushButton("  设置")
        self.about_btn = QPushButton("  关于")
        
        # 设置对象名称
        self.extract_btn.setObjectName("nav_button")
        self.fill_btn.setObjectName("nav_button")
        self.settings_btn.setObjectName("nav_button")
        self.about_btn.setObjectName("nav_button")
        
        # 设置为可选中
        self.extract_btn.setCheckable(True)
        self.fill_btn.setCheckable(True)
        self.settings_btn.setCheckable(True)
        self.about_btn.setCheckable(True)
        
        # 默认选中第一个
        self.extract_btn.setChecked(True)
        
        # 绑定事件
        self.extract_btn.clicked.connect(lambda: self._switch_tab(0))
        self.fill_btn.clicked.connect(lambda: self._switch_tab(1))
        self.settings_btn.clicked.connect(lambda: self._switch_tab(2))
        self.about_btn.clicked.connect(self._show_about)
        
        # 添加到布局
        sidebar_layout.addWidget(self.extract_btn)
        sidebar_layout.addWidget(self.fill_btn)
        sidebar_layout.addWidget(self.settings_btn)
        sidebar_layout.addWidget(self.about_btn)
        sidebar_layout.addStretch()
        
        # 底部版权信息
        copyright_label = QLabel("© 2023 数据工具")
        copyright_label.setObjectName("copyright_label")
        copyright_label.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(copyright_label)
        
        # 添加到主布局
        main_layout.addWidget(sidebar)
    
    def _create_content_area(self, main_layout):
        """
        创建右侧内容区
        
        Args:
            main_layout: 主布局
        """
        # 创建内容区
        content_area = QWidget()
        content_area.setObjectName("content_area")
        content_layout = QVBoxLayout(content_area)
        content_layout.setContentsMargins(20, 20, 20, 20)
        content_layout.setSpacing(15)
        
        # 应用内容区样式
        content_area.setStyleSheet(AppStyles.get_content_style())
        
        # 创建选项卡
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabPosition(QTabWidget.North)
        self.tab_widget.setDocumentMode(True)  # 使用文档模式，更现代的外观
        self.tab_widget.tabBar().setVisible(False)  # 隐藏选项卡栏，由侧边栏控制
        
        # 创建提取和填充选项卡
        extractor_tab = WordExtractorTab(self)
        filler_tab = WordFillerTab(self)
        settings_tab = self._create_settings_tab()
        
        # 添加选项卡
        self.tab_widget.addTab(extractor_tab, "数据提取")
        self.tab_widget.addTab(filler_tab, "数据填充")
        self.tab_widget.addTab(settings_tab, "设置")
        
        # 添加选项卡到内容区
        content_layout.addWidget(self.tab_widget)
        
        # 添加日志显示区
        log_group = QGroupBox("操作日志")
        log_group.setObjectName("log_group")
        log_layout = QVBoxLayout(log_group)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMinimumHeight(120)
        self.log_text.setStyleSheet("font-family: Consolas, Courier New, monospace;")
        log_layout.addWidget(self.log_text)
        
        # 添加日志区到内容区
        content_layout.addWidget(log_group)
        
        # 设置内容区和日志区的比例
        content_layout.setStretch(0, 7)  # 内容区占70%
        content_layout.setStretch(1, 3)  # 日志区占30%
        
        # 添加到主布局
        main_layout.addWidget(content_area)
        
        # 内容区占主要空间
        main_layout.setStretch(0, 1)  # 侧边栏
        main_layout.setStretch(1, 4)  # 内容区
    
    def _create_settings_tab(self):
        """创建设置选项卡"""
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)
        
        # 主题设置
        theme_group = QGroupBox("主题设置")
        theme_layout = QFormLayout(theme_group)
        
        self.theme_combo = QComboBox()
        # 加载可用主题
        for theme_id, theme_name in self.theme_manager.get_available_themes():
            self.theme_combo.addItem(theme_name, theme_id)
        
        # 设置当前主题
        current_theme = self.config_manager.get("theme", "default")
        for i in range(self.theme_combo.count()):
            if self.theme_combo.itemData(i) == current_theme:
                self.theme_combo.setCurrentIndex(i)
                break
        
        self.theme_combo.currentIndexChanged.connect(self._on_theme_changed)
        theme_layout.addRow("选择主题:", self.theme_combo)
        settings_layout.addWidget(theme_group)
        
        # 占位符格式设置
        placeholder_group = QGroupBox("占位符设置")
        placeholder_layout = QFormLayout(placeholder_group)
        
        self.placeholder_formats = QTextEdit()
        current_formats = self.config_manager.get_placeholder_formats()
        self.placeholder_formats.setText("\n".join(current_formats))
        placeholder_layout.addRow("占位符格式:", self.placeholder_formats)
        
        format_hint = QLabel("每行一个格式，使用 %s 表示变量名。例如: {{%s}}, ${%s}, #%s#")
        format_hint.setStyleSheet("color: gray;")
        placeholder_layout.addRow("", format_hint)
        
        settings_layout.addWidget(placeholder_group)
        
        # 图片占位符设置
        image_group = QGroupBox("图片设置")
        image_layout = QFormLayout(image_group)
        
        self.image_width = QLineEdit()
        self.image_width.setText(str(self.config_manager.get("default_image_width", 4.0)))
        image_layout.addRow("默认图片宽度(英寸):", self.image_width)
        
        image_hint = QLabel("使用 img:字段名 作为图片占位符，例如 {{img:照片}} 或 ${img:图片路径}")
        image_hint.setStyleSheet("color: gray;")
        image_layout.addRow("", image_hint)
        
        settings_layout.addWidget(image_group)
        
        # 保存按钮
        save_btn = QPushButton("保存设置")
        save_btn.clicked.connect(self._save_settings)
        settings_layout.addWidget(save_btn)
        
        settings_layout.addStretch()
        return settings_widget
    
    def _on_theme_changed(self, index):
        """主题变更处理"""
        theme_id = self.theme_combo.itemData(index)
        if theme_id:
            self.theme_manager.apply_theme(theme_id)
            self.setStyleSheet(AppStyles.get_main_style())
    
    def _save_settings(self):
        """保存设置"""
        try:
            # 保存主题
            theme_id = self.theme_combo.currentData()
            self.config_manager.set("theme", theme_id)
            
            # 保存占位符格式
            formats_text = self.placeholder_formats.toPlainText()
            formats = [fmt.strip() for fmt in formats_text.split("\n") if fmt.strip()]
            if formats:
                self.config_manager.set("placeholder_formats", formats)
            
            # 保存图片宽度
            try:
                image_width = float(self.image_width.text())
                if 0.5 <= image_width <= 10:
                    self.config_manager.set("default_image_width", image_width)
                else:
                    raise ValueError("图片宽度应在0.5-10英寸范围内")
            except ValueError as e:
                QMessageBox.warning(self, "输入错误", f"图片宽度设置错误: {str(e)}")
                return
            
            # 显示成功消息
            QMessageBox.information(self, "设置已保存", "您的设置已成功保存，部分设置可能需要重启应用后生效。")
            self.logger.info("用户设置已保存")
            self.add_log("设置已保存")
        except Exception as e:
            QMessageBox.critical(self, "保存设置出错", str(e))
            self.logger.error(f"保存设置出错: {str(e)}")
    
    def _switch_tab(self, index):
        """
        切换选项卡
        
        Args:
            index: 选项卡索引
        """
        self.tab_widget.setCurrentIndex(index)
        
        # 更新按钮选中状态
        self.extract_btn.setChecked(index == 0)
        self.fill_btn.setChecked(index == 1)
        self.about_btn.setChecked(False)
    
    def _show_about(self):
        """显示关于对话框"""
        # 取消其他按钮的选中状态
        self.extract_btn.setChecked(False)
        self.fill_btn.setChecked(False)
        self.settings_btn.setChecked(False)
        self.about_btn.setChecked(True)
        
        QMessageBox.about(
            self,
            "关于",
            "<h3>Word数据提取与填充工具 v1.1.0</h3>"
            "<p>一款专业的Word文档数据处理工具，支持：</p>"
            "<ul>"
            "<li>从Word文档中提取数据到Excel</li>"
            "<li>从Excel数据填充到Word模板</li>"
            "<li>支持多种占位符格式: {{变量}}, ${变量}, #变量#</li>"
            "<li>支持图片占位符: {{img:变量}}</li>"
            "<li>主题切换与自定义设置</li>"
            "<li>模板库管理</li>"
            "</ul>"
            "<p>© 2023 数据工具</p>"
        )
    
    def add_log(self, message):
        """
        添加日志消息到日志区
        
        Args:
            message: 日志消息
        """
        self.log_text.append(message)
        # 自动滚动到底部
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def closeEvent(self, event):
        """
        窗口关闭事件
        
        Args:
            event: 关闭事件
        """
        self.logger.info("应用程序关闭")
        event.accept()
    
    def dragEnterEvent(self, event):
        """拖拽进入事件处理"""
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """拖拽放下事件处理"""
        # 获取当前选项卡
        current_tab = self.tab_widget.currentWidget()
        
        # 如果当前选项卡支持拖放，则传递事件
        if hasattr(current_tab, "dropEvent"):
            current_tab.dropEvent(event)
        else:
            event.ignore() 