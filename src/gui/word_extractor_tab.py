"""
Word提取选项卡模块
用于从Word提取数据到Excel的界面
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QLineEdit, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QListWidget, QAbstractItemView,
    QToolButton, QMenu
)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction

from src.core.word_extractor import WordExtractor
from src.utils.file_utils import ensure_dir, get_file_info


class ExtractorThread(QThread):
    """
    提取线程
    用于在后台执行提取操作
    """
    # 信号定义
    progress_signal = Signal(int)
    log_signal = Signal(str)
    finished_signal = Signal(bool, str)
    
    def __init__(self, template_path, target_paths, output_path, merge_output=True, config_manager=None):
        """
        初始化提取线程
        
        Args:
            template_path: 模板文件路径
            target_paths: 目标文件路径列表
            output_path: 输出文件路径
            merge_output: 是否合并输出结果
            config_manager: 配置管理器，可选
        """
        super().__init__()
        self.template_path = template_path
        self.target_paths = target_paths
        self.output_path = output_path
        self.merge_output = merge_output
        self.extractor = WordExtractor(config_manager)
    
    def run(self):
        """执行提取任务"""
        try:
            self.log_signal.emit(f"开始提取数据，共{len(self.target_paths)}个文件...")
            self.progress_signal.emit(5)
            
            # 检测占位符
            self.log_signal.emit("正在检测模板中的占位符...")
            placeholders = self.extractor.detect_placeholders(self.template_path)
            self.log_signal.emit(f"检测到{len(placeholders)}个占位符: {', '.join(placeholders)}")
            self.progress_signal.emit(20)
            
            # 批量处理文件
            self.log_signal.emit("正在批量处理文件...")
            all_data = self.extractor.extract_batch_data_from_files(self.template_path, self.target_paths)
            self.progress_signal.emit(80)
            
            # 导出到Excel
            if not all_data.empty:
                self.log_signal.emit(f"正在导出{len(all_data)}行数据到Excel...")
                self.extractor.export_to_excel(all_data, self.output_path)
                self.progress_signal.emit(100)
                
                self.log_signal.emit(f"提取完成，数据已保存到: {self.output_path}")
                self.finished_signal.emit(True, self.output_path)
            else:
                self.log_signal.emit("未能提取到任何数据")
                self.finished_signal.emit(False, "未能提取到任何数据")
        except Exception as e:
            self.log_signal.emit(f"提取过程中出错: {str(e)}")
            self.finished_signal.emit(False, str(e))


class WordExtractorTab(QWidget):
    """
    Word提取选项卡
    用于从Word提取数据到Excel的界面
    """
    
    def __init__(self, parent=None):
        """
        初始化Word提取选项卡
        
        Args:
            parent: 父窗口
        """
        super().__init__(parent)
        self.parent = parent
        self.target_files = []  # 保存目标文件路径列表
        
        # 获取配置管理器
        self.config_manager = parent.config_manager if hasattr(parent, 'config_manager') else None
        
        # 初始化提取器
        self.extractor = WordExtractor(self.config_manager)
        
        # 初始化界面
        self._init_ui()
    
    def _init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 启用拖放
        self.setAcceptDrops(True)
        
        # 文件选择区域
        file_group = QGroupBox("文件选择")
        file_layout = QFormLayout(file_group)
        
        # 模板文件选择
        template_layout = QHBoxLayout()
        self.template_input = QLineEdit()
        self.template_input.setPlaceholderText("选择包含占位符的Word模板文件")
        
        # 创建按钮菜单
        template_btn_layout = QHBoxLayout()
        template_btn_layout.setSpacing(2)
        
        self.template_btn = QPushButton("浏览...")
        self.template_btn.clicked.connect(self._select_template)
        
        self.template_lib_btn = QPushButton("模板库")
        self.template_lib_btn.clicked.connect(self._open_template_library)
        
        template_btn_layout.addWidget(self.template_btn)
        template_btn_layout.addWidget(self.template_lib_btn)
        
        template_layout.addWidget(self.template_input)
        template_layout.addLayout(template_btn_layout)
        file_layout.addRow("模板文件:", template_layout)
        
        # 目标文件选择 - 改为列表形式
        target_layout = QVBoxLayout()
        target_buttons = QHBoxLayout()
        
        self.target_list = QListWidget()
        self.target_list.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.target_list.setMinimumHeight(120)
        
        self.add_target_btn = QPushButton("添加文件")
        self.add_target_btn.clicked.connect(self._add_target_files)
        
        self.remove_target_btn = QPushButton("删除选中")
        self.remove_target_btn.clicked.connect(self._remove_target_files)
        self.remove_target_btn.setEnabled(False)
        
        self.target_list.itemSelectionChanged.connect(self._on_target_selection_changed)
        
        target_buttons.addWidget(self.add_target_btn)
        target_buttons.addWidget(self.remove_target_btn)
        target_buttons.addStretch()
        
        target_layout.addWidget(self.target_list)
        target_layout.addLayout(target_buttons)
        
        file_layout.addRow("目标文件:", target_layout)
        
        # 输出文件选择
        output_layout = QHBoxLayout()
        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("选择保存提取数据的Excel文件")
        self.output_btn = QPushButton("浏览...")
        self.output_btn.clicked.connect(self._select_output)
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(self.output_btn)
        file_layout.addRow("输出文件:", output_layout)
        
        layout.addWidget(file_group)
        
        # 操作区域
        action_group = QGroupBox("操作")
        action_layout = QVBoxLayout(action_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        action_layout.addWidget(self.progress_bar)
        
        # 按钮
        buttons_layout = QHBoxLayout()
        self.extract_btn = QPushButton("开始提取")
        self.extract_btn.clicked.connect(self._start_extract)
        self.extract_btn.setMinimumHeight(40)
        
        self.open_result_btn = QPushButton("打开结果")
        self.open_result_btn.clicked.connect(self._open_result)
        self.open_result_btn.setEnabled(False)
        self.open_result_btn.setMinimumHeight(40)
        
        buttons_layout.addWidget(self.extract_btn)
        buttons_layout.addWidget(self.open_result_btn)
        action_layout.addLayout(buttons_layout)
        
        layout.addWidget(action_group)
        
        # 添加弹性空间
        layout.addStretch()
    
    def _select_template(self):
        """选择模板文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word模板",
            "",
            "Word 文件 (*.docx *.doc)"
        )
        if file_path:
            self.template_input.setText(file_path)
            self._log(f"已选择模板文件: {file_path}")
    
    def _add_target_files(self):
        """添加目标文件"""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "选择Word目标文件",
            "",
            "Word 文件 (*.docx *.doc)"
        )
        if file_paths:
            for file_path in file_paths:
                # 检查是否已存在
                if file_path not in self.target_files:
                    self.target_files.append(file_path)
                    self.target_list.addItem(os.path.basename(file_path))
            
            self._log(f"已添加{len(file_paths)}个目标文件")
    
    def _remove_target_files(self):
        """删除选中的目标文件"""
        selected_items = self.target_list.selectedItems()
        if not selected_items:
            return
            
        for item in selected_items:
            row = self.target_list.row(item)
            removed_file = self.target_files.pop(row)
            self.target_list.takeItem(row)
            self._log(f"已移除文件: {os.path.basename(removed_file)}")
        
        # 刷新列表
        self.target_list.clearSelection()
    
    def _on_target_selection_changed(self):
        """目标文件选择改变时的处理"""
        self.remove_target_btn.setEnabled(len(self.target_list.selectedItems()) > 0)
    
    def _select_output(self):
        """选择输出文件"""
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "保存Excel文件",
            "",
            "Excel 文件 (*.xlsx *.xls)"
        )
        if file_path:
            # 确保有.xlsx扩展名
            if not file_path.endswith(('.xlsx', '.xls')):
                file_path += '.xlsx'
            self.output_input.setText(file_path)
            self._log(f"已选择输出文件: {file_path}")
    
    def _start_extract(self):
        """开始提取数据"""
        # 获取文件路径
        template_path = self.template_input.text().strip()
        output_path = self.output_input.text().strip()
        
        # 验证文件路径
        if not template_path:
            QMessageBox.warning(self, "警告", "请选择模板文件")
            return
        
        if not self.target_files:
            QMessageBox.warning(self, "警告", "请添加至少一个目标文件")
            return
        
        if not output_path:
            QMessageBox.warning(self, "警告", "请选择输出文件")
            return
        
        if not os.path.exists(template_path):
            QMessageBox.warning(self, "警告", "模板文件不存在")
            return
        
        # 检查目标文件是否存在
        valid_files = []
        for file_path in self.target_files:
            if os.path.exists(file_path):
                valid_files.append(file_path)
            else:
                self._log(f"警告: 文件不存在, 已跳过: {os.path.basename(file_path)}")
        
        if not valid_files:
            QMessageBox.warning(self, "警告", "没有有效的目标文件")
            return
        
        # 确保输出目录存在
        output_dir = os.path.dirname(output_path)
        ensure_dir(output_dir)
        
        # 禁用按钮
        self.extract_btn.setEnabled(False)
        self.open_result_btn.setEnabled(False)
        
        # 创建并启动提取线程
        self.extract_thread = ExtractorThread(
            template_path, 
            valid_files, 
            output_path, 
            True,
            self.config_manager
        )
        self.extract_thread.progress_signal.connect(self.progress_bar.setValue)
        self.extract_thread.log_signal.connect(self._log)
        self.extract_thread.finished_signal.connect(self._extract_finished)
        self.extract_thread.start()
    
    def _extract_finished(self, success, result):
        """
        提取完成回调
        
        Args:
            success: 是否成功
            result: 结果信息
        """
        # 启用按钮
        self.extract_btn.setEnabled(True)
        
        if success:
            self.open_result_btn.setEnabled(True)
            self.result_path = result
            QMessageBox.information(self, "成功", "数据提取完成!")
        else:
            QMessageBox.critical(self, "错误", f"提取失败: {result}")
    
    def _open_result(self):
        """打开结果文件"""
        if hasattr(self, 'result_path') and os.path.exists(self.result_path):
            # 使用系统默认程序打开文件
            os.startfile(self.result_path)
            self._log(f"已打开结果文件: {self.result_path}")
        else:
            QMessageBox.warning(self, "警告", "结果文件不存在")
    
    def _log(self, message):
        """
        添加日志消息
        
        Args:
            message: 日志消息
        """
        if self.parent:
            self.parent.add_log(message)
    
    def _open_template_library(self):
        """打开模板库"""
        # 导入在这里执行，避免循环导入
        from src.gui.template_library import TemplateLibrary
        
        try:
            dialog = TemplateLibrary(self.config_manager, self)
            if dialog.exec():
                template_path = dialog.get_selected_template()
                if template_path and os.path.exists(template_path):
                    self.template_input.setText(template_path)
                    self._log(f"已从模板库选择模板: {template_path}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"打开模板库出错: {str(e)}")
            self._log(f"打开模板库出错: {str(e)}")
    
    def dragEnterEvent(self, event):
        """拖放进入事件处理"""
        # 接受文件拖放
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        """拖放释放事件处理"""
        urls = event.mimeData().urls()
        file_paths = [url.toLocalFile() for url in urls]
        
        word_files = [path for path in file_paths if path.lower().endswith(('.doc', '.docx'))]
        if not word_files:
            return
        
        # 如果模板为空，第一个文件为模板
        if not self.template_input.text() and word_files:
            self.template_input.setText(word_files[0])
            self._log(f"已设置模板文件: {word_files[0]}")
            word_files = word_files[1:]  # 其余文件作为目标文件
        
        # 剩余文件添加为目标文件
        for file_path in word_files:
            self._add_target_file(file_path) 