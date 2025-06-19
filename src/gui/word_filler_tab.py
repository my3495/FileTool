"""
Word填充选项卡模块
用于将Excel数据填充到Word模板的界面
"""
import os
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QLineEdit, QMessageBox, QProgressBar,
    QGroupBox, QFormLayout, QCheckBox, QRadioButton, QButtonGroup
)
from PySide6.QtCore import Qt, QThread, Signal

from src.core.word_filler import WordFiller
from src.utils.file_utils import ensure_dir, get_file_info


class FillerThread(QThread):
    """
    填充线程
    用于在后台执行填充操作
    """
    # 信号定义
    progress_signal = Signal(int)
    log_signal = Signal(str)
    finished_signal = Signal(bool, str)
    
    def __init__(self, template_path, excel_path, output_dir, filename_pattern=None, merge_output=False):
        """
        初始化填充线程
        
        Args:
            template_path: 模板文件路径
            excel_path: Excel数据文件路径
            output_dir: 输出目录
            filename_pattern: 文件名模式
            merge_output: 是否合并输出
        """
        super().__init__()
        self.template_path = template_path
        self.excel_path = excel_path
        self.output_dir = output_dir
        self.filename_pattern = filename_pattern
        self.merge_output = merge_output
        self.filler = WordFiller()
    
    def run(self):
        """执行填充任务"""
        try:
            self.log_signal.emit("开始填充数据...")
            self.progress_signal.emit(10)
            
            # 加载Excel数据
            self.log_signal.emit("正在加载Excel数据...")
            data = self.filler.load_excel_data(self.excel_path)
            self.log_signal.emit(f"加载了{len(data)}行数据")
            self.progress_signal.emit(30)
            
            # 批量填充模板
            self.log_signal.emit("正在填充模板...")
            output_files = self.filler.batch_fill_templates(
                self.template_path,
                self.excel_path,
                self.output_dir,
                self.filename_pattern,
                self.merge_output
            )
            self.progress_signal.emit(100)
            
            result_info = f"生成了{len(output_files)}个文件" if output_files else "未生成任何文件"
            self.log_signal.emit(f"填充完成，{result_info}")
            
            # 返回第一个生成的文件路径作为结果
            result_path = output_files[0] if output_files else ""
            self.finished_signal.emit(True, result_path)
        except Exception as e:
            self.log_signal.emit(f"填充过程中出错: {str(e)}")
            self.finished_signal.emit(False, str(e))


class WordFillerTab(QWidget):
    """
    Word填充选项卡
    用于将Excel数据填充到Word模板的界面
    """
    
    def __init__(self, parent=None):
        """
        初始化Word填充选项卡
        
        Args:
            parent: 父窗口
        """
        super().__init__(parent)
        self.parent = parent
        
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
        
        # Excel文件选择
        excel_layout = QHBoxLayout()
        self.excel_input = QLineEdit()
        self.excel_input.setPlaceholderText("选择包含数据的Excel文件")
        self.excel_btn = QPushButton("浏览...")
        self.excel_btn.clicked.connect(self._select_excel)
        excel_layout.addWidget(self.excel_input)
        excel_layout.addWidget(self.excel_btn)
        file_layout.addRow("Excel文件:", excel_layout)
        
        # 输出目录选择
        output_layout = QHBoxLayout()
        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("选择保存生成Word文件的目录")
        self.output_btn = QPushButton("浏览...")
        self.output_btn.clicked.connect(self._select_output_dir)
        output_layout.addWidget(self.output_input)
        output_layout.addWidget(self.output_btn)
        file_layout.addRow("输出目录:", output_layout)
        
        layout.addWidget(file_group)
        
        # 选项设置区域
        options_group = QGroupBox("选项设置")
        options_layout = QVBoxLayout(options_group)
        
        # 输出模式选择
        output_mode_layout = QHBoxLayout()
        self.separate_radio = QRadioButton("分开保存")
        self.merge_radio = QRadioButton("合并到一个文件")
        self.separate_radio.setChecked(True)  # 默认选择分开保存
        
        output_mode_group = QButtonGroup(self)
        output_mode_group.addButton(self.separate_radio)
        output_mode_group.addButton(self.merge_radio)
        
        output_mode_layout.addWidget(QLabel("输出模式:"))
        output_mode_layout.addWidget(self.separate_radio)
        output_mode_layout.addWidget(self.merge_radio)
        output_mode_layout.addStretch()
        options_layout.addLayout(output_mode_layout)
        
        # 文件名设置
        filename_layout = QHBoxLayout()
        self.filename_input = QLineEdit()
        self.filename_input.setPlaceholderText("例如: 文档_{序号}_{姓名}")
        filename_layout.addWidget(QLabel("文件名模式:"))
        filename_layout.addWidget(self.filename_input)
        options_layout.addLayout(filename_layout)
        
        # 文件名说明
        filename_help = QLabel("注: 使用 {列名} 引用Excel中的列, 如 {姓名}, {部门} 等")
        filename_help.setStyleSheet("color: gray;")
        options_layout.addWidget(filename_help)
        
        layout.addWidget(options_group)
        
        # 操作区域
        action_group = QGroupBox("操作")
        action_layout = QVBoxLayout(action_group)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        action_layout.addWidget(self.progress_bar)
        
        # 按钮
        buttons_layout = QHBoxLayout()
        self.fill_btn = QPushButton("开始填充")
        self.fill_btn.clicked.connect(self._start_fill)
        self.fill_btn.setMinimumHeight(40)
        
        self.open_result_btn = QPushButton("打开结果")
        self.open_result_btn.clicked.connect(self._open_result)
        self.open_result_btn.setEnabled(False)
        self.open_result_btn.setMinimumHeight(40)
        
        buttons_layout.addWidget(self.fill_btn)
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
    
    def _select_excel(self):
        """选择Excel文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Excel数据文件",
            "",
            "Excel 文件 (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_input.setText(file_path)
            self._log(f"已选择Excel文件: {file_path}")
    
    def _select_output_dir(self):
        """选择输出目录"""
        dir_path = QFileDialog.getExistingDirectory(
            self,
            "选择输出目录",
            ""
        )
        if dir_path:
            self.output_input.setText(dir_path)
            self._log(f"已选择输出目录: {dir_path}")
    
    def _start_fill(self):
        """开始填充数据"""
        # 获取文件路径
        template_path = self.template_input.text().strip()
        excel_path = self.excel_input.text().strip()
        output_dir = self.output_input.text().strip()
        
        # 验证文件路径
        if not template_path:
            QMessageBox.warning(self, "警告", "请选择模板文件")
            return
        
        if not excel_path:
            QMessageBox.warning(self, "警告", "请选择Excel数据文件")
            return
        
        if not output_dir:
            QMessageBox.warning(self, "警告", "请选择输出目录")
            return
        
        if not os.path.exists(template_path):
            QMessageBox.warning(self, "警告", "模板文件不存在")
            return
        
        if not os.path.exists(excel_path):
            QMessageBox.warning(self, "警告", "Excel文件不存在")
            return
        
        # 确保输出目录存在
        ensure_dir(output_dir)
        
        # 获取选项
        merge_output = self.merge_radio.isChecked()
        filename_pattern = self.filename_input.text().strip() or None
        
        # 禁用按钮
        self.fill_btn.setEnabled(False)
        self.open_result_btn.setEnabled(False)
        
        # 创建并启动填充线程
        self.fill_thread = FillerThread(
            template_path,
            excel_path,
            output_dir,
            filename_pattern,
            merge_output
        )
        self.fill_thread.progress_signal.connect(self.progress_bar.setValue)
        self.fill_thread.log_signal.connect(self._log)
        self.fill_thread.finished_signal.connect(self._fill_finished)
        self.fill_thread.start()
    
    def _fill_finished(self, success, result):
        """
        填充完成回调
        
        Args:
            success: 是否成功
            result: 结果信息
        """
        # 启用按钮
        self.fill_btn.setEnabled(True)
        
        if success:
            if result:  # 如果有结果文件路径
                self.open_result_btn.setEnabled(True)
                self.result_path = result
            QMessageBox.information(self, "成功", "数据填充完成!")
        else:
            QMessageBox.critical(self, "错误", f"填充失败: {result}")
    
    def _open_result(self):
        """打开结果文件或目录"""
        if self.merge_radio.isChecked():
            # 如果是合并模式，打开结果文件
            if hasattr(self, 'result_path') and os.path.exists(self.result_path):
                os.startfile(self.result_path)
                self._log(f"已打开结果文件: {self.result_path}")
            else:
                QMessageBox.warning(self, "警告", "结果文件不存在")
        else:
            # 如果是分开保存模式，打开输出目录
            output_dir = self.output_input.text().strip()
            if output_dir and os.path.exists(output_dir):
                os.startfile(output_dir)
                self._log(f"已打开输出目录: {output_dir}")
            else:
                QMessageBox.warning(self, "警告", "输出目录不存在")
    
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
            config_manager = self.parent.config_manager
            dialog = TemplateLibrary(config_manager, self)
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
        excel_files = [path for path in file_paths if path.lower().endswith(('.xls', '.xlsx'))]
        
        # 处理Word模板文件
        if word_files and not self.template_input.text():
            self.template_input.setText(word_files[0])
            self._log(f"已设置模板文件: {word_files[0]}")
        
        # 处理Excel数据文件
        if excel_files and not self.excel_input.text():
            self.excel_input.setText(excel_files[0])
            self._log(f"已设置Excel文件: {excel_files[0]}") 