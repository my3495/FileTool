"""
模板库模块
管理Word模板文件
"""
import os
import json
import shutil
from pathlib import Path
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QListWidget, QListWidgetItem, QLabel, QLineEdit,
    QFileDialog, QMessageBox, QInputDialog, QMenu,
    QFrame, QSplitter, QTextEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QAction

from src.core.word_extractor import WordExtractor
from src.utils.file_utils import ensure_dir, get_file_info
from loguru import logger

class TemplateLibrary(QDialog):
    """模板库对话框"""
    
    def __init__(self, config_manager, parent=None):
        """初始化模板库"""
        super().__init__(parent)
        self.config_manager = config_manager
        self.library_dir = config_manager.templates_dir
        self.files_dir = self.library_dir / "files"
        ensure_dir(str(self.files_dir))
        
        self.setWindowTitle("模板库")
        self.setMinimumSize(800, 500)
        
        self._init_ui()
        self._load_templates()
        
        # 初始化提取器
        self.extractor = WordExtractor(config_manager)
    
    def _init_ui(self):
        """初始化用户界面"""
        layout = QVBoxLayout(self)
        
        # 搜索栏
        search_layout = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("搜索模板...")
        self.search_input.textChanged.connect(self._filter_templates)
        search_layout.addWidget(QLabel("搜索:"))
        search_layout.addWidget(self.search_input)
        layout.addLayout(search_layout)
        
        # 分割器
        splitter = QSplitter(Qt.Horizontal)
        
        # 左侧模板列表
        templates_frame = QFrame()
        templates_layout = QVBoxLayout(templates_frame)
        
        self.templates_list = QListWidget()
        self.templates_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.templates_list.customContextMenuRequested.connect(self._show_context_menu)
        self.templates_list.currentItemChanged.connect(self._on_template_selected)
        
        templates_layout.addWidget(QLabel("已保存的模板:"))
        templates_layout.addWidget(self.templates_list)
        
        # 添加按钮
        buttons_layout = QHBoxLayout()
        self.add_btn = QPushButton("添加模板")
        self.add_btn.clicked.connect(self._add_template)
        buttons_layout.addWidget(self.add_btn)
        buttons_layout.addStretch()
        templates_layout.addLayout(buttons_layout)
        
        # 右侧详情区
        details_frame = QFrame()
        details_layout = QVBoxLayout(details_frame)
        
        details_layout.addWidget(QLabel("模板详情:"))
        
        # 基本信息
        self.template_name_label = QLabel("名称: ")
        self.template_path_label = QLabel("路径: ")
        self.template_size_label = QLabel("大小: ")
        
        details_layout.addWidget(self.template_name_label)
        details_layout.addWidget(self.template_path_label)
        details_layout.addWidget(self.template_size_label)
        
        # 占位符预览
        details_layout.addWidget(QLabel("占位符:"))
        self.placeholders_list = QTextEdit()
        self.placeholders_list.setReadOnly(True)
        details_layout.addWidget(self.placeholders_list)
        
        # 添加到分割器
        splitter.addWidget(templates_frame)
        splitter.addWidget(details_frame)
        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 2)
        
        layout.addWidget(splitter)
        
        # 底部按钮
        bottom_layout = QHBoxLayout()
        self.select_btn = QPushButton("选择")
        self.select_btn.clicked.connect(self.accept)
        self.select_btn.setEnabled(False)
        
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.reject)
        
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.select_btn)
        bottom_layout.addWidget(self.cancel_btn)
        
        layout.addLayout(bottom_layout)
    
    def _load_templates(self):
        """加载模板库"""
        self.templates_list.clear()
        
        # 获取模板列表
        template_files = list(self.library_dir.glob("*.json"))
        
        if not template_files:
            item = QListWidgetItem("未找到模板")
            item.setFlags(item.flags() & ~Qt.ItemIsEnabled)
            self.templates_list.addItem(item)
            return
        
        # 加载每个模板
        for template_file in template_files:
            try:
                with open(template_file, 'r', encoding='utf-8') as f:
                    template_info = json.load(f)
                
                item = QListWidgetItem(template_info['name'])
                item.setData(Qt.UserRole, template_info)
                self.templates_list.addItem(item)
            except Exception as e:
                logger.error(f"加载模板出错 {template_file}: {e}")
    
    def _filter_templates(self):
        """筛选模板"""
        search_text = self.search_input.text().lower()
        
        for i in range(self.templates_list.count()):
            item = self.templates_list.item(i)
            if not item.flags() & Qt.ItemIsEnabled:
                continue
                
            template_info = item.data(Qt.UserRole)
            if template_info:
                if search_text in template_info['name'].lower() or search_text in template_info.get('description', '').lower():
                    item.setHidden(False)
                else:
                    item.setHidden(True)
    
    def _show_context_menu(self, pos):
        """显示右键菜单"""
        item = self.templates_list.itemAt(pos)
        if not item:
            return
            
        template_info = item.data(Qt.UserRole)
        if not template_info:
            return
            
        menu = QMenu(self)
        
        # 菜单项
        select_action = QAction("选择", self)
        select_action.triggered.connect(self.accept)
        
        rename_action = QAction("重命名", self)
        rename_action.triggered.connect(lambda: self._rename_template(item))
        
        delete_action = QAction("删除", self)
        delete_action.triggered.connect(lambda: self._delete_template(item))
        
        # 添加菜单项
        menu.addAction(select_action)
        menu.addSeparator()
        menu.addAction(rename_action)
        menu.addAction(delete_action)
        
        # 显示菜单
        menu.exec_(self.templates_list.viewport().mapToGlobal(pos))
    
    def _on_template_selected(self, item):
        """模板选择变化"""
        self.select_btn.setEnabled(False)
        
        if not item:
            self.template_name_label.setText("名称: ")
            self.template_path_label.setText("路径: ")
            self.template_size_label.setText("大小: ")
            self.placeholders_list.clear()
            return
            
        template_info = item.data(Qt.UserRole)
        if not template_info:
            return
            
        # 更新详情
        self.template_name_label.setText(f"名称: {template_info['name']}")
        file_path = template_info.get('file_path', '')
        self.template_path_label.setText(f"路径: {file_path}")
        
        # 获取文件信息
        if file_path and os.path.exists(file_path):
            file_info = get_file_info(file_path)
            self.template_size_label.setText(f"大小: {file_info['size_readable']}")
            
            # 获取占位符
            try:
                placeholders = self.extractor.detect_placeholders(file_path)
                self.placeholders_list.setText("\n".join(placeholders))
            except Exception as e:
                self.placeholders_list.setText(f"获取占位符出错: {str(e)}")
        else:
            self.template_size_label.setText("大小: 文件不存在")
            self.placeholders_list.setText("无法获取占位符信息，文件不存在")
        
        # 使选择按钮可用
        self.select_btn.setEnabled(True)
    
    def _add_template(self):
        """添加模板"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Word模板",
            "",
            "Word 文件 (*.docx *.doc)"
        )
        
        if not file_path:
            return
            
        # 获取模板名称
        name, ok = QInputDialog.getText(
            self, 
            "模板名称",
            "请输入模板名称:",
            text=os.path.basename(file_path).split('.')[0]
        )
        
        if not ok or not name:
            return
            
        # 获取模板描述
        description, ok = QInputDialog.getMultiLineText(
            self,
            "模板描述",
            "请输入模板描述(可选):"
        )
        
        if not ok:
            return
            
        # 复制模板文件到库目录
        try:
            # 检测占位符
            placeholders = self.extractor.detect_placeholders(file_path)
            
            # 创建目标路径
            template_filename = f"{name}_{os.path.basename(file_path)}"
            target_path = os.path.join(self.files_dir, template_filename)
            
            # 复制文件
            shutil.copy2(file_path, target_path)
            
            # 创建模板信息
            template_id = name.lower().replace(' ', '_')
            template_info = {
                'id': template_id,
                'name': name,
                'description': description,
                'file_path': target_path,
                'original_path': file_path,
                'placeholders': placeholders,
                'created_at': str(Path(target_path).stat().st_ctime)
            }
            
            # 保存模板信息
            info_path = os.path.join(self.library_dir, f"{template_id}.json")
            with open(info_path, 'w', encoding='utf-8') as f:
                json.dump(template_info, f, indent=4, ensure_ascii=False)
            
            # 添加到列表
            item = QListWidgetItem(name)
            item.setData(Qt.UserRole, template_info)
            self.templates_list.addItem(item)
            self.templates_list.setCurrentItem(item)
            
            logger.info(f"已添加模板: {name}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"添加模板出错: {str(e)}")
            logger.error(f"添加模板出错: {str(e)}")
    
    def _rename_template(self, item):
        """重命名模板"""
        if not item:
            return
            
        template_info = item.data(Qt.UserRole)
        if not template_info:
            return
            
        # 询问新名称
        new_name, ok = QInputDialog.getText(
            self,
            "重命名模板",
            "请输入新的模板名称:",
            text=template_info['name']
        )
        
        if not ok or not new_name:
            return
            
        try:
            # 更新模板信息
            template_info['name'] = new_name
            
            # 保存模板信息
            info_path = os.path.join(self.library_dir, f"{template_info['id']}.json")
            with open(info_path, 'w', encoding='utf-8') as f:
                json.dump(template_info, f, indent=4, ensure_ascii=False)
            
            # 更新列表项
            item.setText(new_name)
            item.setData(Qt.UserRole, template_info)
            
            logger.info(f"已重命名模板: {template_info['id']} -> {new_name}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"重命名模板出错: {str(e)}")
            logger.error(f"重命名模板出错: {str(e)}")
    
    def _delete_template(self, item):
        """删除模板"""
        if not item:
            return
            
        template_info = item.data(Qt.UserRole)
        if not template_info:
            return
            
        # 询问确认
        reply = QMessageBox.question(
            self,
            "确认删除",
            f"确定要删除模板 {template_info['name']} 吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
            
        try:
            # 删除模板文件
            file_path = template_info.get('file_path', '')
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
            
            # 删除模板信息文件
            info_path = os.path.join(self.library_dir, f"{template_info['id']}.json")
            if os.path.exists(info_path):
                os.remove(info_path)
            
            # 从列表中移除
            self.templates_list.takeItem(self.templates_list.row(item))
            
            logger.info(f"已删除模板: {template_info['name']}")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"删除模板出错: {str(e)}")
            logger.error(f"删除模板出错: {str(e)}")
    
    def get_selected_template(self):
        """获取选中的模板路径"""
        item = self.templates_list.currentItem()
        if not item:
            return None
            
        template_info = item.data(Qt.UserRole)
        if not template_info:
            return None
            
        return template_info.get('file_path') 