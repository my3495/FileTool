"""
主题管理模块
处理应用程序主题
"""
import os
import json
from pathlib import Path
from PySide6.QtGui import QPalette, QColor
from PySide6.QtWidgets import QApplication
from loguru import logger

class ThemeManager:
    """主题管理类"""
    
    def __init__(self, config_manager):
        """初始化主题管理器"""
        self.config_manager = config_manager
        self.themes_dir = config_manager.themes_dir
        
        # 确保默认主题存在
        self._ensure_default_themes()
        
        # 主题缓存
        self._themes = {}
        
        # 加载所有主题
        self._load_all_themes()
    
    def _ensure_default_themes(self):
        """确保默认主题存在"""
        # 创建默认亮色主题
        default_theme = {
            "name": "默认主题",
            "dark_mode": False,
            "colors": {
                "primary": "#2196F3", 
                "secondary": "#03A9F4",
                "accent": "#009688",
                "background": "#FFFFFF",
                "card": "#F5F5F5",
                "text": "#212121",
                "border": "#E0E0E0"
            },
            "styles": {}  # 详细样式会在应用主题时使用 AppStyles 中的方法生成
        }
        
        # 创建默认暗色主题
        dark_theme = {
            "name": "暗色主题",
            "dark_mode": True,
            "colors": {
                "primary": "#1976D2", 
                "secondary": "#0288D1",
                "accent": "#00796B",
                "background": "#121212",
                "card": "#1E1E1E",
                "text": "#FFFFFF",
                "border": "#333333"
            },
            "styles": {}
        }
        
        # 保存默认主题
        self._save_theme("default", default_theme)
        self._save_theme("dark", dark_theme)
    
    def _save_theme(self, theme_id, theme_data):
        """保存主题"""
        theme_file = self.themes_dir / f"{theme_id}.json"
        try:
            with open(theme_file, 'w', encoding='utf-8') as f:
                json.dump(theme_data, f, indent=4, ensure_ascii=False)
        except Exception as e:
            logger.error(f"保存主题出错: {e}")
    
    def _load_all_themes(self):
        """加载所有主题"""
        for theme_file in self.themes_dir.glob("*.json"):
            theme_id = theme_file.stem
            try:
                with open(theme_file, 'r', encoding='utf-8') as f:
                    self._themes[theme_id] = json.load(f)
                    logger.debug(f"已加载主题: {theme_id}")
            except Exception as e:
                logger.error(f"加载主题 {theme_id} 出错: {e}")
    
    def get_available_themes(self):
        """获取可用的主题列表"""
        return [(theme_id, theme_data["name"]) for theme_id, theme_data in self._themes.items()]
    
    def apply_theme(self, theme_id):
        """应用主题"""
        if theme_id not in self._themes:
            logger.warning(f"主题不存在: {theme_id}，使用默认主题")
            theme_id = "default"
        
        theme_data = self._themes[theme_id]
        colors = theme_data["colors"]
        is_dark = theme_data.get("dark_mode", False)
        
        # 更新AppStyles中的颜色
        from src.gui.styles import AppStyles
        AppStyles.update_colors(colors)
        
        # 应用暗色模式 (如果支持)
        app = QApplication.instance()
        if app:
            palette = QPalette()
            if is_dark:
                palette.setColor(QPalette.Window, QColor(colors["background"]))
                palette.setColor(QPalette.WindowText, QColor(colors["text"]))
                palette.setColor(QPalette.Base, QColor(colors["card"]))
                palette.setColor(QPalette.Text, QColor(colors["text"]))
                palette.setColor(QPalette.Button, QColor(colors["card"]))
                palette.setColor(QPalette.ButtonText, QColor(colors["text"]))
                palette.setColor(QPalette.Highlight, QColor(colors["primary"]))
                palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
                app.setPalette(palette)
            else:
                # 恢复默认调色板
                app.setPalette(app.style().standardPalette())
        
        # 保存当前主题
        self.config_manager.set("theme", theme_id)
        
        logger.info(f"已应用主题: {theme_data['name']}")
        return theme_data 