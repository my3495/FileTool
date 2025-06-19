"""
配置管理模块
处理应用程序配置
"""
import os
import json
from pathlib import Path
import appdirs
from loguru import logger

class ConfigManager:
    """配置管理类"""
    
    def __init__(self):
        """初始化配置管理器"""
        # 使用appdirs获取标准配置目录
        self.config_dir = Path(appdirs.user_config_dir("WordTool"))
        self.config_file = self.config_dir / "settings.json"
        self.themes_dir = self.config_dir / "themes"
        self.templates_dir = self.config_dir / "templates"
        
        # 默认配置
        self.default_config = {
            "theme": "default",
            "recent_files": {
                "templates": [],
                "targets": [],
                "excel_files": [],
                "output_dirs": []
            },
            "max_recent_files": 10,
            "placeholder_formats": ["{{%s}}", "${%s}", "#%s#"],
            "default_image_width": 2.0,  # 英寸
            "last_settings": {
                "extractor": {},
                "filler": {}
            }
        }
        
        # 确保配置目录存在
        self._ensure_dirs()
        
        # 加载或创建配置
        self.config = self._load_config()
    
    def _ensure_dirs(self):
        """确保配置目录存在"""
        self.config_dir.mkdir(parents=True, exist_ok=True)
        self.themes_dir.mkdir(exist_ok=True)
        self.templates_dir.mkdir(exist_ok=True)
    
    def _load_config(self):
        """加载配置"""
        if not self.config_file.exists():
            # 创建默认配置
            self._save_config(self.default_config)
            return self.default_config.copy()
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 确保有所有默认配置项
                for key, value in self.default_config.items():
                    if key not in config:
                        config[key] = value
                return config
        except Exception as e:
            logger.error(f"加载配置文件出错: {e}")
            return self.default_config.copy()
    
    def _save_config(self, config):
        """保存配置"""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            logger.info("配置已保存")
        except Exception as e:
            logger.error(f"保存配置文件出错: {e}")
    
    def get(self, key, default=None):
        """获取配置值"""
        return self.config.get(key, default)
    
    def set(self, key, value):
        """设置配置值"""
        self.config[key] = value
        self._save_config(self.config)
    
    def add_recent_file(self, file_type, file_path):
        """添加最近使用的文件"""
        if file_type not in self.config["recent_files"]:
            return
        
        # 确保是绝对路径
        file_path = os.path.abspath(file_path)
        
        # 移除已存在的相同路径
        recent_files = self.config["recent_files"][file_type]
        if file_path in recent_files:
            recent_files.remove(file_path)
        
        # 添加到开头
        recent_files.insert(0, file_path)
        
        # 限制数量
        max_files = self.config["max_recent_files"]
        if len(recent_files) > max_files:
            recent_files = recent_files[:max_files]
        
        self.config["recent_files"][file_type] = recent_files
        self._save_config(self.config)
    
    def get_recent_files(self, file_type):
        """获取最近使用的文件"""
        if file_type not in self.config["recent_files"]:
            return []
        
        # 过滤不存在的文件
        recent_files = self.config["recent_files"][file_type]
        return [f for f in recent_files if os.path.exists(f)]
    
    def save_last_settings(self, tab_name, settings):
        """保存最后一次的设置"""
        self.config["last_settings"][tab_name] = settings
        self._save_config(self.config)
    
    def get_last_settings(self, tab_name):
        """获取最后一次的设置"""
        return self.config["last_settings"].get(tab_name, {})
        
    def get_placeholder_formats(self):
        """获取占位符格式列表"""
        return self.config.get("placeholder_formats", ["{{%s}}"]) 