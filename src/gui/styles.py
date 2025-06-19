"""
样式模块
定义应用程序的主题和样式
"""
import base64


class AppStyles:
    """应用程序样式类"""
    
    # 颜色定义
    PRIMARY_COLOR = "#2C3E50"  # 深蓝色
    SECONDARY_COLOR = "#3498DB"  # 亮蓝色
    ACCENT_COLOR = "#1ABC9C"  # 青绿色
    WARNING_COLOR = "#F39C12"  # 橙色
    ERROR_COLOR = "#E74C3C"  # 红色
    SUCCESS_COLOR = "#2ECC71"  # 绿色
    BACKGROUND_COLOR = "#F5F5F5"  # 浅灰色
    CARD_COLOR = "#FFFFFF"  # 白色
    TEXT_COLOR = "#333333"  # 深灰色
    LIGHT_TEXT_COLOR = "#7F8C8D"  # 浅灰色文本
    BORDER_COLOR = "#BDC3C7"  # 边框颜色
    
    # Base64 编码的 SVG 图标 (白色对勾)
    CHECK_ICON_SVG = "data:image/svg+xml;base64," + base64.b64encode(b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="white"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/></svg>').decode('utf-8')

    # 字体定义
    FONT_FAMILY = "Segoe UI, Arial, sans-serif"
    FONT_SIZE_SMALL = "10px"
    FONT_SIZE_NORMAL = "12px"
    FONT_SIZE_LARGE = "14px"
    FONT_SIZE_XLARGE = "16px"
    FONT_SIZE_XXLARGE = "20px"
    
    @staticmethod
    def _adjust_color(color, factor):
        """
        调整颜色的亮度。
        factor > 1.0 使颜色变亮
        factor < 1.0 使颜色变暗
        """
        if not color.startswith('#'):
            return color
        
        color_hex = color.lstrip('#')
        rgb = tuple(int(color_hex[i:i+2], 16) for i in (0, 2, 4))
        
        new_rgb = tuple(min(255, max(0, int(c * factor))) for c in rgb)
        
        return f"#{new_rgb[0]:02x}{new_rgb[1]:02x}{new_rgb[2]:02x}"
    
    @classmethod
    def update_colors(cls, colors):
        """
        更新样式颜色
        
        Args:
            colors: 颜色字典
        """
        if "primary" in colors:
            cls.PRIMARY_COLOR = colors["primary"]
        if "secondary" in colors:
            cls.SECONDARY_COLOR = colors["secondary"]
        if "accent" in colors:
            cls.ACCENT_COLOR = colors["accent"]
        if "background" in colors:
            cls.BACKGROUND_COLOR = colors["background"]
        if "card" in colors:
            cls.CARD_COLOR = colors["card"]
        if "text" in colors:
            cls.TEXT_COLOR = colors["text"]
        if "border" in colors:
            cls.BORDER_COLOR = colors["border"]
    
    @classmethod
    def get_main_style(cls):
        """获取主要样式表"""
        return f"""
            QWidget {{
                font-family: {cls.FONT_FAMILY};
                font-size: {cls.FONT_SIZE_NORMAL};
                color: {cls.TEXT_COLOR};
            }}
            
            QMainWindow {{
                background-color: {cls.BACKGROUND_COLOR};
            }}
            
            QLabel {{
                color: {cls.TEXT_COLOR};
            }}
            
            QPushButton {{
                background-color: {cls.SECONDARY_COLOR};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
            }}
            
            QPushButton:hover {{
                background-color: {cls._adjust_color(cls.SECONDARY_COLOR, 1.1)};
            }}
            
            QPushButton:pressed {{
                background-color: {cls._adjust_color(cls.SECONDARY_COLOR, 0.9)};
            }}
            
            QPushButton:disabled {{
                background-color: {cls.BORDER_COLOR};
                color: {cls.LIGHT_TEXT_COLOR};
            }}
            
            QLineEdit {{
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 4px;
                padding: 6px;
                background-color: {cls.CARD_COLOR};
                selection-background-color: {cls.SECONDARY_COLOR};
            }}
            
            QLineEdit:focus {{
                border: 1px solid {cls.SECONDARY_COLOR};
            }}
            
            QGroupBox {{
                font-weight: bold;
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 6px;
                margin-top: 12px;
                background-color: {cls.CARD_COLOR};
            }}
            
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 10px;
                padding: 0 5px;
                color: {cls.PRIMARY_COLOR};
            }}
            
            QTabWidget::pane {{
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 4px;
                background-color: {cls.CARD_COLOR};
            }}
            
            QTabBar::tab {{
                background-color: {cls._adjust_color(cls.BACKGROUND_COLOR, 0.98)};
                color: {cls.TEXT_COLOR};
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                padding: 8px 16px;
                margin-right: 2px;
            }}
            
            QTabBar::tab:selected {{
                background-color: {cls.SECONDARY_COLOR};
                color: white;
            }}
            
            QTabBar::tab:hover:!selected {{
                background-color: {cls._adjust_color(cls.BACKGROUND_COLOR, 1.05)};
            }}
            
            QProgressBar {{
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 4px;
                text-align: center;
                background-color: {cls.CARD_COLOR};
            }}
            
            QProgressBar::chunk {{
                background-color: {cls.ACCENT_COLOR};
                border-radius: 3px;
            }}
            
            QListWidget {{
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 4px;
                background-color: {cls.CARD_COLOR};
                alternate-background-color: {cls._adjust_color(cls.BACKGROUND_COLOR, 0.98)};
                padding: 2px;
            }}
            
            QListWidget::item {{
                padding: 5px;
                border-radius: 2px;
            }}
            
            QListWidget::item:selected {{
                background-color: {cls.SECONDARY_COLOR};
                color: white;
            }}
            
            QTextEdit {{
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 4px;
                background-color: {cls.CARD_COLOR};
                selection-background-color: {cls.SECONDARY_COLOR};
            }}
            
            QCheckBox {{
                spacing: 8px;
            }}
            
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 3px;
                background-color: {cls.CARD_COLOR};
            }}
            
            QCheckBox::indicator:hover {{
                border: 1px solid {cls.SECONDARY_COLOR};
            }}

            QCheckBox::indicator:checked {{
                background-color: {cls.SECONDARY_COLOR};
                border: 1px solid {cls.SECONDARY_COLOR};
                image: url({cls.CHECK_ICON_SVG});
            }}
            
            QRadioButton {{
                spacing: 8px;
            }}
            
            QRadioButton::indicator {{
                width: 16px;
                height: 16px;
                border-radius: 8px;
                border: 1px solid {cls.BORDER_COLOR};
                background-color: {cls.CARD_COLOR};
            }}

            QRadioButton::indicator:hover {{
                border: 1px solid {cls.SECONDARY_COLOR};
            }}
            
            QRadioButton::indicator:checked {{
                border: 1px solid {cls.SECONDARY_COLOR};
                background-color: {cls.SECONDARY_COLOR};
                /* The 'dot' is created by padding */
                background-clip: content;
                padding: 4px;
            }}
            
            QComboBox {{
                border: 1px solid {cls.BORDER_COLOR};
                border-radius: 4px;
                padding: 5px;
                background-color: {cls.CARD_COLOR};
            }}
            
            QComboBox:focus {{
                border: 1px solid {cls.SECONDARY_COLOR};
            }}
            
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid {cls.BORDER_COLOR};
                border-top-right-radius: 4px;
                border-bottom-right-radius: 4px;
            }}

            QComboBox::down-arrow {{
                /* You can add a custom arrow SVG here if needed */
            }}
        """
    
    @classmethod
    def get_sidebar_style(cls):
        """获取侧边栏样式表"""
        return f"""
            QWidget#sidebar {{
                background-color: {cls.PRIMARY_COLOR};
                border-radius: 0px;
                border-top-right-radius: 10px;
                border-bottom-right-radius: 10px;
            }}
            
            QLabel#title_label {{
                color: white;
                font-size: {cls.FONT_SIZE_XXLARGE};
                font-weight: bold;
                padding: 20px 0;
            }}
            
            QPushButton#nav_button {{
                background-color: transparent;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 12px;
                text-align: left;
                font-weight: normal;
                font-size: {cls.FONT_SIZE_LARGE};
            }}
            
            QPushButton#nav_button:hover {{
                background-color: {cls._adjust_color(cls.PRIMARY_COLOR, 1.2)};
            }}
            
            QPushButton#nav_button:checked {{
                background-color: {cls._adjust_color(cls.PRIMARY_COLOR, 1.4)};
                font-weight: bold;
            }}
            
            QLabel#copyright_label {{
                color: rgba(255, 255, 255, 0.5);
                font-size: {cls.FONT_SIZE_SMALL};
                padding: 10px 0;
            }}
        """
    
    @classmethod
    def get_content_style(cls):
        """获取内容区样式表"""
        return f"""
            QWidget#content_area {{
                background-color: {cls.BACKGROUND_COLOR};
                border-radius: 10px;
            }}
            
            QGroupBox#action_group {{
                background-color: {cls.CARD_COLOR};
                border-radius: 6px;
            }}
            
            QPushButton#primary_button {{
                background-color: {cls.ACCENT_COLOR};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: {cls.FONT_SIZE_LARGE};
            }}
            
            QPushButton#primary_button:hover {{
                background-color: {cls._adjust_color(cls.ACCENT_COLOR, 1.1)};
            }}
            
            QPushButton#primary_button:pressed {{
                background-color: {cls._adjust_color(cls.ACCENT_COLOR, 0.9)};
            }}
            
            QPushButton#secondary_button {{
                background-color: {cls.SECONDARY_COLOR};
                color: white;
                border: none;
                border-radius: 4px;
                padding: 10px 20px;
                font-weight: bold;
                font-size: {cls.FONT_SIZE_LARGE};
            }}
            
            QPushButton#secondary_button:hover {{
                background-color: {cls._adjust_color(cls.SECONDARY_COLOR, 1.1)};
            }}
            
            QPushButton#secondary_button:pressed {{
                background-color: {cls._adjust_color(cls.SECONDARY_COLOR, 0.9)};
            }}
        """ 