"""
占位符解析模块
用于解析和处理文档中的占位符
"""
import re
from loguru import logger


class BasePlaceholderParser:
    """
    占位符解析器基类
    """
    def __init__(self, config_manager=None):
        self.placeholder_patterns = []
        self.config_manager = config_manager
        logger.debug(f"初始化占位符解析器，支持的格式: {self.placeholder_patterns}")
        
    def contains_placeholders(self, text):
        """
        检查文本中是否包含占位符
        
        Args:
            text: 要检查的文本
            
        Returns:
            bool: 如果包含占位符则为True，否则为False
        """
        for pattern in self.placeholder_patterns:
            if re.search(pattern, text):
                return True
        return False
        
    def find_all_placeholders(self, text):
        """
        查找文本中所有占位符
        
        Args:
            text: 要搜索的文本
            
        Returns:
            list: 占位符列表，每项为 (名称, 完整匹配, 使用的模式)
        """
        placeholders = []
        for pattern in self.placeholder_patterns:
            try:
                matches = re.finditer(pattern, text)
                for match in matches:
                    placeholder_name = match.group(1).strip()  # 确保移除可能的空白字符
                    full_match = match.group(0)
                    placeholders.append((placeholder_name, full_match, pattern))
            except Exception as e:
                logger.error(f"查找占位符时出错, 模式 {pattern}: {str(e)}")
        return placeholders
        
    def extract_placeholders(self, text):
        """
        从文本中提取所有占位符名称
        
        Args:
            text: 要搜索的文本
            
        Returns:
            list: 占位符名称列表
        """
        placeholders = []
        for name, _, _ in self.find_all_placeholders(text):
            if name not in placeholders:
                placeholders.append(name)
        return placeholders
        
    def get_placeholder_name(self, text, raw_match):
        """
        从匹配文本中提取占位符名称，子类应该重写此方法
        
        Args:
            text: 原始文本
            raw_match: 匹配的占位符原始文本
            
        Returns:
            str: 占位符名称
        """
        raise NotImplementedError("子类必须实现get_placeholder_name方法")
        
    def replace(self, text, replacements):
        """
        替换文本中的所有占位符，子类应该重写此方法
        
        Args:
            text: 要处理的文本
            replacements: 替换值字典
            
        Returns:
            str: 替换后的文本
        """
        raise NotImplementedError("子类必须实现replace方法")
        
    def has_placeholder(self, text, placeholder_name):
        """
        检查文本中是否包含指定占位符，子类应该重写此方法
        
        Args:
            text: 要检查的文本
            placeholder_name: 要查找的占位符名称
            
        Returns:
            bool: 如果包含则为True，否则为False
        """
        raise NotImplementedError("子类必须实现has_placeholder方法")

class StandardPlaceholderParser(BasePlaceholderParser):
    """
    标准占位符解析器，支持{{变量}}格式
    """
    def __init__(self, config_manager=None):
        super().__init__(config_manager)
        # 设置占位符模式
        self.placeholder_patterns = [
            r'\{\{(.*?)\}\}',  # {{变量}}
        ]
    
    def get_placeholder_name(self, text, raw_match):
        """
        从匹配文本中提取占位符名称
        
        Args:
            text: 原始文本
            raw_match: 匹配的占位符原始文本
            
        Returns:
            str: 占位符名称
        """
        # 移除大括号并去除空白
        return raw_match.replace('{{', '').replace('}}', '').strip()
    
    def replace(self, text, replacements):
        """
        替换文本中的所有占位符
        
        Args:
            text: 要处理的文本
            replacements: 替换值字典，格式为 {placeholder_name: value}
            
        Returns:
            str: 替换后的文本
        """
        result_text = text
        
        # 查找所有占位符
        placeholders = self.find_all_placeholders(text)
        
        # 替换每个找到的占位符
        for placeholder_name, full_match, _ in placeholders:
            if placeholder_name in replacements:
                value = str(replacements[placeholder_name])
                # 特殊处理：如果值为空白，保留占位符
                if value.strip():
                    result_text = result_text.replace(full_match, value)
                    
        return result_text
        
    def has_placeholder(self, text, placeholder_name):
        """
        检查文本中是否包含指定占位符
        
        Args:
            text: 要检查的文本
            placeholder_name: 要查找的占位符名称
            
        Returns:
            bool: 如果包含则为True，否则为False
        """
        pattern = r'\{\{' + re.escape(placeholder_name) + r'\}\}'
        return bool(re.search(pattern, text))

class DollarPlaceholderParser(BasePlaceholderParser):
    """
    美元占位符解析器，支持${变量}格式
    """
    def __init__(self, config_manager=None):
        super().__init__(config_manager)
        # 设置占位符模式
        self.placeholder_patterns = [
            r'\$\{(.*?)\}',  # ${变量}
        ]
    
    def get_placeholder_name(self, text, raw_match):
        """
        从匹配文本中提取占位符名称
        
        Args:
            text: 原始文本
            raw_match: 匹配的占位符原始文本
            
        Returns:
            str: 占位符名称
        """
        # 移除美元符号和大括号并去除空白
        return raw_match.replace('${', '').replace('}', '').strip()
    
    def replace(self, text, replacements):
        """
        替换文本中的所有占位符
        
        Args:
            text: 要处理的文本
            replacements: 替换值字典，格式为 {placeholder_name: value}
            
        Returns:
            str: 替换后的文本
        """
        result_text = text
        
        # 查找所有占位符
        placeholders = self.find_all_placeholders(text)
        
        # 替换每个找到的占位符
        for placeholder_name, full_match, _ in placeholders:
            if placeholder_name in replacements:
                value = str(replacements[placeholder_name])
                # 特殊处理：如果值为空白，保留占位符
                if value.strip():
                    result_text = result_text.replace(full_match, value)
                    
        return result_text
        
    def has_placeholder(self, text, placeholder_name):
        """
        检查文本中是否包含指定占位符
        
        Args:
            text: 要检查的文本
            placeholder_name: 要查找的占位符名称
            
        Returns:
            bool: 如果包含则为True，否则为False
        """
        pattern = r'\$\{' + re.escape(placeholder_name) + r'\}'
        return bool(re.search(pattern, text))

class HashPlaceholderParser(BasePlaceholderParser):
    """
    井号占位符解析器，支持#变量#格式
    """
    def __init__(self, config_manager=None):
        super().__init__(config_manager)
        # 设置占位符模式
        self.placeholder_patterns = [
            r'#(.*?)#',  # #变量#
        ]
    
    def get_placeholder_name(self, text, raw_match):
        """
        从匹配文本中提取占位符名称
        
        Args:
            text: 原始文本
            raw_match: 匹配的占位符原始文本
            
        Returns:
            str: 占位符名称
        """
        # 移除井号并去除空白
        return raw_match.replace('#', '').strip()
    
    def replace(self, text, replacements):
        """
        替换文本中的所有占位符
        
        Args:
            text: 要处理的文本
            replacements: 替换值字典，格式为 {placeholder_name: value}
            
        Returns:
            str: 替换后的文本
        """
        result_text = text
        
        # 查找所有占位符
        placeholders = self.find_all_placeholders(text)
        
        # 替换每个找到的占位符
        for placeholder_name, full_match, _ in placeholders:
            if placeholder_name in replacements:
                value = str(replacements[placeholder_name])
                # 特殊处理：如果值为空白，保留占位符
                if value.strip():
                    result_text = result_text.replace(full_match, value)
                    
        return result_text
        
    def has_placeholder(self, text, placeholder_name):
        """
        检查文本中是否包含指定占位符
        
        Args:
            text: 要检查的文本
            placeholder_name: 要查找的占位符名称
            
        Returns:
            bool: 如果包含则为True，否则为False
        """
        pattern = r'#' + re.escape(placeholder_name) + r'#'
        return bool(re.search(pattern, text))

class ImagePlaceholderParser(BasePlaceholderParser):
    """
    图片占位符解析器，支持{{img:变量}}格式
    """
    def __init__(self, config_manager=None):
        super().__init__(config_manager)
        self.image_prefix = "img:"
        # 设置占位符模式
        self.placeholder_patterns = [
            r'\{\{' + self.image_prefix + r'(.*?)\}\}',  # {{img:变量}}
        ]
    
    def get_placeholder_name(self, text, raw_match):
        """
        从匹配文本中提取占位符名称
        
        Args:
            text: 原始文本
            raw_match: 匹配的占位符原始文本
            
        Returns:
            str: 占位符名称，不包含前缀
        """
        # 移除大括号和前缀并去除空白
        name = raw_match.replace('{{' + self.image_prefix, '').replace('}}', '').strip()
        return name
    
    def replace(self, text, replacements):
        """
        替换文本中的所有图片占位符
        
        Args:
            text: 要处理的文本
            replacements: 替换值字典，格式为 {placeholder_name: image_path}
            
        Returns:
            str: 替换后的文本，实际上对于图片占位符，这个方法不会改变文本
            但会返回一个需要插入图片的列表
        """
        # 图片占位符不直接替换文本，而是在其他地方处理
        # 这里只是标记需要处理的位置
        return text
        
    def has_placeholder(self, text, placeholder_name):
        """
        检查文本中是否包含指定图片占位符
        
        Args:
            text: 要检查的文本
            placeholder_name: 要查找的占位符名称（不包含前缀）
            
        Returns:
            bool: 如果包含则为True，否则为False
        """
        # 检查完整的占位符格式 {{img:placeholder_name}}
        pattern = r'\{\{' + re.escape(self.image_prefix + placeholder_name) + r'\}\}'
        return bool(re.search(pattern, text))
    
    def is_image_placeholder(self, placeholder_name):
        """
        检查占位符名称是否为图片占位符
        
        Args:
            placeholder_name: 占位符名称
            
        Returns:
            bool: 如果是图片占位符则为True，否则为False
        """
        return placeholder_name.startswith(self.image_prefix) or placeholder_name == "img:照片"
    
    def get_clean_name(self, placeholder_name):
        """
        获取图片占位符的干净名称（不包含前缀）
        
        Args:
            placeholder_name: 占位符名称
            
        Returns:
            str: 不包含前缀的占位符名称
        """
        if self.is_image_placeholder(placeholder_name):
            if placeholder_name.startswith(self.image_prefix):
                return placeholder_name[len(self.image_prefix):].strip()
            elif placeholder_name == "img:照片":
                return "照片"
        return placeholder_name

class PlaceholderParserFactory:
    """
    占位符解析器工厂，用于创建和管理各种格式的占位符解析器
    """
    def __init__(self, config_manager=None):
        self.config_manager = config_manager
        self.parsers = [
            StandardPlaceholderParser(config_manager),  # {{变量}}
            DollarPlaceholderParser(config_manager),    # ${变量}
            HashPlaceholderParser(config_manager),      # #变量#
            ImagePlaceholderParser(config_manager),     # {{img:变量}}
        ]
    
    def get_parser_for_text(self, text):
        """
        获取适用于给定文本的解析器
        
        Args:
            text: 要检查的文本
            
        Returns:
            BasePlaceholderParser: 找到的第一个适用解析器，如果没有找到则返回None
        """
        for parser in self.parsers:
            if parser.contains_placeholders(text):
                return parser
        return None
        
    def find_all_placeholders(self, text):
        """
        在文本中查找所有格式的占位符
        
        Args:
            text: 要搜索的文本
            
        Returns:
            list: 占位符名称列表
        """
        placeholders = set()
        for parser in self.parsers:
            parser_placeholders = parser.extract_placeholders(text)
            placeholders.update(parser_placeholders)
        return list(placeholders)
        
    def replace_all(self, text, replacements):
        """
        替换文本中所有格式的占位符
        
        Args:
            text: 要处理的文本
            replacements: 替换值字典
            
        Returns:
            str: 替换后的文本
        """
        result = text
        for parser in self.parsers:
            if parser.contains_placeholders(result):
                result = parser.replace(result, replacements)
        return result 