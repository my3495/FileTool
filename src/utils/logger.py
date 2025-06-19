"""
日志工具模块
配置日志记录器
"""
import os
import sys
from loguru import logger


def get_app_path():
    """
    获取应用程序的根路径。
    - 在开发环境中，是当前工作目录。
    - 在打包后的可执行文件中，是可执行文件所在的目录。
    """
    if getattr(sys, 'frozen', False):
        # PyInstaller 打包后的路径
        return os.path.dirname(sys.executable)
    else:
        # 开发环境的路径
        return os.path.abspath(".")


def setup_logger(log_dir=None):
    """
    设置日志记录器
    
    Args:
        log_dir: 日志保存目录。如果为 None，则在应用根目录创建 "logs" 文件夹。
    """
    # 移除默认的处理器
    logger.remove()
    
    # 添加控制台处理器，当打包成exe且无控制台时，sys.stdout为None
    if sys.stdout:
        logger.add(
            sys.stdout,
            level="INFO",
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | <cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> - <level>{message}</level>"
        )
    
    try:
        if log_dir is None:
            # 设置默认日志目录为应用根目录下的 "logs" 文件夹
            log_dir = os.path.join(get_app_path(), "logs")

        # 确保日志目录存在
        if not os.path.exists(log_dir):
            os.makedirs(log_dir)
        
        # 日志文件名
        log_path = os.path.join(log_dir, "app.log")
        
        # 添加文件处理器
        logger.add(
            log_path,
            format="{time:YYYY-MM-DD HH:mm:ss.SSS} | {level: <8} | {name}:{function}:{line} - {message}",
            level="DEBUG",
            rotation="1 day",  # 每天轮转一个新文件
            retention="30 days",  # 保留30天的日志
            encoding="utf-8"
        )
        
        logger.info(f"日志文件已配置，保存路径: {log_path}")
    except Exception as e:
        logger.warning(f"无法配置日志文件: {str(e)}，日志将不会写入文件。")
    
    return logger 