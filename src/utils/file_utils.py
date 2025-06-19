"""
文件工具模块
处理文件操作的辅助功能
"""
import os
import shutil
from datetime import datetime
from loguru import logger


def ensure_dir(directory):
    """
    确保目录存在，如果不存在则创建
    
    Args:
        directory: 目录路径
        
    Returns:
        str: 目录路径
    """
    if not os.path.exists(directory):
        os.makedirs(directory)
        logger.info(f"创建目录: {directory}")
    return directory


def generate_filename(base_name, extension, include_timestamp=True):
    """
    生成文件名，可选择是否包含时间戳
    
    Args:
        base_name: 基础文件名
        extension: 文件扩展名（不含点）
        include_timestamp: 是否包含时间戳
        
    Returns:
        str: 生成的文件名
    """
    if include_timestamp:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return f"{base_name}_{timestamp}.{extension}"
    else:
        return f"{base_name}.{extension}"


def backup_file(file_path, backup_dir=None):
    """
    备份文件
    
    Args:
        file_path: 要备份的文件路径
        backup_dir: 备份目录，默认为原目录下的backups子目录
        
    Returns:
        str: 备份文件路径
    """
    if not os.path.exists(file_path):
        logger.warning(f"文件不存在，无法备份: {file_path}")
        return None
        
    # 确定备份目录
    if not backup_dir:
        parent_dir = os.path.dirname(file_path)
        backup_dir = os.path.join(parent_dir, "backups")
    
    # 确保备份目录存在
    ensure_dir(backup_dir)
    
    # 生成备份文件名
    filename = os.path.basename(file_path)
    name, ext = os.path.splitext(filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_filename = f"{name}_{timestamp}{ext}"
    backup_path = os.path.join(backup_dir, backup_filename)
    
    # 复制文件
    shutil.copy2(file_path, backup_path)
    logger.info(f"文件已备份: {backup_path}")
    
    return backup_path


def get_file_info(file_path):
    """
    获取文件信息
    
    Args:
        file_path: 文件路径
        
    Returns:
        dict: 文件信息字典
    """
    if not os.path.exists(file_path):
        logger.warning(f"文件不存在: {file_path}")
        return None
        
    try:
        stat_info = os.stat(file_path)
        file_info = {
            "path": file_path,
            "name": os.path.basename(file_path),
            "size": stat_info.st_size,
            "size_readable": get_readable_file_size(stat_info.st_size),
            "created": datetime.fromtimestamp(stat_info.st_ctime).strftime("%Y-%m-%d %H:%M:%S"),
            "modified": datetime.fromtimestamp(stat_info.st_mtime).strftime("%Y-%m-%d %H:%M:%S"),
            "extension": os.path.splitext(file_path)[1],
        }
        return file_info
    except Exception as e:
        logger.error(f"获取文件信息时出错: {str(e)}")
        return None


def get_readable_file_size(size_in_bytes):
    """
    获取可读的文件大小
    
    Args:
        size_in_bytes: 文件大小（字节）
        
    Returns:
        str: 可读的文件大小
    """
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_in_bytes < 1024.0:
            return f"{size_in_bytes:.2f} {unit}"
        size_in_bytes /= 1024.0
    return f"{size_in_bytes:.2f} PB" 