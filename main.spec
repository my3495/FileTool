# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None

# 确保logs目录存在
if not os.path.exists('logs'):
    os.makedirs('logs')

a = Analysis(
    ['main.py'],  # 应用程序入口点
    pathex=['.'],  # 项目根目录
    binaries=[],
    datas=[
        ('src', 'src'),  # 包含源代码
        ('icon.ico', '.'),  # 图标文件
        ('logs', 'logs'),  # 包含logs目录
    ],
    hiddenimports=['PySide6'],  # 明确包含PySide6模块
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FileTool',  # 生成的可执行文件名称
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  # 禁用控制台窗口(GUI应用)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # 设置应用图标
    version='file_version_info.txt',  # 版本信息文件
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='FileTool',
)

#pyinstaller main.spec