# -*- mode: python ; coding: utf-8 -*-
import sys
import os

# 直接用当前工作目录作为根目录，彻底解决 __file__ 问题
root_path = os.getcwd()

a = Analysis(
    ['app.py'],
    pathex=[root_path],
    binaries=[],
    # 只打包存在的文件夹：templates（static 不存在，直接删掉）
    datas=[
        (os.path.join(root_path, 'templates'), 'templates'),
    ],
    hiddenimports=[
        'fastapi',
        'uvicorn',
        'jinja2',
        'pandas',
        'openpyxl',
        'python-multipart',
        'starlette',
        'aiofiles',
        'numpy'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='无课表工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)