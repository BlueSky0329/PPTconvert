# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 配置：生成 Windows 独立可执行文件（含 Python 与依赖）。
用法：pyinstaller PPTconvert.spec
"""
from PyInstaller.utils.hooks import collect_all

# ttkbootstrap 主题与资源
_datas_ttk, _bins_ttk, _hidden_ttk = collect_all("ttkbootstrap")

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=_bins_ttk,
    datas=_datas_ttk,
    hiddenimports=_hidden_ttk
    + [
        "docx",
        "docx.oxml",
        "docx.opc",
        "pptx",
        "pptx.util",
        "pptx.parts",
        "pptx.enum",
        "pptx.enum.text",
        "pptx.enum.shapes",
        "PIL",
        "PIL._tkinter_finder",
        "lxml",
        "lxml.etree",
        "lxml._elementpath",
        "gui",
        "gui.app",
        "gui.font_data",
        "gui.ui_constants",
        "core",
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="PPTconvert",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
