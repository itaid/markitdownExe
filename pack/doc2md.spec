# -*- mode: python ; coding: utf-8 -*-
"""doc2md Windows 单文件 EXE 打包配置
用法（Windows + Python 3.12 环境）:
    python -m pip install -e ".[gui,ocr]" pyinstaller
    pyinstaller pack/doc2md.spec --noconfirm
"""
import os
from PyInstaller.utils.hooks import collect_all

HERE = os.path.dirname(os.path.abspath(SPEC))  # noqa: F821 (PyInstaller 注入)

datas = []
binaries = []
hiddenimports = []

# markitdown 读取的是包数据（magika ML 模型、扩展名表）；OCR 引擎是原生 DLL
for package in [
    "markitdown", "magika",
    "onnxruntime", "rapidocr_onnxruntime",
    "pypdfium2",
    "cv2", "numpy",
    "pdfminer", "pdfplumber", "mammoth",
    "lxml", "openpyxl", "pandas", "xlrd",
    "olefile", "charset_normalizer",
]:
    try:
        d, b, h = collect_all(package)
        datas += d
        binaries += b
        hiddenimports += h
    except Exception as e:  # noqa: BLE001
        print(f"[doc2md.spec] skip {package}: {e}")

# 图标
icon_path = os.path.join(HERE, "icon.ico")

a = Analysis(
    ["entry_gui.py"],                     # PyInstaller 会把 CWD 切到 spec 所在目录(pack/)
    pathex=[os.path.join(HERE, "..", "src")],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports + ["doc2md_tool", "doc2md_tool.gui",
                                   "doc2md_tool.service", "doc2md_tool.ocr"],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "torch"],   # 精简体积
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, a.binaries, a.datas, [],
    name="doc2md",
    debug=False,
    strip=False,
    upx=False,
    console=False,                       # GUI 程序不弹黑框
    icon=icon_path if os.path.exists(icon_path) else None,
)
