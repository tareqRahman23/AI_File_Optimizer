# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules, copy_metadata


block_cipher = None

datas = [
    ("../assets/app.ico", "assets"),
]
binaries = []
hiddenimports = []

packages = [
    "markitdown",
    "magika",
    "markdownify",
    "bs4",
    "mammoth",
    "openpyxl",
    "pandas",
    "pdfminer",
    "pdfplumber",
    "PIL",
    "pptx",
    "speech_recognition",
    "youtube_transcript_api",
]

for package in packages:
    package_datas, package_binaries, package_hiddenimports = collect_all(package)
    datas += package_datas
    binaries += package_binaries
    hiddenimports += package_hiddenimports

for package in [
    "azure",
    "azure.ai",
    "azure.ai.contentunderstanding",
    "azure.ai.documentintelligence",
    "azure.core",
    "azure.identity",
    "onnxruntime",
]:
    hiddenimports += collect_submodules(package)
    datas += collect_data_files(package)

for package in [
    "markitdown",
    "magika",
    "onnxruntime",
    "pandas",
    "pdfminer.six",
    "pdfplumber",
    "PySide6",
]:
    try:
        datas += copy_metadata(package)
    except Exception:
        pass

a = Analysis(
    ["../app.py"],
    pathex=[".."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    name="File Optimizer for AI",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="../assets/app.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="File Optimizer for AI",
)
