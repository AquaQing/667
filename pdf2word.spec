# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 打包配置 —— PDF 转 Word 工具

用法（在项目目录下执行）：
    pyinstaller pdf2word.spec --clean --noconfirm

产物：dist/pdf2word.exe （Windows）或 dist/pdf2word （Linux/macOS）
"""

block_cipher = None

# pdf2docx 依赖的一些库需要显式声明，否则打包后运行会报 ModuleNotFoundError
hidden_imports = [
    "pdf2docx",
    "fitz",                 # PyMuPDF，pdf2docx 的核心依赖
    "pymupdf",
    "docx",                 # python-docx
    "lxml",
    "lxml.etree",
    "PIL",
    "PIL.Image",
    "fire",
    "termcolor",
    "numpy",
    "cv2",                  # OpenCV，pdf2docx 用于版面分析
]

a = Analysis(
    ["pdf2word.py"],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # 体积优化：排除用不到的重型库
    excludes=[
        "tkinter",
        "matplotlib",
        "scipy",
        "pandas",
        "IPython",
        "notebook",
        "pytest",
        "setuptools",
        "pip",
    ],
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
    name="pdf2word",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    # 关键：保留控制台窗口，才能看到转换进度
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # 单文件模式：整个程序打包成一个 exe，方便拷来拷去
    onefile=True,
)
