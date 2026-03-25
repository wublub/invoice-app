# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = [('invoice_recognizer.py', '.'), ('invoice_icon.ico', '.')]
binaries = []
hiddenimports = ['openpyxl', 'fitz', 'pandas']
tmp_ret = collect_all('pymupdf')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]


a = Analysis(
    ['invoice_app.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['torch', 'scipy', 'matplotlib', 'IPython', 'PySide6', 'shiboken6', 'pyarrow', 'sympy', 'nbformat', 'jedi', 'parso', 'black', 'zmq', 'lxml', 'fsspec', 'jsonschema', 'pygments', 'psutil', 'certifi', 'lark', 'pydantic', 'setuptools', 'streamlit', 'streamlit_sortables'],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='发票识别工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='invoice_icon.ico',
)
