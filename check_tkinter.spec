# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['check_tkinter.py'],
    pathex=[],
    binaries=[('C:/Program Files/wkhtmltopdf/bin/wkhtmltopdf.exe', 'wkhtmltopdf/bin'), ('C:/Program Files/wkhtmltopdf/bin/wkhtmltoimage.exe', 'wkhtmltopdf/bin')],
    datas=[],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='check_tkinter',
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
)
