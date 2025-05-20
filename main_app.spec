# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['main_app.py'],
    pathex=[],
    binaries=[],
    datas=[('WhisperR_icon.png', '.'), ('status_icons', 'status_icons')],
    hiddenimports=[
        'py7zr',
        'pycryptodome',
        'brotli', 
        'zstandard',
        'pyzstd',
        'py7zr.helpers',
        'py7zr.compressor'
    ],
    hookspath=['.', 'hook-py7zr.py'],
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
    [],
    exclude_binaries=True,
    name='main_app',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['WhisperR_icon.png'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main_app',
)
