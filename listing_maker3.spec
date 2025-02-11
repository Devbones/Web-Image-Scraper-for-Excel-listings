# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_data_files
import os

# Specify the paths
cacert_path = 'C:/Users/kusmirek_ar/Downloads/winpython/WPy32-31230/scripts/cacert.pem'
script_dir = 'C:/path/to/your/script'  # Manually set the path to the directory containing listing_maker3.py

a = Analysis(
    ['listing_maker3.py'],
    pathex=[script_dir],  # Manually set the script directory path
    binaries=[],
    datas=[(cacert_path, 'cacert.pem')],
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
    name='listing_maker3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Suppresses terminal window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
