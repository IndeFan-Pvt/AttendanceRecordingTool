# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files
from PyInstaller.utils.hooks import collect_dynamic_libs
from PyInstaller.utils.hooks import collect_submodules

datas = [
    ('.\\docs\\guides\\3分で使う_GUI版.md', '.'),
]
candidate = Path('.\\shift_config.json')
if candidate.exists():
    datas.append((str(candidate), '.'))
for candidate in [
    Path('.\\docs\\guides\\3分で使う_GUI版.html'),
    Path('.\\3分で使う_GUI版.html'),
    Path('.\\exe\\generate_akanecco_shift_gui\\3分で使う_GUI版.html'),
    Path('.\\exe\\generate_akanecco_shift_gui\\_internal\\3分で使う_GUI版.html'),
]:
    if candidate.exists():
        datas.append((str(candidate), '.'))
        break
binaries = []
hiddenimports = ['win32timezone']
datas += collect_data_files('ortools')
binaries += collect_dynamic_libs('ortools')
hiddenimports += collect_submodules('ortools')


a = Analysis(
    ['generate_shift_gui.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
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
    [],
    exclude_binaries=True,
    name='generate_akanecco_shift_gui',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='generate_akanecco_shift_gui',
)
