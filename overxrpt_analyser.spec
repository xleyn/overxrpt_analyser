# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path
base_path = Path('.').resolve()

a = Analysis(
    [str(base_path / 'src' / 'overxrpt_analyser' / 'main.py')],
    pathex=[str(base_path)],
    binaries=[],
    datas=[(str(base_path / 'config'), 'config/'), (str(base_path / 'docs'), 'docs/')],
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
    [],
    exclude_binaries=True,
    name='overxrpt_analyser',
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
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='.',
)
