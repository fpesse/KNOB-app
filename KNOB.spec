# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['win32timezone']
hiddenimports += collect_submodules('pycaw')
hiddenimports += collect_submodules('comtypes')


a = Analysis(
    ['audio_mixer_app.py'],
    pathex=[],
    binaries=[],
    datas=[('assets\\AmazingViews.ttf', 'assets'), ('assets\\knob.ico', 'assets'), ('assets\\knob_256.png', 'assets')],
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
    a.binaries,
    a.datas,
    [],
    name='KNOB',
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
    icon=['assets\\knob.ico'],
)
