# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['pythoncom', 'pywintypes', 'win32timezone']
hiddenimports += ['fitz', 'PIL', 'PIL.Image']
hiddenimports += collect_submodules('win32com')


a = Analysis(
    ['bufs\\hwp_style_mvp.py'],
    pathex=[],
    binaries=[],
    datas=[('bufs\\style-sets.json', 'bufs'), ('bufs\\table-settings.json', 'bufs'), ('bufs\\update-settings.json', 'bufs'), ('bufs\\style-order.json', 'bufs'), ('bufs\\templates', 'bufs\\templates'), ('bufs\\icons', 'bufs\\icons'), ('bufs\\logos', 'bufs\\logos')],
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
    name='BUFS-HWP-Editor',
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
    icon='bufs\\icons\\app.ico',
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BUFS-HWP-Editor',
)
