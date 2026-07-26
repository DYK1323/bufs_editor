# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['pythoncom', 'pywintypes', 'win32timezone']
hiddenimports += ['fitz', 'PIL', 'PIL.Image']
hiddenimports += collect_submodules('win32com')


a = Analysis(
    ['C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\hwp_style_mvp.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\style-sets.json', 'bufs'), ('C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\table-settings.json', 'bufs'), ('C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\style-order.json', 'bufs'), ('C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\templates', 'bufs\\templates'), ('C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\icons', 'bufs\\icons'), ('C:\\Users\\DAYOUNG\\bufs_editor\\bufs\\logos', 'bufs\\logos')],
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
