# -*- mode: python ; coding: utf-8 -*-

import os

a = Analysis(
    [os.path.join(os.path.dirname(__file__), 'web_ui.py')],
    pathex=[],
    binaries=[],
    datas=[('Notifica liq CON RD.doc', '.'), ('Notifica liq sin rd.doc', '.'), ('IEDriverServer.exe', '.')],
    hiddenimports=['main', 'pythoncom', 'win32com.client', 'win32gui', 'win32con'],
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
    name='notificaciones_sitfa',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
