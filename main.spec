# -*- mode: python ; coding: utf-8 -*-

import os

datas = [('Notifica liq CON RD.doc', '.'), ('Notifica liq sin rd.doc', '.')]
driver_path = os.path.join(os.getcwd(), 'IEDriverServer.exe')
if os.path.exists(driver_path):
    datas.append((driver_path, '.'))

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=['selenium', 'selenium.webdriver', 'selenium.webdriver.ie'],
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
    name='main',
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
