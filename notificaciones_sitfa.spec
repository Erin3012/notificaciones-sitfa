# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['C:\\Users\\Eri\\Desktop\\Programación\\notificaciones sitfa\\main.py'],
    pathex=[],
    binaries=[],
    datas=[('Notifica liq CON RD.doc', '.'), ('Notifica liq sin rd.doc', '.')],
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
