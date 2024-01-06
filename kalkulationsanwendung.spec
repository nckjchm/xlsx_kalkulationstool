# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['kalkulationsanwendung.py'],
    pathex=[],
    binaries=[],
    datas=[('resourcen/Kalkulationsvorlage.xlsx', 'resourcen/')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='kalkulationsanwendung',
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
    name='kalkulationsanwendung',
)
app = BUNDLE(
    coll,
    name='kalkulationsanwendung.app',
    icon=None,
    bundle_identifier=None,
)
