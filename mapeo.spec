# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['mapeo.py'],
    pathex=[],
    binaries=[],
    datas=[('C:\\Python\\Mapeo\\data\\litologia.csv', '.'), ('C:\\Python\\Mapeo\\data\\volalt.csv', '.'), ('C:\\Python\\Mapeo\\data\\alteracion.csv', '.'), ('C:\\Python\\Mapeo\\data\\estructuras.csv', '.'), ('C:\\Python\\Mapeo\\data\\mineralizacion.csv', '.'), ('C:\\Python\\Mapeo\\data\\minzone.csv', '.'), ('C:\\Python\\Mapeo\\data\\pisotecho.csv', '.'), ('C:\\Python\\Mapeo\\data\\vetillas.csv', '.'), ('C:\\Python\\Mapeo\\data\\ug.csv', '.')],
    hiddenimports=['comtypes.stream', 'comtypes'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='mapeo',
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
