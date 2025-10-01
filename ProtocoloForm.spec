# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['formulario.py'],
    pathex=[],
    binaries=[],
    datas=[('anexo_geduc_se_requerimento_amparo_legall.docx', '.')],
    hiddenimports=['comtypes', 'win32com', 'pythoncom'],
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
    name='ProtocoloForm',
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
)
