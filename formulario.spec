# formulario.spec
# PyInstaller spec para app Tkinter + pandas + docx + docx2pdf (com Word/COM)
# Gera um .exe único e inclui modelo .docx e config (se existir).

from PyInstaller.utils.hooks import collect_submodules, collect_data_files
import os

# módulos que às vezes o PyInstaller não detecta
hidden = [
    "comtypes",
    # pandas internals:
    "pandas._libs.tslibs.timedeltas",
    "pandas._libs.tslibs.np_datetime",
    "pandas._libs.tslibs.nattype",
    "pandas._libs.interval",
]

# dados do pandas (se necessário):
datas = collect_data_files('pandas', include_py_files=False)

# incluir o modelo docx e o config (se existirem)
if os.path.exists("anexo_geduc_se_requerimento_amparo_legall.docx"):
    datas.append(("anexo_geduc_se_requerimento_amparo_legall.docx", "."))

if os.path.exists("config_form.json"):
    datas.append(("config_form.json", "."))

block_cipher = None

a = Analysis(
    ['formulario.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['pytest','numpy.tests','pandas.tests'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='FormRequerimento',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,      # GUI sem console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if os.path.exists('icon.ico') else None
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='FormRequerimento'
)
