# -*- mode: python ; coding: utf-8 -*-
# Archivo de configuración para PyInstaller - Generador Informe Eléctrico

block_cipher = None

# Módulos a incluir explícitamente (evita errores de import dinámico)
hidden_imports = [
    'tkinter',
    'pandas',
    'openpyxl',
    'openpyxl.cell._writer',
    'openpyxl.styles',
    'openpyxl.utils',
    'requests',
    'tqdm',
    'pyxlsb',
    'win32com',
    'win32com.client',
    'pathlib',
    'zipfile',
    'threading',
    'json',
    'urllib.parse',
]

# Excluir módulos no usados (reduce tamaño)
excludes = ['IPython', 'jupyter', 'notebook', 'tkinter.test']

a = Analysis(
    ['interfaz_informe.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
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
    name='Generador Informe Electrico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Sin ventana de consola (aplicación GUI)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
