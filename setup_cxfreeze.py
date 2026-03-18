# -*- coding: utf-8 -*-
"""
Setup para cx_Freeze - alternativa a PyInstaller.
Genera el ejecutable con: python setup_cxfreeze.py build
El .exe queda en build/Generador_Informe_Electrico/
"""
import sys
from cx_Freeze import setup, Executable

base = None
if sys.platform == "win32":
    base = "Win32GUI"  # Sin ventana de consola

build_exe_options = {
    "packages": ["tkinter", "openpyxl", "pandas", "requests", "tqdm", "pyxlsb", "win32com", "pythoncom", "pywintypes", "core", "app"],
    "excludes": [],
    "include_files": [],
    "build_exe": "build/Generador_Informe_Electrico",  # Carpeta de salida fija
}

setup(
    name="Generador_Informe_Electrico",
    version="1.0",
    description="Generador de Informe Electrico",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            "interfaz_informe.py",
            base=base,
            target_name="Generador_Informe_Electrico.exe",
        )
    ],
)
