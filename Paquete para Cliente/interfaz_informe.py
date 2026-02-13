import sys
import os
from pathlib import Path

# Paquete portable (Python embebido): agregar rutas necesarias ANTES de cualquier import
# 1) Directorio del ejecutable (python_embed) - alli esta _tkinter.pyd, tcl, etc.
_exe_dir = os.path.dirname(os.path.abspath(sys.executable))
if _exe_dir and _exe_dir not in sys.path:
    sys.path.insert(0, _exe_dir)
# 2) Directorio del script - para importar app, core
_script_dir = str(Path(__file__).resolve().parent)
if _script_dir not in sys.path:
    sys.path.insert(0, _script_dir)

from app.gui.informe import InterfazInforme, main

__all__ = ["InterfazInforme", "main"]


if __name__ == "__main__":
    main()
