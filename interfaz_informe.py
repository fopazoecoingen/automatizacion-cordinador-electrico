# Asegurar directorio del script en path (por si se ejecuta desde otra ruta)
import sys
from pathlib import Path
_script_dir = str(Path(__file__).resolve().parent)
if _script_dir not in sys.path:
    sys.path.insert(0, _script_dir)

from app.gui.informe import InterfazInforme, main

__all__ = ["InterfazInforme", "main"]


if __name__ == "__main__":
    main()
