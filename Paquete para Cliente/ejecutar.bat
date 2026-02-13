@echo off
chcp 65001 >nul
cd /d "%~dp0"

python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado.
    echo Descarguelo de https://www.python.org/downloads/
    echo Marque "Add Python to PATH" al instalar.
    pause
    exit /b 1
)

if not exist "venv" (
    echo Instalando dependencias (solo la primera vez)...
    python -m venv venv
    call venv\Scripts\activate.bat
    pip install -r requirements.txt -q
) else (
    call venv\Scripts\activate.bat
)

python interfaz_informe.py
pause
