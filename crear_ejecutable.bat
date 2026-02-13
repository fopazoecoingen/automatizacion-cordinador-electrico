@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
echo ========================================
echo Creando ejecutable / paquete portable
echo Generador Informe Electrico
echo ========================================
echo.
echo El cliente descomprime el ZIP y ejecuta. Sin instalacion. Con win32/Excel.
echo.

set DEST_DIR=Generador_Informe_Electrico_Portable
set ZIP_OUT=Generador_Informe_Electrico_Portable.zip
set WINPYTHON_URL=https://github.com/winpython/winpython/releases/download/17.2.20251222post1/WinPython64-3.13.11.0dot_post1.zip

REM Paso 1: Descargar WinPython si no existe
set WINPY_ZIP=WinPython64-3.13.11.0dot_post1.zip
if not exist "%WINPY_ZIP%" (
    echo Paso 1: Descargando WinPython ~27 MB...
    powershell -NoProfile -Command ^
        "try { " ^
        "  [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " ^
        "  Invoke-WebRequest -Uri '%WINPYTHON_URL%' -OutFile '%WINPY_ZIP%' -UseBasicParsing; " ^
        "  exit 0 " ^
        "} catch { Write-Error $_.Exception.Message; exit 1 }"
    if errorlevel 1 (
        echo ERROR: No se pudo descargar WinPython.
        pause
        exit /b 1
    )
) else (
    echo Paso 1: Usando WinPython ya descargado...
)

echo.
echo Paso 2: Limpiando y preparando...
if exist "%DEST_DIR%" rmdir /s /q "%DEST_DIR%" 2>nul
if exist "%DEST_DIR%" (
    echo ERROR: No se puede borrar %DEST_DIR%. Cierrela y reintente.
    pause
    exit /b 1
)
mkdir "%DEST_DIR%"

echo.
echo Paso 3: Extrayendo WinPython directamente al paquete...
powershell -NoProfile -Command "Expand-Archive -Path '%WINPY_ZIP%' -DestinationPath '%DEST_DIR%' -Force"
if errorlevel 1 (
    echo ERROR al extraer. Pruebe extraer %WINPY_ZIP% manualmente.
    pause
    exit /b 1
)

REM WinPython extrae una carpeta WPy64-313110 o similar. Renombrar a python_embed.
set "WP_FOLDER="
for /d %%d in ("%DEST_DIR%\WPy*") do set "WP_FOLDER=%%d"
for /d %%d in ("%DEST_DIR%\Winpython*") do if not defined WP_FOLDER set "WP_FOLDER=%%d"
for /d %%d in ("%DEST_DIR%\*") do if not defined WP_FOLDER set "WP_FOLDER=%%d"
if defined WP_FOLDER (
    move "!WP_FOLDER!" "%DEST_DIR%\python_embed" >nul 2>&1
)
if not exist "%DEST_DIR%\python_embed" (
    echo ERROR: No se pudo organizar WinPython
    pause
    exit /b 1
)

REM Buscar python.exe principal en el paquete (excluir Scripts)
set "PYTHON_FINAL="
for /f "delims=" %%f in ('dir /s /b "%DEST_DIR%\python_embed\python.exe" 2^>nul ^| findstr /v /i "Scripts"') do (
    if not defined PYTHON_FINAL set "PYTHON_FINAL=%%f" & goto :pip_ready
)
:pip_ready

if not defined PYTHON_FINAL (
    echo ERROR: No se encontro python.exe en el paquete
    pause
    exit /b 1
)

echo.
echo Paso 4: Instalando dependencias...
"%PYTHON_FINAL%" -m pip install --quiet --no-warn-script-location -r "%~dp0requirements_portable.txt"
if errorlevel 1 (
    echo Intentando pip install sin --quiet...
    "%PYTHON_FINAL%" -m pip install -r "%~dp0requirements_portable.txt"
    if errorlevel 1 (
        echo ERROR en pip install
        pause
        exit /b 1
    )
)

echo.
echo Paso 5: Copiando codigo...
xcopy /E /I /Y "app" "%DEST_DIR%\app\" >nul
xcopy /E /I /Y "core" "%DEST_DIR%\core\" >nul
copy "interfaz_informe.py" "%DEST_DIR%\" >nul

if exist "config_empresas.json" copy "config_empresas.json" "%DEST_DIR%\" >nul
REM No incluir plantilla: el cliente ingresa la suya con Examinar

echo.
echo Paso 6: Creando launcher...
set "LAUNCHER=%DEST_DIR%\Generador Informe Electrico.bat"
echo @echo off > "%LAUNCHER%"
echo chcp 65001 ^>nul >> "%LAUNCHER%"
echo cd /d "%%~dp0" >> "%LAUNCHER%"
echo "python_embed\python\python.exe" interfaz_informe.py >> "%LAUNCHER%"
echo if errorlevel 1 pause >> "%LAUNCHER%"

echo.
echo Paso 7: Creando LEEME.txt...
(
echo =====================================================
echo   GENERADOR DE INFORME ELECTRICO - Version Portable
echo =====================================================
echo.
echo SIN INSTALACION: Descomprima y ejecute.
echo REQUISITOS: Windows 10+, Excel, Internet.
echo.
echo USO:
echo 1. Doble clic en "Generador Informe Electrico.bat"
echo 2. Seleccione ano, mes, empresa, barra
echo 3. "Examinar" en Plantilla: elija SU plantilla .xlsx
echo 4. "Examinar" en Destino: donde guardar el informe
echo 5. Clic en "Crear Informe"
echo.
) > "%DEST_DIR%\LEEME.txt"

echo.
echo Paso 8: Creando ZIP...
if exist "%ZIP_OUT%" del "%ZIP_OUT%"
powershell -NoProfile -Command "Compress-Archive -Path '%DEST_DIR%' -DestinationPath '%ZIP_OUT%' -Force"

echo.
echo ========================================
echo   PAQUETE CREADO
echo ========================================
echo.
echo Entregue %ZIP_OUT% al cliente.
echo.
pause
