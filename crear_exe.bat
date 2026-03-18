@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
echo ========================================
echo Creando ejecutable .EXE
echo Generador Informe Electrico
echo ========================================
echo.

set DEST_DIR=Generador_Informe_Electrico_EXE
set ZIP_OUT=Generador_Informe_Electrico_EXE.zip

REM Verificar que PyInstaller está instalado
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller no está instalado. Instalando...
    pip install pyinstaller -q
    if errorlevel 1 (
        echo ERROR: No se pudo instalar PyInstaller.
        pause
        exit /b 1
    )
)

echo Paso 1: Compilando con PyInstaller...
python -m PyInstaller --noconfirm interfaz_informe.spec
if errorlevel 1 (
    echo ERROR: Falló la compilación.
    pause
    exit /b 1
)

echo.
echo Paso 2: Preparando carpeta de entrega...
if exist "%DEST_DIR%" rmdir /s /q "%DEST_DIR%" 2>nul
mkdir "%DEST_DIR%"

REM Copiar el .exe generado (está en dist/)
copy "dist\Generador_Informe_Electrico.exe" "%DEST_DIR%\" >nul
if errorlevel 1 (
    echo ERROR: No se encontró el .exe en dist\
    pause
    exit /b 1
)

REM Copiar archivos de configuración (el cliente puede editarlos)
if exist "config.json" copy "config.json" "%DEST_DIR%\" >nul
if exist "config_empresas.json" copy "config_empresas.json" "%DEST_DIR%\" >nul

REM Si no existen, crear config.json por defecto
if not exist "%DEST_DIR%\config.json" (
    powershell -NoProfile -Command "[IO.File]::WriteAllText((Join-Path (Get-Location) '%DEST_DIR%\config.json'), '{\"path_bd\": \"bd_data\"}')"
)

echo.
echo Paso 3: Creando LEEME.txt...
(
echo =====================================================
echo   GENERADOR DE INFORME ELECTRICO - Version EXE
echo =====================================================
echo.
echo SIN INSTALACION: Descomprima y ejecute el .exe
echo REQUISITOS: Windows 10+, Excel, Internet.
echo.
echo USO:
echo 1. Doble clic en "Generador_Informe_Electrico.exe"
echo 2. Seleccione ano, mes, empresa, barra
echo 3. "Examinar" en Plantilla: elija SU plantilla .xlsx
echo 4. "Examinar" en Destino: donde guardar el informe
echo 5. Clic en "Crear Informe"
echo.
echo CONFIGURACION:
echo - config.json: ruta de base de datos ^(path_bd^)
echo - config_empresas.json: empresas y filtros personalizados
echo.
) > "%DEST_DIR%\LEEME.txt"

echo.
echo Paso 4: Creando ZIP...
if exist "%ZIP_OUT%" del "%ZIP_OUT%"
powershell -NoProfile -Command "Compress-Archive -Path '%DEST_DIR%' -DestinationPath '%ZIP_OUT%' -Force"

echo.
echo ========================================
echo   EJECUTABLE CREADO
echo ========================================
echo.
echo Carpeta: %DEST_DIR%\
echo ZIP: %ZIP_OUT%
echo.
echo Entregue %ZIP_OUT% al cliente. Descomprima y ejecute el .exe.
echo.
echo NOTA: Algunos antivirus pueden marcar ejecutables generados
echo       por PyInstaller. Es normal si el .exe no está firmado.
echo       El cliente puede agregar una excepción si es necesario.
echo.
pause
