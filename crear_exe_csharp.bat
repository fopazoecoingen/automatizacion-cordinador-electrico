@echo off
chcp 65001 >nul
echo ========================================
echo Creando ejecutable C# - Generador Informe Electrico
echo ========================================
echo.

set DEST_DIR=Generador_Informe_Electrico_EXE
set PROYECTO=GeneradorInformeElectrico\GeneradorInformeElectrico.csproj

echo Paso 1: Restaurar y compilar...
dotnet restore "%PROYECTO%"
dotnet publish "%PROYECTO%" -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o publish_temp
if errorlevel 1 (
    echo ERROR: Fallo la compilacion.
    pause
    exit /b 1
)

echo.
echo Paso 2: Preparando carpeta de entrega...
if exist "%DEST_DIR%" rmdir /s /q "%DEST_DIR%" 2>nul
mkdir "%DEST_DIR%"

copy "publish_temp\GeneradorInformeElectrico.exe" "%DEST_DIR%\"
if exist "config.json" copy "config.json" "%DEST_DIR%\"
if exist "config_empresas.json" copy "config_empresas.json" "%DEST_DIR%\"
if not exist "%DEST_DIR%\config.json" (
    echo {"path_bd": "bd_data"} > "%DEST_DIR%\config.json"
)

echo.
echo Paso 3: Creando LEEME.txt...
(
echo =====================================================
echo   GENERADOR DE INFORME ELECTRICO - Version C#
echo =====================================================
echo.
echo EJECUTABLE NATIVO: Doble clic en GeneradorInformeElectrico.exe
echo REQUISITOS: Windows 10+, Excel, Internet.
echo.
echo USO:
echo 1. Doble clic en GeneradorInformeElectrico.exe
echo 2. Seleccione ano, mes, empresa, barra
echo 3. Examinar en Plantilla: elija SU plantilla .xlsx o .xlsm
echo 4. Examinar en Destino: donde guardar el informe
echo 5. Clic en Crear Informe
echo.
echo CONFIGURACION:
echo - config.json: ruta de base de datos ^(path_bd^)
echo - config_empresas.json: empresas y filtros personalizados
echo.
) > "%DEST_DIR%\LEEME.txt"

echo.
echo Paso 4: Limpiando...
rmdir /s /q publish_temp 2>nul

echo.
echo ========================================
echo   EJECUTABLE CREADO
echo ========================================
echo.
echo Carpeta: %DEST_DIR%\
echo.
pause
