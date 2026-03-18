@echo off
chcp 65001 >nul
echo ============================================
echo  Generador Informe Electrico - Modo con Logs
echo ============================================
echo.

set DEST_DIR=%~dp0Generador_Informe_Electrico_EXE
set EXE=%DEST_DIR%\GeneradorInformeElectrico_v3.exe
if not exist "%EXE%" set EXE=%DEST_DIR%\GeneradorInformeElectrico_v2.exe
if not exist "%EXE%" set EXE=%DEST_DIR%\GeneradorInformeElectrico_NUEVO.exe
if not exist "%EXE%" set EXE=%DEST_DIR%\GeneradorInformeElectrico.exe
set LOG=%DEST_DIR%\GeneradorInformeElectrico.log

if not exist "%EXE%" (
    echo ERROR: No se encuentra el ejecutable.
    echo Ejecute primero crear_exe_csharp.bat
    pause
    exit /b 1
)

echo Log: %LOG%
echo.
echo Iniciando aplicacion...
echo Despues de usar la app, el log estara en: %LOG%
echo.

cd /d "%DEST_DIR%"
start "" "%EXE%"

echo.
echo La aplicacion se ha abierto.
echo Cuando haya terminado de usarla (o si ocurre un error), vuelva aqui.
echo.
pause

echo.
if exist "%LOG%" (
    echo --- Ultimas lineas del log ---
    powershell -Command "Get-Content '%LOG%' -Tail 80"
    echo.
    echo Abriendo log completo en Bloc de notas...
    notepad "%LOG%"
) else (
    echo No se encontro el archivo de log en: %LOG%
    echo El log se crea al iniciar la aplicacion.
)
exit /b 0
