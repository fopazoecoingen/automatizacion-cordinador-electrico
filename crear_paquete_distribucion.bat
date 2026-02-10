@echo off
echo ========================================
echo Creando paquete de distribucion
echo ========================================
echo.

REM Verificar que Python este disponible
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado o no esta en el PATH
    echo Por favor instale Python 3.8 o superior
    pause
    exit /b 1
)

echo Verificando Python...
python --version

echo.
echo Verificando PyInstaller...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller no esta instalado. Instalando...
    python -m pip install --upgrade pip
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: No se pudo instalar PyInstaller
        pause
        exit /b 1
    )
    echo PyInstaller instalado correctamente.
) else (
    echo PyInstaller ya esta instalado.
)

echo.
echo Paso 1: Limpiando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Paquete Distribucion" rmdir /s /q "Paquete Distribucion"
if exist "Generador_Informe_Electrico.zip" del /q "Generador_Informe_Electrico.zip"

echo.
echo Paso 2: Construyendo ejecutable...
python -m PyInstaller --name="Generador Informe Electrico" ^
    --onefile ^
    --windowed ^
    --icon=NONE ^
    --add-data "v1;v1" ^
    --hidden-import=tkinter ^
    --hidden-import=pandas ^
    --hidden-import=openpyxl ^
    --hidden-import=requests ^
    --hidden-import=tqdm ^
    --hidden-import=pathlib ^
    --hidden-import=urllib.parse ^
    --hidden-import=zipfile ^
    --hidden-import=threading ^
    --collect-all=pandas ^
    --collect-all=openpyxl ^
    -m app

if errorlevel 1 (
    echo.
    echo ERROR: La construccion del ejecutable fallo
    pause
    exit /b 1
)

echo.
echo Paso 3: Creando carpeta de distribucion...
mkdir "Paquete Distribucion"

echo.
echo Paso 4: Copiando archivos necesarios...
copy "dist\Generador Informe Electrico.exe" "Paquete Distribucion\" >nul
if exist "plantilla_base.xlsx" (
    copy "plantilla_base.xlsx" "Paquete Distribucion\" >nul
    echo   [OK] plantilla_base.xlsx copiado
) else (
    echo   [ADVERTENCIA] plantilla_base.xlsx no encontrado
)

if exist "README_USUARIO_FINAL.txt" (
    copy "README_USUARIO_FINAL.txt" "Paquete Distribucion\" >nul
    echo   [OK] README_USUARIO_FINAL.txt copiado
) else (
    echo   [ADVERTENCIA] README_USUARIO_FINAL.txt no encontrado
    echo   [INFO] Creando README basico...
    echo Instrucciones de uso del Generador de Informe Electrico > "Paquete Distribucion\README_USUARIO_FINAL.txt"
    echo. >> "Paquete Distribucion\README_USUARIO_FINAL.txt"
    echo 1. Ejecute "Generador Informe Electrico.exe" >> "Paquete Distribucion\README_USUARIO_FINAL.txt"
    echo 2. Complete los campos en la interfaz >> "Paquete Distribucion\README_USUARIO_FINAL.txt"
    echo 3. Seleccione el rango de fechas >> "Paquete Distribucion\README_USUARIO_FINAL.txt"
    echo 4. Haga clic en "Crear Informe" >> "Paquete Distribucion\README_USUARIO_FINAL.txt"
)

echo.
echo Paso 5: Creando archivo ZIP de distribucion...
cd "Paquete Distribucion"
powershell -Command "Compress-Archive -Path * -DestinationPath ..\Generador_Informe_Electrico.zip -Force"
cd ..

if exist "Generador_Informe_Electrico.zip" (
    echo.
    echo ========================================
    echo Paquete creado exitosamente!
    echo ========================================
    echo.
    echo Archivos creados:
    echo   - Paquete Distribucion\ (carpeta completa)
    echo   - Generador_Informe_Electrico.zip (archivo para distribuir)
    echo.
    echo El paquete esta listo para entregar al usuario final.
    echo.
) else (
    echo.
    echo ERROR: No se pudo crear el archivo ZIP
    echo El paquete esta en la carpeta "Paquete Distribucion"
)

pause
