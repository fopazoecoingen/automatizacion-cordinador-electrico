@echo off
chcp 65001 >nul
echo ========================================
echo Creando paquete de distribucion
echo Generador Informe Electrico
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
echo Instalando dependencias...
python -m pip install --upgrade pip --quiet
python -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo ERROR: No se pudieron instalar las dependencias
    pause
    exit /b 1
)

echo.
echo Verificando PyInstaller...
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo PyInstaller no esta instalado. Instalando...
    python -m pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: No se pudo instalar PyInstaller
        pause
        exit /b 1
    )
)

echo.
echo Paso 1: Limpiando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Paquete Distribucion" rmdir /s /q "Paquete Distribucion"
if exist "Generador_Informe_Electrico.zip" del /q "Generador_Informe_Electrico.zip"

echo.
echo Paso 2: Construyendo ejecutable (esto puede tardar varios minutos)...
python -m PyInstaller GeneradorInforme.spec --noconfirm

if errorlevel 1 (
    echo.
    echo ERROR: La construccion del ejecutable fallo
    pause
    exit /b 1
)

echo.
echo Paso 3: Creando carpeta de distribucion...
mkdir "Paquete Distribucion" 2>nul

echo.
echo Paso 4: Copiando archivos necesarios...
copy "dist\Generador Informe Electrico.exe" "Paquete Distribucion\" >nul
echo   [OK] Generador Informe Electrico.exe

if exist "plantilla_base.xlsx" (
    copy "plantilla_base.xlsx" "Paquete Distribucion\" >nul
    echo   [OK] plantilla_base.xlsx
) else (
    echo   [ADVERTENCIA] plantilla_base.xlsx no encontrado - el cliente debera tener su propia plantilla
)

if exist "config_empresas.json" (
    copy "config_empresas.json" "Paquete Distribucion\" >nul
    echo   [OK] config_empresas.json
) else (
    echo   [ADVERTENCIA] config_empresas.json no encontrado - el cliente tendra que escribir Empresa/Barra/Medidor manualmente
)

echo.
echo Creando README para el usuario final...
(
echo =====================================================
echo   GENERADOR DE INFORME ELECTRICO - Instrucciones
echo =====================================================
echo.
echo REQUISITOS DEL EQUIPO:
echo   - Windows 10 o superior
echo   - Microsoft Excel instalado ^(para escribir en la plantilla con maxima fidelidad^)
echo   - Conexion a Internet ^(para descargar datos de PLABACOM^)
echo.
echo INSTRUCCIONES DE USO:
echo.
echo 1. Ejecute "Generador Informe Electrico.exe"
echo 2. Seleccione Ano y Mes del informe
echo 3. Seleccione Empresa, Barra y Nombre Medidor de la lista ^(o escriba manualmente si aplica^)
echo 4. Seleccione la plantilla base del cliente
echo 5. Elija la ruta de destino donde guardar el informe
echo 6. Haga clic en "Crear Informe"
echo.
echo El programa descargara automaticamente los datos necesarios
echo y generara el informe con los valores calculados.
echo.
echo BASE DE DATOS INTERNA:
echo Los archivos descargados se almacenan en una carpeta interna del sistema
echo ^(AppData\Local\GeneradorInformeElectrico^) y se van acumulando con el uso.
echo Si un periodo ya fue descargado antes, no se vuelve a descargar.
echo.
echo NOTA: La primera ejecucion puede tardar unos segundos al iniciar.
echo.
) > "Paquete Distribucion\README.txt"

echo.
echo Paso 5: Creando archivo ZIP para distribucion...
cd "Paquete Distribucion"
powershell -Command "Compress-Archive -Path * -DestinationPath '..\Generador_Informe_Electrico.zip' -Force"
cd ..

if exist "Generador_Informe_Electrico.zip" (
    echo.
    echo ========================================
    echo Paquete creado exitosamente!
    echo ========================================
    echo.
    echo Archivos generados:
    echo   - Paquete Distribucion\ ^(carpeta con el ejecutable y documentos^)
    echo   - Generador_Informe_Electrico.zip ^(archivo listo para entregar al cliente^)
    echo.
    echo Entregue el ZIP al cliente. Debe descomprimirlo y ejecutar
    echo "Generador Informe Electrico.exe". Se requiere Excel instalado.
) else (
    echo.
    echo ERROR: No se pudo crear el archivo ZIP
    echo El paquete esta en la carpeta "Paquete Distribucion"
)

echo.
pause
