#!/bin/bash
# Script para crear ejecutable (.app) en macOS
# Uso: ./crear_paquete_distribucion_mac.sh

echo "========================================"
echo "Creando paquete para macOS"
echo "Generador Informe Electrico"
echo "========================================"
echo ""

# Verificar Python
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 no está instalado."
    echo "Instale Python 3.8+ desde python.org o con Homebrew (brew install python)"
    exit 1
fi

echo "Python: $(python3 --version)"

# Instalar dependencias
echo ""
echo "Instalando dependencias..."
python3 -m pip install --upgrade pip -q
python3 -m pip install -r requirements.txt -q

# Instalar PyInstaller si no está
if ! python3 -c "import PyInstaller" 2>/dev/null; then
    echo "Instalando PyInstaller..."
    python3 -m pip install pyinstaller
fi

# Limpiar builds anteriores
echo ""
echo "Limpiando builds anteriores..."
rm -rf build dist "Paquete Distribucion" "Generador_Informe_Electrico.zip"

# Construir
echo ""
echo "Construyendo aplicación (esto puede tardar varios minutos)..."
python3 -m PyInstaller GeneradorInforme.spec --noconfirm

if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Falló la construcción."
    exit 1
fi

# Crear carpeta de distribución
echo ""
echo "Creando paquete de distribución..."
mkdir -p "Paquete Distribucion"
if [ -d "dist/Generador Informe Electrico.app" ]; then
    cp -R "dist/Generador Informe Electrico.app" "Paquete Distribucion/"
else
    cp -R dist/*.app "Paquete Distribucion/" 2>/dev/null || cp -R dist/* "Paquete Distribucion/"
fi

if [ -f "plantilla_base.xlsx" ]; then
    cp "plantilla_base.xlsx" "Paquete Distribucion/"
    echo "  [OK] plantilla_base.xlsx"
fi

# Crear README
cat > "Paquete Distribucion/README.txt" << 'EOF'
=====================================================
  GENERADOR DE INFORME ELÉCTRICO - Instrucciones (Mac)
=====================================================

REQUISITOS:
  - macOS 10.14 o superior
  - Conexión a Internet (para descargar datos de PLABACOM)

INSTRUCCIONES:
1. Abra "Generador Informe Electrico.app" (doble clic)
2. Seleccione Año y Mes del informe
3. Complete Empresa, Barra y Nombre Medidor (si aplica)
4. Seleccione la plantilla base del cliente
5. Elija la ruta de destino del informe
6. Haga clic en "Crear Informe"

NOTA: No se requiere Excel instalado. La plantilla se edita directamente.
EOF

# Crear ZIP
echo ""
echo "Creando ZIP..."
cd "Paquete Distribucion"
zip -r ../Generador_Informe_Electrico_Mac.zip . -q
cd ..

if [ -f "Generador_Informe_Electrico_Mac.zip" ]; then
    echo ""
    echo "========================================"
    echo "Paquete creado exitosamente!"
    echo "========================================"
    echo ""
    echo "Archivos:"
    echo "  - Paquete Distribucion/ (carpeta con la app)"
    echo "  - Generador_Informe_Electrico_Mac.zip"
else
    echo "El paquete está en la carpeta 'Paquete Distribucion'"
fi

echo ""
