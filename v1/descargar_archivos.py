import requests
import zipfile
from pathlib import Path
from tqdm import tqdm
from urllib.parse import quote

meses = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
}

def construir_url(anyo, mes, version="01", tipo="Resultados"):
    """
    Construye la URL del archivo ZIP según el año y mes.
    Estructura: PLABACOM/{año}/{mes}_{nombre_mes}/Energia/Definitivo/v_1/{versión} {tipo}_{año_abrev}{mes}_BD01.zip
    
    Args:
        anyo: Año del archivo
        mes: Mes del archivo (1-12)
        version: Versión del archivo (por defecto "01")
        tipo: Tipo de archivo - "Resultados" o "Bases de Datos" (por defecto "Resultados")
    
    Returns:
        tuple: (url_completa, nombre_archivo)
    """
    anyo_abrev = str(anyo)[-2:]
    mes_str = str(mes).zfill(2)  # Asegurar formato de 2 dígitos
    nombre_mes = meses[mes]
    
    # Construir la ruta según la estructura real del S3
    # PLABACOM/2025/12_Diciembre/Energia/Definitivo/v_1/01 Resultados_2512_BD01.zip
    nombre_archivo = f"{version} {tipo}_{anyo_abrev}{mes_str}_BD01.zip"
    
    # Codificar solo el nombre del archivo (donde están los espacios)
    nombre_archivo_codificado = quote(nombre_archivo, safe='')
    
    # Construir la ruta completa sin codificar las barras
    # Usar mes_str (formato 01, 02, etc.) en lugar de mes (1, 2, etc.)
    ruta_s3 = f"PLABACOM/{anyo}/{mes_str}_{nombre_mes}/Energia/Definitivo/v_1/{nombre_archivo_codificado}"
    
    url_base = "https://cen-plabacom.s3.amazonaws.com/"
    url_completa = url_base + ruta_s3
    
    # Nombre del archivo para guardar localmente (sin codificar)
    # PLABACOM_2025_12_Diciembre_Energia_Definitivo_v_1_01 Resultados_2512_BD01.zip
    nombre_local = f"PLABACOM_{anyo}_{mes}_{nombre_mes}_Energia_Definitivo_v_1_{version} {tipo}_{anyo_abrev}{mes_str}_BD01.zip"
    
    return url_completa, nombre_local

def descargar_archivo(url, ruta_destino, mostrar_progreso=True):
    """
    Descarga un archivo desde una URL con barra de progreso
    
    Returns:
        tuple: (exitoso: bool, codigo_error: int o None, mensaje: str)
    """
    try:
        # Realizar petición con stream para archivos grandes
        # Timeout más largo para archivos grandes (>1GB)
        response = requests.get(url, stream=True, timeout=300)
        response.raise_for_status()
        
        # Obtener tamaño total del archivo
        total_size = int(response.headers.get('content-length', 0))
        
        # Crear barra de progreso
        if mostrar_progreso:
            barra = tqdm(total=total_size, unit='B', unit_scale=True, desc="Descargando")
        
        # Descargar archivo en chunks
        with open(ruta_destino, 'wb') as archivo:
            for chunk in response.iter_content(chunk_size=8192):
                if chunk:
                    archivo.write(chunk)
                    if mostrar_progreso:
                        barra.update(len(chunk))
        
        if mostrar_progreso:
            barra.close()
        
        return True, None, "Descarga completada"
    except requests.exceptions.HTTPError as e:
        codigo_error = e.response.status_code if e.response else None
        if codigo_error == 403:
            mensaje = "El contenido no está disponible (403 Forbidden)"
            print(f"Error 403: {mensaje}")
            return False, 403, mensaje
        else:
            mensaje = f"Error HTTP {codigo_error}: {e}"
            print(f"Error al descargar: {mensaje}")
            return False, codigo_error, mensaje
    except requests.exceptions.RequestException as e:
        print(f"Error al descargar: {e}")
        return False, None, str(e)

def buscar_archivo_existente(anyo, mes, carpeta_zip="bd_data"):
    """
    Busca si existe un archivo ZIP para el año y mes especificados,
    sin importar la versión o si dice 'Resultados' o 'Bases de Datos'.
    
    Args:
        anyo: Año del archivo
        mes: Mes del archivo (1-12)
        carpeta_zip: Carpeta donde buscar
    
    Returns:
        Path: Ruta del archivo encontrado, None si no existe
    """
    carpeta = Path(carpeta_zip)
    if not carpeta.exists():
        return None
    
    # Patrón base para buscar: PLABACOM_AÑO_MES_NombreMes_Energia_Definitivo
    patron_base = f"PLABACOM_{anyo}_{mes}_{meses[mes]}_Energia_Definitivo"
    
    # Buscar archivos que coincidan con el patrón
    for archivo in carpeta.glob("*.zip"):
        if patron_base in archivo.name:
            return archivo
    
    return None

def descomprimir_zip(ruta_zip, carpeta_destino=None, nombre_carpeta=None, mostrar_progreso=True):
    """
    Descomprime un archivo ZIP en una carpeta específica.
    
    Args:
        ruta_zip: Ruta del archivo ZIP a descomprimir
        carpeta_destino: Carpeta base donde descomprimir (si None, usa la carpeta del ZIP)
        nombre_carpeta: Nombre de la carpeta dentro de carpeta_destino (si None, extrae del nombre del ZIP)
        mostrar_progreso: Si mostrar barra de progreso
    
    Returns:
        str: Ruta de la carpeta descomprimida, None si hay error
    """
    try:
        ruta_zip = Path(ruta_zip)
        if not ruta_zip.exists():
            print(f"✗ El archivo ZIP no existe: {ruta_zip}")
            return None
        
        # Determinar carpeta destino
        if carpeta_destino is None:
            carpeta_base = ruta_zip.parent
        else:
            carpeta_base = Path(carpeta_destino)
        
        # Determinar nombre de la carpeta
        if nombre_carpeta is None:
            # Extraer la parte final del nombre: "01 Resultados_2512_BD01" del nombre completo
            # El nombre es: PLABACOM_2025_12_Diciembre_Energia_Definitivo_v_1_01 Resultados_2512_BD01.zip
            # Necesitamos: 01 Resultados_2512_BD01
            nombre_completo = ruta_zip.stem
            # Buscar el patrón "_v_1_" y tomar todo lo que viene después
            if "_v_1_" in nombre_completo:
                nombre_carpeta = nombre_completo.split("_v_1_")[-1]
            else:
                # Si no encuentra el patrón, usar el nombre completo sin extensión
                nombre_carpeta = nombre_completo
        
        carpeta_destino = carpeta_base / nombre_carpeta
        
        # Verificar si ya está descomprimido
        if carpeta_destino.exists() and carpeta_destino.is_dir():
            # Verificar si tiene contenido
            contenido = list(carpeta_destino.iterdir())
            if contenido:
                print(f"✓ El archivo ya está descomprimido: {carpeta_destino}")
                return str(carpeta_destino)
        
        # Crear carpeta destino
        carpeta_destino.mkdir(parents=True, exist_ok=True)
        
        # Descomprimir con barra de progreso
        with zipfile.ZipFile(ruta_zip, 'r') as zip_ref:
            # Obtener lista de archivos
            archivos = zip_ref.namelist()
            total_archivos = len(archivos)
            
            if mostrar_progreso:
                barra = tqdm(total=total_archivos, unit='archivos', desc="Descomprimiendo")
            
            # Extraer archivos
            for archivo in archivos:
                zip_ref.extract(archivo, carpeta_destino)
                if mostrar_progreso:
                    barra.update(1)
            
            if mostrar_progreso:
                barra.close()
        
        print(f"✓ Descompresión completada: {carpeta_destino}")
        return str(carpeta_destino)
    
    except zipfile.BadZipFile:
        print(f"✗ Error: El archivo no es un ZIP válido: {ruta_zip}")
        return None
    except Exception as e:
        print(f"✗ Error al descomprimir: {e}")
        return None

def descargar_y_descomprimir_zip(anyo, mes, carpeta_zip="bd_data", carpeta_descomprimidos=None, descomprimir=True):
    """
    Descarga el archivo ZIP si no existe y opcionalmente lo descomprime.
    
    Args:
        anyo: Año del archivo a descargar
        mes: Mes del archivo a descargar (1-12)
        carpeta_zip: Nombre de la carpeta donde guardar los ZIPs
        carpeta_descomprimidos: Carpeta donde descomprimir (si None, usa carpeta_zip/descomprimidos)
        descomprimir: Si descomprimir automáticamente después de descargar
    
    Returns:
        tuple: (ruta_zip: str o None, ruta_descomprimida: str o None, codigo_error: int o None)
    """
    # Descargar el ZIP
    ruta_zip, codigo_error = descargar_zip_si_no_existe(anyo, mes, carpeta_zip)
    
    if not ruta_zip:
        return None, None, codigo_error
    
    # Descomprimir si se solicita
    ruta_descomprimida = None
    if descomprimir:
        if carpeta_descomprimidos is None:
            # Crear carpeta descomprimidos dentro de carpeta_zip
            carpeta_descomprimidos = Path(carpeta_zip) / "descomprimidos"
        
        # Construir el nombre de la carpeta basado en el año y mes
        anyo_abrev = str(anyo)[-2:]
        mes_str = str(mes).zfill(2)
        # Obtener versión y tipo del nombre del archivo
        nombre_zip = Path(ruta_zip).stem
        # Extraer la parte final: "01 Resultados_2512_BD01"
        if "_v_1_" in nombre_zip:
            nombre_carpeta_descomprimida = nombre_zip.split("_v_1_")[-1]
        else:
            # Fallback: construir desde los parámetros
            nombre_carpeta_descomprimida = f"01 Resultados_{anyo_abrev}{mes_str}_BD01"
        
        ruta_descomprimida = descomprimir_zip(ruta_zip, carpeta_descomprimidos, nombre_carpeta_descomprimida)
    
    return ruta_zip, ruta_descomprimida, None

def descargar_zip_si_no_existe(anyo, mes, carpeta_zip="bd_data"):
    """
    Descarga el archivo ZIP si no existe en la carpeta especificada.
    
    Args:
        anyo: Año del archivo a descargar
        mes: Mes del archivo a descargar (1-12)
        carpeta_zip: Nombre de la carpeta donde guardar los ZIPs
    
    Returns:
        tuple: (ruta: str o None, codigo_error: int o None)
    """
    # Crear carpeta si no existe
    carpeta = Path(carpeta_zip)
    carpeta.mkdir(exist_ok=True)
    
    # Primero buscar si existe algún archivo con ese año y mes (sin importar versión)
    archivo_existente = buscar_archivo_existente(anyo, mes, carpeta_zip)
    
    if archivo_existente:
        tamaño = archivo_existente.stat().st_size / (1024 * 1024)  # Tamaño en MB
        print(f"✓ El archivo ya existe: {archivo_existente}")
        print(f"  Tamaño: {tamaño:.2f} MB")
        return str(archivo_existente), None
    
    # Construir URL y nombre de archivo
    url, nombre_archivo = construir_url(anyo, mes)
    ruta_archivo = carpeta / nombre_archivo
    
    # Verificar también el nombre exacto (por si acaso)
    if ruta_archivo.exists():
        tamaño = ruta_archivo.stat().st_size / (1024 * 1024)  # Tamaño en MB
        print(f"✓ El archivo ya existe: {ruta_archivo}")
        print(f"  Tamaño: {tamaño:.2f} MB")
        return str(ruta_archivo), None
    
    # Si no existe, descargarlo
    print(f"Descargando archivo: {nombre_archivo}")
    print(f"URL: {url}")
    
    exito, codigo_error, mensaje = descargar_archivo(url, ruta_archivo)
    
    if exito:
        tamaño = ruta_archivo.stat().st_size / (1024 * 1024)  # Tamaño en MB
        print(f"✓ Descarga completada: {ruta_archivo}")
        print(f"  Tamaño: {tamaño:.2f} MB")
        return str(ruta_archivo), None
    else:
        if codigo_error == 403:
            print(f"✗ Error 403: El contenido no está disponible para este año/mes")
        else:
            print(f"✗ Error en la descarga: {mensaje}")
        return None, codigo_error

# Ejemplo de uso
if __name__ == "__main__":
    anyo = 2025
    mes = 12
    
    # Descargar y descomprimir automáticamente
    ruta_zip, ruta_descomprimida, codigo_error = descargar_y_descomprimir_zip(anyo, mes)
    
    if ruta_zip:
        print(f"\nArchivo ZIP disponible en: {ruta_zip}")
        if ruta_descomprimida:
            print(f"Archivo descomprimido en: {ruta_descomprimida}")
    else:
        if codigo_error == 403:
            print("\nEl contenido no está disponible para este año/mes (403)")
        else:
            print("\nNo se pudo descargar el archivo")
