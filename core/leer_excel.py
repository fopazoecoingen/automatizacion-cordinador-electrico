"""
Módulo para leer y acceder a datos de archivos Excel de PLABACOM.
"""
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
from typing import Optional, Dict, List, Union

from core.descargar_archivos import meses


def encontrar_archivo_balance(anyo: int, mes: int, carpeta_base: str = "bd_data") -> Optional[Path]:
    """
    Encuentra el archivo Balance_XXYYD.xlsm basado en el año y mes.

    Args:
        anyo: Año del archivo
        mes: Mes del archivo (1-12)
        carpeta_base: Carpeta base donde buscar (por defecto "bd_data")

    Returns:
        Path del archivo encontrado, None si no existe
    """
    # Asegurar que anyo y mes sean enteros
    anyo = int(anyo)
    mes = int(mes)

    # Validar rango del mes
    if mes < 1 or mes > 12:
        print(f"[ERROR] Mes inválido: {mes}. Debe estar entre 1 y 12.")
        return None

    anyo_abrev = str(anyo)[-2:]
    mes_str = str(mes).zfill(2)

    # Construir nombre del archivo: Balance_2512D.xlsm (formato: Balance_YYMMD.xlsm)
    nombre_archivo = f"Balance_{anyo_abrev}{mes_str}D.xlsm"

    print(f"[INFO] Buscando archivo Balance para {mes}/{anyo}")
    print(f"  Nombre esperado: {nombre_archivo}")
    print(f"  Año abreviado: {anyo_abrev}, Mes: {mes_str}")

    # Buscar en la carpeta descomprimidos
    carpeta_descomprimidos = Path(carpeta_base) / "descomprimidos"

    if not carpeta_descomprimidos.exists():
        print(f"[ERROR] La carpeta no existe: {carpeta_descomprimidos.absolute()}")
        return None

    # Buscar en todas las carpetas que coincidan con el patrón
    # El patrón es: "01 Resultados_2512_BD01"
    patron_carpeta = f"*Resultados_{anyo_abrev}{mes_str}_BD01"
    print(f"  Patrón de carpeta: {patron_carpeta}")

    carpetas_encontradas = list(carpeta_descomprimidos.glob(patron_carpeta))
    print(f"  Carpetas encontradas con el patrón: {len(carpetas_encontradas)}")

    for carpeta in carpetas_encontradas:
        archivo_balance = carpeta / nombre_archivo
        if archivo_balance.exists():
            print(f"[OK] Archivo Balance encontrado: {archivo_balance}")
            print(f"  Ruta completa: {archivo_balance.absolute()}")
            return archivo_balance
        else:
            # Buscar también en subcarpetas (por si hay estructura anidada)
            for subcarpeta in carpeta.rglob("*"):
                if subcarpeta.is_dir():
                    archivo_balance = subcarpeta / nombre_archivo
                    if archivo_balance.exists():
                        print(f"[OK] Archivo Balance encontrado en subcarpeta: {archivo_balance}")
                        print(f"  Ruta completa: {archivo_balance.absolute()}")
                        return archivo_balance

    # Si no se encuentra, buscar en cualquier subcarpeta (búsqueda más amplia)
    print(f"  Realizando búsqueda amplia en todas las carpetas...")
    for carpeta in carpeta_descomprimidos.iterdir():
        if carpeta.is_dir():
            # Buscar directamente en la carpeta
            archivo_balance = carpeta / nombre_archivo
            if archivo_balance.exists():
                print(f"[OK] Archivo Balance encontrado: {archivo_balance}")
                print(f"  Ruta completa: {archivo_balance.absolute()}")
                return archivo_balance

            # Buscar en subcarpetas
            for archivo in carpeta.rglob(nombre_archivo):
                if archivo.is_file():
                    print(f"[OK] Archivo Balance encontrado: {archivo}")
                    print(f"  Ruta completa: {archivo.absolute()}")
                    return archivo

    # Listar archivos Balance disponibles para ayudar al usuario
    print(f"[ERROR] No se encontró el archivo Balance: {nombre_archivo}")
    print(f"  Buscado en: {carpeta_descomprimidos.absolute()}")

    # Mostrar archivos Balance disponibles
    archivos_balance = list(carpeta_descomprimidos.rglob("Balance_*.xlsm"))
    if archivos_balance:
        print(f"  Archivos Balance disponibles:")
        for archivo in archivos_balance[:5]:  # Mostrar máximo 5
            print(f"    - {archivo.name} (en {archivo.parent.name})")

    return None


def leer_excel_pandas(ruta_archivo: Union[str, Path], hoja: Optional[str] = None, header: Optional[int] = None) -> pd.DataFrame:
    """
    Lee un archivo Excel usando pandas.

    Args:
        ruta_archivo: Ruta del archivo Excel
        hoja: Nombre de la hoja a leer (si None, lee la primera)
        header: Fila a usar como encabezado (si None, usa la primera fila)

    Returns:
        DataFrame con los datos
    """
    ruta_archivo = Path(ruta_archivo)
    if not ruta_archivo.exists():
        raise FileNotFoundError(f"El archivo no existe: {ruta_archivo}")

    if hoja:
        df = pd.read_excel(ruta_archivo, sheet_name=hoja, header=header)
    else:
        df = pd.read_excel(ruta_archivo, header=header)

    return df


def obtener_hojas_excel(ruta_archivo: Union[str, Path]) -> List[str]:
    """
    Obtiene la lista de hojas disponibles en un archivo Excel.

    Args:
        ruta_archivo: Ruta del archivo Excel

    Returns:
        Lista de nombres de hojas
    """
    ruta_archivo = Path(ruta_archivo)
    if not ruta_archivo.exists():
        raise FileNotFoundError(f"El archivo no existe: {ruta_archivo}")

    # Usar openpyxl para obtener las hojas (mejor para .xlsm)
    wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
    hojas = wb.sheetnames
    wb.close()

    return hojas


def leer_celda_excel(ruta_archivo: Union[str, Path], hoja: str, celda: str):
    """
    Lee el valor de una celda específica de un archivo Excel.

    Args:
        ruta_archivo: Ruta del archivo Excel
        hoja: Nombre de la hoja
        celda: Referencia de celda (ej: "A1", "B5")

    Returns:
        Valor de la celda
    """
    ruta_archivo = Path(ruta_archivo)
    if not ruta_archivo.exists():
        raise FileNotFoundError(f"El archivo no existe: {ruta_archivo}")

    wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
    ws = wb[hoja]
    valor = ws[celda].value
    wb.close()

    return valor


def leer_rango_excel(ruta_archivo: Union[str, Path], hoja: str, rango: str) -> pd.DataFrame:
    """
    Lee un rango específico de celdas de un archivo Excel.

    Args:
        ruta_archivo: Ruta del archivo Excel
        hoja: Nombre de la hoja
        rango: Rango de celdas (ej: "A1:C10")

    Returns:
        DataFrame con los datos del rango
    """
    ruta_archivo = Path(ruta_archivo)
    if not ruta_archivo.exists():
        raise FileNotFoundError(f"El archivo no existe: {ruta_archivo}")

    wb = openpyxl.load_workbook(ruta_archivo, read_only=True, data_only=True)
    ws = wb[hoja]

    # Leer el rango
    datos = []
    for fila in ws[rango]:
        datos.append([celda.value for celda in fila])

    wb.close()

    # Convertir a DataFrame
    df = pd.DataFrame(datos)

    return df


class LectorBalance:
    """
    Clase para leer y acceder a datos del archivo Balance de PLABACOM.
    """

    def __init__(self, anyo: int, mes: int, carpeta_base: str = "bd_data"):
        """
        Inicializa el lector para un año y mes específicos.

        Args:
            anyo: Año del archivo
            mes: Mes del archivo (1-12)
            carpeta_base: Carpeta base donde buscar
        """
        self.anyo = anyo
        self.mes = mes
        self.carpeta_base = carpeta_base
        print(f"\nBuscando archivo Balance para {mes}/{anyo}...")
        self.ruta_archivo = encontrar_archivo_balance(anyo, mes, carpeta_base)

        if self.ruta_archivo is None:
            print(f"[ERROR] No se encontró el archivo Balance para {mes}/{anyo} en {carpeta_base}")
            raise FileNotFoundError(f"No se encontró el archivo Balance para {mes}/{anyo} en {carpeta_base}")

        print(f"[OK] LectorBalance inicializado correctamente")
        print(f"  Archivo: {self.ruta_archivo.name}")
        print(f"  Ubicación: {self.ruta_archivo.parent}")

        self._hojas = None
        self._wb = None

    def obtener_hojas(self) -> List[str]:
        """Obtiene la lista de hojas disponibles."""
        if self._hojas is None:
            self._hojas = obtener_hojas_excel(self.ruta_archivo)
        return self._hojas

    def leer_hoja(self, nombre_hoja: str, header: Optional[int] = None) -> pd.DataFrame:
        """
        Lee una hoja completa del archivo Excel.

        Args:
            nombre_hoja: Nombre de la hoja a leer
            header: Fila a usar como encabezado (si None, usa la primera fila)

        Returns:
            DataFrame con los datos de la hoja
        """
        return leer_excel_pandas(self.ruta_archivo, hoja=nombre_hoja, header=header)

    def leer_celda(self, hoja: str, celda: str):
        """
        Lee el valor de una celda específica.

        Args:
            hoja: Nombre de la hoja
            celda: Referencia de celda (ej: "A1", "B5")

        Returns:
            Valor de la celda
        """
        return leer_celda_excel(self.ruta_archivo, hoja, celda)

    def leer_rango(self, hoja: str, rango: str) -> pd.DataFrame:
        """
        Lee un rango específico de celdas.

        Args:
            hoja: Nombre de la hoja
            rango: Rango de celdas (ej: "A1:C10")

        Returns:
            DataFrame con los datos del rango
        """
        return leer_rango_excel(self.ruta_archivo, hoja, rango)

    def _detectar_fila_encabezados(self, nombre_hoja: str, max_filas: int = 20) -> Optional[int]:
        """
        Detecta automáticamente la fila que contiene los encabezados 'barra' y 'monetario'.

        Args:
            nombre_hoja: Nombre de la hoja a analizar
            max_filas: Número máximo de filas a revisar

        Returns:
            Número de fila que contiene los encabezados, None si no se encuentra
        """
        # Leer las primeras filas sin encabezados para buscar manualmente
        df_temp = leer_excel_pandas(self.ruta_archivo, hoja=nombre_hoja, header=None)

        # Buscar en las primeras filas
        for fila_idx in range(min(max_filas, len(df_temp))):
            fila = df_temp.iloc[fila_idx]
            # Convertir la fila a strings y buscar 'barra' y 'monetario' (insensible a mayúsculas)
            valores_fila = [str(val).lower() if val is not None else "" for val in fila.values]

            tiene_barra = any("barra" in str(val).lower() for val in valores_fila)
            tiene_monetario = any("monetario" in str(val).lower() for val in valores_fila)

            if tiene_barra and tiene_monetario:
                print(f"[OK] Encabezados detectados en la fila {fila_idx}")
                return fila_idx

        return None

    def leer_balance_valorizado(self, header: Optional[int] = None) -> pd.DataFrame:
        """
        Lee la hoja "Balance Valorizado" del archivo Balance.
        Si header es None, detecta automáticamente la fila de encabezados.

        Args:
            header: Fila a usar como encabezado (si None, detecta automáticamente)

        Returns:
            DataFrame con los datos de la hoja "Balance Valorizado"
        """
        nombre_hoja = "Balance Valorizado"
        print(f"\nLeyendo hoja: {nombre_hoja}")

        # Verificar que la hoja existe (búsqueda insensible a mayúsculas)
        hojas_disponibles = self.obtener_hojas()

        # Buscar la hoja (insensible a mayúsculas)
        hoja_encontrada = None
        for hoja in hojas_disponibles:
            if hoja.lower() == nombre_hoja.lower():
                hoja_encontrada = hoja
                break

        if hoja_encontrada is None:
            print(f"[ERROR] Advertencia: La hoja '{nombre_hoja}' no se encontró en el archivo")
            print(f"  Hojas disponibles: {', '.join(hojas_disponibles[:10])}...")
            raise ValueError(f"La hoja '{nombre_hoja}' no existe en el archivo")

        # Usar el nombre exacto de la hoja encontrada
        nombre_hoja = hoja_encontrada

        print(f"[OK] Hoja '{nombre_hoja}' encontrada")

        # Si no se especifica header, detectarlo automáticamente
        if header is None:
            print("  Detectando automáticamente la fila de encabezados...")
            header = self._detectar_fila_encabezados(nombre_hoja)
            if header is None:
                print("[WARNING] No se pudo detectar automáticamente la fila de encabezados, usando fila 0")
                header = 0
            else:
                print(f"  Usando fila {header} como encabezados")

        df = self.leer_hoja(nombre_hoja, header=header)
        print(f"  Filas leídas: {len(df)}")
        print(f"  Columnas: {len(df.columns)}")

        # Mostrar las primeras columnas para verificación
        print(f"  Primeras columnas: {', '.join([str(col) for col in df.columns[:5]])}...")

        # Verificar que tiene las columnas necesarias
        columnas_lower = [str(col).lower() for col in df.columns]
        if "barra" not in columnas_lower:
            print("[WARNING] No se encontró la columna 'barra' en los encabezados")
            print(f"  Columnas disponibles: {', '.join([str(col) for col in df.columns[:10]])}...")
        if "monetario" not in columnas_lower:
            print("[WARNING] No se encontró la columna 'monetario' en los encabezados")
            print(f"  Columnas disponibles: {', '.join([str(col) for col in df.columns[:10]])}...")

        return df

    def obtener_columna(self, df: pd.DataFrame, nombre_columna: str, mostrar_todos: bool = True) -> pd.Series:
        """
        Obtiene una columna específica del DataFrame.

        Args:
            df: DataFrame del cual extraer la columna
            nombre_columna: Nombre de la columna a obtener (búsqueda insensible a mayúsculas)
            mostrar_todos: Si mostrar todos los elementos (True) o solo un resumen (False)

        Returns:
            Serie con los valores de la columna
        """
        # Buscar la columna (insensible a mayúsculas)
        columna_encontrada = None
        for col in df.columns:
            if str(col).lower() == nombre_columna.lower():
                columna_encontrada = col
                break

        if columna_encontrada is None:
            print(f"[ERROR] La columna '{nombre_columna}' no se encontró")
            print(f"  Columnas disponibles: {', '.join([str(c) for c in df.columns[:10]])}...")
            raise ValueError(f"La columna '{nombre_columna}' no existe en el DataFrame")

        print(f"\n[OK] Columna '{columna_encontrada}' encontrada")
        serie = df[columna_encontrada]

        if mostrar_todos:
            print(f"\nTodos los elementos de la columna '{columna_encontrada}':")
            print("=" * 60)
            for idx, valor in enumerate(serie, start=1):
                print(f"{idx:4d}. {valor}")
            print("=" * 60)
            print(f"\nTotal de elementos: {len(serie)}")
            print(f"Valores no nulos: {serie.notna().sum()}")
            print(f"Valores nulos: {serie.isna().sum()}")
        else:
            print(f"\nResumen de la columna '{columna_encontrada}':")
            print(serie.describe())

        return serie

    def buscar_por_barra(self, df: pd.DataFrame, nombre_barra: str) -> pd.DataFrame:
        """
        Busca todos los registros que corresponden a una barra específica y muestra los valores monetarios.

        Args:
            df: DataFrame con los datos del Balance Valorizado
            nombre_barra: Nombre de la barra a buscar (búsqueda insensible a mayúsculas)

        Returns:
            DataFrame filtrado con los registros de esa barra
        """
        # Buscar la columna "barra" (insensible a mayúsculas)
        columna_barra = None
        for col in df.columns:
            if str(col).lower() == "barra":
                columna_barra = col
                break

        if columna_barra is None:
            print("[ERROR] La columna 'barra' no se encontró en el DataFrame")
            raise ValueError("La columna 'barra' no existe en el DataFrame")

        # Filtrar por el nombre de la barra (insensible a mayúsculas)
        df_filtrado = df[df[columna_barra].astype(str).str.lower() == nombre_barra.lower()]

        print(f"\n[OK] Búsqueda de barra: '{nombre_barra}'")
        print(f"  Registros encontrados: {len(df_filtrado)}")

        if len(df_filtrado) == 0:
            print(f"[ERROR] No se encontraron registros para la barra '{nombre_barra}'")
            # Mostrar algunas barras disponibles como sugerencia
            barras_unicas = df[columna_barra].dropna().unique()[:10]
            print(f"  Algunas barras disponibles: {', '.join([str(b) for b in barras_unicas])}...")
            return df_filtrado

        # Buscar la columna "monetario"
        columna_monetario = None
        for col in df.columns:
            if str(col).lower() == "monetario":
                columna_monetario = col
                break

        if columna_monetario:
            print(f"\nValores monetarios para la barra '{nombre_barra}':")
            print("=" * 60)
            valores_monetarios = df_filtrado[columna_monetario]
            for idx, (index_row, valor) in enumerate(valores_monetarios.items(), start=1):
                print(f"{idx:4d}. {valor}")
            print("=" * 60)

            # Estadísticas
            valores_no_nulos = valores_monetarios.dropna()
            if len(valores_no_nulos) > 0:
                suma_total = valores_no_nulos.sum()
                print("\nResumen:")
                print(f"  Total de registros: {len(valores_monetarios)}")
                print(f"  Valores no nulos: {len(valores_no_nulos)}")
                print(f"  Suma total: {suma_total:,.2f}")
                print(f"  Promedio: {valores_no_nulos.mean():,.2f}")
                print(f"  Mínimo: {valores_no_nulos.min():,.2f}")
                print(f"  Máximo: {valores_no_nulos.max():,.2f}")
        else:
            print("[ERROR] La columna 'monetario' no se encontró")

        return df_filtrado

    def guardar_en_plantilla(
        self,
        df_balance: pd.DataFrame,
        ruta_plantilla: str = "plantilla_base.xlsx",
        nombre_barra: Optional[str] = None,
        nombre_empresa: Optional[str] = None,
    ) -> bool:
        """
        Guarda los datos del Balance Valorizado en la plantilla Excel.
        Si se especifica una barra o empresa, guarda solo los datos filtrados (TODAS las columnas).
        Si no se especifica, agrupa por barra y suma los valores monetarios.

        Args:
            df_balance: DataFrame con los datos del Balance Valorizado
            ruta_plantilla: Ruta del archivo plantilla Excel
            nombre_barra: Nombre de la barra específica (opcional)
            nombre_empresa: Nombre de la empresa (busca en columna nombre_corto_empresa) (opcional)

        Returns:
            True si se guardó correctamente, False en caso contrario
        """
        try:
            ruta_plantilla = Path(ruta_plantilla)

            # Crear nombre de la hoja: "mes-año" (ej: "Diciembre-2025")
            nombre_mes = meses[self.mes]
            nombre_hoja = f"{nombre_mes}-{self.anyo}"

            print(f"\n[OK] Guardando datos en plantilla: {ruta_plantilla}")
            print(f"  Hoja: {nombre_hoja}")

            # Buscar columnas necesarias
            columna_barra = None
            columna_monetario = None
            columna_empresa = None

            for col in df_balance.columns:
                col_lower = str(col).lower()
                if col_lower == "barra":
                    columna_barra = col
                elif col_lower == "monetario":
                    columna_monetario = col
                elif col_lower == "nombre_corto_empresa" or col_lower == "nombre corto empresa":
                    columna_empresa = col

            if not columna_barra or not columna_monetario:
                print("[ERROR] No se encontraron las columnas 'barra' o 'monetario'")
                return False

            # Preparar datos
            if nombre_barra or nombre_empresa:
                # Filtrar por barra y/o empresa - guardar TODAS las columnas de las filas filtradas
                df_guardar = df_balance.copy()

                # Aplicar filtro de empresa si se especifica
                if nombre_empresa:
                    if columna_empresa is None:
                        print("[ERROR] No se encontró la columna 'nombre_corto_empresa'")
                        print(
                            f"  Columnas disponibles: {', '.join([str(c) for c in df_balance.columns[:20]])}..."
                        )
                        return False

                    df_guardar = df_guardar[
                        df_guardar[columna_empresa].astype(str).str.lower() == nombre_empresa.lower()
                    ]
                    print(f"  Filtrando por empresa: {nombre_empresa} (columna: {columna_empresa})")
                    print(f"  Filas después de filtrar por empresa: {len(df_guardar)}")

                # Aplicar filtro de barra si se especifica
                if nombre_barra:
                    df_guardar = df_guardar[
                        df_guardar[columna_barra].astype(str).str.lower() == nombre_barra.lower()
                    ]
                    print(f"  Filtrando por barra: {nombre_barra}")

                print(f"  Filas encontradas después de todos los filtros: {len(df_guardar)}")
                print(f"  Columnas a guardar: {len(df_guardar.columns)}")
                print(f"  Columnas: {', '.join([str(col) for col in df_guardar.columns[:15]])}...")

                if len(df_guardar) == 0:
                    print("[WARNING] No se encontraron registros con los filtros especificados")
                    print(f"  Empresa: {nombre_empresa if nombre_empresa else 'Todas'}")
                    print(f"  Barra: {nombre_barra if nombre_barra else 'Todas'}")
            else:
                # Agrupar por barra y sumar valores monetarios
                df_guardar = df_balance.groupby(columna_barra)[columna_monetario].sum().reset_index()
                df_guardar.columns = ["Barra", "Monetario"]
                print("  Agrupando por barra y sumando valores monetarios")

            # Cargar o crear el archivo Excel
            if ruta_plantilla.exists():
                wb = load_workbook(ruta_plantilla)
            else:
                wb = openpyxl.Workbook()
                # Eliminar la hoja por defecto si existe
                if "Sheet" in wb.sheetnames:
                    wb.remove(wb["Sheet"])

            # Crear o sobrescribir la hoja
            if nombre_hoja in wb.sheetnames:
                ws = wb[nombre_hoja]
                # Limpiar la hoja existente
                ws.delete_rows(1, ws.max_row)
                print(f"  Hoja '{nombre_hoja}' ya existe, actualizando...")
            else:
                ws = wb.create_sheet(nombre_hoja)
                print(f"  Creando nueva hoja '{nombre_hoja}'")

            # Escribir encabezados
            encabezados = list(df_guardar.columns)
            print(f"  Guardando {len(encabezados)} columnas: {', '.join([str(col) for col in encabezados[:15]])}...")
            ws.append(encabezados)

            # Formatear encabezados
            from openpyxl.styles import Font, PatternFill

            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")

            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font

            # Escribir datos - asegurar que se escriban todas las filas y columnas
            print(f"  Escribiendo {len(df_guardar)} filas con todas las columnas...")
            for idx, r in enumerate(dataframe_to_rows(df_guardar, index=False, header=False), 1):
                ws.append(r)
                if idx % 100 == 0:
                    print(f"    Procesadas {idx} filas...")

            # Ajustar ancho de columnas
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except Exception:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[column_letter].width = adjusted_width

            # Guardar el archivo
            wb.save(ruta_plantilla)
            wb.close()

            print("[OK] Datos guardados exitosamente")
            print(f"  Total de filas escritas: {len(df_guardar)}")
            print(f"  Total de columnas escritas: {len(df_guardar.columns)}")
            if nombre_barra:
                print("  Todas las columnas de las filas filtradas fueron guardadas")
            print(f"  Archivo: {ruta_plantilla.absolute()}")

            # Verificar que se guardaron todas las columnas
            if (nombre_barra or nombre_empresa) and len(df_guardar.columns) < len(df_balance.columns):
                print(
                    f"[WARNING] Se guardaron {len(df_guardar.columns)} columnas, "
                    f"pero el DataFrame original tiene {len(df_balance.columns)}"
                )
            elif nombre_barra or nombre_empresa:
                print(f"[OK] Se guardaron todas las {len(df_guardar.columns)} columnas del DataFrame original")

            return True

        except Exception as e:
            print(f"[ERROR] Error al guardar en plantilla: {e}")
            import traceback

            traceback.print_exc()
            return False

    def __repr__(self):
        return f"LectorBalance(anyo={self.anyo}, mes={self.mes}, archivo={self.ruta_archivo.name})"


if __name__ == "__main__":
    # Ejemplo de uso básico
    try:
        lector = LectorBalance(2025, 12)
        print(f"Archivo encontrado: {lector.ruta_archivo}")
        print("\nHojas disponibles:")
        for hoja in lector.obtener_hojas():
            print(f"  - {hoja}")

        print("\n" + "=" * 50)
        df_balance = lector.leer_balance_valorizado()
        print("\nPrimeras filas del Balance Valorizado:")
        print(df_balance.head(10))
    except FileNotFoundError as e:
        print(f"Error: {e}")

