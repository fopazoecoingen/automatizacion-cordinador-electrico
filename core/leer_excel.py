"""
Módulo para leer y acceder a datos de archivos Excel de PLABACOM.
"""
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from pathlib import Path
from typing import Optional, Dict, List, Union
from datetime import datetime, date

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


# Mes abreviado para nombre del Anexo Potencia (Ene, Feb, ... Dic)
MESES_ANEXO_POTENCIA = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}


def encontrar_archivo_anexo_potencia(
    anyo: int,
    mes: int,
    carpeta_base: str = "bd_data",
) -> Optional[Path]:
    """
    Encuentra el archivo Anexo 02.b Cuadros de Pago_Potencia_SEN_{Mes}{Year}_Simplificado
    en la carpeta descomprimida del ZIP de Potencia.

    Args:
        anyo: Año del archivo
        mes: Mes del archivo (1-12)
        carpeta_base: Carpeta base (por defecto "bd_data")

    Returns:
        Path del archivo encontrado, None si no existe
    """
    anyo = int(anyo)
    mes = int(mes)
    if mes < 1 or mes > 12:
        return None

    mes_anexo = MESES_ANEXO_POTENCIA.get(mes, "Dic")
    year2 = str(anyo)[-2:]
    # Buscar .xlsb y .xlsx
    nombres_posibles = [
        f"Anexo 02.b Cuadros de Pago_Potencia_SEN_{mes_anexo}{year2}_Simplificado.xlsb",
        f"Anexo 02.b Cuadros de Pago_Potencia_SEN_{mes_anexo}{year2}_Simplificado.xlsx",
    ]

    carpeta_descomprimidos = Path(carpeta_base) / "descomprimidos"
    if not carpeta_descomprimidos.exists():
        return None

    # Buscar en carpetas que contengan "Potencia"
    for carpeta in carpeta_descomprimidos.iterdir():
        if carpeta.is_dir() and "Potencia" in carpeta.name:
            for nombre in nombres_posibles:
                archivo = carpeta / nombre
                if archivo.exists():
                    return archivo
            # Buscar también en subcarpetas
            for archivo in carpeta.rglob("Anexo 02.b*Potencia*Simplificado*"):
                if archivo.suffix.lower() in (".xlsb", ".xlsx"):
                    return archivo

    return None


def encontrar_archivo_cuadros_pago_sscc(
    anyo: int,
    mes: int,
    carpeta_base: str = "bd_data",
) -> Optional[Path]:
    """
    Encuentra EXCEL 1_CUADROS_PAGO_SSCC_{YYMM}_def.xlsx en la carpeta SSCC descomprimida.

    Returns:
        Path del archivo, None si no existe
    """
    anyo = int(anyo)
    mes = int(mes)
    if mes < 1 or mes > 12:
        return None

    yymm = f"{str(anyo)[-2:]}{str(mes).zfill(2)}"
    nombres = [
        f"EXCEL 1_CUADROS_PAGO_SSCC_{yymm}_def.xlsx",
        f"EXCEL 1_CUADROS_PAGO_SSCC_{yymm}_def.xlsb",
    ]

    carpeta_descomprimidos = Path(carpeta_base) / "descomprimidos"
    if not carpeta_descomprimidos.exists():
        return None

    for carpeta in carpeta_descomprimidos.iterdir():
        if carpeta.is_dir() and "SSCC" in carpeta.name:
            for nombre in nombres:
                archivo = carpeta / nombre
                if archivo.exists():
                    return archivo
            for archivo in carpeta.rglob("EXCEL 1_CUADROS_PAGO_SSCC*"):
                if archivo.suffix.lower() in (".xlsx", ".xlsb"):
                    return archivo
    return None


def leer_total_ingresos_sscc(
    anyo: int,
    mes: int,
    nombre_empresa: str = "",
) -> Optional[float]:
    """
    Lee TOTAL INGRESOS POR SSCC CLP desde EXCEL 1_CUADROS_PAGO_SSCC, hoja CPI_.
    Filtra por Nemotecnico Deudor = nombre_empresa y suma columna Monto.

    Returns:
        Valor en CLP, None si no se encuentra
    """
    archivo = encontrar_archivo_cuadros_pago_sscc(anyo, mes)
    if archivo is None:
        print(f"[WARNING] No se encontró EXCEL 1_CUADROS_PAGO_SSCC para {mes}/{anyo}")
        return None

    nombre_empresa_upper = nombre_empresa.strip().upper() if nombre_empresa else ""
    if not nombre_empresa_upper:
        print("[WARNING] nombre_empresa vacío para TOTAL INGRESOS POR SSCC")
        return None

    try:
        kw = {"sheet_name": "CPI_", "header": 0}
        if archivo.suffix.lower() == ".xlsb":
            kw["engine"] = "pyxlsb"
        df = pd.read_excel(archivo, **kw)
    except Exception as e:
        print(f"[WARNING] Error leyendo CPI_: {e}")
        return None

    # Buscar columna Nemotecnico Deudor (puede tener variaciones de nombre)
    col_deudor = None
    col_monto = None
    for c in df.columns:
        c_lower = str(c).lower().replace("ó", "o").replace("í", "i")
        if "nemotecnico" in c_lower and "deudor" in c_lower:
            col_deudor = c
        elif "monto" in c_lower:
            col_monto = c

    if col_deudor is None or col_monto is None:
        print(f"[WARNING] No se encontraron columnas Nemotecnico Deudor o Monto en CPI_")
        return None

    df_filtrado = df[
        df[col_deudor].astype(str).str.strip().str.upper() == nombre_empresa_upper
    ]
    total = df_filtrado[col_monto].apply(
        lambda v: _parsear_valor_monetario(v) or 0
    ).sum()

    print(
        f"[INFO] Leyendo TOTAL INGRESOS POR SSCC CLP desde: {archivo.name} (hoja CPI_)"
    )
    print(f"  -> Dato obtenido ({nombre_empresa}, Monto): {total:,.2f}")
    return float(total)


def leer_compra_venta_energia_gm_holdings(
    anyo: int,
    mes: int,
    nombre_empresa: str = "",
    nombre_barra: str = "",
) -> Optional[float]:
    """
    Lee Compra Venta Energia GM Holdings CLP desde Balance_25XXD, hoja Contratos,
    columna Total CLP. Opcionalmente filtra por empresa/barra si existen esas columnas.

    Returns:
        Valor en CLP, None si no se encuentra
    """
    archivo = encontrar_archivo_balance(anyo, mes)
    if archivo is None:
        print(
            f"[WARNING] No se encontró Balance para leer Compra Venta GM Holdings {mes}/{anyo}"
        )
        return None

    try:
        df = pd.read_excel(archivo, sheet_name="Contratos", header=0)
    except Exception as e:
        print(f"[WARNING] Error leyendo hoja Contratos: {e}")
        return None

    col_total = None
    for c in df.columns:
        c_lower = str(c).strip().lower()
        if "total" in c_lower and "clp" in c_lower:
            col_total = c
            break
    if col_total is None:
        print(f"[WARNING] No se encontró columna Total CLP en Contratos")
        return None

    df_guardar = df.copy()
    if nombre_empresa:
        for c in df.columns:
            c_lower = str(c).lower()
            if "empresa" in c_lower or "nombre_corto" in c_lower or "nemotecnico" in c_lower:
                df_guardar = df_guardar[
                    df_guardar[c].astype(str).str.strip().str.upper()
                    == nombre_empresa.strip().upper()
                ]
                break
    if nombre_barra:
        for c in df.columns:
            if "barra" in str(c).lower():
                df_guardar = df_guardar[
                    df_guardar[c].astype(str).str.strip().str.upper()
                    == nombre_barra.strip().upper()
                ]
                break

    total = df_guardar[col_total].apply(
        lambda v: _parsear_valor_monetario(v) or 0
    ).sum()

    print(
        f"[INFO] Leyendo Compra Venta Energia GM Holdings CLP desde: {archivo.name} (hoja Contratos)"
    )
    print(f"  -> Dato obtenido (Total CLP): {total:,.2f}")
    return float(total)


def leer_valor_concepto_anexo_xlsb(
    ruta_anexo: Path,
    texto_concepto: str,
    nombre_hoja: Optional[str] = None,
) -> Optional[float]:
    """
    Busca 'texto_concepto' en el Anexo y devuelve el valor numérico asociado.
    El valor suele estar en la misma fila, columna adyacente a la derecha.

    Args:
        ruta_anexo: Ruta del archivo .xlsb o .xlsx
        texto_concepto: Texto a buscar (ej: "TOTAL INGRESOS POR POTENCIA FIRME CLP")
        nombre_hoja: Nombre de la hoja donde buscar (ej: "01.BALANCE POTENCIA Dic-25 def").
            Si None, busca en todas las hojas.

    Returns:
        Valor numérico encontrado, None si no se encuentra
    """
    ruta = Path(ruta_anexo)
    if not ruta.exists():
        return None

    texto_upper = texto_concepto.upper()

    if ruta.suffix.lower() == ".xlsb":
        try:
            from pyxlsb import open_workbook
        except ImportError:
            print("[WARNING] pyxlsb no instalado. Ejecute: pip install pyxlsb")
            return None

        with open_workbook(str(ruta)) as wb:
            # Si se especifica hoja, buscar solo ahí; si no, recorrer todas
            if nombre_hoja:
                hojas_a_revisar = [nombre_hoja] if nombre_hoja in wb.sheets else []
            else:
                hojas_a_revisar = list(wb.sheets)
            if nombre_hoja and not hojas_a_revisar:
                # Buscar coincidencia aproximada (insensible a mayúsculas)
                for s in wb.sheets:
                    if nombre_hoja.lower() in str(s).lower():
                        hojas_a_revisar = [s]
                        break

            for sheet_name in hojas_a_revisar:
                if sheet_name not in wb.sheets:
                    continue
                with wb.get_sheet(sheet_name) as sheet:
                    for row in sheet.rows():
                        # Construir dict col -> valor por si las celdas son sparse
                        row_by_col = {}
                        for cell in row:
                            if cell is not None:
                                c = getattr(cell, "c", len(row_by_col))
                                row_by_col[c] = getattr(cell, "v", cell)

                        for col_idx, val in row_by_col.items():
                            valor_str = str(val).strip().upper()
                            if texto_upper in valor_str:
                                # Buscar valor numérico en columnas siguientes de la misma fila
                                cols_orden = sorted(row_by_col.keys())
                                for c in cols_orden:
                                    if c > col_idx:
                                        v = row_by_col[c]
                                        try:
                                            return float(v)
                                        except (TypeError, ValueError):
                                            pass
                                return None
        return None

    # Para .xlsx usar pandas/openpyxl
    try:
        if nombre_hoja:
            df_dict = {nombre_hoja: pd.read_excel(ruta, sheet_name=nombre_hoja, header=None)}
        else:
            df_dict = pd.read_excel(ruta, sheet_name=None, header=None)
    except Exception:
        return None

    for _sheet_name, df in df_dict.items():
        for _, row in df.iterrows():
            row_str = " ".join(str(v).upper() for v in row.dropna().astype(str))
            if texto_upper in row_str:
                for v in row:
                    try:
                        return float(v)
                    except (TypeError, ValueError):
                        continue
                # Buscar primer número en la fila
                for v in row:
                    if isinstance(v, (int, float)) and not pd.isna(v):
                        return float(v)
                    try:
                        return float(v)
                    except (TypeError, ValueError):
                        pass
    return None


def _parsear_valor_monetario(val) -> Optional[float]:
    """Convierte valor con formato 46.709.214 (punto como miles) a float."""
    if val is None:
        return None
    if isinstance(val, (int, float)) and not (isinstance(val, bool)):
        return float(val)
    s = str(val).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def leer_total_ingresos_potencia_firme_anexo(
    ruta_anexo: Path,
    nombre_hoja: str,
    nombre_empresa: str,
) -> Optional[float]:
    """
    Lee el valor TOTAL para una empresa desde la tabla Datos del Anexo Potencia.
    Estructura: col B=Empresa, col C=Potencia SEN, col D=TOTAL.

    Args:
        ruta_anexo: Ruta del archivo .xlsb
        nombre_hoja: Nombre de la hoja (ej: "01.BALANCE POTENCIA Dic-25 def")
        nombre_empresa: Nombre de la empresa (ej: VIENTOS_DE_RENAICO)

    Returns:
        Valor en CLP de la columna TOTAL, None si no se encuentra
    """
    ruta = Path(ruta_anexo)
    if not ruta.exists():
        return None

    nombre_empresa_upper = nombre_empresa.strip().upper()
    if not nombre_empresa_upper:
        return None

    if ruta.suffix.lower() == ".xlsb":
        try:
            from pyxlsb import open_workbook
        except ImportError:
            return None

        with open_workbook(str(ruta)) as wb:
            if nombre_hoja not in wb.sheets:
                for s in wb.sheets:
                    if nombre_hoja.lower() in str(s).lower():
                        nombre_hoja = s
                        break
                else:
                    return None

            with wb.get_sheet(nombre_hoja) as sheet:
                # Col B=Empresa (índice 1), Col D=TOTAL (índice 3)
                # pyxlsb usa columnas 0-based en Cell
                for row in sheet.rows():
                    row_by_col = {}
                    for cell in row:
                        if cell is not None:
                            c = getattr(cell, "c", -1)
                            row_by_col[c] = getattr(cell, "v", cell)

                    # Columna B (índice 1) = Empresa
                    emp_val = row_by_col.get(1) or row_by_col.get(0)
                    if emp_val is None:
                        continue
                    if str(emp_val).strip().upper() != nombre_empresa_upper:
                        continue

                    # Columna D (índice 3) = TOTAL; fallback C (índice 2) = Potencia SEN
                    total_val = row_by_col.get(3) or row_by_col.get(2)
                    if total_val is not None:
                        parsed = _parsear_valor_monetario(total_val)
                        if parsed is not None:
                            return parsed
                return None

    # Fallback xlsx con pandas
    try:
        df = pd.read_excel(ruta, sheet_name=nombre_hoja, header=None)
    except Exception:
        return None
    # Buscar columna Empresa (B=1) y TOTAL (D=3) - pandas usa 0-based
    for _, row in df.iterrows():
        emp_cell = row.iloc[1] if len(row) > 1 else None
        if emp_cell is None:
            continue
        if str(emp_cell).strip().upper() != nombre_empresa_upper:
            continue
        total_cell = row.iloc[3] if len(row) > 3 else row.iloc[2]
        return _parsear_valor_monetario(total_cell)
    return None


def leer_total_ingresos_potencia_firme(
    anyo: int,
    mes: int,
    nombre_empresa: str = "",
) -> Optional[float]:
    """
    Lee el valor TOTAL INGRESOS POR POTENCIA FIRME CLP desde el Anexo 02.b
    de la carpeta Potencia, hoja "01.BALANCE POTENCIA {Mes}-{Year} def".
    Busca la fila donde Empresa = nombre_empresa y devuelve el valor de la columna TOTAL.

    Args:
        anyo: Año
        mes: Mes (1-12)
        nombre_empresa: Nombre de la empresa (ej: VIENTOS_DE_RENAICO). Si vacío, usa fallback.

    Returns:
        Valor en CLP, None si no se encuentra
    """
    archivo = encontrar_archivo_anexo_potencia(anyo, mes)
    if archivo is None:
        print(f"[WARNING] No se encontró Anexo 02.b Potencia para {mes}/{anyo}")
        return None

    mes_anexo = MESES_ANEXO_POTENCIA.get(mes, "Dic")
    year2 = str(anyo)[-2:]
    nombre_hoja = f"01.BALANCE POTENCIA {mes_anexo}-{year2} def"

    if nombre_empresa:
        valor = leer_total_ingresos_potencia_firme_anexo(
            archivo, nombre_hoja, nombre_empresa
        )
        if valor is not None:
            print(
                f"[INFO] Leyendo TOTAL INGRESOS POR POTENCIA FIRME CLP desde: {archivo.name}"
            )
            print(
                f"  -> Dato obtenido ({nombre_empresa}, col TOTAL): {valor:,.2f}"
            )
        return valor

    # Fallback: buscar por texto "TOTAL INGRESOS POR POTENCIA FIRME CLP"
    print(f"[INFO] Leyendo TOTAL INGRESOS POR POTENCIA FIRME CLP desde: {archivo.name}")
    valor = leer_valor_concepto_anexo_xlsb(
        archivo,
        "TOTAL INGRESOS POR POTENCIA FIRME CLP",
        nombre_hoja=nombre_hoja,
    )
    if valor is not None:
        print(f"  -> Dato obtenido (TOTAL INGRESOS POR POTENCIA FIRME CLP): {valor:,.2f}")
    return valor


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

            # Crear nombre de la hoja destino.
            # Para plantillas de cliente se usa siempre la hoja "Resultado".
            nombre_mes = meses[self.mes]
            nombre_hoja = "Resultado"

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

            # Si la plantilla ya tiene una hoja "Resultado", escribir solo los
            # totales en la celda correspondiente al mes, sin destruir el diseño.
            if "Resultado" in wb.sheetnames:
                ws_resultado = wb["Resultado"]
                try:
                    self._escribir_resumen_en_hoja_resultado(
                        ws_resultado,
                        df_guardar,
                        nombre_mes,
                        self.anyo,
                        columna_monetario,
                    )
                    wb.save(ruta_plantilla)
                    wb.close()
                    print("[OK] Datos de resumen escritos en hoja 'Resultado' existente")
                    return True
                except Exception as e:
                    print(
                        "[ERROR] No se pudo escribir en hoja 'Resultado' existente: "
                        f"{e}"
                    )
                    wb.close()
                    return False

            # Si no hay una hoja 'Resultado', usar el comportamiento estándar:
            # crear/actualizar la hoja y volcar la tabla completa.
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
            print(
                "  Guardando "
                f"{len(encabezados)} columnas: "
                f"{', '.join([str(col) for col in encabezados[:15]])}..."
            )
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

    def _escribir_resumen_en_hoja_resultado(
        self,
        ws,
        df_guardar: pd.DataFrame,
        nombre_mes: str,
        anyo: int,
        columna_monetario,
    ) -> None:
        """
        Escribe el total monetario del DataFrame en la hoja 'Resultado' de una
        plantilla existente, en la intersección:
        - Fila del concepto "TOTAL INGRESOS POR POTENCIA FIRME CLP"
        - Columna del mes/año correspondiente (por ejemplo, 'ene-25').
        """
        # Calcular total monetario del DataFrame filtrado/agrupado
        total_monetario = (
            df_guardar[columna_monetario].dropna().astype(float).sum()
        )
        print(f"  Total monetario a escribir en plantilla: {total_monetario:,.2f}")

        # Construir encabezado de mes esperado (ej: 'ene-25')
        meses_abrev = {
            1: "ene",
            2: "feb",
            3: "mar",
            4: "abr",
            5: "may",
            6: "jun",
            7: "jul",
            8: "ago",
            9: "sep",
            10: "oct",
            11: "nov",
            12: "dic",
        }
        encabezado_mes = f"{meses_abrev[self.mes]}-{str(anyo)[-2:]}"
        print(f"  Buscando columna de mes con encabezado: '{encabezado_mes}'")

        # 1) Encontrar columna del mes en las primeras filas (típicamente fila de encabezados)
        col_mes_idx = None
        fila_encabezados_max = 15
        for fila in ws.iter_rows(min_row=1, max_row=fila_encabezados_max):
            for cell in fila:
                raw = cell.value
                # Encabezado como fecha real (p.ej. 01-12-2025)
                if isinstance(raw, (datetime, date)):
                    if raw.year == anyo and raw.month == self.mes:
                        col_mes_idx = cell.column
                        print(f"  Columna de mes (fecha) encontrada en {cell.coordinate}: {raw}")
                        break
                # Encabezado como texto
                else:
                    valor = str(raw).strip()
                    valor_norm = valor.lower()
                    # Coincidencia flexible: empieza con el texto del mes (por si hay sufijos como ' CLP')
                    if valor_norm.startswith(encabezado_mes.lower()):
                        col_mes_idx = cell.column
                        print(f"  Columna de mes (texto) encontrada en {cell.coordinate}: {valor}")
                        break
            if col_mes_idx is not None:
                break

        if col_mes_idx is None:
            # Si no existe la columna, crear una nueva al final de la fila de encabezados
            print(
                f"  No se encontró columna para el mes '{encabezado_mes}'. "
                "Se creará una nueva columna de mes."
            )
            # Buscar la primera fila de encabezados no vacía
            encabezado_row = None
            for fila in ws.iter_rows(min_row=1, max_row=fila_encabezados_max):
                # Considerar fila encabezado si tiene al menos una celda no vacía
                if any(c.value not in (None, "") for c in fila):
                    encabezado_row = fila[0].row
                    break

            if encabezado_row is None:
                raise ValueError(
                    "No se pudo determinar la fila de encabezados para crear la columna de mes"
                )

            # Nueva columna: después de la última columna con datos en esa fila
            last_col = ws.max_column + 1
            header_cell = ws.cell(row=encabezado_row, column=last_col)
            # Usar fecha real para el encabezado; el formato lo maneja la plantilla
            header_cell.value = datetime(anyo, self.mes, 1)
            col_mes_idx = last_col
            print(f"  Columna de mes creada en {header_cell.coordinate} con fecha {header_cell.value}")

        # 2) Encontrar fila del concepto "TOTAL INGRESOS POR POTENCIA FIRME CLP"
        texto_concepto = "TOTAL INGRESOS POR POTENCIA FIRME CLP"
        fila_concepto_idx = None
        fila_busqueda_max = ws.max_row

        for row in ws.iter_rows(min_row=1, max_row=fila_busqueda_max):
            for cell in row:
                valor = str(cell.value).strip().upper() if cell.value is not None else ""
                # Coincidencia flexible: que contenga el texto del concepto
                if texto_concepto in valor:
                    fila_concepto_idx = cell.row
                    print(f"  Fila de concepto encontrada en {cell.coordinate}")
                    break
            if fila_concepto_idx is not None:
                break

        if fila_concepto_idx is None:
            raise ValueError(
                f"No se encontró la fila con el concepto '{texto_concepto}' en la hoja Resultado"
            )

        # 3) Escribir el total monetario en la celda correspondiente
        celda_destino = ws.cell(row=fila_concepto_idx, column=col_mes_idx)
        print(f"  Escribiendo valor en celda {celda_destino.coordinate}")
        celda_destino.value = float(total_monetario)

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

