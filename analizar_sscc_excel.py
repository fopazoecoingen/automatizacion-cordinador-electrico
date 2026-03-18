"""
Script para analizar la hoja CPI_ del archivo 1_CUADROS_PAGO_SSCC.
Muestra estructura, columnas, valores únicos de Deudor y busca coincidencias.
Uso: python analizar_sscc_excel.py [ruta_archivo_o_carpeta_bd]
"""
import sys
from pathlib import Path

# Añadir core al path
sys.path.insert(0, str(Path(__file__).parent))
try:
    from core.leer_excel import encontrar_archivo_cuadros_pago_sscc
except ImportError:
    encontrar_archivo_cuadros_pago_sscc = None

def analizar(ruta_archivo: Path):
    import pandas as pd
    print(f"\n=== Analizando: {ruta_archivo.name} ===\n")
    kw = {"sheet_name": "CPI_", "header": None}
    if ruta_archivo.suffix.lower() == ".xlsb":
        kw["engine"] = "pyxlsb"
    elif ruta_archivo.suffix.lower() in (".xlsx", ".xlsm"):
        kw["engine"] = "openpyxl"
    
    try:
        df_raw = pd.read_excel(ruta_archivo, **kw)
    except Exception as e:
        print(f"Error leyendo: {e}")
        return
    
    print(f"Filas totales: {len(df_raw)}, Columnas: {len(df_raw.columns)}")
    
    # Buscar fila header
    fila_header = None
    for i in range(min(20, len(df_raw))):
        fila_str = " ".join(str(v) for v in df_raw.iloc[i].values if pd.notna(v)).lower()
        fila_str = fila_str.replace("ó", "o").replace("í", "i")
        if "nemotecnico" in fila_str and "deudor" in fila_str and "monto" in fila_str:
            fila_header = i
            break
    
    if fila_header is None:
        print("No se encontró fila de encabezados con Nemotecnico Deudor y Monto")
        print("\nPrimeras 10 filas (valores):")
        for i in range(min(10, len(df_raw))):
            print(f"  Fila {i}: {list(df_raw.iloc[i].values[:12])}")
        return
    
    print(f"\nFila de encabezados: {fila_header} (0-based)")
    headers = df_raw.iloc[fila_header].values
    print(f"Headers (cols 0-11): {list(headers[:12])}")
    
    idx_deudor = idx_monto = None
    for idx, val in enumerate(headers[:10]):
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        c_lower = str(val).lower().replace("ó", "o").replace("í", "i")
        if "nemotecnico" in c_lower and "deudor" in c_lower:
            idx_deudor = idx
        elif "monto" in c_lower and "retencion" not in c_lower:
            idx_monto = idx
    
    if idx_deudor is None or idx_monto is None:
        print("No se encontraron columnas Nemotecnico Deudor o Monto")
        return
    
    print(f"\nCol Deudor (índice): {idx_deudor}, Col Monto (índice): {idx_monto}")
    
    df = df_raw.iloc[fila_header + 1:].copy()
    col_deudor = df.iloc[:, idx_deudor]
    
    def norm(s):
        return str(s).strip().upper().replace(" ", "_") if s is not None and not (isinstance(s, float) and pd.isna(s)) else ""
    
    deudores = col_deudor.apply(norm)
    deudores_uniq = deudores[deudores != ""].unique().tolist()
    
    print(f"\nValores únicos en col Deudor ({len(deudores_uniq)}):")
    for d in sorted(deudores_uniq)[:50]:
        print(f"  '{d}'")
    if len(deudores_uniq) > 50:
        print(f"  ... y {len(deudores_uniq) - 50} más")
    
    # Buscar variantes de VIENTOS/RENAICO
    print("\n--- Valores que contienen 'VIENTOS' o 'RENAICO' ---")
    encontrados = [d for d in deudores_uniq if "VIENTOS" in d or "RENAICO" in d]
    if encontrados:
        for d in encontrados:
            filas = df[deudores == d]
            total_monto = filas.iloc[:, idx_monto].apply(lambda v: _parse(v) or 0).sum()
            print(f"  '{d}' -> {len(filas)} filas, suma Monto: {total_monto:,.0f}")
    else:
        print("  (ninguno)")
    
    print("\n--- Búsqueda exacta 'VIENTOS_DE_RENAICO' ---")
    mask = deudores == "VIENTOS_DE_RENAICO"
    if mask.any():
        filas = df[mask]
        total = filas.iloc[:, idx_monto].apply(lambda v: _parse(v) or 0).sum()
        print(f"  Encontrado: {len(filas)} filas, total Monto: {total:,.2f}")
    else:
        print("  No encontrado")

def _parse(val):
    import pandas as pd
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    import pandas as pd
    s = str(val).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None

if __name__ == "__main__":
    ruta = None
    if len(sys.argv) > 1:
        ruta = Path(sys.argv[1])
    else:
        carpeta_bd = Path(__file__).parent / "bd_data"
        if encontrar_archivo_cuadros_pago_sscc:
            ruta = encontrar_archivo_cuadros_pago_sscc(2025, 12, str(carpeta_bd))
            if ruta:
                ruta = Path(ruta)
    
    if not ruta or not ruta.exists():
        print("No se encontró archivo SSCC. Uso: python analizar_sscc_excel.py <ruta_archivo>")
        print("Ejemplo: python analizar_sscc_excel.py bd_data/descomprimidos/.../1_CUADROS_PAGO_SSCC_2512_def.xlsm")
        sys.exit(1)
    
    if ruta.is_dir():
        for f in ruta.rglob("*CUADROS*SSCC*.xls*"):
            analizar(f)
            break
    else:
        analizar(ruta)
