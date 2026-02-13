"""
Script de diagnóstico para inspeccionar la estructura del archivo BDef Detalle.
Ejecutar: python debug_bdef.py
"""
import sys
from pathlib import Path

# Añadir el directorio del proyecto
sys.path.insert(0, str(Path(__file__).resolve().parent))

import pandas as pd

def main():
    from core.leer_excel import encontrar_archivo_bdef_detalle

    ruta = encontrar_archivo_bdef_detalle(2025, 12, carpeta_base="bd_data")
    if not ruta or not ruta.exists():
        print("ERROR: No se encontró BDef Detalle. ¿Está descomprimido el ZIP de Potencia?")
        return

    print(f"Archivo: {ruta}")
    print("=" * 60)

    try:
        xl = pd.ExcelFile(ruta, engine="openpyxl")
        print(f"Hojas: {xl.sheet_names}")
        print()

        for nombre_hoja in xl.sheet_names:
            if "balance" in nombre_hoja.lower() or "2" in nombre_hoja:
                print(f"\n--- Hoja: {nombre_hoja} ---")
                df = pd.read_excel(ruta, sheet_name=nombre_hoja, header=None, engine="openpyxl")
                print(f"Filas: {len(df)}, Columnas: {len(df.columns)}")
                print()

                # Buscar Empresa, Concepto, Pago PSUF
                for r in range(min(20, len(df))):
                    row = df.iloc[r]
                    for c, val in enumerate(row):
                        if val is None or (isinstance(val, float) and pd.isna(val)):
                            continue
                        v = str(val).strip().lower()
                        if v in ("empresa", "concepto") or ("pago" in v and "psuf" in v):
                            print(f"  Fila {r+1} Col {c}: '{val}'")

                print("\nPrimeras 5 filas - columnas 16-22 (bloque Empresa/Concepto):")
                for r in range(11, min(16, len(df))):
                    row = df.iloc[r]
                    vals = [str(row.iloc[c])[:18] if c < len(row) else "" for c in range(16, min(23, len(row)))]
                    print(f"  Fila Excel {r+1}: {vals}")

                # Buscar VIENTOS_DE_RENAICO en columna 16 (Empresa)
                print("\nFilas con VIENTOS_DE_RENAICO en col Empresa (16):")
                for r in range(len(df)):
                    row = df.iloc[r]
                    if len(row) > 16:
                        emp = str(row.iloc[16]).strip() if row.iloc[16] is not None else ""
                        if "VIENTOS" in emp.upper() and "RENAICO" in emp.upper():
                            con = str(row.iloc[17])[:15] if len(row) > 17 else ""
                            psuf = row.iloc[19] if len(row) > 19 else ""
                            print(f"  Fila {r+1}: Empresa='{emp}' Concepto='{con}' PagoPSUF(19)={psuf}")
                print("\nFilas que contienen 'VIENTOS' o 'RENAICO' o 'Eólica' (búsqueda amplia):")
                count = 0
                for r in range(len(df)):
                    row = df.iloc[r]
                    row_str = " ".join(str(v) for v in row if v is not None)
                    if "VIENTOS" in row_str.upper() or "RENAICO" in row_str.upper() or "EÓLICA" in row_str.upper() or "EOLICA" in row_str.upper():
                        vals = [str(row.iloc[c])[:25] if c < len(row) else "" for c in range(min(6, len(row)))]
                        print(f"  Fila {r+1}: {vals}")
                        count += 1
                        if count >= 10:
                            print("  ...")
                            break

                break  # Solo primera hoja Balance

    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
