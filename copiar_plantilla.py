from pathlib import Path
import win32com.client as win32
from datetime import datetime, date

# === CONFIGURACIÓN ===
NOMBRE_ORIGINAL = "plantilla.xlsx"
NOMBRE_COPIA = "plantilla_copia.xlsx"

# Mes/año que quieres actualizar (ej: enero 2026)
ANYO = 2025
MES = 12   # 1 = ene, 2 = feb, ...

# Valor que quieres pegar (de momento de prueba; luego será el total monetario)
TOTAL_MONETARIO = 123456789.0

# ======================

base_dir = Path(__file__).parent
ruta_original = base_dir / NOMBRE_ORIGINAL
ruta_copia = base_dir / NOMBRE_COPIA

if not ruta_original.exists():
    raise FileNotFoundError(f"No se encontró {ruta_original}")

# 1) Copiar plantilla 1:1
ruta_copia.write_bytes(ruta_original.read_bytes())
print(f"Copia creada: {ruta_copia}")

# 2) Abrir copia en Excel y escribir en hoja Resultado
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False

try:
    wb = excel.Workbooks.Open(str(ruta_copia))
    ws = wb.Worksheets("Resultado")

    used_range = ws.UsedRange
    max_row = used_range.Rows.Count
    max_col = used_range.Columns.Count

    # --- 2.1 Buscar columna del mes dinámicamente ---
    meses_abrev = {
        1: "ene", 2: "feb", 3: "mar", 4: "abr", 5: "may", 6: "jun",
        7: "jul", 8: "ago", 9: "sep", 10: "oct", 11: "nov", 12: "dic"
    }
    encabezado_mes = f"{meses_abrev[MES]}-{str(ANYO)[-2:]}"
    print(f"Buscando columna de mes: {encabezado_mes}")

    col_mes = None
    fila_encabezados_max = min(15, max_row)

    for r in range(1, fila_encabezados_max + 1):
        for c in range(1, max_col + 1):
            raw = ws.Cells(r, c).Value
            if raw is None:
                continue

            # Caso 1: la celda es fecha real (por debajo es 01‑MM‑AAAA)
            if isinstance(raw, (datetime, date)):
                if raw.year == ANYO and raw.month == MES:
                    col_mes = c
                    print(f"Columna mes (fecha) encontrada en ({r}, {c}): {raw}")
                    break
            else:
                # Caso 2: texto 'ene-26', 'feb-26', etc.
                valor = str(raw).strip()
                if valor.lower().startswith(encabezado_mes.lower()):
                    col_mes = c
                    print(f"Columna mes (texto) encontrada en ({r}, {c}): {valor}")
                    break

        if col_mes is not None:
            break

    # Si no existe la columna del mes, crearla copiando la última columna de meses
    if col_mes is None:
        print(f"No encontré columna para el mes {encabezado_mes}, voy a crearla copiando la última columna de meses.")

        # Encontrar la fila de encabezados (primera fila no vacía)
        encabezado_row = None
        for r in range(1, fila_encabezados_max + 1):
            if any(ws.Cells(r, c).Value is not None for c in range(1, max_col + 1)):
                encabezado_row = r
                break

        if encabezado_row is None:
            raise RuntimeError("No pude determinar la fila de encabezados.")

        # Última columna con algo en esa fila (último mes existente)
        last_col = max(
            c for c in range(1, max_col + 1)
            if ws.Cells(encabezado_row, c).Value is not None
        )

        base_col = last_col
        new_col = last_col + 1
        print(f"Copiando reglas/formatos de la columna {base_col} a la nueva columna {new_col}")

        # Copiar toda la columna base (incluye fórmulas y formato) y pegar a la derecha
        ws.Columns(base_col).Copy()
        ws.Columns(new_col).Insert(Shift=0)  # 0 = xlShiftToRight

        # Actualizar encabezado de la nueva columna con la fecha del mes nuevo
        header_cell = ws.Cells(encabezado_row, new_col)
        header_cell.Value = datetime(ANYO, MES, 1)
        col_mes = new_col
        print(f"Nueva columna de mes creada en ({encabezado_row}, {new_col})")

    # --- 2.2 Buscar fila del concepto ---
    texto_concepto = "TOTAL INGRESOS POR POTENCIA FIRME CLP"
    fila_concepto = None

    for r in range(1, max_row + 1):
        raw = ws.Cells(r, 2).Value  # columna B (descripciones)
        if raw is None:
            continue
        valor = str(raw).strip().upper()
        if texto_concepto in valor:
            fila_concepto = r
            print(f"Fila concepto encontrada en B{r}: {valor}")
            break

    if fila_concepto is None:
        raise RuntimeError(f"No encontré fila con texto '{texto_concepto}' en la columna B")

    # --- 2.3 Escribir el valor en la celda (mantiene formato/fórmulas alrededor) ---
    print(f"Escribiendo {TOTAL_MONETARIO} en celda ({fila_concepto}, {col_mes})")
    ws.Cells(fila_concepto, col_mes).Value = float(TOTAL_MONETARIO)

    wb.Save()
    wb.Close(SaveChanges=True)
finally:
    excel.Quit()

print("Listo. Revisa 'plantilla_copia.xlsx' en la hoja Resultado.")