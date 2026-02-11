import re
from pathlib import Path
from datetime import datetime, date
from typing import Union

import win32com.client as win32  # type: ignore[import-untyped]

# Excel constant: pegar solo formatos (evita copiar fechas/valores indeseados como 1919)
XL_PASTE_FORMATS = -4122


def escribir_total_en_resultado(
    ruta_archivo: Union[str, Path],
    anyo: int,
    mes: int,
    total_monetario: float,
    texto_concepto: str = "TOTAL INGRESOS POR POTENCIA FIRME CLP",
) -> None:
    """
    Escribe total_monetario en la hoja 'Resultado' de la plantilla del cliente.

    - Busca la columna del mes (encabezado como fecha o como texto 'ene-26', 'feb-26', etc.).
    - Si no existe la columna, crea una nueva copiando la última columna de meses
      (con sus fórmulas y formatos) y ajusta el encabezado.
    - Busca la fila cuyo texto en la columna B contiene `texto_concepto`.
    - Escribe el valor en la celda (fila concepto, columna mes), manteniendo el formato.
    """
    ruta = Path(ruta_archivo)
    if not ruta.exists():
        raise FileNotFoundError(ruta)

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False

    try:
        wb = excel.Workbooks.Open(str(ruta))
        # Buscar hoja Resultado (insensible a mayúsculas)
        ws_resultado = None
        for sh in wb.Worksheets:
            if str(sh.Name).strip().lower() == "resultado":
                ws_resultado = sh
                break
        if ws_resultado is None:
            raise RuntimeError("No se encontró la hoja 'Resultado' en la plantilla.")
        ws = ws_resultado

        used_range = ws.UsedRange
        max_row = max(used_range.Rows.Count, 50)
        # Buscar en más columnas: UsedRange puede ser reducido en plantillas nuevas
        max_col = max(used_range.Columns.Count, 30)

        # --- Buscar columna del mes ---
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
        # Aceptar dic-25 o dic-2025 como formato de búsqueda
        encabezado_mes_2 = f"{meses_abrev[mes]}-{str(anyo)[-2:]}"
        encabezado_mes_4 = f"{meses_abrev[mes]}-{anyo}"

        col_mes = None
        fila_encabezados_max = min(15, max_row)

        for r in range(1, fila_encabezados_max + 1):
            for c in range(1, max_col + 1):
                raw = ws.Cells(r, c).Value
                if raw is None:
                    continue

                # Fecha real (01‑MM‑AAAA debajo del formato)
                if isinstance(raw, (datetime, date)):
                    if raw.year == anyo and raw.month == mes:
                        col_mes = c
                        break
                else:
                    valor = str(raw).strip().replace(" ", "").lower()
                    ref2 = encabezado_mes_2.lower().replace(" ", "")
                    ref4 = encabezado_mes_4.lower().replace(" ", "")
                    if valor.startswith(ref2) or valor.startswith(ref4):
                        col_mes = c
                        break

            if col_mes is not None:
                break

        # Si no existe la columna del mes, crearla (sin duplicar columna B de conceptos)
        if col_mes is None:
            encabezado_row = None
            for r in range(1, fila_encabezados_max + 1):
                if any(ws.Cells(r, c).Value is not None for c in range(1, max_col + 1)):
                    encabezado_row = r
                    break

            if encabezado_row is None:
                raise RuntimeError("No pude determinar la fila de encabezados.")

            # Buscar columnas que parezcan meses (ene-26, dic-2025, fecha, etc.)
            patron_mes = re.compile(
                r"^(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[\s\-]*\d{2,4}$",
                re.I,
            )
            columnas_mes = []
            for c in range(1, max_col + 1):
                raw = ws.Cells(encabezado_row, c).Value
                if raw is None:
                    continue
                if isinstance(raw, (datetime, date)):
                    columnas_mes.append(c)
                elif raw:
                    val_limpio = str(raw).strip().replace(" ", "")
                    if patron_mes.match(val_limpio):
                        columnas_mes.append(c)

            if columnas_mes:
                # Copiar SOLO formatos de la columna base para evitar valores
                # indeseados (fechas 1919, overflow ########)
                base_col = max(columnas_mes)
                new_col = base_col + 1
                ws.Columns(new_col).Insert(Shift=0)
                ws.Columns(base_col).Copy()
                ws.Columns(new_col).PasteSpecial(Paste=XL_PASTE_FORMATS)
                excel.CutCopyMode = False
                header_cell = ws.Cells(encabezado_row, new_col)
                header_cell.Value = datetime(anyo, mes, 1)
                # Mantener formato de la columna copiada (ene-26 -> dic-25)
                try:
                    base_fmt = ws.Cells(encabezado_row, base_col).NumberFormat
                    if base_fmt and str(base_fmt).strip():
                        header_cell.NumberFormat = base_fmt
                except Exception:
                    pass
            else:
                # Plantilla sin columnas de mes: detectar dónde insertar
                # Si conceptos están en A -> insertar en B. Si en B -> insertar en C.
                col_donde_insertar = 2  # Default: B (conceptos en A)
                for r in range(1, min(max_row + 1, 100)):
                    for col_candidate in (1, 2):
                        raw = ws.Cells(r, col_candidate).Value
                        if raw and "TOTAL INGRESOS" in str(raw).upper():
                            col_donde_insertar = col_candidate + 1
                            break
                    else:
                        continue
                    break
                new_col = col_donde_insertar
                ws.Columns(new_col).Insert(Shift=0)
                header_cell = ws.Cells(encabezado_row, new_col)
                header_cell.Value = datetime(anyo, mes, 1)
                header_cell.NumberFormat = "mmm-yy"  # dic-25

            col_mes = new_col

        # --- Buscar fila del concepto (columna A o B según plantilla) ---
        fila_concepto = None
        texto_concepto_upper = texto_concepto.upper()

        for col_concepto in (2, 1):  # Probar B primero, luego A
            for r in range(1, max_row + 1):
                raw = ws.Cells(r, col_concepto).Value
                if raw is None:
                    continue
                valor = str(raw).strip().upper()
                if texto_concepto_upper in valor:
                    fila_concepto = r
                    break
            if fila_concepto is not None:
                break

        if fila_concepto is None:
            raise RuntimeError(
                f"No encontré fila con texto '{texto_concepto}' en columnas A o B de Resultado"
            )

        # --- Escribir el valor en la celda (mantiene formato) ---
        ws.Cells(fila_concepto, col_mes).Value = float(total_monetario)

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()

