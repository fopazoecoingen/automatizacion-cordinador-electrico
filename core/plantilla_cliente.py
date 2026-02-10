from pathlib import Path
from datetime import datetime, date
from typing import Union

import win32com.client as win32  # type: ignore[import-untyped]


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
        ws = wb.Worksheets("Resultado")

        used_range = ws.UsedRange
        max_row = used_range.Rows.Count
        max_col = used_range.Columns.Count

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
        encabezado_mes = f"{meses_abrev[mes]}-{str(anyo)[-2:]}"

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
                    valor = str(raw).strip()
                    if valor.lower().startswith(encabezado_mes.lower()):
                        col_mes = c
                        break

            if col_mes is not None:
                break

        # Si no existe la columna del mes, crearla copiando la última columna de meses
        if col_mes is None:
            encabezado_row = None
            for r in range(1, fila_encabezados_max + 1):
                if any(ws.Cells(r, c).Value is not None for c in range(1, max_col + 1)):
                    encabezado_row = r
                    break

            if encabezado_row is None:
                raise RuntimeError("No pude determinar la fila de encabezados.")

            last_col = max(
                c for c in range(1, max_col + 1)
                if ws.Cells(encabezado_row, c).Value is not None
            )

            base_col = last_col
            new_col = last_col + 1

            ws.Columns(base_col).Copy()
            ws.Columns(new_col).Insert(Shift=0)  # desplazar a la derecha

            header_cell = ws.Cells(encabezado_row, new_col)
            header_cell.Value = datetime(anyo, mes, 1)
            col_mes = new_col

        # --- Buscar fila del concepto en la columna B ---
        fila_concepto = None
        texto_concepto_upper = texto_concepto.upper()

        for r in range(1, max_row + 1):
            raw = ws.Cells(r, 2).Value  # columna B
            if raw is None:
                continue
            valor = str(raw).strip().upper()
            if texto_concepto_upper in valor:
                fila_concepto = r
                break

        if fila_concepto is None:
            raise RuntimeError(
                f"No encontré fila con texto '{texto_concepto}' en la columna B de Resultado"
            )

        # --- Escribir el valor en la celda (mantiene formato) ---
        ws.Cells(fila_concepto, col_mes).Value = float(total_monetario)

        wb.Save()
        wb.Close(SaveChanges=True)
    finally:
        excel.Quit()

