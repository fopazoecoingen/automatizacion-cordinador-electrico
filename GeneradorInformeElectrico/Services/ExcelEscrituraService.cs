using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace GeneradorInformeElectrico.Services;

/// <summary>
/// Escribe valores en la plantilla mediante Excel COM usando late binding (dynamic).
/// No requiere Microsoft.Office.Interop.Excel - evita FileNotFoundException en publish single-file.
/// </summary>
public static class ExcelEscrituraService
{
    private static readonly Dictionary<int, string> MesesAbrev = new()
    {
        { 1, "ene" }, { 2, "feb" }, { 3, "mar" }, { 4, "abr" },
        { 5, "may" }, { 6, "jun" }, { 7, "jul" }, { 8, "ago" },
        { 9, "sep" }, { 10, "oct" }, { 11, "nov" }, { 12, "dic" }
    };

    // Alineado con core/plantilla_cliente.py VARIANTES_CONCEPTOS y EXCLUIR_AL_BUSCAR
    private static readonly Dictionary<string, List<string>> VariantesConceptos = new()
    {
        { "INGRESOS POR IT POTENCIA", new List<string> { "IT POTENCIA", "INGRESOS IT POTENCIA", "02. IT POTENCIA", "INGRESOS POR IT", "IT Potencia", "IT/POTENCIA", "IT-POTENCIA", "POR IT POTENCIA", "ASIGNACION IT POTENCIA" } },
        { "INGRESOS POR POTENCIA", new List<string> { "INGRESOS POTENCIA", "01. INGRESOS POR POTENCIA", "Ingresos por Potencia", "POR POTENCIA" } },
        // Para otras plantillas de clientes (variantes adicionales)
        { "TOTAL INGRESOS POR POTENCIA FIRME CLP", new List<string> { "POTENCIA FIRME CLP", "TOTAL INGRESOS POTENCIA FIRME", "INGRESOS POR POTENCIA FIRME" } },
        { "TOTAL INGRESOS POR ENERGIA CLP", new List<string> { "INGRESOS POR ENERGIA", "TOTAL INGRESOS ENERGIA", "ENERGIA CLP" } },
        { "TOTAL INGRESOS POR SSCC CLP", new List<string> { "INGRESOS POR SSCC", "SSCC CLP", "TOTAL SSCC" } },
        { "Compra Venta Energia GM Holdings CLP", new List<string> { "COMPRA VENTA ENERGIA", "GM HOLDINGS CLP", "GM Holdings" } },
        { "IMPORTACION MWh", new List<string> { "IMPORTACION", "IMPORTACIÓN", "IMPORTACION MWH", "MWh IMPORTACION" } }
    };

    private static readonly Dictionary<string, List<string>> ExcluirAlBuscar = new()
    {
        { "INGRESOS POR POTENCIA", new List<string> { "FIRME", "POR IT" } }
    };

    private static readonly Regex PatronMes = new(@"^(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)[\s\-]*\d{2,4}$", RegexOptions.IgnoreCase);
    private const int xlShiftToRight = -4161;
    private const int xlPasteFormats = -4122;

    public static void EscribirTodosEnResultado(string rutaArchivo, int anyo, int mes,
        List<(string concepto, double valor)> pares)
    {
        if (pares.Count == 0) return;
        GeneradorInformeElectrico.LogHelper.Log("ExcelEscritura: INICIO");
        var rutaDestino = Path.GetFullPath(rutaArchivo);

        var tempDir = Path.Combine(Path.GetTempPath(), "GeneradorInformeElectrico");
        Directory.CreateDirectory(tempDir);
        var tempFile = Path.Combine(tempDir, $"{Path.GetFileNameWithoutExtension(rutaArchivo)}_{Guid.NewGuid():N}{Path.GetExtension(rutaArchivo)}");
        GeneradorInformeElectrico.LogHelper.Log($"ExcelEscritura: Copiando a temp: {tempFile}");
        try { File.Copy(rutaDestino, tempFile, overwrite: true); }
        catch (Exception ex) { throw new InvalidOperationException($"No se pudo copiar el archivo a la ruta temporal. {ex.Message}", ex); }

        object? excelObj = null;
        try
        {
            GeneradorInformeElectrico.LogHelper.Log("Excel: Obteniendo ProgID Excel.Application...");
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new InvalidOperationException("Excel no está instalado. Instale Microsoft Excel.");

            GeneradorInformeElectrico.LogHelper.Log("Excel: Creando instancia...");
            excelObj = Activator.CreateInstance(excelType);
            if (excelObj == null)
                throw new InvalidOperationException("No se pudo crear la instancia de Excel.");

            dynamic excel = excelObj;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;

            GeneradorInformeElectrico.LogHelper.Log($"Excel: Abriendo {tempFile}");
            ExcelComHelper.RetryOnBusy(() =>
            {
                dynamic workbooks = excel.Workbooks;
                dynamic wb = workbooks.Open(tempFile, 0);
                try
                {
                    EscribirEnWorkbook(excel, wb, tempFile, rutaDestino, anyo, mes, pares);
                }
                finally
                {
                    try { wb.Close(false); } catch { }
                }
            });
            try
            {
                File.Copy(tempFile, rutaDestino, overwrite: true);
            }
            catch (IOException io) when (io.Message.Contains("being used") || io.Message.Contains("está siendo utilizado"))
            {
                GeneradorInformeElectrico.LogHelper.Log("AVISO: No se pudo copiar a destino (archivo en uso). Los datos están en temp.");
                throw new InvalidOperationException(
                    $"El archivo '{Path.GetFileName(rutaDestino)}' está abierto. Cierre Excel y ejecute de nuevo, o guarde el resultado desde la carpeta temporal.", io);
            }
            GeneradorInformeElectrico.LogHelper.Log("Excel: Guardado correcto.");
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException rbe)
        {
            GeneradorInformeElectrico.LogHelper.Log($"RuntimeBinderException: {rbe.Message}");
            throw new InvalidOperationException($"Error de Excel: {rbe.Message}", rbe);
        }
        catch (Exception ex)
        {
            GeneradorInformeElectrico.LogHelper.LogExcepcion(ex);
            throw;
        }
        finally
        {
            if (excelObj != null)
            {
                try
                {
                    dynamic ex = excelObj;
                    ex.ScreenUpdating = true;
                    ex.DisplayAlerts = true;
                    ex.Quit();
                }
                catch { }
                try { Marshal.ReleaseComObject(excelObj); } catch { }
            }
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { }
        }
    }

    private static void EscribirEnWorkbook(dynamic excel, dynamic wb, string tempFile, string rutaDestino, int anyo, int mes, List<(string concepto, double valor)> pares)
    {
            dynamic worksheets = wb.Worksheets;
            int count = worksheets.Count;
            dynamic? ws = null;
            for (var i = 1; i <= count; i++)
            {
                dynamic sheet = worksheets.Item(i);
                string name = (string)sheet.Name;
                if (string.Equals(name.Trim(), "Resultado", StringComparison.OrdinalIgnoreCase))
                {
                    ws = sheet;
                    break;
                }
            }

            if (ws == null)
            {
                wb.Close(false);
                throw new InvalidOperationException("No se encontró la hoja 'Resultado' en la plantilla.");
            }

            dynamic usedRange = ws.UsedRange;
            int rowCount = usedRange.Rows.Count;
            int colCount = usedRange.Columns.Count;
            // Plantilla típica: conceptos hasta fila 79+. UsedRange puede subestimar → margen amplio.
            var maxRow = Math.Max(rowCount, 80) + 80;
            var maxCol = Math.Max(colCount, 30);
            var enc2 = $"{MesesAbrev[mes]}-{(anyo % 100):D2}";
            var enc4 = $"{MesesAbrev[mes]}-{anyo}";

            // Encabezados de mes pueden estar en filas 1-20 (plantilla típica: fila 9)
            int colMes = 0;
            for (var r = 1; r <= Math.Min(20, maxRow); r++)
            {
                for (var c = 1; c <= maxCol; c++)
                {
                    var raw = ws.Cells[r, c].Value;
                    if (raw == null) continue;
                    if (raw is DateTime dt && dt.Year == anyo && dt.Month == mes)
                    {
                        colMes = c;
                        break;
                    }
                    var val = (raw.ToString() ?? "").Trim().Replace(" ", "").ToLower();
                    if (val.StartsWith(enc2.ToLower()) || val.StartsWith(enc4.ToLower()))
                    {
                        colMes = c;
                        break;
                    }
                }
                if (colMes > 0) break;
            }

            if (colMes == 0)
            {
                // Igual que plantilla_cliente.py: encontrar fila encabezados, columnas con patrón mes, insertar con formato
                int encabezadoRow = 0;
                for (var r = 1; r <= Math.Min(20, maxRow); r++)
                {
                    for (var c = 1; c <= maxCol; c++)
                    {
                        if (ws.Cells[r, c].Value != null) { encabezadoRow = r; break; }
                    }
                    if (encabezadoRow > 0) break;
                }
                if (encabezadoRow == 0) encabezadoRow = 1;
                var columnasMes = new List<int>();
                for (var c = 1; c <= maxCol; c++)
                {
                    var raw = ws.Cells[encabezadoRow, c].Value;
                    if (raw == null) continue;
                    if (raw is DateTime) { columnasMes.Add(c); continue; }
                    var v = (raw.ToString() ?? "").Trim().Replace(" ", "");
                    if (PatronMes.IsMatch(v)) columnasMes.Add(c);
                }
                int newCol;
                if (columnasMes.Count > 0)
                {
                    int baseCol = columnasMes.Max();
                    newCol = baseCol + 1;
                    var filasCopiar = Math.Min(maxRow, rowCount + 30);
                    ws.Columns[newCol].Insert(xlShiftToRight);
                    var dirOrigen = $"{ColumnaLetra(baseCol)}1:{ColumnaLetra(baseCol)}{filasCopiar}";
                    var dirDest = $"{ColumnaLetra(newCol)}1:{ColumnaLetra(newCol)}{filasCopiar}";
                    dynamic rngOrigen = ws.Range(dirOrigen);
                    dynamic rngDest = ws.Range(dirDest);
                    rngOrigen.Copy();
                    rngDest.PasteSpecial(xlPasteFormats);
                    excel.CutCopyMode = false;
                    ws.Cells[encabezadoRow, newCol].Value = new DateTime(anyo, mes, 1);
                    try { var fmt = ws.Cells[encabezadoRow, baseCol].NumberFormat; if (!string.IsNullOrEmpty((string)fmt)) ws.Cells[encabezadoRow, newCol].NumberFormat = fmt; } catch { }
                }
                else
                {
                    newCol = 2;
                    for (var r = 1; r <= Math.Min(100, maxRow); r++)
                    {
                        for (var cc = 1; cc <= 2; cc++)
                        {
                            var raw = ws.Cells[r, cc].Value;
                            var s = raw?.ToString() ?? "";
                            if (s.ToUpper().Contains("TOTAL INGRESOS"))
                            { newCol = cc + 1; break; }
                        }
                        if (newCol > 2) break;
                    }
                    ws.Columns[newCol].Insert(xlShiftToRight);
                    ws.Cells[encabezadoRow, newCol].Value = new DateTime(anyo, mes, 1);
                    ws.Cells[encabezadoRow, newCol].NumberFormat = "mmm-yy";
                }
                colMes = newCol;
            }

            GeneradorInformeElectrico.LogHelper.Log($"ExcelEscritura: Escribiendo {pares.Count} conceptos en columna mes {colMes}...");
            foreach (var (textoConcepto, valor) in pares)
            {
                GeneradorInformeElectrico.LogHelper.Log($"  -> {textoConcepto}: {valor:N2}");
                var filaConcepto = BuscarFilaConcepto(ws, textoConcepto, maxRow);
                if (filaConcepto <= 0)
                {
                    GeneradorInformeElectrico.LogHelper.Log($"     [SKIP] No encontrado en plantilla.");
                    continue;
                }
                ws.Cells[filaConcepto, colMes].Value = valor;
                GeneradorInformeElectrico.LogHelper.Log($"     [OK] Escrito en fila {filaConcepto}");
            }

            wb.Save();
    }

    private static string ColumnaLetra(int col)
    {
        if (col <= 0) return "A";
        var s = "";
        while (col > 0) { col--; s = (char)('A' + col % 26) + s; col /= 26; }
        return s;
    }

    private static int BuscarFilaConcepto(dynamic ws, string textoConcepto, int maxRow)
    {
        var textosBuscar = new List<string> { textoConcepto };
        if (VariantesConceptos.TryGetValue(textoConcepto, out var variantes))
            textosBuscar.AddRange(variantes);
        var excluir = ExcluirAlBuscar.GetValueOrDefault(textoConcepto, new List<string>());

        foreach (var textoBuscar in textosBuscar)
        {
            var textoUpper = textoBuscar.ToUpper();
            // plantilla_cliente.py: columnas (2,1,3,4) = B,A,C,D. Ampliado para plantillas con layout distinto.
            foreach (var colConcepto in new[] { 2, 1, 3, 4, 5, 6, 7, 8 })
            {
                for (var r = 1; r <= maxRow; r++)
                {
                    var raw = ws.Cells[r, colConcepto].Value;
                    if (raw == null) continue;
                    var val = (raw.ToString() ?? "").Trim().ToUpper();
                    if (!val.Contains(textoUpper)) continue;
                    if (excluir.Any(e => val.Contains(e.ToUpper()))) continue;
                    return r;
                }
            }
        }
        return 0;
    }
}
