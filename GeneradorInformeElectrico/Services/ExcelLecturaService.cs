#nullable disable
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

namespace GeneradorInformeElectrico.Services;

/// <summary>
/// Servicio de lectura de archivos Excel PLABACOM.
/// Portado desde core/leer_excel.py
/// </summary>
public static class ExcelLecturaService
{
    private static readonly Dictionary<int, string> MesesAbrev = new()
    {
        { 1, "ene" }, { 2, "feb" }, { 3, "mar" }, { 4, "abr" },
        { 5, "may" }, { 6, "jun" }, { 7, "jul" }, { 8, "ago" },
        { 9, "sep" }, { 10, "oct" }, { 11, "nov" }, { 12, "dic" }
    };

    public static string? EncontrarArchivoBalance(int anyo, int mes, string carpetaBase)
    {
        var anyoAbrev = (anyo % 100).ToString("D2");
        var mesStr = mes.ToString("D2");
        var nombreBase = $"Balance_{anyoAbrev}{mesStr}D";
        var patronCarpeta = $"*Resultados_{anyoAbrev}{mesStr}_BD01";
        var carpetaDescomprimidos = Path.Combine(carpetaBase, "descomprimidos");
        if (!Directory.Exists(carpetaDescomprimidos)) return null;
        foreach (var dir in Directory.GetDirectories(carpetaDescomprimidos, patronCarpeta))
        {
            foreach (var ext in new[] { ".xlsm", ".xlsx" })
            {
                var archivo = Path.Combine(dir, nombreBase + ext);
                if (File.Exists(archivo)) return archivo;
            }
            foreach (var sub in Directory.GetDirectories(dir, "*", SearchOption.AllDirectories))
            {
                foreach (var ext in new[] { ".xlsm", ".xlsx" })
                {
                    var archivo = Path.Combine(sub, nombreBase + ext);
                    if (File.Exists(archivo)) return archivo;
                }
            }
        }
        return null;
    }

    public static string? EncontrarArchivoBdefDetalle(int anyo, int mes, string carpetaBase)
    {
        var mesAbrev = MesesAbrev[mes];
        var yymm = (anyo % 100).ToString("D2") + mes.ToString("D2");
        var carpetaDescomprimidos = Path.Combine(carpetaBase, "descomprimidos");
        if (!Directory.Exists(carpetaDescomprimidos)) return null;
        bool MatchBdef(string path)
        {
            var name = Path.GetFileName(path).ToUpperInvariant();
            return (name.Contains("BDEF") && name.Contains("DETALLE") || name.Contains("BDEFDETALLE")) &&
                (name.Contains(mesAbrev.ToUpperInvariant()) || name.Contains((anyo % 100).ToString()));
        }
        foreach (var dir in Directory.GetDirectories(carpetaDescomprimidos))
        {
            if (!dir.Contains("Potencia", StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(dir).Contains(yymm))
                continue;
            var carpetaDetalle = Path.Combine(dir, "07. Detalle por empresa");
            if (!Directory.Exists(carpetaDetalle))
            {
                carpetaDetalle = Path.Combine(dir, "07._Detalle_por_empresa");
                if (!Directory.Exists(carpetaDetalle)) continue;
            }
            foreach (var ext in new[] { "*.xlsx", "*.xlsb", "*.xlsm" })
            {
                foreach (var f in Directory.GetFiles(carpetaDetalle, ext))
                    if (MatchBdef(f)) return f;
            }
        }
        var carpetaDetalleFallback = Path.Combine(carpetaDescomprimidos, "07. Detalle por empresa");
        if (Directory.Exists(carpetaDetalleFallback))
        {
            foreach (var ext in new[] { "*.xlsx", "*.xlsb", "*.xlsm" })
            {
                foreach (var f in Directory.GetFiles(carpetaDetalleFallback, ext))
                    if (MatchBdef(f)) return f;
            }
        }
        foreach (var d in Directory.GetDirectories(carpetaDescomprimidos))
        {
            var sub = Path.Combine(d, "07. Detalle por empresa");
            if (!Directory.Exists(sub)) continue;
            foreach (var ext in new[] { "*.xlsx", "*.xlsb", "*.xlsm" })
            {
                foreach (var f in Directory.GetFiles(sub, ext))
                    if (MatchBdef(f)) return f;
            }
        }
        return null;
    }

    public static string? EncontrarAnexo02b(int anyo, int mes, string carpetaBase)
    {
        var mesNombre = new Dictionary<int, string> {
            { 1, "Ene" }, { 2, "Feb" }, { 3, "Mar" }, { 4, "Abr" },
            { 5, "May" }, { 6, "Jun" }, { 7, "Jul" }, { 8, "Ago" },
            { 9, "Sep" }, { 10, "Oct" }, { 11, "Nov" }, { 12, "Dic" }
        }[mes];
        var anyo2 = (anyo % 100).ToString("D2");
        var periodoEnNombre = $"{mesNombre}{anyo2}".ToLowerInvariant(); // ej: dic25
        var yymm = anyo2 + mes.ToString("D2"); // 2512 para dic 2025
        var carpetaDescomprimidos = Path.Combine(carpetaBase, "descomprimidos");
        if (!Directory.Exists(carpetaDescomprimidos)) return null;
        bool Match(string path)
        {
            var name = Path.GetFileName(path).ToLowerInvariant();
            return name.Contains("02.b") && name.Contains("potencia") && name.Contains("simplificado") &&
                (name.Contains(periodoEnNombre) || (name.Contains(mesNombre.ToLowerInvariant()) && (name.Contains(anyo2) || name.Contains(anyo.ToString()))));
        }
        // Python: buscar primero en carpetas Potencia que contengan yymm (ej: 2512)
        foreach (var dir in Directory.GetDirectories(carpetaDescomprimidos))
        {
            if (!dir.Contains("Potencia", StringComparison.OrdinalIgnoreCase) || !Path.GetFileName(dir).Contains(yymm))
                continue;
            foreach (var f in Directory.GetFiles(dir, "*.xlsx", SearchOption.AllDirectories))
                if (Match(f)) return f;
            foreach (var f in Directory.GetFiles(dir, "*.xlsb", SearchOption.AllDirectories))
                if (Match(f)) return f;
        }
        // Fallback: buscar en todo descomprimidos
        foreach (var f in Directory.GetFiles(carpetaDescomprimidos, "*.xlsx", SearchOption.AllDirectories))
            if (Match(f)) return f;
        foreach (var f in Directory.GetFiles(carpetaDescomprimidos, "*.xlsb", SearchOption.AllDirectories))
            if (Match(f)) return f;
        return null;
    }

    public static string? EncontrarCuadrosPagoSscc(int anyo, int mes, string carpetaBase)
    {
        var yymm = (anyo % 100).ToString("D2") + mes.ToString("D2");
        var mesAbrev = MesesAbrev[mes];
        var anyo2 = (anyo % 100).ToString("D2");
        var periodoAlt = $"{mesAbrev}{anyo2}"; // ej: dic25
        var carpetaDescomprimidos = Path.Combine(carpetaBase, "descomprimidos");
        bool Match(string path)
        {
            var name = Path.GetFileName(path).ToUpperInvariant();
            var tieneSscc = (name.Contains("PAGO") && name.Contains("SSCC")) || (name.Contains("CUADROS") && name.Contains("SSCC"));
            var tienePeriodo = name.Contains(yymm) || name.Contains(mesAbrev.ToUpperInvariant()) || name.Contains(periodoAlt.ToUpperInvariant());
            return tieneSscc && tienePeriodo;
        }
        foreach (var ext in new[] { "*.xlsx", "*.xlsm", "*.xlsb" })
        {
            if (Directory.Exists(carpetaBase))
                foreach (var f in Directory.GetFiles(carpetaBase, ext, SearchOption.AllDirectories))
                    if (Match(f)) return f;
        }
        if (Directory.Exists(carpetaDescomprimidos))
        {
            foreach (var f in Directory.GetFiles(carpetaDescomprimidos, "*.xlsx", SearchOption.AllDirectories))
                if (Match(f)) return f;
            foreach (var f in Directory.GetFiles(carpetaDescomprimidos, "*.xlsm", SearchOption.AllDirectories))
                if (Match(f)) return f;
            foreach (var f in Directory.GetFiles(carpetaDescomprimidos, "*.xlsb", SearchOption.AllDirectories))
                if (Match(f)) return f;
        }
        var carpetaSscc = Path.Combine(carpetaBase, "sscc");
        if (Directory.Exists(carpetaSscc))
        {
            foreach (var f in Directory.GetFiles(carpetaSscc, "*.xlsx", SearchOption.AllDirectories))
                if (Match(f)) return f;
            foreach (var f in Directory.GetFiles(carpetaSscc, "*.xlsb", SearchOption.AllDirectories))
                if (Match(f)) return f;
        }
        return null;
    }

    /// <summary>Lee Balance Valorizado y retorna datos para total monetario, filtros, etc.</summary>
    /// <remarks>Usa Excel COM para monetario (obtiene valor formateado como Python/openpyxl). ClosedXML devuelve raw sin formato.</remarks>
    public static BalanceData? LeerBalanceValorizado(int anyo, int mes, string carpetaBase)
    {
        var archivo = EncontrarArchivoBalance(anyo, mes, carpetaBase);
        if (archivo == null) return null;
        var comResult = LeerBalanceValorizadoCom(archivo);
        if (comResult != null) return comResult;
        try
        {
            using var wb = new XLWorkbook(archivo);
            var ws = wb.Worksheet("Balance Valorizado") ?? wb.Worksheet(1);
            var (filaHeader, cols) = EncontrarEncabezadosBalance(ws);
            if (filaHeader < 0) return null;
            var datos = new List<Dictionary<string, object>>();
            var lastRow = ws.LastRowUsed();
            var maxRowNum = lastRow != null ? lastRow.RowNumber() : filaHeader + 1;
            for (var r = filaHeader + 2; r <= maxRowNum; r++)
            {
                var row = new Dictionary<string, object>();
                foreach (var kv in cols)
                {
                    var cell = ws.Cell(r, kv.Value + 1);
                    object? val;
                    if (string.Equals(kv.Key, "monetario", StringComparison.OrdinalIgnoreCase))
                    {
                        var s = cell.GetFormattedString();
                        if (string.IsNullOrWhiteSpace(s)) s = cell.GetString();
                        var parsed = ParseMonetario(s);
                        if (parsed.HasValue)
                            val = parsed.Value;
                        else if (cell.TryGetValue(out double d))
                            val = RestaurarDecimalesMonetario(d);
                        else if (cell.TryGetValue(out int i))
                            val = RestaurarDecimalesMonetario(i);
                        else
                            val = (object?)s ?? "";
                    }
                    else if (cell.TryGetValue(out double d))
                        val = d;
                    else if (cell.TryGetValue(out int i))
                        val = (double)i;
                    else
                    {
                        var s = cell.GetString();
                        val = ParseDouble(s) ?? (object?)s;
                    }
                    row[kv.Key] = val ?? (object?)"";
                }
                datos.Add(row);
            }
            return new BalanceData(datos, cols);
        }
        catch { return null; }
    }

    /// <summary>Lee Balance con Excel COM para obtener monetario formateado (Text = "325.161.639,078" como Python).</summary>
    private static BalanceData? LeerBalanceValorizadoCom(string archivo)
    {
        if (string.IsNullOrEmpty(archivo) || !File.Exists(archivo)) return null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var excel = Activator.CreateInstance(excelType);
            if (excel == null) return null;
            try
            {
                dynamic ex = excel;
                ex.Visible = false;
                ex.DisplayAlerts = false;
                dynamic wb = ExcelComHelper.RetryOnBusy(() => ex.Workbooks.Open(archivo, 0, true));
                dynamic ws = null;
                for (var i = 1; i <= wb.Worksheets.Count; i++)
                {
                    var name = (string)wb.Worksheets.Item(i).Name;
                    if (name.IndexOf("Balance Valorizado", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ws = wb.Worksheets.Item(i);
                        break;
                    }
                }
                if (ws == null) { wb.Close(false); return null; }
                int filaHeader = -1;
                var cols = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (var r = 1; r <= 20; r++)
                {
                    for (var c = 1; c <= 50; c++)
                    {
                        var val = (ws.Cells[r, c].Text ?? "").ToString().Trim();
                        if (string.IsNullOrWhiteSpace(val)) continue;
                        var norm = val.ToLowerInvariant().Replace(" ", "_");
                        if (norm == "barra") cols["barra"] = c - 1;
                        else if (norm == "monetario") cols["monetario"] = c - 1;
                        else if (norm.Contains("nombre_corto") && norm.Contains("empresa")) cols["nombre_corto_empresa"] = c - 1;
                        else if (norm.Contains("fisico") && norm.Contains("kwh")) cols["fisico_kwh"] = c - 1;
                        else if (norm.Contains("nombre_medidor")) cols["nombre_medidor"] = c - 1;
                    }
                    if (cols.Count >= 2) { filaHeader = r; break; }
                }
                if (filaHeader < 0) { wb.Close(false); return null; }
                var datos = new List<Dictionary<string, object>>();
                var usedR = ws.UsedRange;
                int maxRow = usedR != null ? (int)usedR.Row + (int)usedR.Rows.Count - 1 : filaHeader + 100;
                for (var r = filaHeader + 2; r <= maxRow; r++)
                {
                    var row = new Dictionary<string, object>();
                    foreach (var kv in cols)
                    {
                        var cell = ws.Cells[r, kv.Value + 1];
                        object? val;
                        if (string.Equals(kv.Key, "monetario", StringComparison.OrdinalIgnoreCase))
                        {
                            var text = (cell.Text ?? "").ToString().Trim();
                            var parsed = ParseMonetario(text);
                            if (parsed.HasValue)
                                val = parsed.Value;
                            else
                            {
                                var raw = cell.Value;
                                if (raw is double d) val = RestaurarDecimalesMonetario(d);
                                else if (raw is int i) val = RestaurarDecimalesMonetario(i);
                                else val = (object?)text ?? "";
                            }
                        }
                        else
                        {
                            var raw = cell.Value;
                            if (raw is double d) val = d;
                            else if (raw is int i) val = (double)i;
                            else
                            {
                                var s = (raw?.ToString() ?? "").Trim();
                                val = ParseDouble(s) ?? (object?)s;
                            }
                        }
                        row[kv.Key] = val ?? (object?)"";
                    }
                    datos.Add(row);
                }
                wb.Close(false);
                return new BalanceData(datos, cols);
            }
            finally { try { ((dynamic)excel).Quit(); } catch { } }
        }
        catch { return null; }
    }

    private static (int filaHeader, Dictionary<string, int> cols) EncontrarEncabezadosBalance(IXLWorksheet ws)
    {
        for (var r = 1; r <= 20; r++)
        {
            var cols = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var c = 1; c <= 50; c++)
            {
                var val = ws.Cell(r, c).GetString();
                if (string.IsNullOrWhiteSpace(val)) continue;
                var norm = val.ToLowerInvariant().Replace(" ", "_");
                if (norm == "barra") cols["barra"] = c - 1;
                else if (norm == "monetario") cols["monetario"] = c - 1;
                else if (norm.Contains("nombre_corto") && norm.Contains("empresa")) cols["nombre_corto_empresa"] = c - 1;
                else if (norm.Contains("fisico") && norm.Contains("kwh")) cols["fisico_kwh"] = c - 1;
                else if (norm.Contains("nombre_medidor")) cols["nombre_medidor"] = c - 1;
            }
            if (cols.Count >= 2) return (r, cols);
        }
        return (-1, new Dictionary<string, int>());
    }

    /// <summary>Valor razonable para Pago PSUF: positivo, típicamente &lt; 100e9 CLP.</summary>
    private static bool EsPagoPsufValido(double? v) => v.HasValue && v.Value >= 0 && v.Value < 100_000_000_000;

    public static double? LeerTotalIngresosPotenciaFirme(int anyo, int mes, string carpetaBase,
        string? nombreEmpresa = null, List<string>? conceptoFiltro = null)
    {
        var anexo = EncontrarAnexo02b(anyo, mes, carpetaBase);
        if (!string.IsNullOrWhiteSpace(nombreEmpresa) && anexo != null)
        {
            var vAnexo = LeerTotalIngresosPotenciaFirmeAnexo(anexo, "01.BALANCE POTENCIA", nombreEmpresa.Trim());
            if (vAnexo.HasValue && EsPagoPsufValido(vAnexo))
                return vAnexo.Value;
        }
        var archivo = EncontrarArchivoBdefDetalle(anyo, mes, carpetaBase);
        if (archivo != null)
        {
            var ext = Path.GetExtension(archivo).ToLowerInvariant();
            if (ext == ".xlsx" || ext == ".xlsm")
            {
                try
                {
                    using var wb = new XLWorkbook(archivo);
                    var ws = wb.Worksheet("Balance2") ?? wb.Worksheets.FirstOrDefault(w => w.Name.IndexOf("Balance", StringComparison.OrdinalIgnoreCase) >= 0);
                    if (ws != null)
                    {
                        var v = LeerPagoPsufDesdeBalance2(ws, nombreEmpresa, conceptoFiltro);
                        if (EsPagoPsufValido(v)) return v;
                    }
                }
                catch { }
            }
            var vCom = LeerPagoPsufBdefCom(archivo, nombreEmpresa, conceptoFiltro);
            if (EsPagoPsufValido(vCom)) return vCom;
        }
        if (anexo == null) return null;
        if (!string.IsNullOrWhiteSpace(nombreEmpresa))
        {
            var v = LeerTotalIngresosPotenciaFirmeAnexo(anexo, "01.BALANCE POTENCIA", nombreEmpresa.Trim());
            if (v.HasValue) return v.Value;
        }
        return LeerValorConceptoAnexo(anexo, "01.BALANCE POTENCIA", "TOTAL INGRESOS POR POTENCIA FIRME", excluir: null)
            ?? LeerValorConceptoAnexo(anexo, "01.BALANCE POTENCIA", "POTENCIA FIRME", excluir: null);
    }

    /// <summary>Lee BDef Detalle con Excel COM (para .xlsb o cuando ClosedXML falla). Filtra por Empresa y Concepto (POTENCIA_FIRME).</summary>
    private static double? LeerPagoPsufBdefCom(string ruta, string? nombreEmpresa, List<string>? conceptoFiltro)
    {
        if (string.IsNullOrEmpty(ruta) || !File.Exists(ruta)) return null;
        var conceptos = conceptoFiltro ?? new List<string> { "Eólica" };
        var empNorm = (nombreEmpresa ?? "").Trim().Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var excel = Activator.CreateInstance(excelType);
            if (excel == null) return null;
            try
            {
                dynamic ex = excel;
                ex.Visible = false;
                ex.DisplayAlerts = false;
                dynamic wb = ExcelComHelper.RetryOnBusy(() => ex.Workbooks.Open(ruta, 0, true));
                dynamic ws = null;
                for (var i = 1; i <= wb.Worksheets.Count; i++)
                {
                    var name = (string)wb.Worksheets.Item(i).Name;
                    if (name.IndexOf("Balance", StringComparison.OrdinalIgnoreCase) >= 0 && name.IndexOf("2", StringComparison.Ordinal) >= 0)
                    {
                        ws = wb.Worksheets.Item(i);
                        break;
                    }
                }
                if (ws == null)
                {
                    for (var i = 1; i <= wb.Worksheets.Count; i++)
                    {
                        var name = (string)wb.Worksheets.Item(i).Name;
                        if (name.Equals("Balance2", StringComparison.OrdinalIgnoreCase)) { ws = wb.Worksheets.Item(i); break; }
                    }
                }
                if (ws == null) { wb.Close(false); return null; }
                int colEmpresa = -1, colConcepto = -1, colPago = -1;
                for (var r = 1; r <= 15; r++)
                {
                    for (var c = 1; c <= 30; c++)
                    {
                        var val = (ws.Cells[r, c].Value?.ToString() ?? "").ToUpper().Replace("Ó", "O");
                        if (val.Contains("EMPRESA") && !val.Contains("CONCEPTO")) colEmpresa = c;
                        if (val.Contains("CONCEPTO")) colConcepto = c;
                        if (val.Contains("PAGO") && val.Contains("PSUF")) colPago = c;
                    }
                }
                if (colConcepto <= 0 || colPago <= 0) { wb.Close(false); return null; }
                double total = 0;
                string? ultimaEmpresa = null;
                var usedR = ws.UsedRange;
                int maxR = usedR != null ? (int)usedR.Row + (int)usedR.Rows.Count - 1 : 500;
                if (maxR < 6) maxR = 500;
                for (var r = 6; r <= maxR; r++)
                {
                    var empRaw = ws.Cells[r, colEmpresa].Value;
                    var empVal = (empRaw?.ToString() ?? "").Trim();
                    if (!string.IsNullOrEmpty(empVal)) ultimaEmpresa = empVal;
                    var empresaFiltrar = ultimaEmpresa ?? empVal;
                    if (!string.IsNullOrEmpty(empNorm) && !string.IsNullOrEmpty(empresaFiltrar))
                    {
                        var empCell = empresaFiltrar.Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
                        if (empCell != empNorm && !empCell.Contains(empNorm) && !empNorm.Contains(empCell)) continue;
                    }
                    var conceptoRaw = ws.Cells[r, colConcepto].Value;
                    var concepto = (conceptoRaw?.ToString() ?? "").Trim();
                    if (conceptos.Count > 0)
                    {
                        if (string.IsNullOrWhiteSpace(concepto)) continue;
                        var conceptoNorm = Normalizar(concepto);
                        if (!conceptos.Any(c => conceptoNorm.IndexOf(Normalizar(c ?? ""), StringComparison.OrdinalIgnoreCase) >= 0)) continue;
                    }
                    var pagoRaw = ws.Cells[r, colPago].Value;
                    var pagoVal = ParseMonetarioToDouble(pagoRaw);
                    total += pagoVal;
                }
                wb.Close(false);
                return total;
            }
            finally { try { ((dynamic)excel).Quit(); } catch { } }
        }
        catch { return null; }
    }

    /// <summary>Python _leer_pago_psuf: filtra por Empresa Y Concepto (ej: Eólica).</summary>
    private static double? LeerPagoPsufDesdeBalance2(IXLWorksheet ws, string? nombreEmpresa, List<string>? conceptoFiltro)
    {
        var conceptos = conceptoFiltro ?? new List<string> { "Eólica" };
        int colEmpresa = -1, colConcepto = -1, colPago = -1;
        for (var r = 1; r <= 15; r++)
        {
            for (var c = 1; c <= 30; c++)
            {
                var val = ws.Cell(r, c).GetString().ToUpper().Replace("Ó", "O");
                if (val.Contains("EMPRESA") && !val.Contains("CONCEPTO")) colEmpresa = c;
                if (val.Contains("CONCEPTO")) colConcepto = c;
                if (val.Contains("PAGO") && val.Contains("PSUF")) colPago = c;
            }
        }
        if (colConcepto <= 0 || colPago <= 0) return null;
        var empNorm = (nombreEmpresa ?? "").Trim().Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
        double total = 0;
        string? ultimaEmpresa = null;
        var lru = ws.LastRowUsed();
        var maxR = lru != null ? lru.RowNumber() : 6;
        for (var r = 6; r <= maxR; r++)
        {
            var empVal = colEmpresa > 0 ? ws.Cell(r, colEmpresa).GetString()?.Trim() : "";
            if (!string.IsNullOrEmpty(empVal)) ultimaEmpresa = empVal;
            var empresaFiltrar = ultimaEmpresa ?? empVal;
            if (!string.IsNullOrEmpty(empNorm) && !string.IsNullOrEmpty(empresaFiltrar))
            {
                var empCell = empresaFiltrar.Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
                if (empCell != empNorm && !empCell.Contains(empNorm) && !empNorm.Contains(empCell)) continue;
            }
            var concepto = ws.Cell(r, colConcepto).GetString();
            if (conceptos.Count > 0)
            {
                if (string.IsNullOrWhiteSpace(concepto)) continue;
                var conceptoNorm = Normalizar(concepto);
                if (!conceptos.Any(c => conceptoNorm.IndexOf(Normalizar(c ?? ""), StringComparison.OrdinalIgnoreCase) >= 0)) continue;
            }
            var pago = ParseDouble(ws.Cell(r, colPago).GetString());
            total += pago ?? 0;
        }
        return total;
    }

    public static double? LeerIngresosPorIt(int anyo, int mes, string carpetaBase, string? nombreEmpresa = null)
    {
        var anexo = EncontrarAnexo02b(anyo, mes, carpetaBase);
        if (anexo == null)
        {
            GeneradorInformeElectrico.LogHelper.Log($"  [WARN] Anexo 02.b Potencia no encontrado para {mes}/{anyo} - INGRESOS POR IT no disponible");
            return null;
        }
        if (!string.IsNullOrWhiteSpace(nombreEmpresa))
        {
            var v = LeerValorPorEmpresaYColumna(anexo, "02.IT POTENCIA", nombreEmpresa.Trim(), "Total");
            if (v.HasValue) return Math.Abs(v.Value);
        }
        foreach (var concepto in new[] { "INGRESOS POR IT POTENCIA", "INGRESOS POR IT", "IT POTENCIA", "02.IT POTENCIA" })
        {
            var v = LeerValorConceptoAnexo(anexo, "02.IT POTENCIA", concepto, excluir: new[] { "FIRME" });
            if (v.HasValue) return v;
        }
        GeneradorInformeElectrico.LogHelper.Log($"  [WARN] INGRESOS POR IT: archivo encontrado pero valor no localizado en hoja 02.IT POTENCIA");
        return null;
    }

    public static double? LeerIngresosPorPotencia(int anyo, int mes, string carpetaBase, string? nombreEmpresa = null)
    {
        var anexo = EncontrarAnexo02b(anyo, mes, carpetaBase);
        if (anexo == null) return null;
        if (!string.IsNullOrWhiteSpace(nombreEmpresa))
        {
            var v = LeerTotalIngresosPotenciaFirmeAnexo(anexo, "01.BALANCE POTENCIA", nombreEmpresa.Trim());
            if (v.HasValue) return Math.Abs(v.Value);
            v = LeerValorPorEmpresaYColumna(anexo, "01.BALANCE POTENCIA", nombreEmpresa.Trim(), "TOTAL");
            if (v.HasValue) return Math.Abs(v.Value);
        }
        return LeerValorConceptoAnexo(anexo, "01.BALANCE POTENCIA", "INGRESOS POR POTENCIA", excluir: new[] { "FIRME", "POR IT" });
    }

    /// <summary>Python leer_total_ingresos_potencia_firme_anexo: col B=Empresa, col D=TOTAL, estructura fija.</summary>
    private static double? LeerTotalIngresosPotenciaFirmeAnexo(string ruta, string hojaPatron, string nombreEmpresa)
    {
        if (string.IsNullOrEmpty(ruta) || !File.Exists(ruta)) return null;
        var empNorm = nombreEmpresa.Trim().Replace(" ", "_").ToUpperInvariant();
        if (string.IsNullOrEmpty(empNorm)) return null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var excel = Activator.CreateInstance(excelType);
            if (excel == null) return null;
            try
            {
                dynamic ex = excel;
                ex.Visible = false;
                ex.DisplayAlerts = false;
                dynamic wb = ExcelComHelper.RetryOnBusy(() => ex.Workbooks.Open(ruta, 0, true));
                dynamic ws = null;
                for (var i = 1; i <= wb.Worksheets.Count; i++)
                {
                    var name = (string)wb.Worksheets.Item(i).Name;
                    var nLower = name.ToLowerInvariant();
                    if (nLower.Contains("01") && nLower.Contains("balance") && nLower.Contains("potencia"))
                    {
                        ws = wb.Worksheets.Item(i);
                        break;
                    }
                }
                if (ws == null) { wb.Close(false); return null; }
                for (var r = 1; r <= 600; r++)
                {
                    object ov = ws.Cells[r, 2].Value ?? ws.Cells[r, 1].Value ?? ws.Cells[r, 3].Value;
                    var empVal = (ov?.ToString() ?? "").Trim().Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
                    if (string.IsNullOrEmpty(empVal)) continue;
                    if (!empNorm.Contains(empVal) && !empVal.Contains(empNorm) && !empNorm.StartsWith(empVal) && !empVal.StartsWith(empNorm))
                        continue;
                    var totalVal = ws.Cells[r, 4].Value ?? ws.Cells[r, 3].Value;
                    double? parsed = null;
                    if (totalVal is double d) parsed = d;
                    else if (totalVal != null) parsed = ParseDouble(totalVal.ToString());
                    if (parsed.HasValue) { wb.Close(false); return parsed.Value; }
                }
                var mu = new List<string>();
                for (var rr = 1; rr <= Math.Min(20, 200); rr++)
                {
                    object o = ws.Cells[rr, 2].Value ?? ws.Cells[rr, 1].Value;
                    var s = (o?.ToString() ?? "").Trim();
                    if (!string.IsNullOrEmpty(s)) mu.Add($"'{s}'");
                }
                GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] Balance: No match '{nombreEmpresa}'. Muestras col B/A: {string.Join(", ", mu.Take(5))}");
                wb.Close(false);
                return null;
            }
            finally
            {
                try { ((dynamic)excel).Quit(); } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch { }
            }
        }
        catch { return null; }
    }

    /// <summary>Python _leer_valor_por_empresa_y_columna: fila por USUARIOS=empresa, valor de col Total.</summary>
    private static double? LeerValorPorEmpresaYColumna(string ruta, string hojaPatron, string nombreEmpresa, string colValor)
    {
        if (string.IsNullOrEmpty(ruta) || !File.Exists(ruta)) return null;
        var empNorm = Normalizar(nombreEmpresa.Trim().Replace(" ", "_").ToUpperInvariant());
        if (string.IsNullOrEmpty(empNorm)) return null;
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var excel = Activator.CreateInstance(excelType);
            if (excel == null) return null;
            try
            {
                dynamic ex = excel;
                ex.Visible = false;
                ex.DisplayAlerts = false;
                dynamic wb = ExcelComHelper.RetryOnBusy(() => ex.Workbooks.Open(ruta, 0, true));
                dynamic ws = null;
                var esIt = hojaPatron.IndexOf("IT", StringComparison.OrdinalIgnoreCase) >= 0;
                for (var i = 1; i <= wb.Worksheets.Count; i++)
                {
                    var name = (string)wb.Worksheets.Item(i).Name;
                    var nLower = name.ToLowerInvariant();
                    var match = esIt
                        ? (nLower.Contains("02") && nLower.Contains("it") && nLower.Contains("potencia"))
                        : (nLower.Contains("01") && nLower.Contains("balance") && nLower.Contains("potencia"));
                    if (match)
                    {
                        ws = wb.Worksheets.Item(i);
                        GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] Hoja encontrada: {name}");
                        break;
                    }
                }
                if (ws == null) { GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] No se encontró hoja para patrón {hojaPatron}"); wb.Close(false); return null; }
                int headerRow = -1;
                for (var r = 1; r <= 25; r++)
                {
                    for (var c = 1; c <= 10; c++)
                    {
                        var v = ws.Cells[r, c].Value?.ToString()?.Trim().ToUpperInvariant() ?? "";
                        if (v == "USUARIOS" || v == "EMPRESA" || v.StartsWith("USUARIOS"))
                        {
                            headerRow = r;
                            break;
                        }
                    }
                    if (headerRow > 0) break;
                }
                if (headerRow <= 0) { GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] No se encontró fila USUARIOS/EMPRESA"); wb.Close(false); return null; }
                int colEmpresa = 0;
                int colTarget = -1;
                var maxCols = 260;
                for (var c = 1; c <= maxCols; c++)
                {
                    var h = ws.Cells[headerRow, c].Value?.ToString()?.Trim().ToLowerInvariant() ?? "";
                    if (h == "usuarios" || (colEmpresa == 0 && (h.Contains("usuario") || h.Contains("empresa"))))
                        colEmpresa = c;
                    if (h == "total" && !h.Contains("general"))
                    {
                        colTarget = c;
                        break;
                    }
                }
                if (colTarget < 0)
                    for (var c = 1; c <= maxCols; c++)
                    {
                        var h = (ws.Cells[headerRow, c].Value?.ToString() ?? "").Trim().ToLowerInvariant();
                        if (h.Contains("total") && !h.Contains("general")) { colTarget = c; break; }
                    }
                if (colTarget < 0)
                    for (var c = maxCols; c >= 1; c--)
                    {
                        var h = (ws.Cells[headerRow, c].Value?.ToString() ?? "").Trim().ToLowerInvariant();
                        if (h.Contains("total")) { colTarget = c; break; }
                    }
                if (colTarget < 0 && esIt)
                {
                    for (var c = colEmpresa + 1; c <= Math.Min(colEmpresa + 250, maxCols); c++)
                    {
                        var h = ws.Cells[headerRow, c].Value?.ToString()?.Trim().ToLowerInvariant() ?? "";
                        if (h.Contains("total") || h.Contains("asignacion") || h.Contains("asignación") || h.Contains("valor")) { colTarget = c; break; }
                    }
                }
                if (colTarget < 0 && esIt && colEmpresa > 0)
                    colTarget = colEmpresa + 2;
                if (colEmpresa <= 0 || colTarget < 0)
                {
                    var headers = string.Join(" | ", Enumerable.Range(1, 12).Select(c => $"C{c}={(ws.Cells[headerRow, c].Value?.ToString() ?? "")?.Trim()}"));
                    GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] Columnas no encontradas: emp={colEmpresa}, target={colTarget}. Headers: {headers}");
                    wb.Close(false);
                    return null;
                }
                double total = 0;
                for (var r = headerRow + 1; r <= 600; r++)
                {
                    var empVal = Normalizar(ws.Cells[r, colEmpresa].Value?.ToString()?.Trim().Replace(" ", "_").ToUpperInvariant() ?? "");
                    if (string.IsNullOrEmpty(empVal)) continue;
                    if (!empNorm.Contains(empVal) && !empVal.Contains(empNorm) && !empNorm.StartsWith(empVal) && !empVal.StartsWith(empNorm))
                        continue;
                    var valRaw = ws.Cells[r, colTarget].Value;
                    double? parsed = null;
                    if (valRaw is double d) parsed = d;
                    else if (valRaw != null) parsed = ParseDouble(valRaw.ToString());
                    if (parsed.HasValue) total += parsed.Value;
                }
                if (total == 0)
                {
                    var muestras = new List<string>();
                    for (var rr = headerRow + 1; rr <= Math.Min(headerRow + 15, 200); rr++)
                    {
                        object ov2 = ws.Cells[rr, colEmpresa].Value;
                        var s = (ov2?.ToString() ?? "").Trim();
                        if (!string.IsNullOrEmpty(s)) muestras.Add($"'{s}'");
                    }
                    GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] No match para '{nombreEmpresa}'. Muestras empresa en Excel: {string.Join(", ", muestras.Take(8))}");
                }
                wb.Close(false);
                return total != 0 ? total : null;
            }
            finally
            {
                try { ((dynamic)excel).Quit(); } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch { }
            }
        }
        catch (Exception ex) { GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] Excepción LeerValorPorEmpresa: {ex.Message}"); return null; }
    }

    /// <summary>Busca textoConcepto en la hoja y retorna el valor numérico de la columna adyacente.</summary>
    private static double? LeerValorConceptoAnexo(string ruta, string hojaPatron, string textoConcepto, string[]? excluir = null)
    {
        if (ruta == null || !File.Exists(ruta)) return null;
        var ext = Path.GetExtension(ruta).ToLowerInvariant();
        if (ext == ".xlsb")
            return LeerValorConceptoAnexoCom(ruta, hojaPatron, textoConcepto, excluir);
        try
        {
            using var wb = new XLWorkbook(ruta);
            var ws = wb.Worksheets.FirstOrDefault(w => w.Name.IndexOf(hojaPatron.Split('.')[0], StringComparison.OrdinalIgnoreCase) >= 0 || w.Name.IndexOf(hojaPatron, StringComparison.OrdinalIgnoreCase) >= 0);
            if (ws == null) return null;
            var textoUpper = textoConcepto.ToUpperInvariant();
            var lru = ws.LastRowUsed();
            var lcu = ws.LastColumnUsed();
            var maxR = lru != null ? Math.Min(lru.RowNumber(), 200) : 50;
            var maxC = lcu != null ? Math.Min(lcu.ColumnNumber(), 50) : 20;
            for (var r = 1; r <= maxR; r++)
            {
                for (var c = 1; c <= maxC; c++)
                {
                    var cellVal = (ws.Cell(r, c).GetString() ?? "").Trim().ToUpperInvariant();
                    if (!cellVal.Contains(textoUpper)) continue;
                    if (excluir != null && excluir.Any(e => cellVal.Contains(e.ToUpperInvariant()))) continue;
                    for (var c2 = c + 1; c2 <= maxC + 5; c2++)
                    {
                        var v = ParseDouble(ws.Cell(r, c2).GetString());
                        if (v.HasValue) return Math.Abs(v.Value);
                    }
                }
            }
            return null;
        }
        catch { return null; }
    }

    private static double? LeerValorConceptoAnexoCom(string ruta, string hojaPatron, string textoConcepto, string[]? excluir)
    {
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var excel = Activator.CreateInstance(excelType);
            if (excel == null) return null;
            try
            {
                dynamic ex = excel;
                ex.Visible = false;
                ex.DisplayAlerts = false;
                dynamic wb = ExcelComHelper.RetryOnBusy(() => ex.Workbooks.Open(ruta, 0, true));
                dynamic worksheets = wb.Worksheets;
                dynamic ws = null;
                for (var i = 1; i <= worksheets.Count; i++)
                {
                    var sheet = worksheets.Item(i);
                    var name = (string)sheet.Name;
                    if (name.IndexOf(hojaPatron.Split('.')[0], StringComparison.OrdinalIgnoreCase) >= 0 || name.IndexOf(hojaPatron, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ws = sheet;
                        break;
                    }
                }
                if (ws == null) { wb.Close(false); return null; }
                var textoUpper = textoConcepto.ToUpperInvariant();
                var maxR = 200;
                var maxC = 50;
                for (var r = 1; r <= maxR; r++)
                {
                    for (var c = 1; c <= maxC; c++)
                    {
                        var raw = ws.Cells[r, c].Value;
                        if (raw == null) continue;
                        var cellVal = (raw.ToString() ?? "").Trim().ToUpperInvariant();
                        if (!cellVal.Contains(textoUpper)) continue;
                        if (excluir != null && excluir.Any(e => cellVal.Contains(e.ToUpperInvariant()))) continue;
                        for (var c2 = c + 1; c2 <= maxC + 10; c2++)
                        {
                            var vRaw = ws.Cells[r, c2].Value;
                            if (vRaw == null) continue;
                            var v = ParseDouble(vRaw.ToString());
                            if (v.HasValue) { wb.Close(false); return Math.Abs(v.Value); }
                        }
                    }
                }
                wb.Close(false);
                return null;
            }
            finally
            {
                try { ((dynamic)excel).Quit(); } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch { }
            }
        }
        catch { return null; }
    }

    /// <summary>Python leer_total_ingresos_sscc: 1_CUADROS_PAGO_SSCC, hoja CPI_, Nemotecnico Deudor + Monto.</summary>
    public static double? LeerTotalIngresosSscc(int anyo, int mes, string nombreEmpresa, string carpetaBase)
    {
        var archivo = EncontrarCuadrosPagoSscc(anyo, mes, carpetaBase);
        if (archivo == null)
        {
            GeneradorInformeElectrico.LogHelper.Log($"  [WARN] No se encontró EXCEL 1_CUADROS_PAGO_SSCC para {mes}/{anyo}");
            return null;
        }
        if (string.IsNullOrWhiteSpace(nombreEmpresa))
        {
            GeneradorInformeElectrico.LogHelper.Log("  [WARN] TOTAL INGRESOS POR SSCC: ingrese Empresa para filtrar por Nemotecnico Deudor");
            return null;
        }
        GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] SSCC archivo: {Path.GetFileName(archivo)}");
        var ext = Path.GetExtension(archivo).ToLowerInvariant();
        if (ext == ".xlsb" || ext == ".xlsx" || ext == ".xlsm")
            return LeerTotalIngresosSsccCom(archivo, nombreEmpresa);
        return null;
    }

    private static double? LeerTotalIngresosSsccCom(string archivo, string nombreEmpresa)
    {
        try
        {
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null) return null;
            var excel = Activator.CreateInstance(excelType);
            if (excel == null) return null;
            try
            {
                dynamic ex = excel;
                ex.Visible = false;
                ex.DisplayAlerts = false;
                dynamic wb = ExcelComHelper.RetryOnBusy(() => ex.Workbooks.Open(archivo, 0, true));
                dynamic ws = null;
                for (var i = 1; i <= wb.Worksheets.Count; i++)
                {
                    var name = ((string)wb.Worksheets.Item(i).Name).ToLowerInvariant();
                    if (name == "cpi_" || name.StartsWith("cpi")) { ws = wb.Worksheets.Item(i); break; }
                }
                if (ws == null) { wb.Close(false); return null; }
                var result = LeerSsccDesdeWorksheetCom(ws, nombreEmpresa);
                wb.Close(false);
                return result;
            }
            finally
            {
                try { ((dynamic)excel).Quit(); } catch { }
                try { System.Runtime.InteropServices.Marshal.ReleaseComObject(excel); } catch { }
            }
        }
        catch { return null; }
    }

    private static double? LeerSsccDesdeWorksheet(IXLWorksheet ws, string nombreEmpresa)
    {
        int colDeudor = -1, colMonto = -1, filaHeader = -1;
        for (var r = 1; r <= 15; r++)
        {
            var filaStr = "";
            for (var c = 1; c <= 15; c++)
            {
                var v = ws.Cell(r, c).GetString();
                if (!string.IsNullOrEmpty(v)) filaStr += " " + v;
            }
            filaStr = filaStr.ToLower().Replace("ó", "o").Replace("í", "i");
            if (filaStr.Contains("nemotecnico") && filaStr.Contains("deudor") && filaStr.Contains("monto"))
            {
                filaHeader = r;
                for (var c = 1; c <= 10; c++)
                {
                    var val = ws.Cell(r, c).GetString().ToLower().Replace("ó", "o");
                    if (val.Contains("nemotecnico") && val.Contains("deudor")) colDeudor = c;
                    if (val.Contains("monto") && !val.Contains("retencion")) colMonto = c;
                }
                break;
            }
        }
        if (filaHeader < 0 || colDeudor <= 0 || colMonto <= 0) return null;
        return SumarSsccPorEmpresa(ws, filaHeader, colDeudor, colMonto, nombreEmpresa);
    }

    private static double? LeerSsccDesdeWorksheetCom(dynamic ws, string nombreEmpresa)
    {
        int colDeudor = -1, colMonto = -1, filaHeader = -1;
        for (var r = 1; r <= 15; r++)
        {
            var filaStr = "";
            for (var c = 1; c <= 15; c++)
            {
                var v = ws.Cells[r, c].Value?.ToString() ?? "";
                if (!string.IsNullOrEmpty(v)) filaStr += " " + v;
            }
            filaStr = filaStr.ToLower().Replace("ó", "o").Replace("í", "i");
            if (filaStr.Contains("nemotecnico") && filaStr.Contains("deudor") && filaStr.Contains("monto"))
            {
                filaHeader = r;
                for (var c = 1; c <= 10; c++)
                {
                    var val = (ws.Cells[r, c].Value?.ToString() ?? "").ToLower().Replace("ó", "o").Replace("í", "i");
                    if (colDeudor <= 0 && val.Contains("nemotecnico") && val.Contains("deudor")) colDeudor = c;
                    if (colMonto <= 0 && val.Contains("monto") && !val.Contains("retencion")) colMonto = c;
                }
                break;
            }
        }
        if (filaHeader < 0 || colDeudor <= 0 || colMonto <= 0) return null;
        GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] SSCC CPI_: filaHeader={filaHeader}, colDeudor={colDeudor}, colMonto={colMonto}");
        return SumarSsccPorEmpresaCom(ws, filaHeader, colDeudor, colMonto, nombreEmpresa);
    }

    private static double? SumarSsccPorEmpresa(IXLWorksheet ws, int filaHeader, int colDeudor, int colMonto, string nombreEmpresa)
    {
        var empNorm = nombreEmpresa.Trim().Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
        double total = 0;
        var lru = ws.LastRowUsed();
        var maxR = lru != null ? lru.RowNumber() : filaHeader + 50;
        for (var r = filaHeader + 1; r <= maxR; r++)
        {
            var deudor = (ws.Cell(r, colDeudor).GetString() ?? "").Trim().Replace(" ", "_").Replace("-", "_").ToUpperInvariant();
            if (string.IsNullOrEmpty(deudor)) continue;
            if (deudor != empNorm && !deudor.Contains(empNorm) && !empNorm.Contains(deudor)) continue;
            var monto = ParseDouble(ws.Cell(r, colMonto).GetString());
            if (monto.HasValue) total += monto.Value;
        }
        if (total > 0) total = -total;
        GeneradorInformeElectrico.LogHelper.Log($"  [INFO] SSCC desde CPI_ ({nombreEmpresa}): {total:N2}");
        return total;
    }

    private static double? SumarSsccPorEmpresaCom(dynamic ws, int filaHeader, int colDeudor, int colMonto, string nombreEmpresa)
    {
        var empNorm = nombreEmpresa.Trim().ToUpperInvariant().Replace(" ", "_");
        double total = 0;
        var muestrasDeudor = new List<string>();
        int maxRow = filaHeader + 5000;
        try
        {
            var ur = ws.UsedRange;
            if (ur != null)
            {
                int firstRow = (int)ur.Row;
                int rowCount = (int)ur.Rows.Count;
                int lastRow = firstRow + rowCount - 1;
                if (lastRow > maxRow) maxRow = lastRow;
            }
        }
        catch { }
        GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] SSCC: leyendo filas {filaHeader + 1} a {maxRow}");
        for (var r = filaHeader + 1; r <= maxRow; r++)
        {
            try
            {
                var raw = ws.Cells[r, colDeudor].Value;
                if (raw == null) continue;
                var deudor = (raw.ToString() ?? "").Trim().ToUpperInvariant().Replace(" ", "_");
                if (string.IsNullOrEmpty(deudor)) continue;
                if (muestrasDeudor.Count < 20) muestrasDeudor.Add(deudor);
                if (deudor != empNorm) continue;
                object montoRaw = ws.Cells[r, colMonto].Value;
                double montoVal = ParseMonetarioToDouble(montoRaw);
                if (montoVal != 0)
                {
                    total += montoVal;
                    GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] SSCC match fila {r}: {deudor} -> {montoVal:N0}");
                }
            }
            catch (Exception ex) { GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] SSCC fila {r} error: {ex.Message}"); }
        }
        if (total == 0 && muestrasDeudor.Count > 0)
            GeneradorInformeElectrico.LogHelper.Log($"  [DEBUG] SSCC: buscando '{empNorm}'. Muestras: {string.Join(", ", muestrasDeudor.Distinct().Take(15))}. Configure SSCC_NEMOTECNICO en config_empresas.json si el valor en Excel difiere.");
        if (total > 0) total = -total;
        GeneradorInformeElectrico.LogHelper.Log($"  [INFO] SSCC desde CPI_ ({nombreEmpresa}): {total:N2}");
        return total;
    }

    /// <summary>Convierte valor COM a double (evita dynamic/HasValue).</summary>
    private static double ParseMonetarioToDouble(object val)
    {
        if (val == null) return 0;
        if (val is double d) return d;
        if (val is int i) return i;
        if (val is float f) return f;
        if (val is decimal dec) return (double)dec;
        var s = (val.ToString() ?? "").Trim();
        if (string.IsNullOrEmpty(s)) return 0;
        s = s.Replace(".", "").Replace(",", ".");
        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0;
    }

    /// <summary>Parsea valor monetario tal cual viene de Excel. Soporta: 325161639,078 (coma=decimal) y 32.516.163.908 (puntos=miles, último=decimal).</summary>
    private static double? ParseMonetario(object? val)
    {
        if (val == null) return null;
        if (val is double d && !(val is bool)) return d;
        if (val is int i) return (double)i;
        var s = (val.ToString() ?? "").Trim();
        if (string.IsNullOrEmpty(s)) return null;
        s = s.Replace("$", "").Replace(" ", "").Trim();
        if (s.Contains(","))
        {
            s = s.Replace(".", "").Replace(",", ".");
            return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v1) ? v1 : null;
        }
        if (s.Contains("."))
        {
            var parts = s.Split('.');
            if (parts.Length >= 2)
            {
                var intPart = string.Join("", parts.Take(parts.Length - 1));
                var decPart = parts[^1];
                if (int.TryParse(intPart, NumberStyles.None, CultureInfo.InvariantCulture, out var iVal) &&
                    int.TryParse(decPart, NumberStyles.None, CultureInfo.InvariantCulture, out var dVal))
                {
                    var divisor = Math.Pow(10, decPart.Length);
                    return iVal + dVal / divisor;
                }
            }
        }
        s = s.Replace(".", "").Replace(",", ".");
        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v2))
        {
            // Excel a veces devuelve 325161639078 (sin coma) cuando el valor es 325161639,078 → restaurar decimales
            if (v2 >= 1e9 && Math.Abs(v2 - Math.Floor(v2)) < 1e-9)
                return v2 / 1000;
            return v2;
        }
        return null;
    }

    /// <summary>Cuando el valor crudo es double sin decimales (ej. 325161639078), restaurar 325161639,078.</summary>
    private static double RestaurarDecimalesMonetario(double v)
    {
        if (v >= 1e9 && Math.Abs(v - Math.Floor(v)) < 1e-9)
            return v / 1000;
        return v;
    }

    /// <summary>Python leer_compra_venta: solo sección Físicos (excluir Financieros).</summary>
    public static double? LeerCompraVentaEnergiaGmHoldings(int anyo, int mes, string? nombreEmpresa, string? nombreBarra, string carpetaBase)
    {
        var archivo = EncontrarArchivoBalance(anyo, mes, carpetaBase);
        if (archivo == null) return null;
        try
        {
            using var wb = new XLWorkbook(archivo);
            var ws = wb.Worksheet("Contratos");
            if (ws == null) return null;
            int startDataRow = 1;
            int startCol = 1;
            int endCol = 10;
            for (var r = 1; r <= 50; r++)
            {
                for (var c = 1; c <= 12; c++)
                {
                    var val = ws.Cell(r, c).GetString().ToLower().Replace("í", "i");
                    if (val.Contains("financieros")) continue;
                    if (val.Contains("generadores") && (val.Contains("fisicos") || val.Contains("físicos")))
                    {
                        startDataRow = r;
                        startCol = c;
                        endCol = Math.Min(c + 7, 30);
                        break;
                    }
                }
            }
            int colEmpresa = -1, colBarra = -1, colVenta = -1;
            for (var r = startDataRow; r <= Math.Min(startDataRow + 8, 50); r++)
            {
                for (var c = startCol; c <= endCol; c++)
                {
                    var val = ws.Cell(r, c).GetString().ToLower();
                    if (val.Contains("nombre") && val.Contains("corto") && val.Contains("empresa")) colEmpresa = c;
                    if (val.Contains("barra")) colBarra = c;
                    if (val.Contains("venta") && val.Contains("clp")) colVenta = c;
                }
                if (colVenta > 0) { startDataRow = r + 1; break; }
            }
            if (colVenta <= 0)
            {
                colEmpresa = -1; colBarra = -1; colVenta = -1;
                for (var r = 1; r <= 20; r++)
                {
                    for (var c = 1; c <= 10; c++)
                    {
                        var val = ws.Cell(r, c).GetString().ToLower();
                        if (val.Contains("nombre") && val.Contains("corto") && val.Contains("empresa")) colEmpresa = c;
                        if (val.Contains("barra")) colBarra = c;
                        if (val.Contains("venta") && val.Contains("clp")) colVenta = c;
                    }
                    if (colVenta > 0) { startDataRow = r + 1; break; }
                }
            }
            if (colVenta <= 0) return null;
            var empNorm = (nombreEmpresa ?? "").Trim().Replace(" ", "_").ToUpperInvariant();
            var barNorm = (nombreBarra ?? "").Trim().ToUpperInvariant();
            double sum = 0;
            var lru3 = ws.LastRowUsed();
            var maxR3 = lru3 != null ? lru3.RowNumber() : startDataRow + 10;
            for (var r = startDataRow; r <= maxR3; r++)
            {
                var primeraCelda = ws.Cell(r, startCol).GetString().Trim().ToUpperInvariant();
                if (primeraCelda == "TOTAL" || primeraCelda.Contains("FINANCIEROS")) break;
                if (!string.IsNullOrEmpty(empNorm) && colEmpresa > 0)
                {
                    var emp = ws.Cell(r, colEmpresa).GetString().Trim().Replace(" ", "_").ToUpperInvariant();
                    if (emp != empNorm && !emp.Contains(empNorm) && !empNorm.Contains(emp)) continue;
                }
                if (!string.IsNullOrEmpty(barNorm) && colBarra > 0)
                {
                    var bar = ws.Cell(r, colBarra).GetString().Trim().ToUpperInvariant();
                    if (bar != barNorm && !bar.Contains(barNorm) && !barNorm.Contains(bar)) continue;
                }
                var v = ParseDouble(ws.Cell(r, colVenta).GetString());
                if (v.HasValue) sum += Math.Abs(v.Value);
            }
            return sum;
        }
        catch { return null; }
    }

    private static string Normalizar(string s)
    {
        if (string.IsNullOrEmpty(s)) return "";
        var n = s.ToLowerInvariant();
        foreach (var (o, n2) in new[] { ("í", "i"), ("ó", "o"), ("á", "a"), ("é", "e"), ("ú", "u"), ("ñ", "n") })
            n = n.Replace(o, n2);
        return n;
    }

    private static double? ParseDouble(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        s = s.Replace(",", "").Replace(".", CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
        return double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;
    }

    public record BalanceData(List<Dictionary<string, object?>> Filas, Dictionary<string, int> Columnas);
}
