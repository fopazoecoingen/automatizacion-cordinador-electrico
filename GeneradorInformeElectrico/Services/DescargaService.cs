using System.IO.Compression;
using System.Net.Http;

namespace GeneradorInformeElectrico.Services;

public static class DescargaService
{
    private static readonly HttpClient HttpClient = new()
    {
        Timeout = TimeSpan.FromMinutes(5)
    };

    private static readonly Dictionary<int, string> Meses = new()
    {
        { 1, "Enero" }, { 2, "Febrero" }, { 3, "Marzo" }, { 4, "Abril" },
        { 5, "Mayo" }, { 6, "Junio" }, { 7, "Julio" }, { 8, "Agosto" },
        { 9, "Septiembre" }, { 10, "Octubre" }, { 11, "Noviembre" }, { 12, "Diciembre" }
    };

    private static readonly Dictionary<int, string> MesesAbrev = new()
    {
        { 1, "ene" }, { 2, "feb" }, { 3, "mar" }, { 4, "abr" },
        { 5, "may" }, { 6, "jun" }, { 7, "jul" }, { 8, "ago" },
        { 9, "sep" }, { 10, "oct" }, { 11, "nov" }, { 12, "dic" }
    };

    public static readonly Dictionary<string, string> TiposArchivo = new()
    {
        { "energia_resultados", "01 Resultados (Energía)" },
        { "energia_antecedentes", "02 Antecedentes de Cálculo" },
        { "sscc", "Balance SSCC" },
        { "potencia", "Balance Psuf (Potencia)" }
    };

    public static string? BuscarArchivoExistente(int anyo, int mes, string tipo, string carpetaZip)
    {
        var patronBase = $"PLABACOM_{anyo}_{mes}_{Meses[mes]}";
        var patronExtra = tipo switch
        {
            "energia_resultados" => "Energia_Definitivo",
            "energia_antecedentes" => "Antecedentes",
            "sscc" => "SSCC",
            "potencia" => "Potencia",
            _ => null
        };
        if (patronExtra == null || !Directory.Exists(carpetaZip)) return null;
        foreach (var f in Directory.GetFiles(carpetaZip, "*.zip"))
        {
            var name = Path.GetFileName(f);
            if (name.Contains(patronBase) && name.Contains(patronExtra))
                return f;
        }
        return null;
    }

    public static (string url, string nombreLocal) ConstruirUrl(int anyo, int mes, string tipo)
    {
        var mesStr = mes.ToString("D2");
        var nombreMes = Meses[mes];
        var anyoAbrev = (anyo % 100).ToString("D2");
        var mesAbrev = MesesAbrev[mes];
        var baseS3 = $"PLABACOM/{anyo}/{mesStr}_{nombreMes}";
        var urlBase = "https://cen-plabacom.s3.amazonaws.com/";

        return tipo switch
        {
            "energia_resultados" => (
                $"{urlBase}{baseS3}/Energia/Definitivo/v_1/{Uri.EscapeDataString($"01 Resultados_{anyoAbrev}{mesStr}_BD01.zip")}",
                $"PLABACOM_{anyo}_{mes}_{nombreMes}_Energia_Definitivo_v_1_01 Resultados_{anyoAbrev}{mesStr}_BD01.zip"
            ),
            "sscc" => (
                $"{urlBase}{baseS3}/SSCC/Definitivo/v_1/{Uri.EscapeDataString($"Balance_SSCC_{anyo}_{mesAbrev}_def.zip")}",
                $"PLABACOM_{anyo}_{mes}_{nombreMes}_SSCC_Balance_SSCC_{anyo}_{mesAbrev}_def.zip"
            ),
            "potencia" => (
                $"{urlBase}{baseS3}/Potencia/Definitivo/v_1/{Uri.EscapeDataString($"Balance_Psuf_{anyoAbrev}{mesStr}_def.zip")}",
                $"PLABACOM_{anyo}_{mes}_{nombreMes}_Potencia_Balance_Psuf_{anyoAbrev}{mesStr}_def.zip"
            ),
            _ => throw new ArgumentException($"Tipo no soportado: {tipo}")
        };
    }

    public static async Task<(string? rutaZip, string? rutaDescomprimida, int? codigoError)> DescargarYDescomprimirAsync(
        int anyo, int mes, string tipo, string carpetaZip, string carpetaDescomprimidos,
        IProgress<string>? progress = null, CancellationToken ct = default)
    {
        Directory.CreateDirectory(carpetaZip);
        Directory.CreateDirectory(carpetaDescomprimidos);

        var existente = BuscarArchivoExistente(anyo, mes, tipo, carpetaZip);
        if (existente != null)
        {
            progress?.Report($"[OK] Archivo ya existe: {Path.GetFileName(existente)}");
            var rutaDescomprimida = DescomprimirZip(existente, carpetaDescomprimidos, progress);
            return (existente, rutaDescomprimida, null);
        }

        var (url, nombreLocal) = ConstruirUrl(anyo, mes, tipo);
        var rutaArchivo = Path.Combine(carpetaZip, nombreLocal);

        progress?.Report($"Descargando {TiposArchivo.GetValueOrDefault(tipo, tipo)}...");
        try
        {
            using var response = await HttpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, ct);
            if (response.StatusCode == System.Net.HttpStatusCode.Forbidden)
                return (null, null, 403);

            response.EnsureSuccessStatusCode();
            await using var fs = File.Create(rutaArchivo);
            await response.Content.CopyToAsync(fs, ct);
        }
        catch (HttpRequestException ex) when (ex.Message.Contains("403"))
        {
            return (null, null, 403);
        }
        catch (Exception ex)
        {
            progress?.Report($"Error descarga: {ex.Message}");
            return (null, null, null);
        }

        progress?.Report("Descomprimiendo...");
        var descomprimida = DescomprimirZip(rutaArchivo, carpetaDescomprimidos, progress);
        return (rutaArchivo, descomprimida, null);
    }

    private static string? DescomprimirZip(string rutaZip, string carpetaDestino, IProgress<string>? progress)
    {
        try
        {
            var nombreCarpeta = Path.GetFileNameWithoutExtension(rutaZip);
            if (nombreCarpeta.Contains("_v_1_"))
                nombreCarpeta = nombreCarpeta.Split(new[] { "_v_1_" }, StringSplitOptions.None).Last();
            var destino = Path.Combine(carpetaDestino, nombreCarpeta);
            if (Directory.Exists(destino) && Directory.EnumerateFileSystemEntries(destino).Any())
            {
                progress?.Report($"[OK] Ya descomprimido: {destino}");
                return destino;
            }
            Directory.CreateDirectory(destino);
            ZipFile.ExtractToDirectory(rutaZip, destino);
            progress?.Report($"[OK] Descomprimido en {destino}");
            return destino;
        }
        catch (Exception ex)
        {
            progress?.Report($"Error descomprimir: {ex.Message}");
            return null;
        }
    }
}
