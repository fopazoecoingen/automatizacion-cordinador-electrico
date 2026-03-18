using System.IO;
using GeneradorInformeElectrico.Models;
using Newtonsoft.Json;

namespace GeneradorInformeElectrico.Services;

public static class ConfigService
{
    public static string GetBaseDirectory()
    {
        return Path.GetDirectoryName(
            Environment.ProcessPath ?? AppContext.BaseDirectory
        ) ?? Directory.GetCurrentDirectory();
    }

    public static string GetCarpetaBaseDatos()
    {
        var baseDir = GetBaseDirectory();
        try
        {
            var configPath = Path.Combine(baseDir, "config.json");
            if (File.Exists(configPath))
            {
                var json = File.ReadAllText(configPath);
                var config = JsonConvert.DeserializeObject<Config>(json);
                var pathBd = config?.PathBd?.Trim();
                if (!string.IsNullOrEmpty(pathBd))
                {
                    var p = Path.IsPathRooted(pathBd)
                        ? Path.GetFullPath(pathBd)
                        : Path.GetFullPath(Path.Combine(baseDir, pathBd));
                    return p;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[WARNING] config.json: {ex.Message}");
        }
        return Path.GetFullPath(Path.Combine(baseDir, "bd_data"));
    }

    public static void GuardarCarpetaBaseDatos(string ruta)
    {
        var baseDir = GetBaseDirectory();
        var configPath = Path.Combine(baseDir, "config.json");
        try
        {
            Config? config = null;
            if (File.Exists(configPath))
            {
                var json = File.ReadAllText(configPath);
                config = JsonConvert.DeserializeObject<Config>(json);
            }
            config ??= new Config();
            config.PathBd = string.IsNullOrWhiteSpace(ruta) ? null : ruta.Trim();
            File.WriteAllText(configPath, JsonConvert.SerializeObject(new { path_bd = config.PathBd ?? "bd_data" }, Formatting.Indented));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[WARNING] Guardar config: {ex.Message}");
        }
    }

    public static ConfigEmpresas CargarConfigEmpresas()
    {
        var baseDir = GetBaseDirectory();
        var path = Path.Combine(baseDir, "config_empresas.json");
        if (!File.Exists(path)) return new ConfigEmpresas();
        try
        {
            var json = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<ConfigEmpresas>(json) ?? new ConfigEmpresas();
        }
        catch { return new ConfigEmpresas(); }
    }

    public static UltimosDatos? CargarUltimosDatos()
    {
        var baseDir = GetBaseDirectory();
        var path = Path.Combine(baseDir, "config_ultimos_datos.json");
        if (!File.Exists(path)) return null;
        try
        {
            var json = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<UltimosDatos>(json);
        }
        catch { return null; }
    }

    public static void GuardarUltimosDatos(UltimosDatos datos)
    {
        var baseDir = GetBaseDirectory();
        var path = Path.Combine(baseDir, "config_ultimos_datos.json");
        try
        {
            File.WriteAllText(path, JsonConvert.SerializeObject(datos, Formatting.Indented));
        }
        catch { }
    }
}
