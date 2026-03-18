namespace GeneradorInformeElectrico;

/// <summary>
/// Escribe logs a archivo para depuración cuando se ejecuta desde script.
/// Archivo: GeneradorInformeElectrico.log (junto al .exe o en la carpeta actual)
/// </summary>
public static class LogHelper
{
    private static string? _rutaLog;
    private static readonly object _lock = new();

    public static string RutaLog
    {
        get
        {
            if (_rutaLog != null) return _rutaLog;
            try
            {
                var dir = Path.GetDirectoryName(Environment.ProcessPath)
                    ?? Path.GetDirectoryName(AppContext.BaseDirectory)
                    ?? Path.GetTempPath();
                _rutaLog = Path.Combine(dir, "GeneradorInformeElectrico.log");
                return _rutaLog;
            }
            catch { return Path.Combine(Path.GetTempPath(), "GeneradorInformeElectrico.log"); }
        }
    }

    public static void Log(string mensaje)
    {
        var linea = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {mensaje}";
        lock (_lock)
        {
            try { File.AppendAllText(RutaLog, linea + Environment.NewLine); } catch { }
        }
    }

    public static void LogExcepcion(Exception ex)
    {
        try
        {
            Log($"ERROR Tipo: {ex?.GetType()?.Name ?? "null"}");
            Log($"  Mensaje: {ex?.Message ?? "(vacio)"}");
            if (ex is System.Runtime.InteropServices.COMException comEx)
                Log($"  HRESULT: 0x{comEx.HResult:X}");
            var inner = ex?.InnerException;
            while (inner != null)
            {
                Log($"  Inner ({inner.GetType().Name}): {inner.Message ?? "(vacio)"}");
                inner = inner.InnerException;
            }
            if (ex?.StackTrace != null) Log($"  Stack: {ex.StackTrace}");
            Log($"  ToString: {ex}");
        }
        catch (Exception logEx) { Log($"Fallo al loguear: {logEx.Message}"); }
    }

    public static void IniciarSesion()
    {
        Log("=== Iniciando aplicación (v2 late-binding) ===");
        Log($"Log en: {RutaLog}");
    }
}
