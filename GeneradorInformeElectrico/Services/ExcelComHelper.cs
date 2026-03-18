using System.Runtime.InteropServices;

namespace GeneradorInformeElectrico.Services;

/// <summary>
/// Ayuda a manejar errores COM de Excel (RPC_E_SERVERCALL_RETRYLATER cuando Excel está ocupado).
/// También captura RuntimeBinderException con "0x8001010A" (dispatch ID, Count, etc.).
/// </summary>
public static class ExcelComHelper
{
    private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

    private static bool EsExcelOcupado(Exception ex)
    {
        if (ex is COMException com && com.HResult == RPC_E_SERVERCALL_RETRYLATER) return true;
        var msg = (ex.Message ?? "") + (ex.InnerException?.Message ?? "");
        return msg.Contains("0x8001010A") || msg.Contains("8001010A");
    }

    /// <summary>Ejecuta una acción COM y reintenta si Excel responde que está ocupado.</summary>
    public static T RetryOnBusy<T>(Func<T> action, int maxRetries = 5)
    {
        for (var i = 0; i < maxRetries; i++)
        {
            try
            {
                return action();
            }
            catch (Exception ex) when (EsExcelOcupado(ex))
            {
                if (i == maxRetries - 1) throw;
                var delay = 1500 + (i * 500);
                GeneradorInformeElectrico.LogHelper.Log($"  [INFO] Excel ocupado, reintentando en {delay}ms...");
                Thread.Sleep(delay);
            }
        }
        throw new InvalidOperationException("Excel no respondió tras varios intentos. Cierre otras ventanas de Excel e intente de nuevo.");
    }

    /// <summary>Ejecuta una acción COM y reintenta si Excel responde que está ocupado.</summary>
    public static void RetryOnBusy(Action action, int maxRetries = 5)
    {
        RetryOnBusy(() => { action(); return 0; }, maxRetries);
    }
}
