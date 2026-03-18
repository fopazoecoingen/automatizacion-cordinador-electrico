using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace GeneradorInformeElectrico.Models;

public class ConfigEmpresas
{
    public List<EmpresaConfig> Empresas { get; set; } = new();
}

public class EmpresaConfig
{
    [JsonProperty("nombreEmpresa")]
    public string? NombreEmpresa { get; set; }

    [JsonProperty("Barras")]
    public List<string>? Barras { get; set; }

    [JsonProperty("IMPORTACION_MWh")]
    public string? IMPORTACION_MWh { get; set; }

    [JsonProperty("TOTAL_INGRESOS_POR_ENERGIA_CLP")]
    public object? TOTAL_INGRESOS_POR_ENERGIA_CLP { get; set; }

    [JsonProperty("POTENCIA_FIRME")]
    public object? POTENCIA_FIRME { get; set; }

    /// <summary>Opcional: Nemotécnico para filtrar SSCC si difiere de nombreEmpresa.</summary>
    [JsonProperty("SSCC_NEMOTECNICO")]
    public string? SSCC_NEMOTECNICO { get; set; }

    /// <summary>Opcional: Medidores para GM Holdings si aplica.</summary>
    [JsonProperty("GM_HOLDINGS_MEDIDORES")]
    public object? GM_HOLDINGS_MEDIDORES { get; set; }

    public List<string> GetMedidoresEnergia()
    {
        if (TOTAL_INGRESOS_POR_ENERGIA_CLP == null) return new();
        if (TOTAL_INGRESOS_POR_ENERGIA_CLP is JArray arr)
            return arr.Select(x => x.ToString().Trim()).Where(x => !string.IsNullOrEmpty(x)).ToList();
        if (TOTAL_INGRESOS_POR_ENERGIA_CLP is string s && !string.IsNullOrWhiteSpace(s))
            return new List<string> { s.Trim() };
        return new();
    }

    public List<string> GetConceptoPotenciaFirme()
    {
        if (POTENCIA_FIRME == null) return new List<string> { "Eólica" };
        if (POTENCIA_FIRME is JArray arr)
            return arr.Select(x => x.ToString().Trim()).Where(x => !string.IsNullOrEmpty(x)).ToList();
        if (POTENCIA_FIRME is string s && !string.IsNullOrWhiteSpace(s))
            return new List<string> { s.Trim() };
        return new List<string> { "Eólica" };
    }
}
