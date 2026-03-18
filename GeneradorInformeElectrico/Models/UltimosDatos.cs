using Newtonsoft.Json;

namespace GeneradorInformeElectrico.Models;

public class UltimosDatos
{
    [JsonProperty("anyo")]
    public int? Anyo { get; set; }

    [JsonProperty("mes")]
    public int? Mes { get; set; }

    [JsonProperty("empresa")]
    public string? Empresa { get; set; }

    [JsonProperty("barra")]
    public string? Barra { get; set; }

    [JsonProperty("plantilla")]
    public string? Plantilla { get; set; }

    [JsonProperty("destino")]
    public string? Destino { get; set; }
}
