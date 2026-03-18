using Newtonsoft.Json;

namespace GeneradorInformeElectrico.Models;

public class Config
{
    [JsonProperty("path_bd")]
    public string? PathBd { get; set; }
}
