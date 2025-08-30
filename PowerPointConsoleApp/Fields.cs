using System.Text.Json.Serialization;

public class Fields
{
    [JsonPropertyName("System.Title")]
    public string? SystemTitle { get; set; }

    [JsonPropertyName("System.Description")]
    public string? SystemDescription { get; set; }
}