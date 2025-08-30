using System.Text.Json.Serialization;

public class Sprint
{
    [JsonPropertyName("id")]
    public string? Id { get; set; }
    [JsonPropertyName("name")]
    public string? Name { get; set; }
    [JsonPropertyName("path")]
    public string? Path { get; set; }
}