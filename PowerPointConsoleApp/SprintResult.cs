using System.Text.Json.Serialization;

public class SprintResult
{
    [JsonPropertyName("value")]
    public Sprint[]? Value { get; set; }
}