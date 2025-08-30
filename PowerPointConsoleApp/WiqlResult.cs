using System.Text.Json.Serialization;

public class WiqlResult
{
    [JsonPropertyName("workItems")]
    public WorkItemRef[]? WorkItems { get; set; }
}