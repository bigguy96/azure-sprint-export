using System.Text.Json.Serialization;

public class WorkItemsResult
{
    [JsonPropertyName("value")]
    public WorkItemDetail[]? Value { get; set; }
}