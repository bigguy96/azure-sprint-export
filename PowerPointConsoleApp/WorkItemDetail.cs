using System.Text.Json.Serialization;

public class WorkItemDetail
{
    [JsonPropertyName("id")]
    public int Id { get; set; }
    [JsonPropertyName("fields")]
    public Fields? Fields { get; set; }
}