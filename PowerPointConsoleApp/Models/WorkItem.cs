using System.Text.Json.Serialization;

namespace PowerPointConsoleApp.Models;

public class WorkItem
{
    [JsonPropertyName("id")]
    public int Id { get; set; }

    [JsonPropertyName("title")]
    public string? Title { get; set; }

    [JsonPropertyName("description")]
    public string? Description { get; set; }

    [JsonPropertyName("workItemType")]
    public string? WorkItemType { get; set; }

    [JsonPropertyName("parentId")]
    public int? ParentId { get; set; }
}
