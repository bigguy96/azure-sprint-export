using System.Text.Json.Serialization;

namespace PowerPointConsoleApp.Models;

public class Fields
{
    [JsonPropertyName("System.Title")]
    public string? SystemTitle { get; set; }

    [JsonPropertyName("System.Description")]
    public string? SystemDescription { get; set; }

    [JsonPropertyName("System.WorkItemType")]
    public string? SystemWorkItemType { get; set; }

    [JsonPropertyName("System.Parent")]
    public int? SystemParent { get; set; }
}
