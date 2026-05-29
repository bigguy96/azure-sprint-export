using System.Text.Json.Serialization;

namespace PowerPointConsoleApp.Models;

public class WorkItemRef
{
    [JsonPropertyName("id")]
    public int Id { get; set; }
}
