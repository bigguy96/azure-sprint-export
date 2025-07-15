using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace SprintItemsApp.Models;

public class WorkItemList
{
    [JsonPropertyName("value")]
    public List<WorkItem> Value { get; set; }
}