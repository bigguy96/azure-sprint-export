using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace SprintItemsApp.Models;

public class WorkItemRelation
{
    [JsonPropertyName("rel")]
    public string RelationType { get; set; }

    [JsonPropertyName("url")]
    public string Url { get; set; }

    [JsonPropertyName("attributes")]
    public Dictionary<string, object> Attributes { get; set; }

    // Extract work item ID from URL
    public int? TargetId
    {
        get
        {
            if (string.IsNullOrEmpty(Url))
                return null;
            var segments = Url.Split('/');
            return int.TryParse(segments.Last(), out var id) ? id : (int?)null;
        }
    }
}