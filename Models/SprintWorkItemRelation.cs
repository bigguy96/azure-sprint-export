using System.Text.Json.Serialization;

namespace SprintItemsApp.Models
{
    public class SprintWorkItemRelation
    {
        [JsonPropertyName("rel")]
        public string RelationType { get; set; }

        [JsonPropertyName("source")]
        public WorkItemTarget Source { get; set; }

        [JsonPropertyName("target")]
        public WorkItemTarget Target { get; set; }
    }
}