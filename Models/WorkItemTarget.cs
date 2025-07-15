using System.Text.Json.Serialization;

namespace SprintItemsApp.Models
{
    public class WorkItemTarget
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }
        [JsonPropertyName("url")]
        public string Url { get; set; }
    }
}