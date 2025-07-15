using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace SprintItemsApp.Models
{
    public class SprintList
    {
        [JsonPropertyName("value")]
        public List<Sprint> Value { get; set; }
        [JsonPropertyName("count")]
        public int Count { get; set; }
    }
}