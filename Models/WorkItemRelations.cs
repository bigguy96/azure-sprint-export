using System.Text.Json.Serialization;
using System.Collections.Generic;

namespace SprintItemsApp.Models
{
    public class WorkItemRelations
    {
        [JsonPropertyName("workItemRelations")]
        public List<SprintWorkItemRelation> Relations { get; set; }
    }
}