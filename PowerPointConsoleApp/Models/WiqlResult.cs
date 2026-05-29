using System.Text.Json.Serialization;

namespace PowerPointConsoleApp.Models;

public record WiqlResult([property: JsonPropertyName("workItems")] WorkItemRef[]? WorkItems);
