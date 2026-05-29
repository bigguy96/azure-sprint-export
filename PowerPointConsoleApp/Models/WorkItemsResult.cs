using System.Text.Json.Serialization;

namespace PowerPointConsoleApp.Models;

public record WorkItemsResult([property: JsonPropertyName("value")] WorkItemDetail[]? Value);
