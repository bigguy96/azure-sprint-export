using System.Text.Json.Serialization;

namespace PowerPointConsoleApp.Models;

public record SprintResult([property: JsonPropertyName("value")] Sprint[]? Value);
