using System;
using System.Text.Json.Serialization;

namespace SprintItemsApp.Models;

public class Sprint
{
    [JsonPropertyName("id")]
    public string Id { get; set; }
    [JsonPropertyName("name")]
    public string Name { get; set; }
    [JsonPropertyName("path")]
    public string Path { get; set; }
    [JsonPropertyName("attributes")]
    public SprintAttributes Attributes { get; set; }
}

public class SprintAttributes
{
    [JsonPropertyName("startDate")]
    public DateTime? StartDate { get; set; }
    [JsonPropertyName("finishDate")]
    public DateTime? FinishDate { get; set; }
}