using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;

namespace AzureDevOpsPptWpf.Services;

public class AzureDevOpsService
{
    private readonly HttpClient _httpClient;

    public AzureDevOpsService(string personalAccessToken)
    {
        _httpClient = new HttpClient();

        var authToken = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($":{personalAccessToken}"));
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);
    }

    // Get teams for a project (to populate team dropdown)
    public async Task<List<Team>> GetTeamsAsync(string organization, string project)
    {
        string url = $"https://dev.azure.com/{organization}/_apis/projects/{project}/teams?api-version=7.0";

        var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();

        var teamsResponse = JsonSerializer.Deserialize<TeamsResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        return teamsResponse?.Value ?? new List<Team>();
    }

    // Get iterations (sprints) for a specific team
    public async Task<List<Iteration>> GetSprintsAsync(string organization, string project, string teamName)
    {
        var url = $"https://dev.azure.com/{organization}/{project}/{teamName}/_apis/work/teamsettings/iterations?api-version=7.0";

        var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();

        var iterationsResponse = JsonSerializer.Deserialize<IterationsResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        return iterationsResponse?.Value ?? new List<Iteration>();
    }

    // Get work item IDs in sprint for team — only Product Backlog Item and Bug
    public async Task<List<int>> GetWorkItemIdsInSprintAsync(string organization, string project, string teamName, string iterationPath)
    {
        var wiql = new
        {
            query = $@"
                SELECT [System.Id]
                FROM WorkItems
                WHERE
                    [System.TeamProject] = '{project}'
                    AND [System.IterationPath] UNDER '{iterationPath}'
                    AND [System.WorkItemType] IN ('Product Backlog Item', 'Bug')
                    AND [System.State] <> 'Removed'"
        };

        var url = $"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql?api-version=7.0";

        var content = new StringContent(JsonSerializer.Serialize(wiql), System.Text.Encoding.UTF8, "application/json");
        var response = await _httpClient.PostAsync(url, content);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();

        var wiqlResponse = JsonSerializer.Deserialize<WiqlResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        var ids = new List<int>();
        if (wiqlResponse?.WorkItems != null)
        {
            foreach (var wi in wiqlResponse.WorkItems)
                ids.Add(wi.Id);
        }

        return ids;
    }

    // Batch get work item details including description, title, assigned to, state
    public async Task<List<WorkItem>> GetWorkItemsAsync(string organization, List<int> ids)
    {
        if (ids == null || ids.Count == 0)
            return new List<WorkItem>();

        var idsString = string.Join(",", ids);
        var fields = new[]
        {
            "System.Id",
            "System.Title",
            "System.Description",
            "System.AssignedTo",
            "System.State",
            "System.WorkItemType",
            "System.IterationPath"
        };

        var url = $"https://dev.azure.com/{organization}/_apis/wit/workitems?ids={idsString}&fields={string.Join(",", fields)}&api-version=7.0";

        var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();

        var workItemsResponse = JsonSerializer.Deserialize<WorkItemsResponse>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

        return workItemsResponse?.Value ?? new List<WorkItem>();
    }


    // Data models

    public class TeamsResponse
    {
        public int Count { get; set; }
        public List<Team> Value { get; set; }
    }
    public class Team
    {
        public string Id { get; set; }
        public string Name { get; set; }
    }

    public class IterationsResponse
    {
        public int Count { get; set; }
        public List<Iteration> Value { get; set; }
    }
    public class Iteration
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Path { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? FinishDate { get; set; }
        public string TimeFrame { get; set; }
    }

    public class WiqlResponse
    {
        public List<WorkItemReference> WorkItems { get; set; }
    }
    public class WorkItemReference
    {
        public int Id { get; set; }
        public string Url { get; set; }
    }

    public class WorkItemsResponse
    {
        public int Count { get; set; }
        public List<WorkItem> Value { get; set; }
    }
    public class WorkItem
    {
        public int Id { get; set; }
        public Dictionary<string, object> Fields { get; set; }
        public string Url { get; set; }

        // Helpers for common fields
        public string Title => Fields.TryGetValue("System.Title", out var title) ? title?.ToString() : string.Empty;
        public string Description => Fields.TryGetValue("System.Description", out var desc) ? desc?.ToString() : string.Empty;
        public string AssignedTo
        {
            get
            {
                if (Fields.TryGetValue("System.AssignedTo", out var assigned))
                {
                    // AssignedTo can be complex object
                    if (assigned is JsonElement elem && elem.ValueKind == JsonValueKind.Object)
                    {
                        if (elem.TryGetProperty("displayName", out var displayName))
                            return displayName.GetString();
                    }
                    else
                    {
                        return assigned.ToString();
                    }
                }
                return "";
            }
        }
        public string State => Fields.TryGetValue("System.State", out var state) ? state?.ToString() : string.Empty;
        public string WorkItemType => Fields.TryGetValue("System.WorkItemType", out var type) ? type?.ToString() : string.Empty;
    }
}