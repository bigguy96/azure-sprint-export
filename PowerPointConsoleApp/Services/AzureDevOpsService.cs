using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.Extensions.Options;
using PowerPointConsoleApp.Models;
using PowerPointConsoleApp.Options;

namespace PowerPointConsoleApp.Services;

// Retrieves sprint and work item data from the Azure DevOps REST API.
// The HttpClient is supplied by IHttpClientFactory (named client "AzDo") and is pre-configured
// with the Basic auth header and Accept: application/json in Program.cs.
public sealed class AzureDevOpsService(
    IHttpClientFactory httpClientFactory,
    IOptions<AzureDevOpsOptions> options) : IAzureDevOpsService
{
    private readonly AzureDevOpsOptions _options = options.Value;

    // Returns all sprint iterations for the configured team, or null when deserialization fails.
    public async Task<Sprint[]?> GetSprintsAsync(string? teamName)
    {
        // Azure DevOps Iterations REST endpoint for the configured organisation, project, and team.
        var url = $"https://dev.azure.com/{_options.Org}/{_options.Project}/{teamName}/_apis/work/teamsettings/iterations?api-version=7.0";

        var client = httpClientFactory.CreateClient("AzDo");
        var response = await client.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var json = await response.Content.ReadAsStringAsync();
        return JsonSerializer.Deserialize<SprintResult>(json)?.Value;
    }

    // Returns the work items relevant to the given sprint iteration path.
    // Performs five steps (see comments inline):
    //   1. Fetch PBIs and Bugs assigned directly to the sprint.
    //   2. Fetch Tasks assigned to the sprint.
    //   3. Resolve each task's parent to surface items from earlier sprints.
    //   4. Union sprint items with any missing parents.
    //   5. Batch-fetch full field details for all collected IDs.
    // Tasks themselves are excluded — only parent items get slides.
    public async Task<WorkItem[]?> GetWorkItemsForSprintAsync(string? sprintPath)
    {
        if (string.IsNullOrWhiteSpace(sprintPath))
            return [];

        // Escape single quotes in the path so the WIQL string literal is valid.
        var escapedPath = sprintPath.Replace("'", "''");

        // WIQL endpoint – accepts a query in the request body and returns matching work item IDs.
        var wiqlUrl = $"https://dev.azure.com/{_options.Org}/{_options.Project}/_apis/wit/wiql?api-version=7.0";
        var client = httpClientFactory.CreateClient("AzDo");

        // 1. Fetch PBIs and Bugs directly assigned to the sprint.
        var sprintItemIds = await RunWiqlAsync(client, wiqlUrl, $@"
            SELECT [System.Id]
            FROM WorkItems
            WHERE [System.IterationPath] = '{escapedPath}'
              AND [System.WorkItemType] IN ('Product Backlog Item', 'Bug')
            ORDER BY [System.Id]");

        // Store IDs in a HashSet for O(1) membership checks when filtering missing parents.
        var sprintItemIdSet = sprintItemIds.ToHashSet();

        // 2. Fetch Tasks assigned to the sprint.
        var taskIds = await RunWiqlAsync(client, wiqlUrl, $@"
            SELECT [System.Id]
            FROM WorkItems
            WHERE [System.IterationPath] = '{escapedPath}'
              AND [System.WorkItemType] = 'Task'
            ORDER BY [System.Id]");

        // 3. Fetch task details to resolve their parent IDs.
        var missingParentIds = new HashSet<int>();
        if (taskIds.Length > 0)
        {
            // Batch-request only System.Id and System.Parent to keep the payload small.
            var taskIdsStr = string.Join(",", taskIds);
            var taskDetailsUrl = $"https://dev.azure.com/{_options.Org}/{_options.Project}/_apis/wit/workitems?ids={taskIdsStr}&api-version=7.0&fields=System.Id,System.Parent";
            var taskDetailsResponse = await client.GetAsync(taskDetailsUrl);
            taskDetailsResponse.EnsureSuccessStatusCode();
            var taskDetailsResult = JsonSerializer.Deserialize<WorkItemsResult>(await taskDetailsResponse.Content.ReadAsStringAsync());

            foreach (var task in taskDetailsResult?.Value ?? [])
            {
                var parentId = task.Fields?.SystemParent;
                // Add the parent only if it is not already in the current sprint.
                if (parentId.HasValue && !sprintItemIdSet.Contains(parentId.Value))
                    missingParentIds.Add(parentId.Value);
            }
        }

        // 4. Combine sprint PBIs/Bugs + parents from previous sprints.
        var allIds = sprintItemIdSet.Union(missingParentIds).ToArray();
        if (allIds.Length == 0)
            return [];

        // 5. Batch-fetch full details for all items.
        // $expand=Fields returns every field value, including Description, not available from WIQL alone.
        var idsStr = string.Join(",", allIds);
        var workItemsUrl = $"https://dev.azure.com/{_options.Org}/{_options.Project}/_apis/wit/workitems?ids={idsStr}&api-version=7.0&$expand=Fields";
        var workItemsResponse = await client.GetAsync(workItemsUrl);
        workItemsResponse.EnsureSuccessStatusCode();

        var workItemsResult = JsonSerializer.Deserialize<WorkItemsResult>(await workItemsResponse.Content.ReadAsStringAsync());

        // Project the raw API response into the application model, ordered by ID for consistent slide ordering.
        return (workItemsResult?.Value ?? [])
            .OrderBy(wi => wi.Id)
            .Select(wi => new WorkItem
            {
                Id = wi.Id,
                Title = wi.Fields?.SystemTitle,
                Description = wi.Fields?.SystemDescription,
                WorkItemType = wi.Fields?.SystemWorkItemType,
                ParentId = wi.Fields?.SystemParent
            }).ToArray();
    }

    // Executes a WIQL query and returns the array of matching work item IDs.
    private static async Task<int[]> RunWiqlAsync(HttpClient client, string wiqlUrl, string query)
    {
        var content = new StringContent(JsonSerializer.Serialize(new { query }));
        content.Headers.ContentType = new MediaTypeHeaderValue("application/json");

        var response = await client.PostAsync(wiqlUrl, content);
        response.EnsureSuccessStatusCode();

        var result = JsonSerializer.Deserialize<WiqlResult>(await response.Content.ReadAsStringAsync());
        return (result?.WorkItems ?? []).Select(wi => wi.Id).ToArray();
    }
}
