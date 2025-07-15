using System.Net.Http;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using SprintItemsApp.Models;
using System;

namespace SprintItemsApp.Services
{
    public class AzureDevOpsService
    {
        private readonly string _bearerToken;
        private readonly string _organization;
        private readonly string _project;
        private readonly string _team;

        public AzureDevOpsService(string bearerToken, string organization, string project, string team)
        {
            _bearerToken = bearerToken ?? throw new ArgumentNullException(nameof(bearerToken));
            _organization = NormalizeName(organization) ?? throw new ArgumentNullException(nameof(organization));
            _project = NormalizeName(project) ?? throw new ArgumentNullException(nameof(project));
            _team = NormalizeName(team) ?? throw new ArgumentNullException(nameof(team));
            System.Diagnostics.Debug.WriteLine($"Normalized Azure DevOps Config: Organization={_organization}, Project={_project}, Team={_team}");
        }

        private static string NormalizeName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;
            return string.Join("-", name.Split(' ', StringSplitOptions.RemoveEmptyEntries)).Trim();
        }

        public async Task<List<Sprint>> GetSprintsAsync()
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(
                "Basic", Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($":{_bearerToken}")));

            var url = $"https://dev.azure.com/{_organization}/{_project}/{_team}/_apis/work/teamsettings/iterations?api-version=7.0";
            System.Diagnostics.Debug.WriteLine($"Sprints URL: {url}");
            var response = await client.GetAsync(url);

            if (response.IsSuccessStatusCode)
            {
                var json = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine("Sprints JSON: " + json);
                var sprintList = JsonSerializer.Deserialize<SprintList>(json);
                return sprintList.Value ?? new List<Sprint>();
            }
            else
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"Sprints API Error: {response.StatusCode}, URL: {url}, Content: {errorContent}");
                var fallbackUrl = $"https://dev.azure.com/{_organization}/{_project}/_apis/work/iterations?api-version=7.0";
                System.Diagnostics.Debug.WriteLine($"Trying Fallback Sprints URL: {fallbackUrl}");
                var fallbackResponse = await client.GetAsync(fallbackUrl);

                if (fallbackResponse.IsSuccessStatusCode)
                {
                    var json = await fallbackResponse.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine("Fallback Sprints JSON: " + json);
                    var sprintList = JsonSerializer.Deserialize<SprintList>(json);
                    return sprintList.Value ?? new List<Sprint>();
                }
                else
                {
                    var fallbackErrorContent = await fallbackResponse.Content.ReadAsStringAsync();
                    System.Diagnostics.Debug.WriteLine($"Fallback Sprints API Error: {fallbackResponse.StatusCode}, URL: {fallbackUrl}, Content: {fallbackErrorContent}");
                    throw new Exception($"Failed to retrieve sprints: {response.ReasonPhrase}, Content: {errorContent}; Fallback failed: {fallbackResponse.ReasonPhrase}, Content: {fallbackErrorContent}");
                }
            }
        }

        public async Task<List<WorkItem>> GetWorkItemsForSprintAsync(string iterationId)
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(
                "Basic", Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($":{_bearerToken}")));

            var url = $"https://dev.azure.com/{_organization}/{_project}/{_team}/_apis/work/teamsettings/iterations/{iterationId}/workitems?api-version=7.0";
            System.Diagnostics.Debug.WriteLine($"Work Items URL: {url}");
            var response = await client.GetAsync(url);

            if (!response.IsSuccessStatusCode)
            {
                var errorContent = await response.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"Work Item Relations API Error: {response.StatusCode}, Content: {errorContent}");
                throw new Exception($"Failed to retrieve work items: {response.ReasonPhrase}, Content: {errorContent}");
            }

            var json = await response.Content.ReadAsStringAsync();
            System.Diagnostics.Debug.WriteLine("Work Item Relations JSON: " + json);
            var relations = JsonSerializer.Deserialize<WorkItemRelations>(json);
            var workItemIds = relations.Relations.Select(r => r.Target.Id).ToList();

            if (!workItemIds.Any())
            {
                System.Diagnostics.Debug.WriteLine("No work item IDs found for the sprint.");
                return new List<WorkItem>();
            }

            var ids = string.Join(",", workItemIds);
            var detailsUrl = $"https://dev.azure.com/{_organization}/{_project}/_apis/wit/workitems?ids={ids}&$expand=relations&api-version=7.0";
            System.Diagnostics.Debug.WriteLine($"Work Item Details URL: {detailsUrl}");
            var detailsResponse = await client.GetAsync(detailsUrl);

            if (detailsResponse.IsSuccessStatusCode)
            {
                var detailsJson = await detailsResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine("Work Items JSON: " + detailsJson);
                var workItemsList = JsonSerializer.Deserialize<WorkItemList>(detailsJson);
                var workItems = workItemsList.Value ?? new List<WorkItem>();

                // Log work item counts for debugging
                System.Diagnostics.Debug.WriteLine($"Work Items Count: {workItems.Count}");
                foreach (var workItem in workItems)
                {
                    var childIds = workItem.Relations
                        .Where(r => r.RelationType == "System.LinkTypes.Hierarchy-Forward")
                        .Select(r => r.TargetId)
                        .Where(id => id.HasValue)
                        .Select(id => id.Value)
                        .ToList();
                    workItem.Children = workItems.Where(wi => childIds.Contains(wi.Id)).ToList();
                    System.Diagnostics.Debug.WriteLine($"WorkItem ID {workItem.Id}: Children Count = {workItem.Children.Count}");
                }

                var topLevelWorkItems = workItems.Where(wi => !wi.ParentId.HasValue || !workItemIds.Contains(wi.ParentId.Value)).ToList();
                System.Diagnostics.Debug.WriteLine($"Top-Level Work Items Count: {topLevelWorkItems.Count}");
                return topLevelWorkItems;
            }
            else
            {
                var errorContent = await detailsResponse.Content.ReadAsStringAsync();
                System.Diagnostics.Debug.WriteLine($"Work Items API Error: {detailsResponse.StatusCode}, Content: {errorContent}");
                throw new Exception($"Failed to retrieve work item details: {detailsResponse.ReasonPhrase}, Content: {errorContent}");
            }
        }
    }
}