using PowerPointConsoleApp.Models;

namespace PowerPointConsoleApp.Services;

// Defines the contract for retrieving data from the Azure DevOps REST API.
public interface IAzureDevOpsService
{
    // Returns all sprint iterations for the configured team, or null when deserialization fails.
    Task<Sprint[]?> GetSprintsAsync(string? teamName);

    // Returns the work items relevant to the given sprint iteration path.
    // Tasks are resolved to their parent PBIs/Bugs; tasks themselves are excluded from the result.
    Task<WorkItem[]?> GetWorkItemsForSprintAsync(string? sprintPath);
}
