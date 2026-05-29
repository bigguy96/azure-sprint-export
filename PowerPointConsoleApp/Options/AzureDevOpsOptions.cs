namespace PowerPointConsoleApp.Options;

// Strongly-typed binding for the AzDo:* configuration section.
// Values are loaded from User Secrets, environment variables, or Windows Credential Manager.
public sealed class AzureDevOpsOptions
{
    public const string SectionName = "AzDo";

    // Azure DevOps organisation name as it appears in the URL (e.g. "mycompany").
    public string Org { get; init; } = string.Empty;

    // Azure DevOps project name (e.g. "MyProject").
    public string Project { get; init; } = string.Empty;

    // Azure DevOps team name (e.g. "MyProject Team").
    public string Team { get; init; } = string.Empty;

    // Personal Access Token with read access to Work Items and Iterations.
    public string Token { get; init; } = string.Empty;
}
