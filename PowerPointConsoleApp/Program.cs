using System.Net.Http.Headers;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using PowerPointConsoleApp.Infrastructure;
using PowerPointConsoleApp.Options;
using PowerPointConsoleApp.Services;

// Build the generic host, which wires up configuration, DI, and the HTTP client factory.
// Host.CreateDefaultBuilder already adds:
//   - appsettings.json / appsettings.{env}.json
//   - User Secrets (in Development)
//   - Environment variables
// We extend it with the Windows Credential Manager source.
var host = Host.CreateDefaultBuilder(args)
    .ConfigureAppConfiguration((_, config) =>
    {
        // Always load User Secrets regardless of environment (CreateDefaultBuilder only loads them in Development).
        config.AddUserSecrets<Program>();

        // Supplement the default sources with Windows Credential Manager for secure on-machine PAT storage.
        config.AddWindowsCredentialManager(["AzDo:Org", "AzDo:Project", "AzDo:Team", "AzDo:Token"]);
    })
    .ConfigureServices((ctx, services) =>
    {
        // Bind strongly-typed options from configuration.
        services.Configure<AzureDevOpsOptions>(ctx.Configuration.GetSection(AzureDevOpsOptions.SectionName));
        services.Configure<AppOptions>(ctx.Configuration.GetSection(AppOptions.SectionName));

        // Register the named HttpClient "AzDo" with Basic auth and JSON Accept headers.
        // IHttpClientFactory manages the underlying handler lifetime, avoiding socket exhaustion.
        services.AddHttpClient("AzDo", (sp, client) =>
        {
            var opts = sp.GetRequiredService<IOptions<AzureDevOpsOptions>>().Value;

            // Fail fast if the PAT is missing — every API call would fail without it.
            ArgumentException.ThrowIfNullOrWhiteSpace(opts.Token, nameof(opts.Token));

            // Azure DevOps Basic auth: base64-encode ":<PAT>" (username is intentionally empty).
            var authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($":{opts.Token}"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        });

        // Register application services against their interfaces for easy testing/swapping.
        services.AddSingleton<IAzureDevOpsService, AzureDevOpsService>();
        services.AddSingleton<IPowerPointService, PowerPointService>();
    })
    .Build();

// Resolve services from the DI container.
var devOps = host.Services.GetRequiredService<IAzureDevOpsService>();
var pptService = host.Services.GetRequiredService<IPowerPointService>();
var azDoOptions = host.Services.GetRequiredService<IOptions<AzureDevOpsOptions>>().Value;

// --- Orchestration ---
Console.WriteLine("Fetching sprints...");
var sprints = await devOps.GetSprintsAsync(azDoOptions.Team);

// Abort early if the API returned no sprints for this team.
if (sprints is null or { Length: 0 })
{
    Console.WriteLine("No sprints found.");
    return;
}

// Display a numbered list of available sprints so the user can pick one.
Console.WriteLine("Available sprints:");
for (var i = 0; i < sprints.Length; i++)
    Console.WriteLine($"{i + 1}. {sprints[i].Name}");

// Read and validate the user's selection (must be a number within the displayed range).
Console.Write("Select a sprint by number: ");
if (!int.TryParse(Console.ReadLine(), out var sel) || sel < 1 || sel > sprints.Length)
{
    Console.WriteLine("Invalid selection.");
    return;
}

// Convert the 1-based user input to a 0-based array index.
var sprint = sprints[sel - 1];
Console.WriteLine($"Selected sprint: {sprint.Name}");

Console.WriteLine("Fetching work items...");
var workItems = await devOps.GetWorkItemsForSprintAsync(sprint.Path);

// Nothing to generate if the sprint has no relevant work items.
if (workItems is null or { Length: 0 })
{
    Console.WriteLine("No work items found in this sprint.");
    return;
}

// Delegate slide generation to the PowerPoint service.
var outputPath = await pptService.GenerateAsync(workItems);
Console.WriteLine($"Presentation created at: {outputPath}");