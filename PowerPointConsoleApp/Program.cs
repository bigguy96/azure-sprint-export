using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.RegularExpressions;
using Microsoft.Extensions.Configuration;
using PowerPointConsoleApp;
using A = DocumentFormat.OpenXml.Drawing;

var docsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
var templatePath = Path.Combine(docsPath, "template.pptx"); // Your PowerPoint template file path
var outputPath = Path.Combine(docsPath, "SprintWorkItems.pptx");
var config = new ConfigurationBuilder()
    .AddUserSecrets(typeof(Program).Assembly, optional: true)
    .AddEnvironmentVariables(prefix: "AZDO_")
    .AddWindowsCredentialManager(["AzDo:Org", "AzDo:Project", "AzDo:Team", "AzDo:Token"])
    .Build();
var organization = config["AzDo:Org"];
var project = config["AzDo:Project"];
var team = config["AzDo:Team"];
var pat = config["AzDo:Token"];

await Main();
return;

async Task Main()
{
    Console.WriteLine("Fetching sprints...");
    var sprints = await GetSprints(team);
    if (sprints is { Length: 0 })
    {
        Console.WriteLine("No sprints found.");
        return;
    }

    Console.WriteLine("Available sprints:");
    for (var i = 0; i < sprints!.Length; i++)
        Console.WriteLine($"{i + 1}. {sprints[i].Name}");

    Console.Write("Select a sprint by number: ");
    if (!int.TryParse(Console.ReadLine(), out var sel) || sel < 1 || sel > sprints.Length)
    {
        Console.WriteLine("Invalid selection.");
        return;
    }

    var sprint = sprints[sel - 1];
    Console.WriteLine($"Selected sprint: {sprint.Name}");

    Console.WriteLine("Fetching work items...");
    var workItems = await GetWorkItemsForSprint(sprint.Path);

    if (workItems is { Length: 0 })
    {
        Console.WriteLine("No work items found in this sprint.");
        return;
    }

    // Copy template to output file
    File.Copy(templatePath, outputPath, true);

    using var ppt = PresentationDocument.Open(outputPath, true);
    var presentationPart = ppt.PresentationPart;
    var presentation = presentationPart?.Presentation;

    var slideIdList = presentation!.SlideIdList ??= new SlideIdList();
    var maxSlideId = slideIdList.ChildElements.OfType<SlideId>().Select(s => s.Id.Value).DefaultIfEmpty((uint)255).Max();

    // Use first slide as template
    var sourceSlidePart = presentationPart!.SlideParts.First();
    var sourceLayoutPart = sourceSlidePart.SlideLayoutPart;

    foreach (var item in workItems!)
    {
        var newSlidePart = presentationPart.AddNewPart<SlidePart>();
        // Clone slide content from template slide
        sourceSlidePart.Slide.Save(newSlidePart.GetStream(FileMode.Create));

        // Add layout part to new slide
        newSlidePart.AddPart(sourceLayoutPart);

        // Replace placeholders
        ReplacePlaceholderTextByOrder(newSlidePart, true, $"{item.Id} - {item.Title}");
        var desc = string.IsNullOrWhiteSpace(item.Description) ? "(No Description)" : StripHtmlTags(item.Description);
        ReplacePlaceholderTextByOrder(newSlidePart, false, desc);

        maxSlideId++;
        var relId = presentationPart.GetIdOfPart(newSlidePart);
        slideIdList.AppendChild(new SlideId { Id = maxSlideId, RelationshipId = relId });
    }

    presentation.Save();
    Console.WriteLine($"Presentation created at: {outputPath}");
}

static void ReplacePlaceholderTextByOrder(SlidePart slidePart, bool isTitle, string text)
{
    var shapesWithText = slidePart.Slide.Descendants<Shape>()
        .Where(s => s.TextBody != null)
        .ToList();

    Shape? shape;

    if (isTitle)
    {
        shape = shapesWithText.FirstOrDefault();
    }
    else
    {
        shape = shapesWithText.Skip(1).FirstOrDefault();
    }

    var textBody = shape?.TextBody;
    if (textBody == null)
        return;

    textBody.RemoveAllChildren<A.Paragraph>();

    var para = new A.Paragraph();
    foreach (var line in text.Split(['\n'], StringSplitOptions.None))
    {
        var run = new A.Run(
            new A.RunProperties { FontSize = 2400, Language = "en-US", Dirty = false }
        );
        run.RunProperties?.AppendChild(new A.LatinFont() { Typeface = "Garamond" });
        run.AppendChild(new A.Text(line));
        para.AppendChild(run);
        para.AppendChild(new A.Break());
    }
    textBody.AppendChild(para);
}

static string StripHtmlTags(string source)
{
    if (string.IsNullOrEmpty(source))
        return string.Empty;

    if (source.Contains("&nbsp;"))
    {
        source = source.Replace("&nbsp;", " ");
    }

    return Regex.Replace(source, "<.*?>", string.Empty);
}

async Task<Sprint[]?> GetSprints(string? teamName)
{
    using var client = CreateHttpClient();

    var url = $"https://dev.azure.com/{organization}/{project}/{teamName}/_apis/work/teamsettings/iterations?api-version=7.0";

    var response = await client.GetAsync(url);
    response.EnsureSuccessStatusCode();

    var json = await response.Content.ReadAsStringAsync();
    var sprintResult = JsonSerializer.Deserialize<SprintResult>(json);

    return sprintResult?.Value;
}

async Task<WorkItem[]?> GetWorkItemsForSprint(string? sprintPath)
{
    using var client = CreateHttpClient();

    // 1. WIQL query to get work item IDs (only Product Backlog Item and Bug)
    var wiql = new
    {
        query = $@"
                SELECT [System.Id]
                FROM WorkItems
                WHERE [System.IterationPath] = '{sprintPath?.Replace("'", "''")}'
                  AND [System.WorkItemType] IN ('Product Backlog Item', 'Bug')
                ORDER BY [System.Id]"
    };

    var wiqlContent = new StringContent(JsonSerializer.Serialize(wiql));
    wiqlContent.Headers.ContentType = new MediaTypeHeaderValue("application/json");

    var wiqlUrl = $"https://dev.azure.com/{organization}/{project}/_apis/wit/wiql?api-version=7.0";

    var wiqlResponse = await client.PostAsync(wiqlUrl, wiqlContent);
    wiqlResponse.EnsureSuccessStatusCode();

    var wiqlJson = await wiqlResponse.Content.ReadAsStringAsync();
    var wiqlResult = JsonSerializer.Deserialize<WiqlResult>(wiqlJson);
    var ids = (wiqlResult?.WorkItems ?? []).Select(wi => wi.Id).ToArray();

    if (ids is { Length: 0 })
        return [];

    // 2. Batch request to get work item details including Description
    var idsStr = string.Join(",", ids!);
    var workItemsUrl = $"https://dev.azure.com/{organization}/{project}/_apis/wit/workitems?ids={idsStr}&api-version=7.0&$expand=Fields";

    var workItemsResponse = await client.GetAsync(workItemsUrl);
    workItemsResponse.EnsureSuccessStatusCode();

    var workItemsJson = await workItemsResponse.Content.ReadAsStringAsync();
    var workItemsResult = JsonSerializer.Deserialize<WorkItemsResult>(workItemsJson);

    return (workItemsResult?.Value ?? []).Select(wi => new WorkItem
    {
        Id = wi.Id,
        Title = wi.Fields?.SystemTitle,
        Description = wi.Fields?.SystemDescription
    }).ToArray();
}

HttpClient CreateHttpClient()
{
    var client = new HttpClient();
    var authToken = Convert.ToBase64String(System.Text.Encoding.ASCII.GetBytes($":{pat}"));
    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);
    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
    return client;
}