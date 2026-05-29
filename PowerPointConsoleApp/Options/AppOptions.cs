namespace PowerPointConsoleApp.Options;

// Strongly-typed binding for file-path settings used by the PowerPoint generator.
public sealed class AppOptions
{
    public const string SectionName = "App";

    // Full path to the .pptx template file. Defaults to template.pptx in the user's Documents folder.
    public string TemplatePath { get; init; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "template.pptx");

    // Full path where the generated presentation will be written.
    public string OutputPath { get; init; } =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "SprintWorkItems.pptx");
}
