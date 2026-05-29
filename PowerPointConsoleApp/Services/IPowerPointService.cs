using PowerPointConsoleApp.Models;

namespace PowerPointConsoleApp.Services;

// Defines the contract for generating a PowerPoint presentation from a set of work items.
public interface IPowerPointService
{
    // Generates a .pptx file at the configured output path.
    // One slide is created per work item, using the configured template file as the visual prototype.
    // Returns the full path of the generated file.
    Task<string> GenerateAsync(WorkItem[] workItems);
}
