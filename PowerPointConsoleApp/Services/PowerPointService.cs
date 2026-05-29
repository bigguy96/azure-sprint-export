using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Extensions.Options;
using PowerPointConsoleApp.Models;
using PowerPointConsoleApp.Options;
using System.Text.RegularExpressions;
using A = DocumentFormat.OpenXml.Drawing;

namespace PowerPointConsoleApp.Services;

// Generates a PowerPoint presentation from a list of Azure DevOps work items.
// Each work item produces one slide cloned from the first slide of the template file.
public sealed partial class PowerPointService(IOptions<AppOptions> options) : IPowerPointService
{
    private readonly AppOptions _options = options.Value;

    // Generates a .pptx file at the configured output path.
    // One slide is created per work item, using the configured template file as the visual prototype.
    // Returns the full path of the generated file.
    public Task<string> GenerateAsync(WorkItem[] workItems)
    {
        // Copy the template so the original is never modified. Overwrites any existing output file.
        File.Copy(_options.TemplatePath, _options.OutputPath, overwrite: true);

        using var ppt = PresentationDocument.Open(_options.OutputPath, isEditable: true);
        var presentationPart = ppt.PresentationPart
            ?? throw new InvalidOperationException("Template has no PresentationPart.");
        var presentation = presentationPart.Presentation
            ?? throw new InvalidOperationException("Template has no Presentation element.");

        // Ensure the SlideIdList element exists; create it if the template omitted it.
        var slideIdList = presentation.SlideIdList ??= new SlideIdList();

        // Find the highest existing slide ID so new slides get unique, incrementing IDs.
        // The Open XML spec requires IDs >= 256, so 255 is the safe floor when no slides exist yet.
        var maxSlideId = slideIdList.ChildElements
            .OfType<SlideId>()
            .Select(s => s.Id!.Value)
            .DefaultIfEmpty((uint)255)
            .Max();

        // Use the first slide in the template as the visual prototype for all generated slides.
        var sourceSlidePart = presentationPart.SlideParts.First();
        var sourceLayoutPart = sourceSlidePart.SlideLayoutPart
            ?? throw new InvalidOperationException("Template slide has no SlideLayoutPart.");

        // Generate one slide per work item.
        foreach (var item in workItems)
        {
            // Add a blank slide part to the presentation package.
            var newSlidePart = presentationPart.AddNewPart<SlidePart>();

            // Stamp the template slide's XML onto the new part, preserving all formatting.
            sourceSlidePart.Slide.Save(newSlidePart.GetStream(FileMode.Create));

            // Link the same layout so the new slide inherits the template's theme and placeholders.
            newSlidePart.AddPart(sourceLayoutPart);

            // Write the work item type, ID, and title into the title placeholder (first text shape).
            ReplacePlaceholderText(newSlidePart, isTitle: true,
                $"{item.WorkItemType} {item.Id} - {item.Title}");

            // Strip HTML from the description (Azure DevOps stores it as HTML).
            // Fall back to a placeholder string when the description is empty.
            var desc = string.IsNullOrWhiteSpace(item.Description)
                ? "(No Description)"
                : StripHtmlTags(item.Description);
            ReplacePlaceholderText(newSlidePart, isTitle: false, desc);

            // Register the new slide in the presentation's slide list with a unique ID.
            maxSlideId++;
            var relId = presentationPart.GetIdOfPart(newSlidePart);
            slideIdList.AppendChild(new SlideId { Id = maxSlideId, RelationshipId = relId });
        }

        // Flush all changes to the Open XML package.
        presentation.Save();

        return Task.FromResult(_options.OutputPath);
    }

    // Replaces text in either the title or body placeholder of a slide.
    // Shapes are selected by position: the first text shape is the title, the second is the body.
    // Each line of 'text' becomes a separate run followed by a line break.
    private static void ReplacePlaceholderText(SlidePart slidePart, bool isTitle, string text)
    {
        // Collect all shapes that have a text body, in document order.
        var shapesWithText = slidePart.Slide.Descendants<Shape>()
            .Where(s => s.TextBody != null)
            .ToList();

        // Pick the target shape: first for title, second for body.
        var shape = isTitle
            ? shapesWithText.FirstOrDefault()
            : shapesWithText.Skip(1).FirstOrDefault();

        var textBody = shape?.TextBody;
        if (textBody is null)
            return;

        // Clear any existing paragraphs so we start with a clean slate.
        textBody.RemoveAllChildren<A.Paragraph>();

        // Build a single paragraph that contains one run per line of text.
        var para = new A.Paragraph();
        foreach (var line in text.Split(['\n'], StringSplitOptions.None))
        {
            // Apply consistent font styling: 24pt Garamond, English locale, not dirty (no re-layout needed).
            var run = new A.Run(new A.RunProperties { FontSize = 2400, Language = "en-US", Dirty = false });
            run.RunProperties?.AppendChild(new A.LatinFont { Typeface = "Garamond" });
            run.AppendChild(new A.Text(line));
            para.AppendChild(run);

            // Insert an explicit line break after each run so multi-line text renders correctly.
            para.AppendChild(new A.Break());
        }
        textBody.AppendChild(para);
    }

    // Removes all HTML tags from 'source' and decodes common HTML entities.
    // Returns string.Empty when the input is null or empty.
    private static string StripHtmlTags(string source)
    {
        if (string.IsNullOrEmpty(source))
            return string.Empty;

        // Decode the non-breaking space entity before stripping tags.
        source = source.Replace("&nbsp;", " ");

        // Use the source-generated regex to strip every HTML tag (e.g. <p>, </div>, <br />).
        return HtmlTagRegex().Replace(source, string.Empty);
    }

    /// <summary>
    /// Source-generated, compiled regular expression that matches any HTML tag
    /// (e.g. &lt;p&gt;, &lt;/div&gt;, &lt;br /&gt;).
    /// Using <see cref="GeneratedRegexAttribute"/> avoids runtime compilation overhead
    /// on every call to <see cref="StripHtmlTags"/>.
    /// </summary>
    [GeneratedRegex("<.*?>", RegexOptions.Compiled)]
    private static partial Regex HtmlTagRegex();
}
