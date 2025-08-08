using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace AzureDevOpsPptWpf.Services;

public static class PowerPointGenerator
{
    public static void GeneratePresentation(string filePath, List<AzureDevOpsService.WorkItem> workItems)
    {
        using (var presentationDoc = PresentationDocument.Create(filePath, DocumentFormat.OpenXml.PresentationDocumentType.Presentation))
        {
            var presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation();

            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(new CommonSlideData(new ShapeTree()));

            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));

            slideMasterPart.SlideMaster.Append(new SlideLayoutIdList(new SlideLayoutId() { Id = 1U, RelationshipId = presentationPart.GetIdOfPart(slideLayoutPart) }));
            presentationPart.Presentation.Append(new SlideIdList());

            uint slideId = 256;

            foreach (var wi in workItems)
            {
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = CreateSlide(wi);

                presentationPart.Presentation.SlideIdList.Append(new SlideId()
                {
                    Id = slideId++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                });
            }

            presentationPart.Presentation.Save();
        }
    }

    private static Slide CreateSlide(AzureDevOpsService.WorkItem wi)
    {
        var slide = new Slide(new CommonSlideData(new ShapeTree()));

        var shapeTree = slide.CommonSlideData.ShapeTree;

        // Add title shape (Work Item ID and Title)
        shapeTree.AppendChild(CreateTextShape(1, "Title", $"{wi.Id}: {wi.Title}", 0, 0, 8000000, 1000000, 44));

        // Add Assigned To
        shapeTree.AppendChild(CreateTextShape(2, "AssignedTo", $"Assigned To: {wi.AssignedTo}", 0, 1100000, 8000000, 800000, 24));

        // Add State
        shapeTree.AppendChild(CreateTextShape(3, "State", $"Status: {wi.State}", 0, 1900000, 8000000, 800000, 24));

        // Add Description
        shapeTree.AppendChild(CreateTextShape(4, "Description", wi.Description ?? "(No Description)", 0, 2700000, 8000000, 4000000, 18, true));

        return slide;
    }

    private static Shape CreateTextShape(uint id, string name, string text, long offsetX, long offsetY, long width, long height, int fontSize, bool wrapText = false)
    {
        var shape = new Shape(
            new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() { Id = id, Name = name },
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
                new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
            new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                new A.Transform2D(
                    new A.Offset() { X = offsetX, Y = offsetY },
                    new A.Extents() { Cx = width, Cy = height })),
            new DocumentFormat.OpenXml.Presentation.TextBody(
                new A.BodyProperties() { Wrap = wrapText ? A.TextWrappingValues.Square : A.TextWrappingValues.None },
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.Text(text))
                    {
                        RunProperties = new A.RunProperties() { FontSize = fontSize * 100 }
                    })));

        return shape;
    }

}