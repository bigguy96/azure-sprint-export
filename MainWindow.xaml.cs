﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using Microsoft.Extensions.Configuration;
using SprintItemsApp.Models;
using SprintItemsApp.Services;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using System.Windows.Controls.Primitives;
using System.IO;
using DocumentFormat.OpenXml.Drawing;

namespace SprintItemsApp;

public partial class MainWindow : Window
{
    private readonly AzureDevOpsService _service;
    private string _selectedColumn;

    public MainWindow()
    {
        InitializeComponent();
        var config = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json")
            .Build();
        var credentialService = new CredentialService();
        if (!credentialService.TryGetCredential("SprintItemsApp.AzureDevOps", out var username, out var bearerToken))
        {
            MessageBox.Show("Failed to load Azure DevOps token from Credential Manager.");
            return;
        }
        var organization = config["AzureDevOps:Organization"];
        var project = config["AzureDevOps:Project"];
        var team = config["AzureDevOps:Team"];
        _service = new AzureDevOpsService(bearerToken, organization, project, team);
        LoadSprintsAsync();
    }

    private async void LoadSprintsAsync()
    {
        try
        {
            var sprints = await _service.GetSprintsAsync();
            SprintComboBox.ItemsSource = sprints;
            SprintComboBox.SelectedIndex = -1;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading sprints: {ex.Message}");
        }
    }

    private async void SprintComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (SprintComboBox.SelectedItem is Sprint selectedSprint)
        {
            try
            {
                var workItems = await _service.GetWorkItemsForSprintAsync(selectedSprint.Id);
                WorkItemsGrid.ItemsSource = workItems;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading work items: {ex.Message}");
            }
        }
    }

    private void WorkItemsGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        var selectedWorkItems = WorkItemsGrid.SelectedItems.Cast<WorkItem>().ToList();
        if (selectedWorkItems.Any())
        {
            // Can add logic if needed
        }
    }

    private void WorkItemsGrid_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
    {
        if (e.OriginalSource is FrameworkElement element && element.DataContext is DataGridColumnHeader header)
        {
            _selectedColumn = header.Column.Header.ToString();
            if (WorkItemsGrid.ItemsSource is IEnumerable<WorkItem> workItems)
            {
                foreach (var item in workItems)
                {
                    item.IsHighlighted = _selectedColumn == "Type" && item.WorkItemType == "Bug";
                    foreach (var child in item.Children)
                    {
                        child.IsHighlighted = _selectedColumn == "Type" && child.WorkItemType == "Bug";
                    }
                }
                WorkItemsGrid.Items.Refresh();
            }
        }
    }

    private void ChildDataGrid_Loaded(object sender, RoutedEventArgs e)
    {
        ResizeDataGridColumns(sender as DataGrid);
    }

    private void WorkItemsGrid_RowDetailsVisibilityChanged(object sender, DataGridRowDetailsEventArgs e)
    {
        if (e.DetailsElement is DataGrid dataGrid && e.Row.DetailsVisibility == Visibility.Visible)
        {
            ResizeDataGridColumns(dataGrid);
        }
    }

    private void ResizeDataGridColumns(DataGrid dataGrid)
    {
        if (dataGrid != null)
        {
            double totalMinWidth = 0;
            foreach (var column in dataGrid.Columns)
            {
                column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                dataGrid.UpdateLayout();
                var desiredWidth = column.ActualWidth;
                if (column.MinWidth > 0 && desiredWidth < column.MinWidth)
                {
                    desiredWidth = column.MinWidth;
                }
                totalMinWidth += desiredWidth;
                if (column.Width.IsAuto && !column.Width.IsSizeToCells)
                {
                    column.Width = new DataGridLength(desiredWidth, DataGridLengthUnitType.Pixel);
                }
            }

            var starColumns = dataGrid.Columns.Where(c => c.Width.IsStar).ToList();
            if (starColumns.Any())
            {
                var availableWidth = dataGrid.ActualWidth - totalMinWidth;
                if (availableWidth > 0)
                {
                    var starWidth = availableWidth / starColumns.Count;
                    foreach (var column in starColumns)
                    {
                        var desiredWidth = Math.Max(column.MinWidth, starWidth);
                        column.Width = new DataGridLength(desiredWidth, DataGridLengthUnitType.Pixel);
                    }
                }
            }
        }
    }

    private void ExportToPowerPoint_Click(object sender, RoutedEventArgs e)
    {
        var saveFileDialog = new SaveFileDialog
        {
            Filter = "PowerPoint Files (*.pptx)|*.pptx",
            DefaultExt = "pptx",
            FileName = "WorkItemsExport.pptx"
        };

        if (saveFileDialog.ShowDialog() == true)
        {
            try
            {
                ExportToPowerPoint(saveFileDialog.FileName);
                MessageBox.Show("PowerPoint file created successfully!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to PowerPoint: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void ExportToPowerPoint(string filePath)
    {
        try
        {
            using var presentationDoc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation);
            // Create presentation part
            var presentationPart = presentationDoc.AddPresentationPart();
            presentationPart.Presentation = new Presentation(new SlideIdList(), new SlideMasterIdList());

            // Create theme part
            var themePart = presentationPart.AddNewPart<ThemePart>();
            themePart.Theme = new Theme(
                    new ThemeElements(
                        new ColorScheme(
                                new Dark1Color(new SystemColor { Val = SystemColorValues.WindowText, LastColor = "000000" }),
                                new Light1Color(new SystemColor { Val = SystemColorValues.Window, LastColor = "FFFFFF" }),
                                new Dark2Color(new RgbColorModelHex { Val = "1F497D" }),
                                new Light2Color(new RgbColorModelHex { Val = "EEECE1" }),
                                new Accent1Color(new RgbColorModelHex { Val = "4F81BD" }),
                                new Accent2Color(new RgbColorModelHex { Val = "C0504D" }),
                                new Accent3Color(new RgbColorModelHex { Val = "9BBB59" }),
                                new Accent4Color(new RgbColorModelHex { Val = "8064A2" }),
                                new Accent5Color(new RgbColorModelHex { Val = "4BACC6" }),
                                new Accent6Color(new RgbColorModelHex { Val = "F79646" }),
                                new Hyperlink(new RgbColorModelHex { Val = "0000FF" }),
                                new FollowedHyperlinkColor(new RgbColorModelHex { Val = "800080" })
                            )
                            { Name = "Office" },
                        new FontScheme(
                                new MajorFont(new LatinFont { Typeface = "Calibri" }),
                                new MinorFont(new LatinFont { Typeface = "Calibri" })
                            )
                            { Name = "Office" },
                        new FormatScheme(
                                new FillStyleList(
                                    new SolidFill(new SchemeColor { Val = SchemeColorValues.PhColor }),
                                    new GradientFill(),
                                    new NoFill()
                                ),
                                new LineStyleList(new Outline(), new Outline(), new Outline()),
                                new EffectStyleList(new EffectStyle()),
                                new BackgroundFillStyleList(
                                    new SolidFill(new SchemeColor { Val = SchemeColorValues.PhColor }),
                                    new GradientFill(),
                                    new GradientFill()
                                )
                            )
                            { Name = "Office" }
                    )
                )
                { Name = "Office Theme" };
            themePart.Theme.Save();

            // Create slide master part
            var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
            slideMasterPart.SlideMaster = new SlideMaster(
                new CommonSlideData(
                    new ShapeTree()
                ),
                new DocumentFormat.OpenXml.Presentation.ColorMap
                {
                    Background1 = ColorSchemeIndexValues.Dark1,
                    Text1 = ColorSchemeIndexValues.Light1,
                    Background2 = ColorSchemeIndexValues.Light2,
                    Text2 = ColorSchemeIndexValues.Dark2,
                    Accent1 = ColorSchemeIndexValues.Accent1,
                    Accent2 = ColorSchemeIndexValues.Accent2,
                    Accent3 = ColorSchemeIndexValues.Accent3,
                    Accent4 = ColorSchemeIndexValues.Accent4,
                    Accent5 = ColorSchemeIndexValues.Accent5,
                    Accent6 = ColorSchemeIndexValues.Accent6
                },
                new SlideLayoutIdList()
            );
            slideMasterPart.SlideMaster.Save();

            // Create slide layout part
            var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
            slideLayoutPart.SlideLayout = new SlideLayout(
                new CommonSlideData(
                        new ShapeTree()
                    )
                    { Name = "Blank" }
            );
            slideLayoutPart.SlideLayout.Save();

            // Link slide master to theme and layout
            slideMasterPart.AddPart(themePart);
            var slideMasterId = new SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            };
            presentationPart.Presentation.SlideMasterIdList?.Append(slideMasterId);

            var slideLayoutId = new SlideLayoutId
            {
                Id = 2147483649U,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            };
            slideMasterPart.SlideMaster.SlideLayoutIdList?.Append(slideLayoutId);

            // Collect selected work items
            if (WorkItemsGrid.ItemsSource is not IEnumerable<WorkItem> workItems || !workItems.Any(w => w.IsSelected))
            {
                throw new InvalidOperationException("No work items selected for export.");
            }

            // Create a slide for each selected parent work item
            var slideIdValue = 256U;
            foreach (var workItem in workItems.Where(w => w.IsSelected))
            {
                // Create slide part
                var slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(
                    new CommonSlideData(
                        new ShapeTree()
                    )
                );
                slidePart.AddPart(slideLayoutPart);

                // Link slide to presentation
                var slideId = new SlideId
                {
                    Id = slideIdValue++,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };
                presentationPart.Presentation.SlideIdList?.Append(slideId);

                // Create table
                var table = new Table();
                var tableProps = new TableProperties();
                table.Append(tableProps);

                // Table grid
                var tableGrid = new TableGrid();
                for (var i = 0; i < 5; i++)
                {
                    tableGrid.Append(new GridColumn { Width = 1905000 }); // ~2 inches per column
                }
                table.Append(tableGrid);

                // Header row
                var headerRow = new TableRow { Height = 370000 }; // ~0.5 inches
                AddTableCell(headerRow, "ID", true);
                AddTableCell(headerRow, "Title", true);
                AddTableCell(headerRow, "State", true);
                AddTableCell(headerRow, "Type", true);
                AddTableCell(headerRow, "Assignee", true);
                table.Append(headerRow);

                // Parent row
                var parentRow = new TableRow { Height = 370000 };
                AddTableCell(parentRow, workItem.Id.ToString(), false);
                AddTableCell(parentRow, workItem.Title ?? "", false);
                AddTableCell(parentRow, workItem.State ?? "", false);
                AddTableCell(parentRow, workItem.WorkItemType ?? "", false);
                AddTableCell(parentRow, workItem.Assignee ?? "", false);
                table.Append(parentRow);

                // Child rows
                foreach (var child in workItem.Children ?? [])
                {
                    var childRow = new TableRow { Height = 370000 };
                    AddTableCell(childRow, child.Id.ToString(), false, 190500); // Indent ~0.25 inches
                    AddTableCell(childRow, child.Title ?? "", false);
                    AddTableCell(childRow, child.State ?? "", false);
                    AddTableCell(childRow, child.WorkItemType ?? "", false);
                    AddTableCell(childRow, child.Assignee ?? "", false);
                    table.Append(childRow);
                }

                // Position table on slide using GraphicFrame
                var graphicFrame = new A.GraphicFrame(
                    new A.NonVisualGraphicFrameProperties(
                        new A.NonVisualDrawingProperties { Id = 2U, Name = "Table" },
                        new A.NonVisualGraphicFrameDrawingProperties()
                    ),
                    new Transform(
                        new Offset { X = 500000, Y = 500000 }, // ~0.5 inches from top-left
                        new Extents { Cx = 5 * 1905000, Cy = (1 + (workItem.Children?.Count ?? 0)) * 370000 } // Width: 5 columns, Height: rows
                    ),
                    new Graphic(
                        new GraphicData(table) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/table" }
                    )
                );
                slidePart.Slide.CommonSlideData?.ShapeTree?.Append(graphicFrame);

                // Save slide
                slidePart.Slide.Save();
            }

            // Save presentation
            presentationPart.Presentation.Save();
        }
        catch (Exception ex)
        {
            File.WriteAllText("export_error.log", ex.ToString());
            throw; // Rethrow to show in UI
        }
    }

    // private void ExportToPowerPoint(string filePath)
    // {
    //     using (PresentationDocument presentationDoc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
    //     {
    //         // Create presentation part
    //         PresentationPart presentationPart = presentationDoc.AddPresentationPart();
    //         presentationPart.Presentation = new Presentation(new SlideIdList(), new SlideMasterIdList());
    //         presentationPart.Presentation.Save();

    //         // Create slide master part
    //         SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
    //         slideMasterPart.SlideMaster = new SlideMaster(
    //             new CommonSlideData(new ShapeTree()),
    //             new ColorMap
    //             {
    //                 Background1 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Dark1),
    //                 Text1 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Light1),
    //                 Background2 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Light2),
    //                 Text2 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Dark2),
    //                 Accent1 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Accent1),
    //                 Accent2 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Accent2),
    //                 Accent3 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Accent3),
    //                 Accent4 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Accent4),
    //                 Accent5 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Accent5),
    //                 Accent6 = new EnumValue<A.ColorSchemeIndexValues>(A.ColorSchemeIndexValues.Accent6)
    //             },
    //             new SlideLayoutIdList()
    //         );
    //         slideMasterPart.SlideMaster.Save();

    //         // Create slide layout part
    //         SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
    //         slideLayoutPart.SlideLayout = new SlideLayout(
    //             new CommonSlideData(new ShapeTree()) { Name = new CommonSlideDataName { Value = "Blank" } },
    //             new ColorMapOverride()
    //         );
    //         slideLayoutPart.SlideLayout.Save();

    //         // Link slide master to layout
    //         SlideMasterId slideMasterId = new SlideMasterId
    //         {
    //             Id = 2147483648U,
    //             RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
    //         };
    //         presentationPart.Presentation.SlideMasterIdList.Append(slideMasterId);

    //         // Link layout to master
    //         SlideLayoutId slideLayoutId = new SlideLayoutId
    //         {
    //             Id = 2147483649U,
    //             RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
    //         };
    //         slideMasterPart.SlideMaster.SlideLayoutIdList.Append(slideLayoutId);

    //         // Collect selected work items
    //         var workItems = WorkItemsGrid.ItemsSource as IEnumerable<WorkItem>;
    //         if (workItems == null) return;

    //         // Create a slide for each selected parent work item
    //         uint slideIdValue = 256U;
    //         foreach (var workItem in workItems.Where(w => w.IsSelected))
    //         {
    //             // Create slide part
    //             SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
    //             slidePart.Slide = new Slide(
    //                 new CommonSlideData(new ShapeTree()),
    //                 new ColorMapOverride()
    //             );
    //             slidePart.AddPart(slideLayoutPart);

    //             // Link slide to presentation
    //             SlideId slideId = new SlideId
    //             {
    //                 Id = slideIdValue++,
    //                 RelationshipId = presentationPart.GetIdOfPart(slidePart)
    //             };
    //             presentationPart.Presentation.SlideIdList.Append(slideId);

    //             // Create table
    //             A.Table table = new A.Table();
    //             A.TableProperties tableProps = new A.TableProperties();
    //             table.Append(tableProps);

    //             // Table grid
    //             A.TableGrid tableGrid = new A.TableGrid();
    //             for (int i = 0; i < 5; i++)
    //             {
    //                 tableGrid.Append(new A.GridColumn { Width = 1905000 }); // ~2 inches per column
    //             }
    //             table.Append(tableGrid);

    //             // Header row
    //             A.TableRow headerRow = new A.TableRow { Height = 370000 }; // ~0.5 inches
    //             AddTableCell(headerRow, "ID", true);
    //             AddTableCell(headerRow, "Title", true);
    //             AddTableCell(headerRow, "State", true);
    //             AddTableCell(headerRow, "Type", true);
    //             AddTableCell(headerRow, "Assignee", true);
    //             table.Append(headerRow);

    //             // Parent row
    //             A.TableRow parentRow = new A.TableRow { Height = 370000 };
    //             AddTableCell(parentRow, workItem.Id.ToString(), false);
    //             AddTableCell(parentRow, workItem.Title, false);
    //             AddTableCell(parentRow, workItem.State, false);
    //             AddTableCell(parentRow, workItem.WorkItemType, false);
    //             AddTableCell(parentRow, workItem.Assignee, false);
    //             table.Append(parentRow);

    //             // Child rows
    //             foreach (var child in workItem.Children)
    //             {
    //                 A.TableRow childRow = new A.TableRow { Height = 370000 };
    //                 AddTableCell(childRow, child.Id.ToString(), false, 190500); // Indent ~0.25 inches
    //                 AddTableCell(childRow, child.Title, false);
    //                 AddTableCell(childRow, child.State, false);
    //                 AddTableCell(childRow, child.WorkItemType, false);
    //                 AddTableCell(childRow, child.Assignee, false);
    //                 table.Append(childRow);
    //             }

    //             // Add table to slide
    //             slidePart.Slide.CommonSlideData.ShapeTree.Append(table);
    //             slidePart.Slide.Save();
    //         }

    //         // Save presentation
    //         presentationPart.Presentation.Save();
    //     }
    // }

    private void AddTableCell(TableRow row, string text, bool isHeader, int indent = 0)
    {
        var cell = new TableCell();
        var textBody = new A.TextBody();
        var bodyProps = new BodyProperties();
        var paragraph = new Paragraph();
        var paraProps = new ParagraphProperties();
        if (indent > 0)
        {
            paraProps.LeftMargin = indent;
        }
        var run = new Run();
        var runProps = new RunProperties { FontSize = isHeader ? 1800 : 1600 }; // 18pt, 16pt
        runProps.Append(new LatinFont { Typeface = "Calibri" });
        var runText = new A.Text { Text = text ?? string.Empty };
        run.Append(runProps);
        run.Append(runText);
        paragraph.Append(paraProps);
        paragraph.Append(run);
        textBody.Append(bodyProps);
        textBody.Append(paragraph);
        cell.Append(textBody);
        var cellProps = new TableCellProperties();
        cell.Append(cellProps);
        row.Append(cell);
    }
}

internal class CommonSlideDataName : StringValue
{
    public string Value { get; set; }
}