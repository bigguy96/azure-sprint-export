using System;
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

namespace SprintItemsApp
{
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
            if (!credentialService.TryGetCredential("SprintItemsApp.AzureDevOps", out string username, out string bearerToken))
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
                var workItems = WorkItemsGrid.ItemsSource as IEnumerable<WorkItem>;
                if (workItems != null)
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
                    // Auto-size all columns initially to content
                    column.Width = new DataGridLength(1, DataGridLengthUnitType.Auto);
                    dataGrid.UpdateLayout();
                    var desiredWidth = column.ActualWidth;
                    if (column.MinWidth > 0 && desiredWidth < column.MinWidth)
                    {
                        desiredWidth = column.MinWidth;
                    }
                    totalMinWidth += desiredWidth;
                    // Set fixed width for Auto columns, leave Star columns to adjust
                    if (column.Width.IsAuto && !column.Width.IsSizeToCells)
                    {
                        column.Width = new DataGridLength(desiredWidth, DataGridLengthUnitType.Pixel);
                    }
                }

                // Adjust Star columns to share remaining space
                var starColumns = dataGrid.Columns.Where(c => c.Width.IsStar).ToList();
                if (starColumns.Any())
                {
                    double availableWidth = dataGrid.ActualWidth - totalMinWidth;
                    if (availableWidth > 0)
                    {
                        double starWidth = availableWidth / starColumns.Count;
                        foreach (var column in starColumns)
                        {
                            double desiredWidth = Math.Max(column.MinWidth, starWidth);
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
            using (PresentationDocument presentationDoc = PresentationDocument.Create(filePath, PresentationDocumentType.Presentation))
            {
                // Create presentation part
                PresentationPart presentationPart = presentationDoc.AddPresentationPart();
                presentationPart.Presentation = new Presentation(new SlideIdList(), new SlideMasterIdList());
                presentationPart.Presentation.Save();

                // Create slide master part
                SlideMasterPart slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
                slideMasterPart.SlideMaster = new SlideMaster(
                    new CommonSlideData(new ShapeTree()),
                    new ColorMapOverride(),
                    new SlideLayoutIdList()
                );
                slideMasterPart.SlideMaster.Save();

                // Create slide layout part
                SlideLayoutPart slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
                slideLayoutPart.SlideLayout = new SlideLayout(new CommonSlideData(new ShapeTree()));
                slideLayoutPart.SlideLayout.Save();

                // Link slide master to layout
                SlideMasterId slideMasterId = new SlideMasterId
                {
                    Id = 2147483648U,
                    RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
                };
                presentationPart.Presentation.SlideMasterIdList.Append(slideMasterId);

                // Create slide part
                SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
                slidePart.Slide = new Slide(new CommonSlideData(new ShapeTree()));
                slidePart.AddPart(slideLayoutPart);
                slidePart.Slide.Save();

                // Link slide to presentation
                SlideId slideId = new SlideId
                {
                    Id = 256U,
                    RelationshipId = presentationPart.GetIdOfPart(slidePart)
                };
                presentationPart.Presentation.SlideIdList.Append(slideId);

                // Create table
                A.Table table = new A.Table();
                A.TableProperties tableProps = new A.TableProperties();
                table.Append(tableProps);

                // Table grid
                A.TableGrid tableGrid = new A.TableGrid();
                for (int i = 0; i < 5; i++)
                {
                    tableGrid.Append(new A.GridColumn { Width = 1905000 }); // ~2 inches per column
                }
                table.Append(tableGrid);

                // Collect selected work items
                var workItems = WorkItemsGrid.ItemsSource as IEnumerable<WorkItem>;
                if (workItems == null) return;

                // Header row
                A.TableRow headerRow = new A.TableRow { Height = 370000 }; // ~0.5 inches
                AddTableCell(headerRow, "ID", true);
                AddTableCell(headerRow, "Title", true);
                AddTableCell(headerRow, "State", true);
                AddTableCell(headerRow, "Type", true);
                AddTableCell(headerRow, "Assignee", true);
                table.Append(headerRow);

                // Data rows: Only include parent work items that are selected
                foreach (var workItem in workItems.Where(w => w.IsSelected))
                {
                    A.TableRow row = new A.TableRow { Height = 370000 };
                    AddTableCell(row, workItem.Id.ToString(), false);
                    AddTableCell(row, workItem.Title, false);
                    AddTableCell(row, workItem.State, false);
                    AddTableCell(row, workItem.WorkItemType, false);
                    AddTableCell(row, workItem.Assignee, false);
                    table.Append(row);

                    // Include ALL child work items, regardless of IsSelected
                    foreach (var child in workItem.Children)
                    {
                        A.TableRow childRow = new A.TableRow { Height = 370000 };
                        AddTableCell(childRow, child.Id.ToString(), false, 190500); // Indent ~0.25 inches
                        AddTableCell(childRow, child.Title, false);
                        AddTableCell(childRow, child.State, false);
                        AddTableCell(childRow, child.WorkItemType, false);
                        AddTableCell(childRow, child.Assignee, false);
                        table.Append(childRow);
                    }
                }

                // Add table to slide
                slidePart.Slide.CommonSlideData.ShapeTree.Append(table);
                slidePart.Slide.Save();

                // Save presentation
                presentationPart.Presentation.Save();
            }
        }

        private void AddTableCell(A.TableRow row, string text, bool isHeader, int indent = 0)
        {
            A.TableCell cell = new A.TableCell();
            A.TextBody textBody = new A.TextBody();
            A.BodyProperties bodyProps = new A.BodyProperties();
            A.Paragraph paragraph = new A.Paragraph();
            A.ParagraphProperties paraProps = new A.ParagraphProperties();
            if (indent > 0)
            {
                paraProps.LeftMargin = indent;
            }
            A.Run run = new A.Run();
            A.RunProperties runProps = new A.RunProperties { FontSize = isHeader ? 1800 : 1600 }; // 18pt, 16pt
            runProps.Append(new A.LatinFont { Typeface = "Calibri" });
            A.Text runText = new A.Text { Text = text ?? string.Empty };
            run.Append(runProps);
            run.Append(runText);
            paragraph.Append(paraProps);
            paragraph.Append(run);
            textBody.Append(bodyProps);
            textBody.Append(paragraph);
            cell.Append(textBody);
            A.TableCellProperties cellProps = new A.TableCellProperties();
            cell.Append(cellProps);
            row.Append(cell);
        }
    }
}