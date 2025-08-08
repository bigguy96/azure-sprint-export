using AzureDevOpsPptWpf.Models;
using AzureDevOpsPptWpf.Services;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace AzureDevOpsPptWpf.Views;

public partial class MainWindow : Window
{
    private AzureDevOpsService _devOpsService;
    private readonly ObservableCollection<AzureDevOpsService.WorkItem> _workItems = [];
    private List<AzureDevOpsService.Team> _teams = [];
    private List<AzureDevOpsService.Iteration> _sprints = [];

    public MainWindow()
    {
        InitializeComponent();
        WorkItemsDataGrid.ItemsSource = _workItems;

        //WindowState = WindowState.Maximized;
        WindowStartupLocation = WindowStartupLocation.CenterScreen;
    }

    private async void LoadTeams_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            const string yourPersonalAccessToken = "tset";
            _devOpsService = new AzureDevOpsService(yourPersonalAccessToken);

            string org = OrgTextBox.Text;
            string project = ProjectTextBox.Text;

            _teams = await _devOpsService.GetTeamsAsync(org, project);
            TeamComboBox.ItemsSource = _teams;
            TeamComboBox.DisplayMemberPath = "Name";
        }
        catch (System.Exception ex)
        {
            MessageBox.Show($"Error loading teams: {ex.Message}");
        }
    }

    private async void TeamComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (TeamComboBox.SelectedItem is AzureDevOpsService.Team selectedTeam)
        {
            try
            {
                string org = OrgTextBox.Text;
                string project = ProjectTextBox.Text;

                _sprints = await _devOpsService.GetSprintsAsync(org, project, selectedTeam.Name);
                SprintComboBox.ItemsSource = _sprints;
                SprintComboBox.DisplayMemberPath = "Name";
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error loading sprints: {ex.Message}");
            }
        }
    }

    private async void LoadWorkItems_Click(object sender, RoutedEventArgs e)
    {
        if (TeamComboBox.SelectedItem is AzureDevOpsService.Team selectedTeam && SprintComboBox.SelectedItem is AzureDevOpsService.Iteration selectedSprint)
        {
            try
            {
                string org = OrgTextBox.Text;
                string project = ProjectTextBox.Text;

                var workItemIds = await _devOpsService.GetWorkItemIdsInSprintAsync(org, project, selectedTeam.Name, selectedSprint.Path);
                var workItems = await _devOpsService.GetWorkItemsAsync(org, workItemIds);

                _workItems.Clear();
                foreach (var wi in workItems)
                    _workItems.Add(wi);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error loading work items: {ex.Message}");
            }
        }
        else
        {
            MessageBox.Show("Please select both Team and Sprint.");
        }
    }

    private void WorkItemsDataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Handle logic on selection change if needed
        var dataGrid = sender as DataGrid;
        if (dataGrid == null) return;

        var selectedWorkItem = dataGrid.SelectedItem as WorkItemViewModel;
        if (selectedWorkItem != null)
        {
            // Example: Enable/disable buttons or show preview
        }
    }

    private void SprintComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        // Optional: Do something when sprint changes, like clearing work items
        _workItems.Clear();
    }

    // Highlight selected row (built-in DataGrid selection)
    // On double click, show detail window
    private void WorkItemsDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
    {
        if (WorkItemsDataGrid.SelectedItem is AzureDevOpsService.WorkItem selectedItem)
        {
            var detailWindow = new WorkItemDetailWindow(selectedItem);
            detailWindow.Owner = this;
            detailWindow.ShowDialog();
        }
    }
}