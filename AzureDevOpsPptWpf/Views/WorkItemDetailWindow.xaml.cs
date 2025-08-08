using System.Windows;
using AzureDevOpsPptWpf.Services;

namespace AzureDevOpsPptWpf.Views;

public partial class WorkItemDetailWindow : Window
{
    public AzureDevOpsService.WorkItem WorkItem { get; }

    public WorkItemDetailWindow(AzureDevOpsService.WorkItem workItem)
    {
        InitializeComponent();
        WorkItem = workItem;
        DataContext = WorkItem;
    }
}