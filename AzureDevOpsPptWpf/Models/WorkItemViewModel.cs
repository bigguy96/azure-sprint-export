using System.Collections.ObjectModel;
using System.ComponentModel;

namespace AzureDevOpsPptWpf.Models;

public class WorkItemViewModel : INotifyPropertyChanged
{
    private bool _isSelected;

    public int Id { get; set; }
    public string Title { get; set; }
    public string WorkItemType { get; set; }
    public string AssignedTo { get; set; }
    public string State { get; set; }
    public string Description { get; set; }

    public ObservableCollection<WorkItemViewModel> Children { get; set; } = new();

    public bool IsSelected
    {
        get => _isSelected;
        set { _isSelected = value; PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(IsSelected))); }
    }

    public event PropertyChangedEventHandler PropertyChanged;
}