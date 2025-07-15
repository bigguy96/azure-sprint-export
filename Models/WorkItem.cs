using System.Text.Json.Serialization;
using System.Text.Json;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Linq;

namespace SprintItemsApp.Models
{
    public class WorkItem : INotifyPropertyChanged
    {
        private bool _isSelected;
        private bool _isHighlighted;

        [JsonPropertyName("id")]
        public int Id { get; set; }

        [JsonPropertyName("fields")]
        public Dictionary<string, object> Fields { get; set; }

        [JsonPropertyName("relations")]
        public List<WorkItemRelation> Relations { get; set; } = new List<WorkItemRelation>();

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged();
                }
            }
        }

        public bool IsHighlighted
        {
            get => _isHighlighted;
            set
            {
                if (_isHighlighted != value)
                {
                    _isHighlighted = value;
                    OnPropertyChanged();
                }
            }
        }

        public string Title => Fields.TryGetValue("System.Title", out var title) ? title.ToString() : string.Empty;
        public string State => Fields.TryGetValue("System.State", out var state) ? state.ToString() : string.Empty;
        public string WorkItemType => Fields.TryGetValue("System.WorkItemType", out var type) ? type.ToString() : string.Empty;
        public string Assignee
        {
            get
            {
                if (Fields.TryGetValue("System.AssignedTo", out var assignedTo) && assignedTo != null)
                {
                    if (assignedTo is JsonElement jsonElement)
                    {
                        if (jsonElement.TryGetProperty("displayName", out var displayNameElement))
                        {
                            return displayNameElement.GetString() ?? string.Empty;
                        }
                        System.Diagnostics.Debug.WriteLine($"WorkItem ID {Id}: System.AssignedTo is JsonElement but missing displayName: {JsonSerializer.Serialize(jsonElement)}");
                    }
                    else if (assignedTo is Dictionary<string, object> dict)
                    {
                        if (dict.TryGetValue("displayName", out var displayName))
                        {
                            return displayName.ToString();
                        }
                        System.Diagnostics.Debug.WriteLine($"WorkItem ID {Id}: System.AssignedTo is Dictionary but missing displayName: {JsonSerializer.Serialize(dict)}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"WorkItem ID {Id}: System.AssignedTo is unexpected type: {assignedTo?.GetType().FullName}, Value: {JsonSerializer.Serialize(assignedTo)}");
                    }
                }
                return string.Empty;
            }
        }

        public int? ParentId => Relations.FirstOrDefault(r => r.RelationType == "System.LinkTypes.Hierarchy-Reverse")?.TargetId;
        public List<WorkItem> Children { get; set; } = new List<WorkItem>();

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}