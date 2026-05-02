using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp;

public partial class ParticipantHintsWindow : Window
{
    public ParticipantHintsWindow(string participantName, IEnumerable<ParticipantHintEntry> hints)
    {
        InitializeComponent();
        WindowTitle = $"Hinweise für {participantName}";

        foreach (var item in hints.Select(ParticipantHintEditorItem.FromEntry))
        {
            if (item.IsActive)
            {
                ActiveItems.Add(item);
            }
            else
            {
                DoneItems.Add(item);
            }
        }

        DataContext = this;
    }

    public string WindowTitle { get; }
    public ObservableCollection<ParticipantHintEditorItem> ActiveItems { get; } = new();
    public ObservableCollection<ParticipantHintEditorItem> DoneItems { get; } = new();

    public IReadOnlyList<ParticipantHintEditorItem> GetAllItems()
    {
        return ActiveItems.Concat(DoneItems).ToList();
    }

    private void AddExitButton_OnClick(object sender, RoutedEventArgs e)
    {
        ActiveItems.Add(CreateItem(ParticipantHintTypes.Exit));
    }

    private void AddAmButton_OnClick(object sender, RoutedEventArgs e)
    {
        ActiveItems.Add(CreateItem(ParticipantHintTypes.AmReport));
    }

    private void AddStellwerkButton_OnClick(object sender, RoutedEventArgs e)
    {
        ActiveItems.Add(CreateItem(ParticipantHintTypes.StellwerkTest));
    }

    private void AddFreeButton_OnClick(object sender, RoutedEventArgs e)
    {
        ActiveItems.Add(CreateItem(ParticipantHintTypes.Free));
    }

    private static ParticipantHintEditorItem CreateItem(string type)
    {
        var now = DateTime.Today;
        return new ParticipantHintEditorItem
        {
            Type = type,
            Status = ParticipantHintStatuses.Active,
            Date = type is ParticipantHintTypes.Exit or ParticipantHintTypes.StellwerkTest
                ? now.ToString("yyyy-MM-dd")
                : string.Empty,
            Month = type == ParticipantHintTypes.AmReport
                ? now.ToString("yyyy-MM")
                : string.Empty
        };
    }

    private void MarkDoneButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ParticipantHintEditorItem item } && ActiveItems.Remove(item))
        {
            item.Status = ParticipantHintStatuses.Done;
            DoneItems.Add(item);
        }
    }

    private void ReactivateHintButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is Button { Tag: ParticipantHintEditorItem item } && DoneItems.Remove(item))
        {
            item.Status = ParticipantHintStatuses.Active;
            ActiveItems.Add(item);
        }
    }

    private void DeleteHintButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (sender is not Button { Tag: ParticipantHintEditorItem item })
        {
            return;
        }

        ActiveItems.Remove(item);
        DoneItems.Remove(item);
    }

    private void SaveButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
    }

    private void CancelButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
    }
}
