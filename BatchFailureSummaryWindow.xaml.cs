using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace VerlaufsakteApp;

public partial class BatchFailureSummaryWindow : Window
{
    public BatchFailureSummaryWindow(IReadOnlyList<string> lines, int totalFailures, string batchLabel = "Batch")
    {
        InitializeComponent();

        Title = $"{batchLabel} mit Fehlern abgeschlossen";
        TitleTextBlock.Text = $"{batchLabel} mit Fehlern abgeschlossen";
        SummaryTextBlock.Text = totalFailures == 1
            ? "Ein Eintrag konnte nicht verarbeitet werden."
            : $"{totalFailures} Einträge konnten nicht verarbeitet werden.";

        FailureItemsControl.ItemsSource = lines.ToList();
    }

    private void OkButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}
