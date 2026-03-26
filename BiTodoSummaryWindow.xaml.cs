using System.Collections.Generic;
using System.Linq;
using System.Windows;
using VerlaufsakteApp.Models;

namespace VerlaufsakteApp;

public partial class BiTodoSummaryWindow : Window
{
    public BiTodoSummaryWindow(BiTodoCollectSummary summary)
    {
        InitializeComponent();

        SummaryTextBlock.Text = summary.FailureCount == 0
            ? $"{summary.SuccessCount} Teilnehmer erfolgreich übernommen."
            : $"{summary.SuccessCount} erfolgreich, {summary.FailureCount} mit Hinweisen oder Fehlern.";

        DocumentStateTextBlock.Text = summary.DocumentOpened
            ? "Das Sammeldokument wurde geöffnet."
            : "Es konnte kein Sammeldokument geöffnet werden.";

        ResultItemsControl.ItemsSource = summary.Results.ToList();
    }

    private void OkButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}
