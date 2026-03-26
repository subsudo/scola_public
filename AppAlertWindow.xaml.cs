using System.Windows;
using System.Windows.Media;

namespace VerlaufsakteApp;

public enum AppAlertKind
{
    Info,
    Warning,
    Error
}

public partial class AppAlertWindow : Window
{
    public AppAlertWindow(string title, string lead, string body, AppAlertKind kind, string? footnote = null)
    {
        InitializeComponent();

        Title = title;
        HeadingTextBlock.Text = title;
        LeadTextBlock.Text = lead;
        BodyTextBlock.Text = body;
        FootnoteTextBlock.Text = footnote ?? string.Empty;
        FootnoteTextBlock.Visibility = string.IsNullOrWhiteSpace(footnote)
            ? Visibility.Collapsed
            : Visibility.Visible;
        ApplyKind(kind);
    }

    private void ApplyKind(AppAlertKind kind)
    {
        string bgHex;
        string borderHex;
        string glyph;

        switch (kind)
        {
            case AppAlertKind.Error:
                bgHex = "#3A1B1B";
                borderHex = "#D17878";
                glyph = "×";
                break;
            case AppAlertKind.Info:
                bgHex = "#1B2A3A";
                borderHex = "#5B9BD5";
                glyph = "i";
                break;
            default:
                bgHex = "#3A331B";
                borderHex = "#C8A96C";
                glyph = "!";
                break;
        }

        IconCircle.Background = BrushFromHex(bgHex);
        IconCircle.BorderBrush = BrushFromHex(borderHex);
        IconGlyph.Foreground = BrushFromHex(borderHex);
        IconGlyph.Text = glyph;
    }

    private static SolidColorBrush BrushFromHex(string hex)
    {
        var color = (Color)ColorConverter.ConvertFromString(hex);
        return new SolidColorBrush(color);
    }

    private void OkButton_OnClick(object sender, RoutedEventArgs e)
    {
        DialogResult = true;
        Close();
    }
}
