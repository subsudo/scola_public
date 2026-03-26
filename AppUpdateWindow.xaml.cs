using System.ComponentModel;
using System.Windows;
using VerlaufsakteApp.Models;
using VerlaufsakteApp.Services;

namespace VerlaufsakteApp;

public partial class AppUpdateWindow : Window
{
    private readonly AppUpdateService _appUpdateService;
    private readonly GitHubReleaseInfo _releaseInfo;
    private bool _isDownloading;

    public AppUpdateWindow(AppUpdateService appUpdateService, GitHubReleaseInfo releaseInfo)
    {
        InitializeComponent();

        _appUpdateService = appUpdateService;
        _releaseInfo = releaseInfo;

        CurrentVersionTextBlock.Text = _appUpdateService.CurrentVersionDisplay;
        NewVersionTextBlock.Text = releaseInfo.VersionString;
        ReleaseNotesTextBlock.Text = string.IsNullOrWhiteSpace(releaseInfo.ReleaseNotes)
            ? "Keine zusätzlichen Hinweise verfügbar."
            : releaseInfo.ReleaseNotes;
        ProgressStatusTextBlock.Text = "Lade Update herunter...";
    }

    public DownloadedUpdateInfo? DownloadedUpdate { get; private set; }
    public bool WasDeferred { get; private set; }

    private void LaterButton_OnClick(object sender, RoutedEventArgs e)
    {
        WasDeferred = true;
        DialogResult = false;
        Close();
    }

    private async void UpdateButton_OnClick(object sender, RoutedEventArgs e)
    {
        if (_isDownloading)
        {
            return;
        }

        try
        {
            SetDownloadingState(true);
            var progress = new Progress<UpdateDownloadProgress>(updateProgress =>
            {
                ProgressStatusTextBlock.Text = updateProgress.BuildStatusText();
                if (updateProgress.Percentage is int percentage)
                {
                    DownloadProgressBar.IsIndeterminate = false;
                    DownloadProgressBar.Value = Math.Clamp(percentage, 0, 100);
                }
                else
                {
                    DownloadProgressBar.IsIndeterminate = true;
                }
            });

            DownloadedUpdate = await _appUpdateService.DownloadUpdateAsync(
                _releaseInfo,
                progress,
                CancellationToken.None);

            _isDownloading = false;
            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"Updater-Dialog: Download fehlgeschlagen: {ex.Message}");
            ErrorTextBlock.Text = $"Update konnte nicht heruntergeladen werden: {ex.Message}";
            ErrorTextBlock.Visibility = Visibility.Visible;
            SetDownloadingState(false);
        }
    }

    protected override void OnClosing(CancelEventArgs e)
    {
        if (_isDownloading)
        {
            e.Cancel = true;
            return;
        }

        if (!DialogResult.HasValue && DownloadedUpdate is null)
        {
            WasDeferred = true;
        }

        base.OnClosing(e);
    }

    private void SetDownloadingState(bool isDownloading)
    {
        _isDownloading = isDownloading;
        DownloadProgressBar.Visibility = isDownloading ? Visibility.Visible : Visibility.Collapsed;
        ProgressStatusTextBlock.Visibility = isDownloading ? Visibility.Visible : Visibility.Collapsed;
        ErrorTextBlock.Visibility = Visibility.Collapsed;
        UpdateButton.IsEnabled = !isDownloading;
        LaterButton.IsEnabled = !isDownloading;
        if (isDownloading)
        {
            DownloadProgressBar.Value = 0;
            DownloadProgressBar.IsIndeterminate = true;
            ProgressStatusTextBlock.Text = "Lade Update herunter...";
        }
    }
}
