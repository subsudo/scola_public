namespace VerlaufsakteApp.Models;

public sealed class DownloadedUpdateInfo
{
    public required GitHubReleaseInfo Release { get; init; }
    public required string LocalFilePath { get; init; }
}
