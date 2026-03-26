namespace VerlaufsakteApp.Models;

public sealed class GitHubReleaseInfo
{
    public required string TagName { get; init; }
    public required string VersionString { get; init; }
    public required Version Version { get; init; }
    public required string ReleaseTitle { get; init; }
    public required string ReleaseNotes { get; init; }
    public required string DownloadUrl { get; init; }
    public DateTimeOffset? PublishedAtUtc { get; init; }
}
