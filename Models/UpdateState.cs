namespace VerlaufsakteApp.Models;

public sealed class UpdateState
{
    public DateTimeOffset LastCheckedUtc { get; set; }
    public string? LastOfferedVersion { get; set; }
    public string? SkippedVersion { get; set; }
    public DateTimeOffset? SkipUntilUtc { get; set; }
    public string? CachedLatestVersion { get; set; }
    public string? CachedTagName { get; set; }
    public string? CachedReleaseTitle { get; set; }
    public string? CachedReleaseNotes { get; set; }
    public string? CachedDownloadUrl { get; set; }
    public DateTimeOffset? CachedPublishedAtUtc { get; set; }
}
