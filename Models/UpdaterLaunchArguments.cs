namespace VerlaufsakteApp.Models;

public sealed class UpdaterLaunchArguments
{
    public required string TargetExecutablePath { get; init; }
    public required string DownloadedExecutablePath { get; init; }
    public required int SourceProcessId { get; init; }
    public required string VersionString { get; init; }
}
