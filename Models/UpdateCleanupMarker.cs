namespace VerlaufsakteApp.Models;

public sealed class UpdateCleanupMarker
{
    public string? TargetVersion { get; set; }
    public DateTimeOffset CreatedUtc { get; set; }
}
