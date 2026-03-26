namespace VerlaufsakteApp.Models;

public sealed class HeaderMetadataCacheEntry
{
    public string DocumentPath { get; set; } = string.Empty;
    public DateTime LastWriteTimeUtc { get; set; }
    public long Length { get; set; }
    public DateTime LastSeenUtc { get; set; }
    public string OdooUrl { get; set; } = string.Empty;
}
