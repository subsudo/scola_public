namespace VerlaufsakteApp.Models;

public sealed class HeaderMetadataCacheEntry
{
    public string DocumentPath { get; set; } = string.Empty;
    public DateTime LastWriteTimeUtc { get; set; }
    public long Length { get; set; }
    public DateTime LastSeenUtc { get; set; }
    public string HeaderSignature { get; set; } = string.Empty;
    public string OdooUrl { get; set; } = string.Empty;
    public string CounselorInitials { get; set; } = string.Empty;
}
