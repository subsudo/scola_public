namespace VerlaufsakteApp.Models;

public sealed class HeaderMetadataCacheDocument
{
    public int Version { get; set; }
    public List<HeaderMetadataCacheEntry> Entries { get; set; } = new();
}
