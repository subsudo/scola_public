namespace VerlaufsakteApp.Models;

public sealed record HeaderMetadata(string OdooUrl, string CounselorInitials)
{
    public static HeaderMetadata Empty { get; } = new(string.Empty, string.Empty);
}
