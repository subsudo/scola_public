namespace VerlaufsakteApp.Models;

public sealed record HeaderMetadata(string OdooUrl)
{
    public static HeaderMetadata Empty { get; } = new(string.Empty);
}
