namespace VerlaufsakteApp.Models;

public static class DisplayDensityMode
{
    public const string Standard = "Standard";
    public const string Compact = "Kompakt";

    public static string Normalize(string? value)
    {
        return value switch
        {
            Standard => Standard,
            Compact => Compact,
            "Klassisch" => Standard,
            _ => Standard
        };
    }
}
