using System.Text.RegularExpressions;

namespace VerlaufsakteApp.Services;

public static class SearchTextUtility
{
    public static IReadOnlyList<string> Tokenize(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Array.Empty<string>();
        }

        return Regex.Matches(value.ToLowerInvariant(), @"[\p{L}\p{N}]+")
            .Select(match => match.Value.Trim())
            .Where(token => token.Length > 0)
            .ToArray();
    }

    public static string ReplaceUmlauts(string value)
    {
        return value
            .Replace("ä", "ae", StringComparison.OrdinalIgnoreCase)
            .Replace("ö", "oe", StringComparison.OrdinalIgnoreCase)
            .Replace("ü", "ue", StringComparison.OrdinalIgnoreCase)
            .Replace("ß", "ss", StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsRobustToken(string token)
    {
        return !string.IsNullOrWhiteSpace(token) && token.Length >= 2 && token.Any(char.IsLetterOrDigit);
    }
}
