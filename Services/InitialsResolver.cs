using System.IO;
using System.Text.RegularExpressions;

namespace VerlaufsakteApp.Services;

public class InitialsResolver
{
    private static readonly Regex SuffixRegex = new(@"^[\p{L}\p{N}]{2,8}$", RegexOptions.Compiled);

    public string TryResolveFromDocumentPath(string? docPath)
    {
        if (string.IsNullOrWhiteSpace(docPath))
        {
            return string.Empty;
        }

        var fileName = Path.GetFileNameWithoutExtension(docPath);
        if (string.IsNullOrWhiteSpace(fileName) || !fileName.Contains('_'))
        {
            return string.Empty;
        }

        var suffix = fileName.Split('_', StringSplitOptions.RemoveEmptyEntries).LastOrDefault()?.Trim() ?? string.Empty;
        return SuffixRegex.IsMatch(suffix) ? suffix : string.Empty;
    }
}
