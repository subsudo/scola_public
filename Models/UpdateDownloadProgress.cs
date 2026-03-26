namespace VerlaufsakteApp.Models;

public sealed class UpdateDownloadProgress
{
    public long BytesReceived { get; init; }
    public long? TotalBytes { get; init; }
    public int? Percentage { get; init; }

    public string BuildStatusText()
    {
        var receivedText = FormatBytes(BytesReceived);
        if (TotalBytes is not long totalBytes || totalBytes <= 0)
        {
            return $"Lade Update herunter... {receivedText}";
        }

        var percentageText = Percentage is int percentage
            ? $"{percentage}%"
            : "Fortschritt unbekannt";
        return $"Lade Update herunter... {percentageText} ({receivedText} / {FormatBytes(totalBytes)})";
    }

    private static string FormatBytes(long bytes)
    {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        var unitIndex = 0;

        while (value >= 1024 && unitIndex < units.Length - 1)
        {
            value /= 1024;
            unitIndex++;
        }

        return $"{value:0.#} {units[unitIndex]}";
    }
}
