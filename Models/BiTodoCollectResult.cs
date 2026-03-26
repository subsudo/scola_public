namespace VerlaufsakteApp.Models;

public sealed class BiTodoCollectRequest
{
    public string FullName { get; set; } = string.Empty;
    public string Initials { get; set; } = string.Empty;
    public string DocumentPath { get; set; } = string.Empty;
    public string FailureMessage { get; set; } = string.Empty;
}

public sealed class BiTodoCollectResult
{
    public string Name { get; set; } = string.Empty;
    public string Initials { get; set; } = string.Empty;
    public bool IsSuccess { get; set; }
    public string Message { get; set; } = string.Empty;

    public string DisplayName => string.IsNullOrWhiteSpace(Initials)
        ? Name
        : $"{Name} ({Initials})";
}

public sealed class BiTodoCollectSummary
{
    public List<BiTodoCollectResult> Results { get; set; } = new();
    public bool DocumentOpened { get; set; }
    public int SuccessCount => Results.Count(r => r.IsSuccess);
    public int FailureCount => Results.Count - SuccessCount;
}
