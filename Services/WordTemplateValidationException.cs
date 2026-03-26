namespace VerlaufsakteApp.Services;

internal enum WordTemplateValidationErrorKind
{
    BookmarkMissing,
    StructuredEntryTableInvalid,
    BiTodoTableInvalid,
    BiTodoContentInvalid
}

internal sealed class WordTemplateValidationException : InvalidOperationException
{
    public WordTemplateValidationException(
        WordTemplateValidationErrorKind kind,
        string bookmarkName,
        string userMessage)
        : base(userMessage)
    {
        Kind = kind;
        BookmarkName = bookmarkName;
        UserMessage = userMessage;
    }

    public WordTemplateValidationErrorKind Kind { get; }

    public string BookmarkName { get; }

    public string UserMessage { get; }
}
