namespace TXTextControl.Document.Classification.Models;

internal sealed record DocumentClassificationProfile(string Name, IReadOnlyList<KeywordRule> Rules);
