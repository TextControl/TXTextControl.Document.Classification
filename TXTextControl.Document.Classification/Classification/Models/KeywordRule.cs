namespace TXTextControl.Document.Classification.Models;

internal sealed record KeywordRule(
    string Term,
    double Weight = 1.0,
    KeywordMatchMode MatchMode = KeywordMatchMode.WholeWord,
    bool CaseSensitive = false,
    DocumentRegion? TargetRegion = null,
    KeywordStrength Strength = KeywordStrength.Normal);
