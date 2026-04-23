namespace TXTextControl.Document.Classification.Models;

internal sealed record CategoryScore(
    string Category,
    double Score,
    IReadOnlyList<KeywordHit> KeywordHits);
