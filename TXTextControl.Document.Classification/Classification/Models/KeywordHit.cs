namespace TXTextControl.Document.Classification.Models;

internal sealed record KeywordHit(
    string Keyword,
    DocumentRegion Region,
    int Occurrences,
    double WeightedScore);
