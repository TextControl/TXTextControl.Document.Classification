namespace TXTextControl.Document.Classification.Models;

internal sealed record DocumentClassificationResult(
    string DocumentPath,
    string PredictedCategory,
    double Confidence,
    IReadOnlyList<CategoryScore> Scores);
