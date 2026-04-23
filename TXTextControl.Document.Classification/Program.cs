using TXTextControl.Document.Classification;
using TXTextControl.Document.Classification.Models;

var defaultDocumentPath = "Documents/Sample Financial.docx";

var documentPath = args.Length > 0 && !string.IsNullOrWhiteSpace(args[0])
    ? args[0].Trim()
    : defaultDocumentPath;

if (string.IsNullOrWhiteSpace(documentPath))
{
    Console.WriteLine("Usage: dotnet run -- <path-to-docx>");
    Console.WriteLine("No valid document path available.");
    return;
}

if (args.Length == 0)
{
    Console.WriteLine($"No argument provided. Using default document path: {documentPath}");
}

IReadOnlyList<DocumentClassificationProfile> profiles;

try
{
    profiles = DocumentClassificationProfileLoader.LoadFromFile();
}
catch (Exception ex)
{
    Console.WriteLine($"Failed to load classification profiles: {ex.Message}");
    return;
}

var classifier = new DocxKeywordClassifier(profiles);
var result = classifier.Classify(documentPath);

Console.WriteLine($"Document: {result.DocumentPath}");
Console.WriteLine($"Classification: {result.PredictedCategory}");
Console.WriteLine($"Confidence: {result.Confidence:P2}");
Console.WriteLine("Scores:");

foreach (var score in result.Scores.OrderByDescending(s => s.Score))
{
    Console.WriteLine($"- {score.Category}: {score.Score:F2}");

    foreach (var hit in score.KeywordHits.OrderByDescending(h => h.WeightedScore))
    {
        Console.WriteLine(
            $"  • [{hit.Region}] '{hit.Keyword}' x{hit.Occurrences} => {hit.WeightedScore:F2}");
    }
}
