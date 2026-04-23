namespace TXTextControl.Document.Classification.Models;

internal sealed record DocumentSegment(DocumentRegion Region, string Text, int Order);
