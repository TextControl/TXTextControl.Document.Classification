using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using TXTextControl;
using TXTextControl.Document.Classification.Models;

namespace TXTextControl.Document.Classification;

internal sealed class DocxKeywordClassifier
{
    private readonly IReadOnlyList<DocumentClassificationProfile> _profiles;

    public DocxKeywordClassifier(IEnumerable<DocumentClassificationProfile> profiles)
    {
        _profiles = profiles?.Where(p => p.Rules.Count > 0).ToArray()
            ?? throw new ArgumentNullException(nameof(profiles));

        if (_profiles.Count == 0)
        {
            throw new ArgumentException("At least one profile with rules is required.", nameof(profiles));
        }
    }

    public DocumentClassificationResult Classify(string docxPath)
    {
        if (string.IsNullOrWhiteSpace(docxPath))
        {
            throw new ArgumentException("A document path is required.", nameof(docxPath));
        }

        if (!File.Exists(docxPath))
        {
            throw new FileNotFoundException("The specified DOCX file does not exist.", docxPath);
        }

        if (!string.Equals(Path.GetExtension(docxPath), ".docx", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Only .docx files are supported.");
        }

        var segments = ExtractSegments(docxPath);
        if (segments.Count == 0)
        {
            return new DocumentClassificationResult(docxPath, "Unknown", 0.0, []);
        }

        var scores = _profiles
            .Select(profile => ScoreProfile(profile, segments))
            .OrderByDescending(s => s.Score)
            .ToArray();

        var best = scores.First();
        if (best.Score <= 0)
        {
            return new DocumentClassificationResult(docxPath, "Unknown", 0.0, scores);
        }

        var second = scores.Length > 1 ? scores[1].Score : 0.0;
        var separation = best.Score <= 0 ? 0 : (best.Score - second) / best.Score;
        var coverage = best.KeywordHits.Count / (double)_profiles.First(p => p.Name == best.Category).Rules.Count;
        var confidence = Clamp((separation * 0.7) + (coverage * 0.3), 0.0, 1.0);

        return new DocumentClassificationResult(docxPath, best.Category, confidence, scores);
    }

    private static IReadOnlyList<DocumentSegment> ExtractSegments(string docxPath)
    {
        using var textControl = new ServerTextControl();
        textControl.Create();
        textControl.Load(docxPath, StreamType.WordprocessingML);

        var segments = new List<DocumentSegment>();
        var segmentOrder = 0;

        var paragraphIndex = 0;
        foreach (Paragraph paragraph in textControl.Paragraphs)
        {
            var text = Normalize(paragraph.Text);
            if (!string.IsNullOrWhiteSpace(text))
            {
                var region = ResolveBodyRegion(paragraph, paragraphIndex, text);
                segments.Add(new DocumentSegment(region, text, segmentOrder++));
            }

            paragraphIndex++;
        }

        foreach (Section section in textControl.Sections)
        {
            foreach (HeaderFooter headerFooter in section.HeadersAndFooters)
            {
                var region = headerFooter.Type.ToString().Contains("Footer", StringComparison.OrdinalIgnoreCase)
                    ? DocumentRegion.Footer
                    : DocumentRegion.Header;

                foreach (Paragraph paragraph in headerFooter.Paragraphs)
                {
                    var text = Normalize(paragraph.Text);
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        segments.Add(new DocumentSegment(region, text, segmentOrder++));
                    }
                }
            }
        }

        if (segments.Count == 0)
        {
            var fallback = Normalize(textControl.Text);
            if (!string.IsNullOrWhiteSpace(fallback))
            {
                segments.Add(new DocumentSegment(DocumentRegion.Body, fallback, segmentOrder++));
            }
        }

        return segments;
    }

    private static CategoryScore ScoreProfile(
        DocumentClassificationProfile profile,
        IReadOnlyList<DocumentSegment> segments)
    {
        var hits = new List<KeywordHit>();
        var total = 0.0;
        var segmentCount = Math.Max(segments.Count, 1);

        foreach (var rule in profile.Rules)
        {
            var candidateSegments = rule.TargetRegion is null
                ? segments
                : segments.Where(s => s.Region == rule.TargetRegion.Value).ToArray();

            if (candidateSegments.Count == 0)
            {
                continue;
            }

            var occurrencesByRegion = new Dictionary<DocumentRegion, (int RawOccurrences, double WeightedOccurrences)>();

            foreach (var segment in candidateSegments)
            {
                var matches = CountMatches(segment.Text, rule);
                if (matches == 0)
                {
                    continue;
                }

                var positionWeight = PositionWeight(segment, segmentCount);
                var weightedOccurrences = matches * positionWeight;

                occurrencesByRegion[segment.Region] =
                    occurrencesByRegion.TryGetValue(segment.Region, out var existing)
                        ? (existing.RawOccurrences + matches, existing.WeightedOccurrences + weightedOccurrences)
                        : (matches, weightedOccurrences);
            }

            foreach (var occurrence in occurrencesByRegion)
            {
                // Dampen repeated matches so one frequent generic term (for example "analysis")
                // does not dominate the whole category score.
                var frequencyFactor = Math.Sqrt(occurrence.Value.WeightedOccurrences);
                var weighted = frequencyFactor
                    * rule.Weight
                    * RegionWeight(occurrence.Key)
                    * KeywordStrengthWeight(rule);

                total += weighted;
                hits.Add(new KeywordHit(rule.Term, occurrence.Key, occurrence.Value.RawOccurrences, weighted));
            }
        }

        return new CategoryScore(profile.Name, total, hits);
    }

    private static int CountMatches(string text, KeywordRule rule)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(rule.Term))
        {
            return 0;
        }

        var options = RegexOptions.CultureInvariant | RegexOptions.Compiled;
        if (!rule.CaseSensitive)
        {
            options |= RegexOptions.IgnoreCase;
        }

        var pattern = rule.MatchMode switch
        {
            KeywordMatchMode.Exact => $"^{Regex.Escape(rule.Term)}$",
            KeywordMatchMode.Phrase => Regex.Escape(rule.Term),
            _ => $@"\b{Regex.Escape(rule.Term)}\b"
        };

        return Regex.Matches(text, pattern, options).Count;
    }

    private static string Normalize(string? input)
    {
        if (string.IsNullOrWhiteSpace(input))
        {
            return string.Empty;
        }

        var normalized = input.Normalize(NormalizationForm.FormD);
        var sb = new StringBuilder(normalized.Length);

        foreach (var c in normalized)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(c);
            if (category != UnicodeCategory.NonSpacingMark)
            {
                sb.Append(char.IsControl(c) ? ' ' : c);
            }
        }

        return Regex.Replace(sb.ToString().Normalize(NormalizationForm.FormC), @"\s+", " ").Trim();
    }

    private static double RegionWeight(DocumentRegion region) => region switch
    {
        DocumentRegion.Title => 1.65,
        DocumentRegion.Heading => 1.45,
        DocumentRegion.Header => 1.20,
        DocumentRegion.Footer => 1.10,
        _ => 1.00
    };

    private static double PositionWeight(DocumentSegment segment, int totalSegments)
    {
        if (segment.Region is DocumentRegion.Header or DocumentRegion.Footer)
        {
            return 1.0;
        }

        if (totalSegments <= 1)
        {
            return 1.0;
        }

        var normalizedPosition = (double)segment.Order / (totalSegments - 1);

        return normalizedPosition switch
        {
            <= 0.05 => 1.35,
            <= 0.20 => 1.20,
            <= 0.50 => 1.08,
            _ => 1.00
        };
    }

    private static double KeywordStrengthWeight(KeywordRule rule)
    {
        if (rule.Strength != KeywordStrength.Normal)
        {
            return rule.Strength switch
            {
                KeywordStrength.Weak => 0.75,
                KeywordStrength.Strong => 1.35,
                _ => 1.0
            };
        }

        // Heuristic fallback when profile doesn't explicitly set strength.
        if (rule.MatchMode == KeywordMatchMode.Exact)
        {
            return 1.35;
        }

        if (rule.MatchMode == KeywordMatchMode.Phrase && rule.Term.Contains(' ', StringComparison.Ordinal))
        {
            return 1.20;
        }

        if (rule.MatchMode == KeywordMatchMode.WholeWord && rule.Term.Length <= 5)
        {
            return 0.85;
        }

        return 1.0;
    }

    private static DocumentRegion ResolveBodyRegion(Paragraph paragraph, int paragraphIndex, string normalizedText)
    {
        var styleName = paragraph.FormattingStyle?.Trim() ?? string.Empty;

        if (IsTitleStyle(styleName) || (paragraphIndex == 0 && LooksLikeTitle(normalizedText)))
        {
            return DocumentRegion.Title;
        }

        if (IsHeadingStyle(styleName))
        {
            return DocumentRegion.Heading;
        }

        return DocumentRegion.Body;
    }

    private static bool IsHeadingStyle(string styleName)
        => styleName.Contains("heading", StringComparison.OrdinalIgnoreCase);

    private static bool IsTitleStyle(string styleName)
        => styleName.Contains("title", StringComparison.OrdinalIgnoreCase);

    private static bool LooksLikeTitle(string text)
    {
        if (string.IsNullOrWhiteSpace(text) || text.Length > 120)
        {
            return false;
        }

        var words = text.Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        return words.Length is >= 1 and <= 10;
    }

    private static double Clamp(double value, double min, double max)
        => value < min ? min : value > max ? max : value;
}
