using System.Text.Json;
using System.Text.Json.Serialization;
using TXTextControl.Document.Classification.Models;

namespace TXTextControl.Document.Classification;

internal static class DocumentClassificationProfileLoader
{
    private const string DefaultFileName = "classification-profiles.json";

    public static IReadOnlyList<DocumentClassificationProfile> LoadFromFile(string? filePath = null)
    {
        var effectivePath = ResolvePath(filePath);
        if (!File.Exists(effectivePath))
        {
            throw new FileNotFoundException($"Profile file was not found: {effectivePath}", effectivePath);
        }

        var json = File.ReadAllText(effectivePath);
        var profiles = JsonSerializer.Deserialize<List<DocumentClassificationProfile>>(json, CreateSerializerOptions())
            ?? throw new InvalidOperationException("Profile file is empty or invalid.");

        Validate(profiles);
        return profiles;
    }

    private static string ResolvePath(string? filePath)
    {
        if (!string.IsNullOrWhiteSpace(filePath))
        {
            return Path.GetFullPath(filePath);
        }

        return Path.Combine(AppContext.BaseDirectory, DefaultFileName);
    }

    private static JsonSerializerOptions CreateSerializerOptions() => new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        Converters = { new JsonStringEnumConverter() }
    };

    private static void Validate(IReadOnlyList<DocumentClassificationProfile> profiles)
    {
        if (profiles.Count == 0)
        {
            throw new InvalidOperationException("At least one profile is required.");
        }

        var duplicateProfile = profiles
            .GroupBy(p => p.Name, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(g => g.Count() > 1);

        if (duplicateProfile is not null)
        {
            throw new InvalidOperationException($"Duplicate profile name found: '{duplicateProfile.Key}'.");
        }

        foreach (var profile in profiles)
        {
            if (string.IsNullOrWhiteSpace(profile.Name))
            {
                throw new InvalidOperationException("Profile name cannot be empty.");
            }

            if (profile.Rules.Count == 0)
            {
                throw new InvalidOperationException($"Profile '{profile.Name}' must have at least one rule.");
            }

            var duplicateKeyword = profile.Rules
                .GroupBy(r => r.Term, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault(g => g.Count() > 1);

            if (duplicateKeyword is not null)
            {
                throw new InvalidOperationException(
                    $"Profile '{profile.Name}' contains duplicate keyword '{duplicateKeyword.Key}'.");
            }

            foreach (var rule in profile.Rules)
            {
                if (string.IsNullOrWhiteSpace(rule.Term))
                {
                    throw new InvalidOperationException($"Profile '{profile.Name}' has an empty keyword rule.");
                }

                if (rule.Weight <= 0)
                {
                    throw new InvalidOperationException(
                        $"Profile '{profile.Name}' has non-positive weight for keyword '{rule.Term}'.");
                }
            }
        }
    }
}
