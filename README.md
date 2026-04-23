# TXTextControl.Document.Classification

A .NET 10 console application that classifies `.docx` documents into predefined categories using weighted keyword profiles.

## What this project does

The classifier:

- Loads classification profiles from `classification-profiles.json`.
- Extracts text from:
  - body paragraphs,
  - each section's headers and footers.
- Detects structure-aware regions (`Title`, `Heading`, `Body`, `Header`, `Footer`).
- Scores category rules using:
  - keyword match mode (`WholeWord`, `Phrase`, `Exact`),
  - keyword `strength` (`Weak`, `Normal`, `Strong`),
  - region weight,
  - position-aware weighting,
  - diminishing returns for repeated matches.
- Returns:
  - predicted category,
  - confidence,
  - per-category score details and keyword hits.

## Tech stack

- Target framework: `.NET 10` (`net10.0`)
- Main package: `TXTextControl.TextControl.Core.SDK` (`34.0.3`)

## Project structure

- `Program.cs` - CLI entry point.
- `Classification/DocxKeywordClassifier.cs` - extraction + scoring pipeline.
- `Classification/DocumentClassificationProfileLoader.cs` - profile load/validation.
- `Classification/Models/*` - core model types.
- `classification-profiles.json` - category and rule configuration.
- `Documents/` - sample `.docx` files.

## Run the classifier

> Pass the document path as the first argument.

```powershell
cd TXTextControl.Document.Classification
dotnet run --project TXTextControl.Document.Classification.csproj -- "Documents/Sample Resume.docx"
```

Or from repo root:

```powershell
dotnet run --project "TXTextControl.Document.Classification/TXTextControl.Document.Classification.csproj" -- "TXTextControl.Document.Classification/Documents/Sample Resume.docx"
```

## Example output

```text
Document: TXTextControl.Document.Classification/Documents/Sample Resume.docx
Classification: Resume
Confidence: 68.26%
Scores:
- Resume: 28.33
- Report: 8.74
- Financial: 7.11
...
```

## Profile format (`classification-profiles.json`)

Each profile has a `name` and `rules` collection. Each rule supports:

- `term` (string)
- `weight` (number > 0)
- `matchMode` (`WholeWord`, `Phrase`, `Exact`)
- `strength` (`Weak`, `Normal`, `Strong`)
- optional: `caseSensitive` (bool)
- optional: `targetRegion` (`Title`, `Heading`, `Body`, `Header`, `Footer`)

Minimal example:

```json
[
  {
    "name": "Resume",
    "rules": [
      {
        "term": "work experience",
        "weight": 3.0,
        "matchMode": "Phrase",
        "strength": "Strong"
      }
    ]
  }
]
```

## Notes

- Only `.docx` files are supported.
- Profiles are validated at startup (duplicate categories/terms, empty terms, non-positive weights).
- If no profile scores above zero, result category is `Unknown`.
