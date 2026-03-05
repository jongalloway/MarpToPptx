# Release Validation

This document defines the heavier pre-release validation flow for `MarpToPptx`.

## Hosted Release Gate

Use the GitHub Actions workflow `.github/workflows/release-gate.yml` for the automated hosted pre-release gate.

Workflow name: `Release Gate`

Purpose:

- run restore, build, tests, and packaging in the selected configuration
- generate the full sample set, including the remote-assets sample
- validate every generated PPTX with the Open XML SDK helper
- import every generated PPTX into LibreOffice Impress and export it to PDF
- upload the generated PPTX, PDF, and package artifacts for manual inspection

This gate is stronger than normal CI because it checks every sample deck and forces a second PPTX consumer to load the generated files.

This gate is still not equivalent to PowerPoint Desktop compatibility.

## Manual PowerPoint Review

PowerPoint Desktop review is a manual release checklist, not an automated GitHub-hosted gate.

Reason:

- GitHub-hosted runners do not provide a supported PowerPoint Desktop environment for COM-based validation.
- The most reliable PowerPoint check is still a human opening the generated decks in PowerPoint and reviewing them visually.

### Generate Review Artifacts

From the repo root on a Windows machine with PowerPoint installed:

```powershell
pwsh ./scripts/Generate-SamplePptxSet.ps1 -Configuration Release -OutputDirectory artifacts/manual-review -IncludeRemoteSamples -Force
```

If you want the standard Open XML validation at the same time for a specific deck, run the existing smoke helper:

```powershell
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md -Configuration Release
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/06-remote-assets.md -Configuration Release -AllowRemoteAssets
```

### Review Checklist

Open each generated deck from `artifacts/manual-review/` in PowerPoint Desktop and confirm:

- the file opens without a repair dialog, corruption message, or content-loss warning
- slide count matches the source markdown expectations
- text layout is readable and does not visibly overflow or overlap in unexpected ways
- background color and background image behavior match the sample intent
- headers, footers, and page numbers appear on the expected slides
- class-variant theme styling appears correctly on `03-theme-css.pptx`
- code blocks, lists, and tables remain editable and visually coherent
- local media placeholders and embedded media objects appear as expected on `04-content-coverage.pptx`
- remote asset slides render correctly on `06-remote-assets.pptx`

### Sample Focus Areas

- `01-minimal.pptx`: baseline slide generation
- `02-directives.pptx`: front matter, carry-forward directives, background image and size behavior, headers, footers, paginate
- `03-theme-css.pptx`: theme parsing, heading hierarchy, class variants, layout-sensitive typography
- `04-content-coverage.pptx`: images, tables, code, MP3, M4A, video
- `05-compatibility-gaps.pptx`: current non-goals and known-approximation behavior
- `06-remote-assets.pptx`: real HTTP(S) image fetches

## Current Exclusions

- Generated-deck-as-template validation is not part of the release gate while [issue #66](https://github.com/jongalloway/MarpToPptx/issues/66) is open.
