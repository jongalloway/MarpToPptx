# Release Validation

This document defines the heavier pre-release validation flow for `MarpToPptx`.

## CI Smoke Test Gate

Every pull request and push to `main` runs the `CI` workflow (`.github/workflows/ci.yml`), which includes:

1. **Build, test, and pack** — restore, build, unit tests, NuGet pack.
2. **Smoke test** — generate PPTX from every non-remote sample in `samples/` using `Invoke-AllPptxSmokeTests.ps1`, then validate each result with the Open XML SDK helper.
3. **Contrast audit** — run `Invoke-ContrastAudit.ps1` across the generated smoke-test PPTX files using the `MarpToPptx.ContrastAuditor` tool. This step is informational for pull requests (`continue-on-error: true`) so it reports regressions without blocking merges, but audit results and the JSON report are always uploaded as artifacts.

The contrast audit catches text/background color-contrast regressions (WCAG thresholds) that the Open XML schema validator cannot detect.

## Hosted Release Gate

Use the GitHub Actions workflow `.github/workflows/release-gate.yml` for the automated hosted pre-release gate.

Workflow name: `Release Gate`

The release gate runs two parallel jobs.

### `hosted-release-gate` (Ubuntu)

Purpose:

- run restore, build, tests, and packaging in the selected configuration
- generate the full sample set, including the remote-assets sample and the opt-in theme fixtures under `samples/themes/`
- validate every generated PPTX with the Open XML SDK helper (`--no-build`)
- **run the contrast auditor on every generated PPTX** — this is a hard failure gate at release time
- import every generated PPTX into LibreOffice Impress and export it to PDF
- upload the generated PPTX, contrast audit report (`contrast-audit-report.json`), PDF, and package artifacts for manual inspection

This gate is stronger than normal CI because it checks every sample deck, enforces the contrast audit, and forces a second PPTX consumer (LibreOffice) to load the generated files.

### `windows-powerpoint-validation` (Windows)

Purpose:

- validate cross-platform build and generation on a Windows runner
- generate the full sample set (excluding remote-asset samples for determinism)
- validate every generated PPTX with the Open XML SDK helper
- run the contrast auditor on every generated PPTX
- attempt to open each PPTX in PowerPoint Desktop via COM automation (skipped automatically when PowerPoint is not available on the runner)
- export slide images from each PPTX via PowerPoint COM automation (skipped automatically when PowerPoint is not available)
- upload all generated PPTX, contrast audit report, and slide exports as artifacts

On GitHub-hosted `windows-latest` runners, PowerPoint is not installed, so the COM steps are skipped gracefully. If you have a self-hosted Windows runner with PowerPoint installed, these steps run fully and produce slide export images alongside the contrast audit report.

## Manual PowerPoint Review

PowerPoint Desktop review is a manual release checklist, not an automated GitHub-hosted gate.

Reason:

- GitHub-hosted runners do not provide a supported PowerPoint Desktop environment for COM-based validation.
- The most reliable PowerPoint check is still a human opening the generated decks in PowerPoint and reviewing them visually.

### Generate Review Artifacts

From the repo root on a Windows machine with PowerPoint installed:

```powershell
pwsh ./scripts/Generate-SamplePptxSet.ps1 -Configuration Release -OutputDirectory artifacts/manual-review -IncludeThemeSamples -IncludeRemoteSamples -Force
```

To also run the contrast audit locally before opening in PowerPoint:

```powershell
pwsh ./scripts/Invoke-ContrastAudit.ps1 -PptxDirectory artifacts/manual-review -Configuration Release -ReportPath artifacts/manual-review/contrast-audit-report.json -ContinueOnError
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
- presenter notes appear on the expected slides in `07-presenter-notes.pptx`, and the control slide has no notes

### Sample Focus Areas

- `01-minimal.pptx`: baseline slide generation
- `02-directives.pptx`: front matter, carry-forward directives, background image and size behavior, headers, footers, paginate
- `03-theme-css.pptx`: theme parsing, heading hierarchy, class variants, layout-sensitive typography
- `04-content-coverage.pptx`: images, tables, code, MP3, M4A, video
- `05-compatibility-gaps.pptx`: current non-goals, known-approximation behavior, and recently implemented compatibility regression checks
- `06-remote-assets.pptx`: real HTTP(S) image fetches
- `07-presenter-notes.pptx`: explicit presenter notes packaging and slide-to-notes attachment

### Theme Fixture Focus Areas

- `09-community-beam.pptx`: academic-style theme fixture with assertive heading and contrast-slide styling
- `10-community-graph-paper.pptx`: local SVG patterned background fixture for theme background fidelity
- `11-community-wave.pptx`: local SVG/art background fixture for bold section-band rendering
- `12-community-dracula.pptx`: dark-theme fixture stressing contrast and code-heavy slide treatment
- `13-popular-gaia.pptx`: bright-theme fixture stressing large-scale typography and clean-surface layout

## Current Exclusions

- Generated-deck-as-template validation is not part of the release gate while [issue #66](https://github.com/jongalloway/MarpToPptx/issues/66) is open.
