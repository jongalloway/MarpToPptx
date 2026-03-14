# Scripts

This directory contains the main PowerShell scripts that were useful during PPTX testing and troubleshooting.

## Available Helpers

### `Invoke-TestConcierge.ps1`

Interactive front door for the existing script set. It asks which workflow you want, gathers the relevant arguments, and then invokes the underlying helper script.

```powershell
pwsh ./scripts/Invoke-TestConcierge.ps1
```

For example, to export `samples/04-content-coverage.md`, launch the concierge, choose `Generate one PPTX from Markdown`, then pick `samples/04-content-coverage.md` from the sample list and accept the default output path if it works for you.

### `Generate-LocalPptx.ps1`

Generate a PPTX with the local CLI project instead of the published `dnx` tool.

```powershell
pwsh ./scripts/Generate-LocalPptx.ps1 -InputMarkdown samples/01-minimal.md -OutputPath artifacts/samples/01-minimal-scripted.pptx
pwsh ./scripts/Generate-LocalPptx.ps1 -InputMarkdown samples/06-remote-assets.md -OutputPath artifacts/samples/06-remote-assets-scripted.pptx -AllowRemoteAssets
```

### `Generate-SamplePptxSet.ps1`

Generate PPTX output for each sample deck in `samples/`, writing results to `artifacts/samples/`. The script skips `samples/README.md` and will auto-pick a companion CSS file such as `samples/03-theme.css` for `samples/03-theme-css.md`.

```powershell
pwsh ./scripts/Generate-SamplePptxSet.ps1
pwsh ./scripts/Generate-SamplePptxSet.ps1 -Configuration Release -Force
pwsh ./scripts/Generate-SamplePptxSet.ps1 -Configuration Release -Force -IncludeRemoteSamples
```

### `Expand-Pptx.ps1`

Expand a `.pptx` package into a normal directory for inspection.

```powershell
pwsh ./scripts/Expand-Pptx.ps1 -PptxPath artifacts/samples/01-minimal-scripted.pptx -Force
```

### `Test-PowerPointOpen.ps1`

Open a PPTX through PowerPoint COM automation and optionally save a copy. This is useful when `OpenXmlValidator` passes but PowerPoint compatibility is still in question.

```powershell
pwsh ./scripts/Test-PowerPointOpen.ps1 -PptxPath artifacts/samples/01-minimal-scripted.pptx
pwsh ./scripts/Test-PowerPointOpen.ps1 -PptxPath artifacts/samples/01-minimal-scripted.pptx -SaveCopyAs artifacts/samples/01-minimal-scripted-roundtrip.pptx
```

### `Compare-PptxStructure.ps1`

Compare two PPTX packages or two expanded package directories. The script reports file inventory differences and content differences in key package XML files.

```powershell
pwsh ./scripts/Compare-PptxStructure.ps1 -PathA artifacts/samples/01-minimal-scripted.pptx -PathB artifacts/samples/01-minimal-marp-cli.pptx
```

### `Export-PptxSlides.ps1`

Export each slide in a `.pptx` file as an image using PowerPoint COM automation. Slide images are written to `artifacts/slide-exports/<pptx-name>/` by default, with stable filenames such as `slide-001.png`.

```powershell
pwsh ./scripts/Export-PptxSlides.ps1 -PptxPath artifacts/smoke-tests/01-minimal-generated-debug.pptx
pwsh ./scripts/Export-PptxSlides.ps1 -PptxPath artifacts/smoke-tests/01-minimal-generated-debug.pptx -Format jpg
pwsh ./scripts/Export-PptxSlides.ps1 -PptxPath artifacts/smoke-tests/01-minimal-generated-debug.pptx -OutputDirectory artifacts/slide-exports/my-review
```

The script requires Microsoft PowerPoint to be installed. Pass `-CiSafe` to skip the export silently when PowerPoint is unavailable rather than failing.

To generate a PPTX and export its slide images in one command, use the `-ExportSlides` flag on `Invoke-PptxSmokeTest.ps1`:

```powershell
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md -ExportSlides
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/04-content-coverage.md -ExportSlides -SlideExportFormat jpg -CiSafe
```

### `Invoke-PptxSmokeTest.ps1`

Run the main local PPTX smoke-test flow in one command: generate with the local CLI project, validate with the .NET Open XML helper, audit rendered text contrast, and then open the result in PowerPoint.

When you do not provide `-OutputPath`, the script now writes default files using explicit names:

- generated output: `artifacts/smoke-tests/<sample>-generated-<configuration>.pptx`
- contrast audit report: `artifacts/smoke-tests/<sample>-generated-<configuration>-contrast-audit.txt`
- PowerPoint-resaved copy: `artifacts/smoke-tests/<sample>-powerpoint-resaved-<configuration>.pptx`

Add `-ExportSlides` to automatically export slide images after the smoke test completes. Use `-SlideExportFormat` to choose `jpg` or `png` (default `png`), and `-SlideExportDirectory` to override the output location.

```powershell
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/03-theme-css.md -ThemeCss samples/03-theme.css -Configuration Release
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md -CiSafe
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/06-remote-assets.md -Configuration Release -AllowRemoteAssets -CiSafe
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md -ExportSlides
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/04-content-coverage.md -ExportSlides -SlideExportFormat jpg -CiSafe
```

### `Invoke-AllPptxSmokeTests.ps1`

Run the smoke-test flow for each sample deck in `samples/`. By default, the script skips `samples/README.md`, the compatibility-gap repro sample, and remote-asset samples unless they are explicitly enabled.

The aggregate smoke suite uses contrast auditing in `Selected` mode by default. That currently audits `01-minimal.md`, `04-content-coverage.md`, and `07-presenter-notes.md`, which gives CI a stable rendering-focused subset without failing on decks that intentionally exercise contrast-unstable theme or directive combinations. Use `-ContrastAuditMode All` to audit every selected sample or `-ContrastAuditMode None` to disable contrast checks for the batch run.

```powershell
pwsh ./scripts/Invoke-AllPptxSmokeTests.ps1 -Configuration Release -CiSafe
pwsh ./scripts/Invoke-AllPptxSmokeTests.ps1 -Configuration Release -CiSafe -IncludeRemoteSamples
pwsh ./scripts/Invoke-AllPptxSmokeTests.ps1 -Configuration Release -CiSafe -OnlyRemoteSamples
pwsh ./scripts/Invoke-AllPptxSmokeTests.ps1 -Configuration Release -CiSafe -IncludeRemoteSamples -ContinueOnError
```

## Notes

- These scripts are intended for local Windows-based troubleshooting.
- `Test-PowerPointOpen.ps1` requires Microsoft PowerPoint to be installed and available through COM interop.
- `Export-PptxSlides.ps1` requires Microsoft PowerPoint to be installed and available through COM interop.
- `Generate-LocalPptx.ps1` is the preferred path for renderer debugging because it executes the current workspace code.
- `Invoke-TestConcierge.ps1` is the easiest way to discover the local test and export flows without remembering script parameters.
- `Invoke-PptxSmokeTest.ps1` is the quickest end-to-end check before or after renderer/package changes.
- `Invoke-AllPptxSmokeTests.ps1` is the quickest way to run the full local smoke suite across the sample directory.
- `Invoke-PptxSmokeTest.ps1` runs a contrast audit by default and fails the smoke test when rendered text does not meet the auditor's thresholds. Use `-SkipContrastAudit` only when contrast is not the behavior under investigation.
- `Invoke-PptxSmokeTest.ps1 -CiSafe` keeps the PowerPoint step when COM automation is available, but automatically falls back to generation plus .NET-hosted Open XML validation on CI agents or other environments without PowerPoint.
- `Export-PptxSlides.ps1 -CiSafe` and `Invoke-PptxSmokeTest.ps1 -ExportSlides -CiSafe` skip the slide export step gracefully when PowerPoint is unavailable.
- Remote asset smoke coverage is opt-in in the sample-generation helpers so the default local flow stays deterministic when working offline.
- The manual pre-release PowerPoint review checklist is documented in `doc/release-validation.md`.
