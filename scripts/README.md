# Scripts

This directory contains the main PowerShell scripts that were useful during PPTX testing and troubleshooting.

## Available Helpers

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

### `Invoke-PptxSmokeTest.ps1`

Run the main local PPTX smoke-test flow in one command: generate with the local CLI project, validate with the .NET Open XML helper, and then open the result in PowerPoint.

When you do not provide `-OutputPath`, the script now writes default files using explicit names:

- generated output: `artifacts/smoke-tests/<sample>-generated-<configuration>.pptx`
- PowerPoint-resaved copy: `artifacts/smoke-tests/<sample>-powerpoint-resaved-<configuration>.pptx`

```powershell
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/03-theme-css.md -ThemeCss samples/03-theme.css -Configuration Release
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/01-minimal.md -CiSafe
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown samples/06-remote-assets.md -Configuration Release -AllowRemoteAssets -CiSafe
```

## Notes

- These scripts are intended for local Windows-based troubleshooting.
- `Test-PowerPointOpen.ps1` requires Microsoft PowerPoint to be installed and available through COM interop.
- `Generate-LocalPptx.ps1` is the preferred path for renderer debugging because it executes the current workspace code.
- `Invoke-PptxSmokeTest.ps1` is the quickest end-to-end check before or after renderer/package changes.
- `Invoke-PptxSmokeTest.ps1 -CiSafe` keeps the PowerPoint step when COM automation is available, but automatically falls back to generation plus .NET-hosted Open XML validation on CI agents or other environments without PowerPoint.
- Remote asset smoke coverage is opt-in in the sample-generation helpers so the default local flow stays deterministic when working offline.
- The manual pre-release PowerPoint review checklist is documented in `doc/release-validation.md`.
