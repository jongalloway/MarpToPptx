# Contributing to MarpToPptx

## Repository Structure

- `src/MarpToPptx.Core` — semantic slide model, Markdown parsing, theme parsing, layout planning
- `src/MarpToPptx.Pptx` — Open XML PPTX rendering and template-aware presentation generation
- `src/MarpToPptx.Cli` — `marp2pptx` command-line entrypoint
- `src/MarpToPptx.OpenXmlValidator` — .NET validation helper used by smoke tests and CI
- `scripts/` — PowerShell helpers for local generation, smoke tests, package inspection, and PowerPoint troubleshooting
- `tests/MarpToPptx.Tests` — xUnit v3 tests running on Microsoft Testing Platform
- `samples/` — Marp-style sample decks for smoke tests, feature coverage, theme parsing, and compatibility-gap repros
- `.github/copilot-instructions.md` — repo-specific Copilot guidance for PPTX compatibility, testing flow, and reference sources
- `.github/prompts/` — reusable prompts for PPTX compatibility investigation and implementation review
- `.github/workflows/ci.yml` — Ubuntu build/test/pack plus an Ubuntu CI-safe PPTX smoke-test job

## Conventions

- Solution format: `MarpToPptx.slnx`
- Centralized package management: `Directory.Packages.props`
- Test framework: xUnit v3
- Test runner: Microsoft Testing Platform via `global.json`
- CLI packaging: `marp2pptx` as a .NET tool, with single-file publish as an alternate deployment mode

## Running From Source

```bash
dotnet run --project src/MarpToPptx.Cli -- input.md -o output.pptx
dotnet run --project src/MarpToPptx.Cli -- input.md --template theme.pptx
dotnet run --project src/MarpToPptx.Cli -- input.md --theme-css theme.css
```

## Local Packaging

Build a local tool package:

```bash
dotnet pack src/MarpToPptx.Cli -c Release
```

This produces a tool package under `artifacts/nupkg/` with package ID `MarpToPptx` and command name `marp2pptx`.

Run it with `dnx` from the local package source:

```bash
dnx MarpToPptx --add-source ./artifacts/nupkg sample.md -o sample.pptx
```

Or install it as a local tool:

```bash
dotnet new tool-manifest
dotnet tool install MarpToPptx --add-source ./artifacts/nupkg
dotnet tool run marp2pptx sample.md -o sample.pptx
```

## Releases

NuGet publishing is handled by `.github/workflows/publish.yml` using nuget.org Trusted Publishing with GitHub OIDC.

Pre-release validation is available through `.github/workflows/release-gate.yml`.

- Versioning is tag-based via `MinVer`.
- Stable release tags use the form `v1.2.3`.
- The publish workflow builds, tests, packs, and pushes the tool package from `artifacts/nupkg/`.

To cut a release:

```bash
git tag v1.2.3
git push origin v1.2.3
```

Or use the helper script:

```powershell
pwsh ./scripts/New-ReleaseTag.ps1 -Version 1.2.3
```

The script requires a clean, up-to-date `main` branch and asks you to confirm that release validation is complete before it creates and pushes the tag. Use `-Force` only when you intentionally need to bypass those checks.

After the workflow finishes and NuGet indexing completes, install or run the published tool with:

```bash
dotnet tool install --global MarpToPptx
dnx MarpToPptx sample.md -o sample.pptx
```

## Reference Documentation

- [Marp Markdown behavior and directives](doc/marp-markdown.md)
- [PPTX compatibility notes](doc/pptx-compatibility-notes.md)
- [`DocumentFormat.OpenXml` 3.4.1 audit](doc/openxml-3.4.1-audit.md)
- [VS Code workflow integration](doc/vscode-workflow.md)
- [Release validation and checklist](doc/release-validation.md)
- [Script helpers](scripts/README.md)
