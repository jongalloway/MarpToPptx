# MarpToPptx

MarpToPptx is a .NET 10 CLI and library for compiling Marp-flavored Markdown into editable PowerPoint presentations.

## Current Structure

- `src/MarpToPptx.Core`: semantic slide model, Markdown parsing, theme parsing, layout planning
- `src/MarpToPptx.Pptx`: Open XML PPTX rendering and template-aware presentation generation
- `src/MarpToPptx.Cli`: `marp2pptx` command-line entrypoint
- `tests/MarpToPptx.Tests`: xUnit v3 tests running on Microsoft Testing Platform

## Usage

```bash
dotnet run --project src/MarpToPptx.Cli -- input.md -o output.pptx
dotnet run --project src/MarpToPptx.Cli -- input.md --template theme.pptx
dotnet run --project src/MarpToPptx.Cli -- input.md --theme-css theme.css
```

## Local Tool And Dnx

Build a local tool package:

```bash
dotnet pack src/MarpToPptx.Cli -c Release
```

That produces a tool package under `artifacts/nupkg/` with package ID `MarpToPptx` and command name `marp2pptx`.

Run it with `dnx` from the local package source:

```bash
dnx MarpToPptx --add-source ./artifacts/nupkg sample.md -o sample.pptx
```

You can also install it as a local tool from the package output:

```bash
dotnet new tool-manifest
dotnet tool install MarpToPptx --add-source ./artifacts/nupkg
dotnet tool run marp2pptx sample.md -o sample.pptx
```

Public `dnx MarpToPptx ...` usage depends on publishing the package to a NuGet feed. That follow-up is tracked as a separate GitHub issue.

## Releases

NuGet publishing is handled by `.github/workflows/publish.yml` using nuget.org Trusted Publishing with GitHub OIDC.

- Versioning is tag-based via `MinVer`.
- Stable release tags should use the form `v1.2.3`.
- The publish workflow builds, tests, packs, and then pushes the tool package from `artifacts/nupkg/`.

Before the first publish, configure nuget.org Trusted Publishing for:

- Owner: `jongalloway`
- Repository: `MarpToPptx`
- Workflow file: `publish.yml`

The workflow also expects a repository secret named `NUGET_USER` containing the nuget.org profile name used for Trusted Publishing.

To cut a release:

```bash
git tag v1.2.3
git push origin v1.2.3
```

After the workflow finishes and NuGet indexing completes, install or run the published tool with:

```bash
dotnet tool install MarpToPptx --global
dnx MarpToPptx sample.md -o sample.pptx
```

## Repository Conventions

- Solution format: `MarpToPptx.slnx`
- Centralized package management: `Directory.Packages.props`
- Test framework: xUnit v3
- Test runner: Microsoft Testing Platform via `global.json`
- CLI packaging direction: `marp2pptx` as a .NET tool, while preserving single-file publish as a deployment mode

## Current Milestone

- Marp-style front matter and directive parsing
- Slide splitting on `---`
- Semantic slide model independent from PPTX
- Basic theme extraction for font families, font sizes, colors, and padding
- PPTX generation for headings, paragraphs, bullet lists, images, and code blocks
- Table content fallback rendered as editable text while native PPTX table generation remains a product requirement
- Template-copy workflow for reusing an existing `.pptx` theme/master

## Steering Decisions

- `ImageSharp` is intentionally not used for image sizing unless it is explicitly re-approved after licensing review
- Intrinsic image sizing should prefer built-in platform capabilities or a minimal in-project metadata reader
- Remaining roadmap work should be evaluated against the current implemented baseline rather than an empty starting point

## Roadmap

- Improve CSS coverage for more Marp theme features
- Refine layout heuristics for denser or highly designed decks
- Expand template integration to map multiple layouts intelligently
- Add native PPTX table generation and richer table styling
- Add code block syntax highlighting
- Support remote assets and additional image formats
