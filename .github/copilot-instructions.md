# MarpToPptx Workspace Instructions

## Project focus

- MarpToPptx is a .NET 10 CLI and library that compiles Marp-flavored Markdown into editable PowerPoint `.pptx` files.
- Treat `DocumentFormat.OpenXml` as the primary implementation surface for PPTX generation.
- Use ShapeCrawler as a reference source for PPTX-specific patterns when useful, not as an automatic dependency choice.

## Important repo context

- PowerPoint compatibility is the real success criterion. A package can pass `OpenXmlValidator` and still fail to open or trigger repair in PowerPoint.
- When changing PPTX generation, preserve package structure and relationship invariants documented in `doc/pptx-compatibility-notes.md`.
- For `DocumentFormat.OpenXml` `3.4.1`, consult `doc/openxml-3.4.1-audit.md` before assuming newer SDK surface removes manual package-shaping work.
- Prefer minimal, compatibility-focused changes over broad renderer refactors.

## Validation and testing workflow

- For renderer or package changes, follow this order:
  1. Run targeted tests in `tests/MarpToPptx.Tests`.
  2. Generate a sample deck with the local CLI project.
  3. Validate with `src/MarpToPptx.OpenXmlValidator`.
  4. If compatibility risk remains, run the PowerPoint smoke flow on Windows.
- Use `dotnet run --project src/MarpToPptx.Cli -- ...` for local renderer debugging. Do not rely on `dnx` unless you explicitly want the published tool path.
- Use `pwsh ./scripts/Invoke-PptxSmokeTest.ps1 -InputMarkdown ...` for the quickest end-to-end local check.
- If golden package baselines intentionally change, regenerate them with `UPDATE_GOLDEN_PACKAGES=1` during the test run.

## Useful scripts and references

- Start with `scripts/README.md` for generation, expansion, structure comparison, and PowerPoint-open verification helpers.
- Use `doc/pptx-compatibility-notes.md` when debugging repaired or non-opening PPTX output.
- Use `doc/marp-markdown.md` for repo-specific Marp behavior and directive expectations.
- Use these external references only when needed for authoritative format or API details:
  - Open XML SDK: https://github.com/dotnet/Open-XML-SDK
  - Open XML presentation overview: https://learn.microsoft.com/en-us/office/open-xml/presentation/overview
  - ShapeCrawler source: https://github.com/ShapeCrawler/ShapeCrawler

## Working guidance for agents

- If validation passes but PowerPoint fails, inspect package structure, relationships, and content types before changing slide content logic.
- Compare generated output against a known-good PowerPoint or reference package when troubleshooting package-shape issues.
- Keep instructions files small. Put specialized workflows into prompts or skills instead of expanding this file.