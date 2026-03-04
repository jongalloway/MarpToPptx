# Sample Decks

This directory contains Marp-style sample decks for manual testing, debugging, and compatibility checks.

## Suggested Progression

1. `01-minimal.md`
Basic smoke test for slide splitting, headings, paragraphs, and lists.

2. `02-directives.md`
Covers front matter plus the directive subset currently supported by `MarpToPptx`.

3. `03-theme-css.md`
Uses `03-theme.css` to exercise theme parsing for fonts, colors, padding, headings, and code styles.

4. `04-content-coverage.md`
Combines images, code blocks, ordered lists, unordered lists, and tables in one deck.

5. `05-compatibility-gaps.md`
Intentionally includes Marp features that are not fully implemented yet so behavior gaps are easy to reproduce and discuss.

## Running The Samples

Run any sample with the published tool:

```bash
dnx MarpToPptx samples/01-minimal.md -o artifacts/samples/01-minimal.pptx
```

Or use the local source project:

```bash
dotnet run --project src/MarpToPptx.Cli -- samples/01-minimal.md -o artifacts/samples/01-minimal.pptx
```

For the themed sample, pass the companion CSS file:

```bash
dotnet run --project src/MarpToPptx.Cli -- samples/03-theme-css.md --theme-css samples/03-theme.css -o artifacts/samples/03-theme-css.pptx
```

## Notes

- Asset references are relative to each sample Markdown file.
- The compatibility-gap sample is useful for validating current limitations without needing to invent ad hoc repros.
- Output paths above target `artifacts/samples/`, but any writable location will work.
