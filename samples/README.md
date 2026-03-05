# Sample Decks

This directory contains Marp-style sample decks for manual testing, debugging, and compatibility checks.

## Suggested Progression

1. `01-minimal.md`
Basic smoke test for slide splitting, headings, paragraphs, and lists.

2. `02-directives.md`
Covers front matter plus the supported slide directives, including background, header, footer, and pagination overrides.

3. `03-theme-css.md`
Uses `03-theme.css` to exercise theme parsing for fonts, colors, padding, background images, typography, headings, and code styles.

4. `04-content-coverage.md`
Combines images, syntax-highlighted code blocks, local MP3/M4A audio, local video, ordered lists, unordered lists, and native tables in one deck.

5. `05-compatibility-gaps.md`
Intentionally includes Marp features that are not fully implemented yet so behavior gaps are easy to reproduce and discuss.

6. `06-remote-assets.md`
Opt-in integration smoke deck for real HTTP(S) image fetches using commit-pinned raw GitHub URLs.

7. `07-presenter-notes.md`
Dedicated smoke deck for presenter notes packaging, including note-bearing slides, a no-notes control slide, and a slide that mixes directive comments with presenter notes.

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
- `04-content-coverage.md` depends on the small local media fixtures under `samples/assets/`.
- `06-remote-assets.md` is intended for integration smoke testing and should be run with remote assets explicitly enabled.
- `07-presenter-notes.md` is the explicit smoke deck for speaker notes and PowerPoint-open compatibility of emitted notes parts.
- The compatibility-gap sample is useful for validating current limitations without needing to invent ad hoc repros.
- Output paths above target `artifacts/samples/`, but any writable location will work.
