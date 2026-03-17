# Sample Decks

This directory contains the default Marp-style sample decks for manual testing, debugging, and compatibility checks.

Theme-oriented example decks live under `samples/themes/`. They are intentionally outside the default top-level sample scan so the default smoke scripts and CI behavior stay unchanged, but release validation can opt in to them explicitly.

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
Mixed compatibility and regression deck that keeps unsupported Marp features easy to reproduce while also checking recently implemented directive behavior.

6. `06-remote-assets.md`
Opt-in integration smoke deck for real HTTP(S) image fetches using commit-pinned raw GitHub URLs.

7. `07-presenter-notes.md`
Dedicated smoke deck for presenter notes packaging, including note-bearing slides, a no-notes control slide, and a slide that mixes directive comments with presenter notes.

8. `08-showcase.md`
Speaker-style showcase deck generated from repo content, using `08-showcase.css` plus local Marp SVG assets to cover the Marp ecosystem, MarpToPptx capabilities, presenter notes, and the recommended VS Code task configuration.

9. `09-diagrams.md`
Diagram-focused sample deck that mixes Mermaid fences and `diagram` fences to exercise DiagramForge-backed flowchart, block, state, mindmap, matrix, pyramid, funnel, and radial output with the companion `09-diagrams.css` theme.

## Theme Decks

These live under `samples/themes/` and are intended as a second-tier theme/example suite rather than part of the default root smoke set.

1. `themes/09-community-beam.md`
Academic-style smoke deck inspired by community Beam and Beamer-like Marp themes, using `themes/09-community-beam.css` to stress assertive headings, contrast slides, and appendix-style typography.

2. `themes/10-community-graph-paper.md`
Notebook-style smoke deck inspired by community Graph Paper themes, using `themes/10-community-graph-paper.css` plus a local SVG grid background to validate patterned backgrounds without remote assets.

3. `themes/11-community-wave.md`
Conference-talk smoke deck inspired by community Wave-style themes, using `themes/11-community-wave.css` plus local background art to exercise bold section bands and compact closing-slide typography.

4. `themes/12-community-dracula.md`
Dark-theme smoke deck inspired by the community Dracula theme, using `themes/12-community-dracula.css` to stress high-contrast text, saturated accents, and code-heavy slides.

5. `themes/13-popular-gaia.md`
Bright-theme smoke deck inspired by Marp's popular built-in Gaia direction, using `themes/13-popular-gaia.css` to exercise large-scale typography, clean surfaces, and bold section breaks.

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
- `08-showcase.md` is the branded speaker-style sample that exercises the batch smoke runner with a richer content deck and companion CSS.
- `09-diagrams.md` is the focused DiagramForge sample deck for Mermaid plus conceptual diagram coverage using the companion `09-diagrams.css` theme.
- `samples/themes/09-community-beam.md`, `samples/themes/10-community-graph-paper.md`, `samples/themes/11-community-wave.md`, and `samples/themes/12-community-dracula.md` are repo-authored fixtures inspired by themes listed in Awesome Marp's community themes section; they are not vendored copies of upstream sample decks.
- `samples/themes/13-popular-gaia.md` is a repo-authored fixture inspired by Marp's built-in Gaia theme direction rather than a community theme listing.
- The current generation and smoke scripts only scan top-level Markdown files under `samples/` by default, so `samples/themes/` stays out of the default CI/local smoke suite unless a caller opts in explicitly.
- The compatibility watchlist sample is useful for validating current limitations and guarding recently implemented compatibility features without needing ad hoc repro decks.
- Output paths above target `artifacts/samples/`, but any writable location will work.
