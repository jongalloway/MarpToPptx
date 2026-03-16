---
name: export-current-deck-with-theme-css
description: Export the current Marp Markdown deck to PPTX with a Marp-compatible CSS theme by using the published MarpToPptx CLI.
---

# Export Current Deck With Theme CSS

Use this skill when the active Marp deck should be exported with a companion CSS theme instead of a PowerPoint template.

## Purpose

Run the published `MarpToPptx` CLI with `--theme-css` so the output deck follows the content repo's Marp theme styling.

## Preferred Commands

```bash
dnx MarpToPptx path/to/deck.md --theme-css path/to/theme.css -o path/to/deck.pptx
```

```bash
marp2pptx path/to/deck.md --theme-css path/to/theme.css -o path/to/deck.pptx
```

## Guidance

1. Confirm the Markdown deck path.
2. Look for an explicit CSS path from the user first.
3. If none is provided, check for a nearby companion stylesheet that matches the deck name or the repo's `themes/` folder convention.
4. Keep the output path adjacent to the source deck by default.
5. Use the published CLI contract only.

## Important Rules

- Prefer `--theme-css` when the user's styling source is Marp CSS rather than a PowerPoint template.
- Do not assume a CSS file exists; verify it or ask for the path.
- Do not build `MarpToPptx` from source for this end-user workflow.

## Output Expectations

- Report the exact command used.
- Report the CSS file that was applied.
- Report the generated `.pptx` path.