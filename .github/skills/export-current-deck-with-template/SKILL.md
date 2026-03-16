---
name: export-current-deck-with-template
description: Export the current Marp Markdown deck to PPTX by applying an existing PowerPoint template with the published MarpToPptx CLI.
---

# Export Current Deck With Template

Use this skill when the user wants the active Marp deck rendered into an editable `.pptx` that inherits masters and layouts from an existing PowerPoint template.

## Purpose

Run the published `MarpToPptx` CLI with `--template` so the generated deck uses a shared PowerPoint design.

## Preferred Commands

```bash
dnx MarpToPptx path/to/deck.md --template path/to/theme.pptx -o path/to/deck.pptx
```

```bash
marp2pptx path/to/deck.md --template path/to/theme.pptx -o path/to/deck.pptx
```

## Guidance

1. Confirm the Markdown deck path and template `.pptx` path.
2. Use the published CLI rather than cloning or building the `MarpToPptx` source repo.
3. Keep the output path next to the source deck by default unless the user wants a different artifact location.
4. If the deck uses template-directed layouts such as `layout: Title and Content` or `<!-- _layout: Template[1] -->`, preserve that authoring model rather than rewriting the source.
5. If the user is unsure whether they need a template or CSS theming, use a template when they want PowerPoint masters/layouts and use `--theme-css` when they want Marp-style CSS theming.

## Important Rules

- Do not treat a previous generated deck as the same thing as a rendering template.
- Do not switch to local-source execution or repo-maintainer smoke scripts for normal end-user export.
- If a template path is missing, ask for it rather than inventing one.

## Output Expectations

- Report the exact command used.
- Report the template path that was applied.
- Report the generated `.pptx` path.