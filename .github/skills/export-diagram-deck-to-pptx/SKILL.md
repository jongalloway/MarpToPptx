---
name: export-diagram-deck-to-pptx
description: Export a Marp deck that uses Mermaid or diagram fences to PPTX with the published MarpToPptx CLI.
---

# Export Diagram Deck To PPTX

Use this skill when the user is working on a Marp deck that contains Mermaid fences or `diagram` fences and wants an editable PowerPoint export.

## Purpose

Export a diagram-heavy deck through the published `MarpToPptx` CLI while preserving the normal Marp authoring flow.

## Detection

Treat the deck as diagram-focused when it contains fences such as:

````md
```mermaid
flowchart LR
```
````

````md
```diagram
diagram: matrix
```
````

## Preferred Commands

If the deck has a companion CSS theme:

```bash
dnx MarpToPptx path/to/deck.md --theme-css path/to/deck.css -o path/to/deck.pptx
```

Without a CSS theme:

```bash
dnx MarpToPptx path/to/deck.md -o path/to/deck.pptx
```

Installed-tool variants are also fine:

```bash
marp2pptx path/to/deck.md --theme-css path/to/deck.css -o path/to/deck.pptx
```

## Guidance

1. Confirm the active file is a Markdown deck.
2. Check whether it contains `mermaid` or `diagram` fences.
3. If a companion CSS file exists and is part of the author's workflow, pass it with `--theme-css`.
4. Preserve any deck-level diagram options such as `diagram-theme` in Markdown rather than trying to rewrite the deck automatically.
5. After export, recommend opening the `.pptx` to review diagram layout and readability.

## Important Rules

- Use the published CLI path, not this repo's local PowerShell smoke scripts, for normal end-user authoring workflows.
- Do not remove or rewrite Mermaid or `diagram` fences just to make export succeed.
- If the user reports styling problems, first confirm whether the correct CSS file was supplied.
- If the user wants remote assets in the same deck, mention `--allow-remote-assets` explicitly rather than assuming it.

## Output Expectations

- Report whether Mermaid or `diagram` fences were detected.
- Report the exact command used.
- Report whether a CSS file was applied.
- Report the generated `.pptx` path.