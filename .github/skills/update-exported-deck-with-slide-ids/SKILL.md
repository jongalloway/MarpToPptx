---
name: update-exported-deck-with-slide-ids
description: Update a previously exported managed deck by using slideId directives, --write-slide-ids, and --update-existing with the published MarpToPptx CLI.
---

# Update Exported Deck With Slide IDs

Use this skill when the user wants to iteratively update a previously exported `MarpToPptx` deck while keeping slide identity stable through `slideId` directives.

## Purpose

Help the user work through the published CLI workflow for:

1. bootstrapping `slideId` directives into source when needed
2. regenerating against a previously exported managed deck with `--update-existing`
3. keeping the previous deck and rendering template as separate concepts

## Preferred Commands

Bootstrap missing `slideId` directives into the Markdown source:

```bash
dnx MarpToPptx deck.md --write-slide-ids -o deck.pptx
```

Update against a previously exported managed deck:

```bash
dnx MarpToPptx deck.md --update-existing previous-deck.pptx -o updated-deck.pptx
```

Update against a previously exported managed deck while still using a template for rendering:

```bash
dnx MarpToPptx deck.md --update-existing previous-deck.pptx --template conference-template.pptx -o updated-deck.pptx
```

Installed-tool variants such as `marp2pptx` and `dotnet tool run marp2pptx` are also fine.

## Guidance

1. Treat the previously exported deck passed to `--update-existing` as the reconciliation source.
2. Treat `--template` as the rendering source for layouts, masters, and `Template[n]` behavior.
3. If the deck does not yet have explicit `slideId` directives, suggest `--write-slide-ids` as the bootstrap step.
4. Prefer updating from a previous managed deck produced by `MarpToPptx`, not from an arbitrary PowerPoint file.
5. Use the published CLI workflow only. Do not switch to local-source execution for normal end-user usage.

## Important Rules

- Do not confuse `--template` with `--update-existing`; they serve different roles.
- Do not claim that manual edits inside a changed managed slide will be merged shape-by-shape.
- Explain that unmanaged slides are preserved, while changed managed slides are replaced wholesale.
- If the user wants to verify the repo's maintainer smoke flow, that is a separate maintainer-oriented prompt and not this end-user skill.

## Output Expectations

- Report whether `slideId` bootstrapping is needed.
- Report the exact command used.
- State which file is the previous managed deck and which file is the rendering template, if any.
- Summarize what the user should expect to be preserved versus replaced during update.