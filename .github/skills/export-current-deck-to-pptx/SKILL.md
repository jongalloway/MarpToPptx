---
name: export-current-deck-to-pptx
description: Export the current Marp Markdown deck to an editable PowerPoint file with the published MarpToPptx CLI.
---

# Export Current Deck To PPTX

Use this skill when the user wants to turn the current Marp Markdown file into an editable `.pptx` without building `MarpToPptx` from source.

## Purpose

Export the active deck to a `.pptx` file next to the source Markdown file by using the published `MarpToPptx` package.

## Preferred Commands

Use one of these published-tool entry points:

```bash
dnx MarpToPptx path/to/deck.md -o path/to/deck.pptx
```

```bash
marp2pptx path/to/deck.md -o path/to/deck.pptx
```

```bash
dotnet tool run marp2pptx path/to/deck.md -o path/to/deck.pptx
```

## Guidance

1. Confirm that the active file is a Markdown deck.
2. Default the output to the same directory and basename as the source file unless the user asks for a different path.
3. Prefer `dnx MarpToPptx` for low-setup content repositories.
4. If the user already has a global or local tool install, using `marp2pptx` or `dotnet tool run marp2pptx` is also fine.
5. If a matching output file already exists, ask before overwriting it unless the user has already asked for regeneration.

## Important Rules

- Do not build `MarpToPptx` from source for this workflow.
- Do not switch to repo-maintainer PowerShell utilities unless the user explicitly wants maintainer-only troubleshooting.
- If export fails because the deck uses remote URLs, mention `--allow-remote-assets` as the next published-CLI option to try.
- If the user asks for contrast auditing, add `--contrast-warnings summary` or `--contrast-warnings detailed` and optionally `--contrast-report <path>`.

## Output Expectations

- Report the exact command used.
- Report the generated `.pptx` path.
- If the command fails, summarize the failure and the next published-CLI adjustment to try.