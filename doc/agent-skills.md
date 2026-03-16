# Agent Skills

This repository includes example end-user Agent Skills under `.github/skills/` that wrap the published `MarpToPptx` CLI.

- `export-current-deck-to-pptx` — export the active Marp deck to an editable `.pptx`.
- `export-current-deck-with-template` — export by applying a PowerPoint template.
- `export-current-deck-with-theme-css` — export by applying a Marp-compatible CSS theme.
- `export-diagram-deck-to-pptx` — export a deck that uses Mermaid or `diagram` fences.
- `update-exported-deck-with-slide-ids` — update a previously exported managed deck by using `slideId`, `--write-slide-ids`, and `--update-existing`.

These are example assets intended to be copied into another content repository rather than used as maintainer-only repo automation.

Agent Skills are part of the broader skills ecosystem and are designed to be portable across compatible agents and tools, including VS Code, GitHub Copilot CLI, and other skills-compatible environments.

For how to add Agent Skills to your own repository, see the official VS Code documentation:

- [Use Agent Skills in VS Code](https://code.visualstudio.com/docs/copilot/customization/agent-skills)