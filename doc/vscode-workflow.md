# VS Code Authoring Workflow With MarpToPptx

## Purpose

This document describes how to integrate `MarpToPptx` into a VS Code authoring workflow in a content repository. It covers installing the tool from NuGet, setting up VS Code tasks, and using an edit / preview / export loop with the Marp extension.

These instructions assume the `MarpToPptx` tool is already published to NuGet. You do not need to build this repository from source.

## Prerequisites

- [.NET 10 SDK](https://dotnet.microsoft.com/download)
- VS Code with the [Marp for VS Code](https://marketplace.visualstudio.com/items?itemName=marp-team.marp-vscode) extension installed
- A content repository containing Marp Markdown files (`.md`)

## dnx Versus dotnet tool run

Both `dnx` and `dotnet tool run` invoke the published tool, but they are intended for different scenarios.

| Scenario | Recommended invocation |
| --- | --- |
| Quick one-off export or VS Code task in a content repo | `dnx MarpToPptx` |
| Persistent install on a developer workstation | `dotnet tool install --global MarpToPptx` and then `marp2pptx` |
| Local tool pinned to a `.config/dotnet-tools.json` manifest | `dotnet tool restore` followed by `dotnet tool run marp2pptx` |

`dnx` requires no prior install step and no tool manifest in the content repository. It resolves the latest stable version from NuGet at run time. Use it when you want zero setup for contributors who only author Marp content and never need to build `MarpToPptx` itself.

`dotnet tool run` is useful when you want to pin a specific version via a tool manifest. It requires running `dotnet tool restore` first. Prefer this approach for CI pipelines where reproducibility matters.

## Setting Up VS Code Tasks

Add a `.vscode/tasks.json` file to your content repository. The examples below use `dnx` so that contributors only need the .NET SDK; no additional install step is required.

### Minimal tasks.json

This example adds two tasks: one that exports the currently open file to a `.pptx` with the same base name, and one that exports using a PowerPoint template.

```json
{
  "version": "2.0.0",
  "tasks": [
    {
      "label": "Export to PPTX",
      "type": "shell",
      "command": "dnx",
      "args": [
        "MarpToPptx",
        "${file}",
        "-o",
        "${fileDirname}/${fileBasenameNoExtension}.pptx"
      ],
      "group": "build",
      "presentation": {
        "reveal": "always",
        "panel": "shared"
      },
      "problemMatcher": []
    },
    {
      "label": "Export to PPTX (with template)",
      "type": "shell",
      "command": "dnx",
      "args": [
        "MarpToPptx",
        "${file}",
        "--template",
        "${workspaceFolder}/templates/theme.pptx",
        "-o",
        "${fileDirname}/${fileBasenameNoExtension}.pptx"
      ],
      "group": "build",
      "presentation": {
        "reveal": "always",
        "panel": "shared"
      },
      "problemMatcher": []
    }
  ]
}
```

Run either task from **Terminal → Run Task** or by pressing `Ctrl+Shift+B` (Windows / Linux) or `Cmd+Shift+B` (macOS) if you mark it as the default build task.

### Variable reference

| VS Code variable | Resolves to |
| --- | --- |
| `${file}` | Absolute path to the currently open file |
| `${fileDirname}` | Directory containing the current file |
| `${fileBasenameNoExtension}` | Filename without the extension |
| `${workspaceFolder}` | Root folder of the VS Code workspace |

### Theme CSS variant

If you use a CSS file for theme styling instead of a `.pptx` template:

```json
{
  "label": "Export to PPTX (theme CSS)",
  "type": "shell",
  "command": "dnx",
  "args": [
    "MarpToPptx",
    "${file}",
    "--theme-css",
    "${workspaceFolder}/themes/custom.css",
    "-o",
    "${fileDirname}/${fileBasenameNoExtension}.pptx"
  ],
  "group": "build",
  "presentation": {
    "reveal": "always",
    "panel": "shared"
  },
  "problemMatcher": []
}
```

## Sharing tasks.json Across A Team

Commit `.vscode/tasks.json` to your content repository. Because `dnx` resolves the tool from NuGet at run time, all contributors get the same experience as long as they have the .NET 10 SDK installed. There is no need to run any install command or maintain a tool manifest.

Typical content repository layout:

```text
my-content-repo/
├── .vscode/
│   └── tasks.json          ← commit this
├── templates/
│   └── theme.pptx          ← optional shared template
├── themes/
│   └── custom.css          ← optional shared theme CSS
├── slides/
│   ├── deck-one.md
│   └── deck-two.md
└── README.md
```

## Example Agent Skills

This repository includes example Agent Skills under `.github/skills/` for users who prefer agent-driven export flows in addition to VS Code tasks.

The current examples cover:

- exporting the current deck to PPTX
- exporting with a PowerPoint template
- exporting with a theme CSS file
- exporting diagram-heavy decks that use Mermaid or `diagram` fences
- updating a previously exported managed deck with `slideId`, `--write-slide-ids`, and `--update-existing`

These are example assets intended to be copied into a separate content repository, typically under that repository's own `.github/skills/` folder.

Important boundaries:

- These example Agent Skills target the published CLI surface such as `dnx MarpToPptx`, `marp2pptx`, and `dotnet tool run marp2pptx`.
- They do not assume access to this source repository.
- They do not depend on the maintainer-focused PowerShell helpers in `scripts/`.
- They are meant to stay portable across skills-compatible agents and tools rather than being VS Code-only repo automation.

Use tasks when you want a deterministic one-click export button. Use skills when you want an agent to choose the right published CLI invocation, detect a companion CSS file, or recognize that a deck is diagram-focused.

For a short index of the example Agent Skills in this repository, see [doc/agent-skills.md](../doc/agent-skills.md).

## The Edit / Preview / Export Loop

With the [Marp for VS Code](https://marketplace.visualstudio.com/items?itemName=marp-team.marp-vscode) extension installed, the typical workflow is:

1. **Edit** — Open a `.md` file and author slides in Marp Markdown. Use `---` to separate slides.
2. **Preview** — Click the Marp preview icon in the VS Code editor toolbar, or run **Marp: Open Preview to the Side** from the Command Palette. The preview renders the HTML representation of your slides in real time.
3. **Export** — When the deck is ready, run the **Export to PPTX** task from **Terminal → Run Task**. The `.pptx` file is written next to the source Markdown file.
4. **Review** — Open the generated `.pptx` in PowerPoint or another compatible viewer to confirm layout and content before sharing.

The preview step uses the Marp for VS Code extension and renders HTML; it does not use `MarpToPptx`. The two tools complement each other: the extension provides a live visual preview during authoring, while `MarpToPptx` produces the editable PPTX output for distribution.

## Notes On dnx Version Resolution

`dnx` without a version constraint resolves the latest stable version of the package from NuGet at run time. If your team needs a specific version, add a `--version` flag:

```json
"args": [
  "MarpToPptx",
  "--version",
  "1.2.3",
  "${file}",
  "-o",
  "${fileDirname}/${fileBasenameNoExtension}.pptx"
]
```

Pinning a version ensures consistent output across machines and avoids unexpected behavior if a new release changes rendering defaults.

## Relevant Links

- [MarpToPptx on NuGet](https://www.nuget.org/packages/MarpToPptx/)
- [Marp for VS Code extension](https://marketplace.visualstudio.com/items?itemName=marp-team.marp-vscode)
- [CLI options reference](../doc/marp-markdown.md#cli-surface)
