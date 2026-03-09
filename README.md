# MarpToPptx

[![Build](https://github.com/jongalloway/MarpToPptx/actions/workflows/ci.yml/badge.svg)](https://github.com/jongalloway/MarpToPptx/actions/workflows/ci.yml)
[![NuGet Version](https://img.shields.io/nuget/v/MarpToPptx?logo=nuget)](https://www.nuget.org/packages/MarpToPptx/)
[![NuGet Downloads](https://img.shields.io/nuget/dt/MarpToPptx?logo=nuget)](https://www.nuget.org/packages/MarpToPptx/)
[![.NET 10](https://img.shields.io/badge/.NET-10-512BD4?logo=dotnet)](https://dotnet.microsoft.com/)

**✨ Turn your Marp Markdown into real, editable PowerPoint files.**

So you'd like to write your presentation slides in Markdown? It's lightweight, easy to version, and works great with AI-powered workflows. The amazing [Marp](https://marp.app/) ecosystem has you covered with a mature, open-source solution for authoring beautiful slide decks in plain text:

- [**Marp**](https://marp.app/) — the Markdown Presentation Ecosystem
- [**Marp for VS Code**](https://marketplace.visualstudio.com/items?itemName=marp-team.marp-vscode) — live preview and export right in your editor
- [**Marp CLI**](https://github.com/marp-team/marp-cli) — command-line conversion to HTML, PDF, and PPTX
- [**awesome-marp**](https://github.com/marp-team/awesome-marp) — community themes, tools, and examples

There's just one hangup. When you're asked to turn in a PowerPoint deck at a conference, need to share editable slides with a colleague, or want to integrate into an existing corporate deck — you need your slides in real PowerPoint format. Unfortunately, Marp's PPTX export produces uneditable image-per-slide output that you can't select, edit, or restyle.

**That's where MarpToPptx comes in. 🎉** It reads your Marp-flavored Markdown and produces native Open XML PowerPoint files where every heading, bullet, table, and code block is a real, selectable, editable PowerPoint shape. The output opens cleanly in PowerPoint — no repair prompts, no surprises.

```mermaid
flowchart LR
    A["📝 Marp Markdown"] --> B["⚙️ MarpToPptx"]
    C["🎨 CSS Theme"] -.-> B
    D["📊 .pptx Template"] -.-> B
    B --> E["✅ Editable .pptx"]
    style A fill:#4a9eff,color:#fff
    style B fill:#7c3aed,color:#fff
    style C fill:#f59e0b,color:#fff
    style D fill:#f59e0b,color:#fff
    style E fill:#10b981,color:#fff
```

## 🚀 Quick Start

MarpToPptx requires [.NET 10](https://dotnet.microsoft.com/download). The fastest way to try it — no install needed:

```bash
dnx MarpToPptx slides.md -o slides.pptx
```

Or install it globally as a .NET tool:

```bash
dotnet tool install --global MarpToPptx
marp2pptx slides.md -o slides.pptx
```

### Apply a Theme or Template

Use a CSS theme file for Marp-style theming:

```bash
marp2pptx slides.md --theme-css brand.css -o slides.pptx
```

Or reuse an existing PowerPoint template to inherit your organization's masters and layouts:

```bash
marp2pptx slides.md --template corporate.pptx -o slides.pptx
```

To target a specific template layout from markdown, use `layout` in front matter for the default content layout, or `layout` / `_layout` HTML-comment directives for sticky or single-slide overrides:

```md
---
layout: Title and Content
---

<!-- _layout: Section Header -->
# Agenda
```

If the template's branded title slide is authored as slide content instead of a reusable layout, target the actual template slide directly:

```md
<!-- _layout: Template[1] -->
```

See [doc/using-templates.md](doc/using-templates.md) for the quickstart and [doc/template-authoring-guidelines.md](doc/template-authoring-guidelines.md) for the technical details.

## 📋 Features

| Category | What's supported |
|---|---|
| **Slide structure** | Front matter directives, `---` slide splitting, presenter notes |
| **Text content** | Headings, paragraphs, ordered and unordered lists, bold/italic/code spans |
| **Rich content** | Local images, syntax-highlighted code blocks, native tables |
| **Media** | Embedded MP3/M4A audio, embedded video |
| **Theming** | CSS-based Marp themes (fonts, colors, padding, backgrounds, typography) |
| **Templates** | Copy masters and layouts from an existing `.pptx` |
| **Directives** | `backgroundColor`, `backgroundImage`, `header`, `footer`, `paginate`, `layout`, scoped overrides |
| **Output quality** | Open XML validated, tested to open without repair prompts in PowerPoint |
| **Platform** | Runs anywhere .NET 10 runs — CI-tested on Ubuntu, works on Windows and macOS |

## 💻 VS Code Integration

Add a one-click export task to any content repository with a `.vscode/tasks.json`:

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
      "presentation": { "reveal": "always", "panel": "shared" },
      "problemMatcher": []
    }
  ]
}
```

Run **Terminal → Run Task → Export to PPTX** while editing a Markdown file. The `.pptx` appears next to your source file.

For template-based export, version pinning, team sharing, and integrating the Marp for VS Code preview extension, see [`doc/vscode-workflow.md`](doc/vscode-workflow.md).

## 🎯 Sample Decks

The [`samples/`](samples/) directory contains ready-to-run Marp decks that exercise different features:

```bash
dnx MarpToPptx samples/01-minimal.md -o artifacts/samples/01-minimal.pptx
dnx MarpToPptx samples/04-content-coverage.md -o artifacts/samples/04-content-coverage.pptx
dnx MarpToPptx samples/03-theme-css.md --theme-css samples/03-theme.css -o artifacts/samples/03-theme-css.pptx
```

See [`samples/README.md`](samples/README.md) for the full list and suggested progression.

## 🗺️ Roadmap

- Broader CSS coverage for advanced Marp theme features
- Smarter layout heuristics for dense or highly designed slides
- Multi-layout template mapping
- Improved table styling fidelity
- Expanded syntax highlighting themes
- Remote asset support

## 🤝 Contributing

See [`CONTRIBUTING.md`](CONTRIBUTING.md) for repository structure, conventions, building from source, packaging, and release process.

## 📖 Documentation

- [Using a PowerPoint template](doc/using-templates.md)
- [Marp Markdown behavior and directives](doc/marp-markdown.md)
- [Template authoring guidelines](doc/template-authoring-guidelines.md)
- [PPTX compatibility notes](doc/pptx-compatibility-notes.md)
- [VS Code workflow integration](doc/vscode-workflow.md)
- [Release validation checklist](doc/release-validation.md)
