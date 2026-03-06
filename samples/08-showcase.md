---
theme: marp-showcase
paginate: true
lang: en-US
header: Marp Ecosystem Showcase
footer: MarpToPptx smoke sample
---

<!-- class: lead -->
<!-- _backgroundImage: url(assets/marp.svg) -->
<!-- _backgroundSize: contain -->
# MarpToPptx Showcase

Markdown authoring with editable PowerPoint output.

This sample is written like a short conference talk about the Marp ecosystem and where MarpToPptx fits.

<!-- Speaker framing: introduce the ecosystem first, then position MarpToPptx as the editable-PPTX path for teams that still want Marp-style authoring. -->

---

<!-- class: ecosystem -->
## What Marp Is

- **Marp** stands for **Markdown Presentation Ecosystem**.
- It is built around a CommonMark-based authoring model with simple horizontal-rule slide breaks.
- Marp adds directives, image shortcuts, built-in themes, and CSS theming.
- The official toolset can export decks to HTML, PDF, and PowerPoint.

The value proposition from the official Marp site is simple: create beautiful slide decks while staying focused on the story in Markdown.

<!-- Speaker note: this is the framing slide for people who have heard the name but do not know the product boundaries yet. -->

---

## The Marp Ecosystem In One Story

- **Marpit** supplies the slide framework and plugin surface.
- **Marp Core** adds practical authoring features and built-in themes.
- **Marp for VS Code** makes the write-preview loop fast.
- **Marp CLI** handles export, watch mode, and automation.
- **MarpToPptx** adds editable PowerPoint handoff.

The pitch to a speaker is simple: keep writing in Markdown, but hand off a deck that PowerPoint users can still edit slide by slide.

<!-- Speaker note: emphasize coexistence rather than competition. The audience should hear that MarpToPptx is additive to existing Marp workflows. -->

---

## Why The Ecosystem Works

- The ecosystem is fully open-source and MIT licensed.
- The tooling stays close to plain Markdown instead of inventing a separate slide language.
- Theme CSS remains the main customization surface for design systems.
- The official tools cover editing, preview, export, and automation.
- The stack is composable enough for custom workflows and integrations.

---

<!-- class: contrast -->
## What MarpToPptx Adds

- Keeps Markdown as the authoring source while producing editable PowerPoint shapes.
- Maps theme, background, header, footer, and pagination intent into PPTX constructs.
- Uses Open XML validation plus PowerPoint open checks as compatibility gates.
- Fits teams that want Markdown discipline without giving up ordinary Office collaboration.

This is why the repo treats PowerPoint compatibility as the real success metric, not just schema-valid XML.

---

## Current Repo Structure

- **MarpToPptx.Core**: parsing, theme extraction, slide model, and layout planning.
- **MarpToPptx.Pptx**: Open XML rendering and presentation generation.
- **MarpToPptx.Cli**: the marp2pptx command entrypoint.
- **MarpToPptx.OpenXmlValidator**: package validation used by smoke tests and CI.
- **MarpToPptx.Tests**: xUnit v3 coverage on Microsoft Testing Platform.

---

## Capabilities Worth Showing Live

1. Front matter and slide directives for themes, pagination, classes, backgrounds, headers, and footers.
2. Built-in Marp ideas such as theme CSS, directive-driven layout changes, and Markdown-first authoring.
3. Editable PPTX generation for headings, paragraphs, bullet lists, images, code blocks, native tables, and header/footer text.
4. Local audio and video embedding for supported media formats.
5. Template-copy workflow for reusing an existing `.pptx` theme or master.
6. Local and CI-friendly smoke scripts for generation, validation, and PowerPoint-open checks.

<!-- Speaker note: this slide is the “why should a team care?” moment. Tie these bullets back to handoff, review, and release workflows. -->

---

<!-- class: compact -->
## VS Code Export Task

The repo README shows this inside `tasks.json`:

```json
{
  "label": "Export to PPTX",
  "type": "shell",
  "command": "dnx",
  "args": ["MarpToPptx", "${file}", "-o", "deck.pptx"]
}
```

- Edit Markdown.
- Preview in VS Code.
- Run the task and hand off the PPTX.

---

## Release And Validation Narrative

- Unit tests catch parser and renderer regressions.
- **Invoke-PptxSmokeTest.ps1** and **Invoke-AllPptxSmokeTests.ps1** provide end-to-end local validation.
- The validator project checks Open XML correctness.
- PowerPoint open and round-trip save remain the final compatibility gate.

The repo’s release workflow separates fast CI coverage from heavier release-gate validation.

---

## Marp Themes And Authoring Conventions

- Marp ships with built-in themes such as **default**, **gaia**, and **uncover**.
- Directives are YAML-shaped, so front matter and HTML comments fit naturally into ordinary Markdown workflows.
- Theme CSS stays readable and close to normal web styling, which keeps decks maintainable.
- The same authoring approach works well for previews, automation, and source control review.

That design discipline is a big reason MarpToPptx can map source intent into a more editable PowerPoint structure.

---

## Roadmap To Mention In The Q&A

- Improve CSS coverage for more Marp theme features.
- Refine layout heuristics for denser or more designed decks.
- Expand template integration to map multiple layouts more intelligently.
- Improve native PPTX table styling and layout fidelity.
- Expand code block highlighting coverage and theme fidelity.
- Support remote assets and additional image formats.

Closing point: MarpToPptx is for teams that want Markdown authoring discipline without giving up editable PowerPoint deliverables.

<!-- Closing note: end with the interoperability message. The tool matters because it lets Markdown-authored decks participate in ordinary Office collaboration. -->