# Prompt for Copilot CLI: Marp‑Compatible Markdown → Editable PPTX Compiler (.NET 10)

Create a new .NET 10 project that implements a **Marp‑compatible Markdown → editable PPTX compiler**, built as a standalone tool that works *alongside* the Marp ecosystem. The goal is to accept Marp‑flavored Markdown, parse it into a semantic slide model, and generate a fully editable `.pptx` using the Open XML SDK (with optional ShapeCrawler helpers). The project should integrate cleanly with VS Code Marp workflows while producing real, editable PowerPoint files rather than rasterized slide images.

This PRD captures both the long-term product target and the repository steering decisions already adopted during implementation.

---

## Repository Steering

- Solution file format: `SLNX` (`MarpToPptx.slnx`)
- Centralized package management: `Directory.Packages.props`
- Test framework: xUnit v3
- Test runner: Microsoft Testing Platform
- `.NET 10` `dotnet test` integration: `global.json` with `"test": { "runner": "Microsoft.Testing.Platform" }`
- CLI packaging direction: ship `marp2pptx` as a .NET tool while preserving support for single-file publish

These conventions are part of the expected implementation, not optional follow-up work.

---

## Functional Requirements

### 1. Input Parsing (Markdig + Marp Extensions)
Use **Markdig** as the Markdown engine and implement a custom extension to support Marp features:

- Slide separators: `---`
- Front‑matter
- Marp directives inside HTML comments:
  - `<!-- theme: default -->`
  - `<!-- paginate: true -->`
  - `<!-- class: lead -->`
  - `<!-- backgroundImage: url(...) -->`
- Standard Markdown elements:
  - Headings
  - Paragraphs
  - Bullet/numbered lists
  - Images
  - Tables
  - Code blocks

The parser should output a **semantic slide model**, not HTML.

---

## Semantic Slide Model

Define a strongly typed model that represents slides independently of PPTX or HTML:

```csharp
class SlideDeck { List<Slide> Slides; Theme Theme; }
class Slide { List<ISlideElement> Elements; SlideStyle Style; }

interface ISlideElement {}

class Heading : ISlideElement { /* level, text */ }
class Paragraph : ISlideElement { /* text */ }
class BulletList : ISlideElement { /* items */ }
class ImageElement : ISlideElement { /* src, alt */ }
class CodeBlock : ISlideElement { /* language, code */ }
class TableElement : ISlideElement { /* rows, columns */ }
```

This model should be deterministic, testable, and independent of rendering concerns.

---

## Theme + Layout Resolution

Marp themes are CSS‑based. Implement a minimal CSS extractor that supports:

- Font families
- Font sizes
- Colors
- Backgrounds
- Margins and padding
- Code block styling
- Heading size hierarchy

Map these to PPTX theme elements:

- Slide master color scheme
- Text box defaults
- Background fills
- Font families and sizes
- Layout spacing

You do **not** need full CSS support—only the subset Marp themes rely on.

---

## PPTX Generation (Open XML SDK + optional ShapeCrawler)

Use **Open XML SDK** as the primary engine for generating editable PPTX files. Optionally layer **ShapeCrawler** for convenience when working with text boxes, images, and tables.

Implementation steering:

- Avoid `ImageSharp` for intrinsic image sizing unless it is explicitly re-approved after licensing review.
- Prefer built-in platform capabilities or a minimal in-project metadata reader for image dimensions when aspect-ratio-aware placement is required.

Requirements:

- Generate editable text boxes for headings, paragraphs, and lists.
- Generate real PPTX tables (not images).
- Insert images with correct sizing and aspect ratio.
- Render code blocks as styled text boxes.
- Support slide backgrounds (solid, gradient, or image).
- Support user‑supplied `.pptx` templates to control slide masters and themes.

---

## CLI Tool (`marp2pptx`)

Create a CLI that mirrors Marp’s workflow:

```
marp2pptx input.md -o output.pptx
marp2pptx input.md --template theme.pptx
```

The CLI should integrate cleanly with VS Code tasks and file watchers.

The CLI should also be packageable as a .NET tool named `marp2pptx`.

---

## Non‑Functional Requirements

- .NET 10
- Cross‑platform single‑file publish
- Deterministic, testable pipeline
- Clear separation of concerns:
  - Markdown parsing
  - Slide model
  - Theme resolution
  - PPTX rendering
  - CLI interface
- Repository-level build and test conventions should remain aligned with CPM, `SLNX`, xUnit v3, and Microsoft Testing Platform.

---

## Implementation Recommendations

- Use **Markdig** with a custom Marp extension for parsing directives and slide boundaries.
- Build a `MarpThemeParser` that extracts only the CSS properties needed for PPTX mapping.
- Implement a `LayoutEngine` that maps semantic slide elements to PPTX shapes.
- Use a `.pptx` template to avoid manually constructing slide masters.
- Add automated tests using xUnit v3 on Microsoft Testing Platform for:
  - Slide splitting
  - Directive parsing
  - Theme extraction
  - PPTX element generation

Current milestone note:

- Native PPTX tables remain a product requirement.
- Until that lands, semantic table content may use an editable text fallback in the initial milestone implementation, but this should not be treated as the final state.

---

## Deliverables

1. A .NET 10 solution with the following projects:
   - `MarpToPptx.Core` (parsing, model, theme, layout)
   - `MarpToPptx.Pptx` (Open XML rendering)
   - `MarpToPptx.Cli` (command‑line interface)
   - `MarpToPptx.Tests` (unit tests)

2. Initial implementation of:
   - Markdown → SlideDeck parser
   - Basic PPTX generator supporting:
     - Title slide
     - Heading + paragraph
     - Bullet lists
     - Images
  - Code blocks as styled editable text boxes
  - Basic theme extraction for fonts, sizes, colors, and padding
  - Template-copy workflow for reusing an existing `.pptx` theme/master

3. A roadmap for:
   - Tables
   - Code blocks
   - Theme mapping
   - Template support

---

## Current Status Guidance

At the time of this update, the repository has already adopted the steering choices above and has an initial implementation in place for:

- semantic slide parsing
- directive parsing
- minimal theme extraction
- PPTX output for headings, paragraphs, bullet lists, images, and code blocks
- template-copy support

Remaining roadmap items should be evaluated against this baseline rather than against an empty starting point.

