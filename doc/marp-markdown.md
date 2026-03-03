# Marp Markdown In MarpToPptx

## Definition

For this repository, `Marp Markdown` is best understood as a layered, implementation-defined format rather than a single standalone specification.

1. `CommonMark` provides the base Markdown model.
2. `Marpit Markdown` adds slide authoring syntax such as slide separators, directives, extended image syntax, fragmented lists, and theme CSS conventions.
3. `Marp Core` adds Marp-specific defaults and extensions such as built-in themes, `size`, math, emoji, auto-scaling, and safer HTML defaults.
4. `Marp CLI` and editors such as Marp for VS Code add tooling behavior such as preview, metadata handling, theme loading, and output conversion.

In this repository, when we say `Marp Markdown`, we mean:

`Marpit Markdown` plus the Marp-oriented subset that `MarpToPptx` currently parses and renders.

That distinction matters because `MarpToPptx` does not implement the full Marp toolchain. It implements a focused parser and editable PPTX renderer for a practical subset.

## Documentation Model

Use the following structure when documenting syntax or compatibility for this project:

1. `Base Markdown`
CommonMark-style block and inline Markdown accepted by the parser.

2. `Slide Authoring Syntax`
Marpit concepts such as slide splitting, front matter, directives, and slide backgrounds.

3. `Marp Extensions`
Marp-specific features that may exist upstream, with explicit notes about whether `MarpToPptx` supports them.

4. `Tooling And Output Behavior`
CLI options, theme CSS loading, template usage, and PPTX rendering behavior. These are implementation details, not core Markdown syntax.

## Feature Matrix

| Layer | Feature | Upstream Meaning | MarpToPptx Status | Notes |
| --- | --- | --- | --- | --- |
| CommonMark / Markdown | Headings | Standard Markdown headings | Supported | Parsed into `HeadingElement`. |
| CommonMark / Markdown | Paragraphs | Standard paragraphs | Supported | Parsed into `ParagraphElement`. |
| CommonMark / Markdown | Bullet and ordered lists | Standard Markdown lists | Supported | Parsed into `BulletListElement`; nesting depth is preserved numerically. |
| CommonMark / Markdown | Fenced code blocks | Standard fenced code blocks | Supported | Parsed into `CodeBlockElement`. |
| CommonMark / Markdown | Tables | GFM-style tables | Supported with fallback rendering | Parsed into `TableElement`, rendered as editable text rather than native PPTX tables. |
| CommonMark / Markdown | Inline emphasis and links | Standard inline formatting | Partially supported | Text content is preserved, but inline styling and hyperlinks are flattened to plain text in the semantic model. |
| CommonMark / Markdown | Raw HTML | CommonMark / Marp HTML behavior | Limited | No explicit HTML element modeling beyond comment directives. |
| Marpit | Slide separator `---` | Splits slides on horizontal rules | Supported subset | `SlideTokenizer` splits on a line that is exactly `---`, except inside fenced code blocks. |
| Marpit | Alternate separators `***`, `___`, `- - -` | Alternate CommonMark rulers that can separate slides | Not supported | Only literal `---` is recognized as a slide boundary. |
| Marpit | YAML front matter | Deck-level metadata | Supported subset | Parsed as simple `key: value` pairs only. |
| Marpit | Directives in HTML comments | Slide and deck directives | Supported subset | Supports `theme`, `paginate`, `class`, `backgroundImage`, and `backgroundColor`. |
| Marpit | Spot directives with `_` prefix | Apply a local directive to one slide only | Not supported | `_paginate`, `_class`, and similar are not recognized specially. |
| Marpit | `headingDivider` | Split slides before headings automatically | Not supported | No heading-based slide splitting. |
| Marpit | Header / footer directives | Per-slide repeated content | Not supported | Parsed neither into style nor slide elements. |
| Marpit | Extended image syntax | Width, height, filters, `bg`, split backgrounds | Partially supported | Normal Markdown images are parsed; Marpit image keywords are treated as alt text, not structured options. |
| Marpit | Background image syntax via image alt text | `![bg](...)` and related | Not supported | Backgrounds are supported only through directives, not image syntax. |
| Marpit | Fragmented lists | Incremental list reveal | Not supported | No fragment model in parser or renderer. |
| Marpit | Theme CSS | CSS-driven slide theming | Supported subset | Minimal CSS extraction for fonts, sizes, colors, padding, and code style. |
| Marp Core | Built-in theme names | Themes such as `default`, `gaia`, `uncover` | Name-only unless CSS is supplied | Theme name is stored, but real styling comes from parsed CSS or defaults. |
| Marp Core | `size` directive | Slide size preset selection | Not supported | Renderer uses a fixed 16:9 slide size. |
| Marp Core | Math | MathJax / KaTeX rendering | Not supported | No math parsing or render pipeline. |
| Marp Core | Emoji conversion | Shortcodes and Twemoji behavior | Not supported | Emoji are treated as normal text characters. |
| Marp Core | `fit` / auto-scaling | Resize headings and oversized blocks | Not supported | Layout is heuristic, but there is no Marp `fit` syntax support. |
| Marp Core | Safer HTML defaults | Controlled HTML allowlist | Not explicitly implemented | Behavior comes mostly from Markdig parsing choices, not Marp Core compatibility logic. |
| Marp CLI / tooling | Metadata directives | `title`, `description`, `author`, etc. | Not supported | Front matter values are preserved generically, but not interpreted as output metadata. |
| Marp CLI / tooling | Theme loading | `--theme`, theme sets, theme resolution | Partial equivalent | `--theme-css` loads a CSS file for minimal theme extraction. |
| Marp CLI / tooling | Output conversion modes | HTML, PDF, PPTX, images, notes | Not supported | This tool produces editable PPTX only. |
| Marp CLI / tooling | Browser-based rendering behavior | Preview, browser output, local-file rules | Not supported | `MarpToPptx` does not use a browser renderer. |
| MarpToPptx-specific | PPTX template reuse | Copy masters/themes from an existing deck | Supported | `--template` copies an existing `.pptx` before rendering slides. |

## Current MarpToPptx Compatibility Profile

### Supported Authoring Features

- YAML front matter with simple scalar values
- Slide splitting on literal `---`
- HTML comment directives for:
  - `theme`
  - `paginate`
  - `class`
  - `backgroundImage`
  - `backgroundColor`
- Headings
- Paragraphs
- Bullet lists and ordered lists
- Images with local file paths
- Fenced and indented code blocks
- GFM-style tables at the semantic-model level
- Minimal theme CSS extraction from `section`, `:root`, `body`, `h1`-`h6`, `pre`, and `code`

### Supported Rendering Features

- Editable PPTX text boxes for headings, paragraphs, lists, and code blocks
- Editable text fallback for tables
- Local image embedding with aspect-ratio-aware placement
- Solid slide background color
- Full-slide background image via directive
- Optional template-copy workflow via `--template`

### Not Yet Supported

- Native Marpit background image syntax such as `![bg](...)`
- `headingDivider`
- header / footer directives
- fragmented lists
- spot directives using `_`
- Marp Core `size`
- math
- Marp `fit`
- native PPTX table generation
- Marp CLI metadata semantics
- browser-based preview or HTML/PDF/image outputs

## Implementation Map

This is the repo-specific map from the documentation model to code.

### Base Markdown

- `src/MarpToPptx.Core/Parsing/MarpMarkdownParser.cs`
- Uses `Markdig` with `UseAdvancedExtensions()`.
- Converts Markdown blocks into semantic slide elements:
  - `HeadingElement`
  - `ParagraphElement`
  - `BulletListElement`
  - `ImageElement`
  - `CodeBlockElement`
  - `TableElement`

Important current behavior:

- Inline emphasis is flattened to text.
- Non-image links keep their visible text but not hyperlink semantics.
- Images inside paragraphs are emitted as separate `ImageElement` records.

### Slide Authoring Syntax

- `src/MarpToPptx.Core/Parsing/FrontMatterParser.cs`
- `src/MarpToPptx.Core/Parsing/SlideTokenizer.cs`
- `src/MarpToPptx.Core/Parsing/MarpDirectiveParser.cs`

Important current behavior:

- Front matter is parsed only when it appears first and is delimited by `---`.
- Front matter parsing is intentionally simple and does not implement full YAML features such as multiline blocks or nested structures.
- Slide splitting ignores `---` inside fenced code blocks.
- Directive parsing only recognizes HTML comments that match `<!-- key: value -->`.

### Theme Resolution

- `src/MarpToPptx.Core/Themes/MarpThemeParser.cs`
- `src/MarpToPptx.Core/Themes/ThemeDefinition.cs`

Important current behavior:

- CSS support is intentionally narrow.
- Only a small set of selectors and declarations are mapped into the PPTX theme model.
- Theme names from front matter are preserved even when no corresponding upstream Marp theme bundle is loaded.

### Semantic Model

- `src/MarpToPptx.Core/Models/SlideDeck.cs`

The semantic model is the contract between parsing and rendering. It is intentionally simpler than Marpit's HTML-oriented runtime model.

### PPTX Rendering

- `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs`

Important current behavior:

- Slides are rendered into a fixed 16:9 PowerPoint presentation.
- Tables are currently rendered as editable text, not native table shapes.
- Missing images fall back to a visible text placeholder.
- Background images and template reuse are renderer features, not Markdown syntax features.

### CLI Surface

- `src/MarpToPptx.Cli/Program.cs`

Current CLI options:

- input Markdown path
- `-o`, `--output`
- `--template`
- `--theme-css`

These options are intentionally narrower than Marp CLI. They expose only the parts that `MarpToPptx` currently implements.

## Suggested Terminology For Future Docs

Use these terms consistently:

- `Marp Markdown`: the practical authoring format accepted by this tool
- `Marpit syntax`: upstream slide-authoring constructs defined by Marpit
- `Marp Core features`: upstream Marp-specific extensions that may or may not be implemented here
- `MarpToPptx compatibility`: the exact subset this repository parses and renders today

That wording avoids implying full compatibility when the repository currently implements a targeted subset.
