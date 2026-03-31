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
| Marpit | Directives in HTML comments | Slide and deck directives | Supported subset | Supports `theme`, `paginate`, `class`, `backgroundImage`, `backgroundSize`, `backgroundColor`, `color`, `header`, and `footer`. |
| Marpit | Spot directives with `_` prefix | Apply a local directive to one slide only | Supported | All recognised directive keys work with a `_` prefix (e.g. `_class`, `_paginate`, `_backgroundColor`, `_color`). Spot directives apply to the current slide only and do not carry forward to subsequent slides. |
| Marpit | `headingDivider` | Split slides before headings automatically | Supported subset | Parsed from front matter as an integer 1–6; slides are split before headings at or above that level. |
| Marpit | Header / footer directives | Per-slide repeated content | Supported | `header` and `footer` string values are stored in `SlideStyle` and emitted into PPTX text shapes on each slide. |
| Marpit | Extended image syntax | Width, height, filters, `bg`, split backgrounds | Partially supported | Normal Markdown images are parsed; `![bg](url)` is promoted to a slide background. Other Marpit image keywords (width, height, percentage, split) are treated as alt text with no effect on sizing or layout. |
| Marpit | Background image syntax via image alt text | `![bg](...)` and related | Supported subset | `![bg](url)` sets the slide background image. Modifiers in the alt text (`bg cover`, `bg contain`, `bg left`, `bg right`, percentage sizing) are not yet parsed; the image is always rendered full-bleed. A `backgroundImage` directive on the same slide takes precedence over `![bg](...)`. See [Background Image Precedence](#background-image-precedence) below. |
| Marpit | Fragmented lists | Incremental list reveal | Not supported | No fragment model in parser or renderer. |
| Marpit | Theme CSS | CSS-driven slide theming | Supported subset | CSS extraction for fonts, sizes, colors, padding, background, line-height, letter-spacing, text-transform, font-weight, and code style. See the CSS property reference below. |
| Marp Core | Built-in theme names | Themes such as `default`, `gaia`, `uncover` | Name-only unless CSS is supplied | Theme name is stored, but real styling comes from parsed CSS or defaults. |
| Marp Core | `size` directive | Slide size preset selection | Not supported | Renderer uses a fixed 16:9 slide size. |
| Marp Core | Math | MathJax / KaTeX rendering | Not supported | No math parsing or render pipeline. |
| Marp Core | Emoji conversion | Shortcodes and Twemoji behavior | Not supported | Emoji are treated as normal text characters. |
| Marp Core | `fit` / auto-scaling | Resize headings and oversized blocks | Not supported | Layout is heuristic, but there is no Marp `fit` syntax support. |
| Marp Core | Safer HTML defaults | Controlled HTML allowlist | Not explicitly implemented | Behavior comes mostly from Markdig parsing choices, not Marp Core compatibility logic. |
| Marp CLI / tooling | Metadata directives | `title`, `description`, `author`, etc. | Not supported | Front matter values are preserved generically, but not interpreted as output metadata. |
| Marp CLI / tooling | Theme loading | `--theme`, theme sets, theme resolution | Partial equivalent | `--theme-css` loads a CSS file for minimal theme extraction. |
| Marp CLI / tooling | Template layout selection | Choose named `.pptx` layouts or a specific template slide from markdown | Supported | Front matter `layout` sets the default content layout; `layout` and `_layout` comment directives select named template layouts per slide. `Template[n]` targets an authored template slide directly. Matching a named layout or template slide suppresses theme/class/background/header/footer styling for that slide so the template can dominate. |
| Marp CLI / tooling | Output conversion modes | HTML, PDF, PPTX, images, notes | Not supported | This tool produces editable PPTX only. |
| Marp CLI / tooling | Browser-based rendering behavior | Preview, browser output, local-file rules | Not supported | `MarpToPptx` does not use a browser renderer. |
| MarpToPptx-specific | PPTX template reuse | Copy masters/themes from an existing deck | Supported | `--template` copies an existing `.pptx` before rendering slides. |

## Directive Behavior

Marpit directives can be authored in YAML front matter or as HTML comments inside slide content. `MarpToPptx` recognises three scopes:

### Global Directives (front matter only)

Front matter key-value pairs set deck-level defaults that apply to all slides unless overridden. The following keys are interpreted and mapped to `SlideStyle`:

| Key | Type | Description |
| --- | --- | --- |
| `theme` | string | Theme name; used for CSS class-variant lookup |
| `paginate` | boolean | Enable slide-number rendering |
| `class` | string | CSS class applied to all slides |
| `backgroundImage` | string | URL of background image for all slides |
| `backgroundSize` | string | Background sizing hint (`cover`, `contain`, etc.) |
| `backgroundPosition` | string | Background image position (`center`, `top`, `bottom`, `left`, `right`, or two-keyword combinations) |
| `backgroundColor` | string | Background fill color for all slides |
| `color` | string | Default text color for slide text shapes |
| `header` | string | Repeated header text on every slide |
| `footer` | string | Repeated footer text on every slide |
| `headingDivider` | integer (1–6) | Split slides before headings at or above that level |
| `lang` | BCP-47 string | Language tag written to PPTX document metadata |
| `style` | CSS string | Inline CSS merged with any `--theme-css` CSS |
| `transition` | string | Default slide transition for all slides (see [Transition Directive](#transition-directive)) |
| `diagram-theme` | string | Default DiagramForge theme for all `mermaid` and `diagram` fences (see [Diagram Theme Directive](#diagram-theme-directive)) |

All other front matter keys are stored in `SlideDeck.FrontMatter` but not interpreted.

### Local Directives (HTML comments, carried forward)

HTML comments of the form `<!-- key: value -->` inside slide content set a local override for the current slide. The new value **carries forward** to every subsequent slide unless overridden again. Local directives update the carry-forward style.

```markdown
<!-- backgroundColor: #102A43 -->
<!-- class: contrast -->
# This slide and all slides after it use the dark background
```

The same keys listed in the global table above are recognised as local directives.

### Spot Directives (HTML comments, current slide only)

A leading `_` on the key makes a directive a **spot directive**: it applies only to the current slide and does not carry forward. Spot directives are layered on top of the inherited local style for that slide.

```markdown
<!-- _paginate: false -->
<!-- _backgroundColor: #FFD700 -->
# This slide only: pagination off, gold background
```

All recognised directive keys support the `_` prefix. The next slide reverts to the inherited carry-forward value.

### Transition Directive

The `transition` directive adds a PowerPoint slide transition to the slide that advances into it. It can be set globally in front matter (deck default), as a local directive (carries forward), or as a spot directive (current slide only).

**Syntax:**

```markdown
transition: <type> [dir:<direction>] [dur:<milliseconds>]
```

**Supported transition types:**

| Directive value | PowerPoint transition | Direction supported |
|---|---|---|
| `fade` | Fade | No |
| `push` | Push | Yes (`left`, `right`, `up`, `down`) |
| `wipe` | Wipe | Yes (`left`, `right`, `up`, `down`) |
| `cut` | Cut (instant) | No |
| `cover` | Cover | Yes (`left`, `right`, `up`, `down`) |
| `pull` | Pull (Uncover) | Yes (`left`, `right`, `up`, `down`) |
| `random-bar` | Random Bar | `horizontal` (default), `vertical` |
| `morph` | Morph (fallback: fade) | No |

**Optional parameters:**

- `dir:<value>` — sets the direction for transitions that support it. Ignored otherwise.
- `dur:<ms>` — duration in milliseconds. Mapped to the `spd` attribute bands: `≤300ms` → fast, `≤700ms` → medium, `>700ms` → slow.

> **Morph note:** Morph is an Office 2019+ feature that requires an `mc:AlternateContent` wrapper in the PPTX. The current implementation emits a `fade` as a compatible fallback for all PowerPoint versions. Full morph support is tracked as a follow-up.

**Examples:**

```markdown
---
transition: fade
---

# First slide (inherits fade)

---

<!-- transition: push dir:right -->
# Second slide (push right, carries forward)

---

<!-- _transition: wipe dur:500 -->
# Third slide (wipe, spot override — does not carry forward)

---

# Fourth slide (back to push right from slide 2)
```

### Diagram Theme Directive

The `diagram-theme` front matter directive sets a deck-level preferred DiagramForge theme name for all `mermaid` and `diagram` fenced code blocks in the presentation. This lets you apply a consistent diagram look without repeating front matter inside every fence.

**Syntax (front matter only):**

```yaml
diagram-theme: <theme-name>
```

Supported DiagramForge theme names include `default`, `prism`, `dracula`, and `presentation`.

**Precedence:**

1. **Fence-level `theme:`** — a `theme:` key inside the fenced block's own YAML front matter always wins.
2. **Deck-level `diagram-theme`** — applied when the fence does not specify its own theme.
3. **Neither set** — DiagramForge falls back to its default color mapping derived from the Marp slide theme.

**Examples:**

Set a deck-wide default:

```yaml
---
diagram-theme: prism
---
```

Override a single diagram back to a different theme:

````markdown
```mermaid
---
theme: dracula
---
flowchart LR
  A --> B
```
````

In the example above, the Mermaid diagram renders with the `dracula` theme despite the deck-level `prism` setting, because the fence-level `theme:` takes precedence.

### Presenter Notes

HTML comments that do **not** match the `<!-- key: value -->` pattern are treated as presenter notes. They are stripped from the slide content and stored in `Slide.Notes`, then emitted as PPTX speaker notes.

Presenter notes preserve line breaks and support Markdown-style formatting for bold, italic, strikethrough, inline code, and fenced code blocks in the emitted PPTX notes text. Non-directive note text is still captured from HTML comments, not from normal slide Markdown blocks.

Compatibility note: during manual PowerPoint testing, inline-code and fenced-code note text may not consistently honor the requested monospace font face in the notes pane, even though the PPTX run structure is emitted with code-oriented formatting. Treat note code formatting as best-effort PowerPoint compatibility rather than a strict visual guarantee.

````markdown
<!-- This is a presenter note, not a directive. -->

<!-- **Bold** and *italic* and `code` inside notes. -->

<!--
```csharp
Console.WriteLine("notes code block");
```
-->
````

## Current MarpToPptx Compatibility Profile

### Supported Authoring Features

- YAML front matter with simple `key: value` scalars, plus `|` literal block scalars (for `style`)
  - `lang` — sets the BCP-47 language tag for document metadata
  - `style` — inline CSS merged with any external theme CSS
  - `diagram-theme` — sets a deck-level preferred DiagramForge theme for all `mermaid` and `diagram` fences
- Slide splitting on literal `---`
- `headingDivider` in front matter (integer 1–6) — also splits slides before headings at or above that level
- HTML comment directives (local and spot) for:
  - `theme`
  - `paginate`
  - `class`
  - `backgroundImage`
  - `backgroundSize`
  - `backgroundPosition`
  - `backgroundColor`
  - `header`
  - `footer`
  - `transition` — slide transition type with optional `dir:` and `dur:` parameters
- Spot directives — all of the above keys work with a `_` prefix (e.g. `_class`, `_paginate`); they apply only to the current slide and do not carry forward
- Headings
- Paragraphs
- Bullet lists and ordered lists
- Images with local file paths
- Explicit visible captions via Markdown image title attribute: `![alt](url "Caption text")`
- Fenced and indented code blocks
- GFM-style tables at the semantic-model level
- Theme CSS extraction from `section`, `:root`, `body`, `h1`-`h6`, `pre`, and `code` — see the CSS property reference below

### Supported Rendering Features

- Editable PPTX text boxes for headings, paragraphs, lists, and code blocks
- Editable text fallback for tables
- Local image embedding with aspect-ratio-aware placement
- Explicit visible image captions: use the Markdown image title attribute `![alt text](url "Caption text")` to render a visible caption below the image; the caption is styled smaller than body text to distinguish it from regular content; alt text remains as accessibility metadata on the image shape and is not shown visibly
- Solid slide background color
- Full-slide background image via directive, `![bg](url)` image syntax, or theme `background-image`
- Line spacing applied from CSS `line-height`
- Letter spacing applied from CSS `letter-spacing`
- Text transform applied from CSS `text-transform` (`uppercase`, `lowercase`, `capitalize`)
- Optional template-copy workflow via `--template`
- Named template layout selection via front matter `layout`, comment directives `layout` / `_layout`, and authored template-slide selection via `Template[n]`
- Slide transitions via the `transition` directive (`fade`, `push`, `wipe`, `cut`, `cover`, `pull`, `random-bar`, `morph` with fade fallback); optional `dir:` and `dur:` parameters

### Not Yet Supported

- `![bg](url)` modifiers: `bg cover`, `bg contain`, `bg left`, `bg right`, and percentage sizing (`bg 50%`) — the image is always rendered full-bleed when using bg syntax
- Split-background layouts from multiple `![bg](url)` images on one slide (only the first is used)
- fragmented lists
- Marp Core `size`
- math
- Marp `fit`
- native PPTX table generation
- Marp CLI metadata semantics
- browser-based preview or HTML/PDF/image outputs
- CSS `background-attachment`, `background-blend-mode`, `background-origin`, `background-clip`
- CSS `font-variant`, `font-stretch`, `font-style`
- CSS `text-shadow`, `box-shadow`
- CSS `border`, `border-radius` on shapes
- CSS `opacity`, `filter`
- CSS custom properties (`--var-name`) and `var()` references
- CSS nesting, pseudo-classes (`:hover`, `:first-child`, etc.), or combinators
- CSS `@media`, `@keyframes`, or other at-rules

## Background Image Precedence

When multiple background image sources are present on the same slide, the following precedence applies (highest to lowest):

1. **`backgroundImage` directive** — a `<!-- backgroundImage: url -->` comment directive or front-matter `backgroundImage` key on the slide. Directives always win.
2. **`![bg](url)` image syntax** — a Markdown image whose alt text is exactly `bg` (case-insensitive). Applied only if no directive has set a background image for that slide.
3. **Theme `background-image`** — the background image defined in the active theme CSS. Applied at render time only if neither a directive nor `![bg](...)` syntax has set a background image.

When multiple `![bg](url)` images appear on the same slide, only the first one is used. The others are silently discarded (split-background layouts are not yet supported).

## Theme CSS Property Reference

The following table documents every CSS property that `MarpToPptx` recognises, which selectors it applies to, and what it maps to in the PPTX output.

### Supported Selectors

`section`, `:root`, `body`, `h1`–`h6`, `pre`, `code`

All other selectors are silently ignored.

### Supported Properties Per Selector

| CSS Property | Selectors | PPTX Mapping | Notes |
| --- | --- | --- | --- |
| `font-family` | `section`, `:root`, `body`, `h1`–`h6`, `pre`, `code` | Run font (Latin typeface) | First family in the comma-separated list is used. Quotes are stripped. |
| `font-size` | `section`, `:root`, `body`, `h1`–`h6`, `pre`, `code` | Run font size | Accepts `px`, `rem`, and unitless pt. `1px` = 0.75pt. `1rem` = 12pt. |
| `font-weight` | `section`, `:root`, `body`, `h1`–`h6` | Run bold flag | `bold`, `bolder`, or numeric ≥ 600 → bold. `normal`, `lighter`, or numeric < 600 → not bold. |
| `color` | `section`, `:root`, `body`, `h1`–`h6`, `pre`, `code` | Run fill color | Hex or `rgb()`/`rgba()` values. |
| `background-color` | `section`, `:root`, `body` | Slide background fill | Hex or `rgb()`/`rgba()` values. |
| `background` | `section`, `:root`, `body`, `pre`, `code` | Slide or code background | Color and URL extracted from shorthand. Other shorthand tokens (e.g. `no-repeat`) are ignored. |
| `background-image` | `section`, `:root`, `body` | Theme-level background image | `url(...)` value. Applied as full-bleed image behind all slides unless overridden per slide. |
| `background-size` | `section`, `:root`, `body` | Background image sizing | `cover` (default) scales the image to fill the slide, cropping as needed. `contain` scales it to fit within the slide with letterboxing. Other values fall back to cover behavior. |
| `background-position` | `section`, `:root`, `body` | Background image position | Supported keyword values: `center` (default), `top`, `bottom`, `left`, `right`, and two-keyword combinations such as `top left`, `bottom right`, `top center`. Percentage and length values are not supported and fall back to centered. |
| `padding` | `section`, `:root`, `body` | Slide padding inset | 1–4 value shorthand. Accepts `px`, `rem`, and unitless pt. |
| `line-height` | `section`, `:root`, `body`, `h1`–`h6`, `pre`, `code` | Paragraph line spacing (percent) | Unitless or `%` values treated as multipliers. `1.5` and `150%` both map to 150 % line spacing. |
| `letter-spacing` | `section`, `:root`, `body`, `h1`–`h6`, `pre`, `code` | Run character spacing | Accepts `px`, `rem`, and unitless pt. Converted to pt before storage. |
| `text-transform` | `section`, `:root`, `body`, `h1`–`h6` | Text content transform | `uppercase`, `lowercase`, and `capitalize` applied at render time. `none` or unrecognised values leave text unchanged. |

### Explicitly Unsupported Properties

The following properties are present in many Marp theme files but are silently ignored by the parser.

| CSS Property | Reason Not Supported |
| --- | --- |
| `background-attachment` | No PPTX equivalent. |
| `background-blend-mode` | No PPTX equivalent. |
| `background-repeat` | Background images are always rendered as full-bleed fills. |
| `background-clip` | No PPTX equivalent. |
| `font-style` | Italic run property not yet mapped. |
| `font-variant` | No PPTX equivalent for small-caps. |
| `font-stretch` | No PPTX equivalent. |
| `text-shadow` | No PPTX equivalent. |
| `box-shadow` | No PPTX equivalent. |
| `border` / `border-radius` | Shape borders not modelled per text element. |
| `opacity` | No PPTX per-shape opacity mapping. |
| `margin` / `margin-top` / `margin-bottom` | Slide layout uses a fixed padding inset, not per-element margins. |
| `text-decoration` | Underline/strikethrough run property not yet mapped. |
| `CSS custom properties` (`--name`, `var()`) | Variable resolution is not implemented. |

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
- Image alt text is accessibility metadata stored on the picture shape; it is not rendered as visible slide text.
- An explicit visible caption can be added via the Markdown title attribute `![alt](url "caption")`; the caption renders below the image in a smaller font size.

### Slide Authoring Syntax

- `src/MarpToPptx.Core/Parsing/FrontMatterParser.cs`
- `src/MarpToPptx.Core/Parsing/SlideTokenizer.cs`
- `src/MarpToPptx.Core/Parsing/MarpDirectiveParser.cs`

Important current behavior:

- Front matter is parsed only when it appears first and is delimited by `---`.
- Front matter parsing is intentionally simple and does not implement full YAML features such as multiline blocks or nested structures.
- Slide splitting ignores `---` inside fenced code blocks.
- `headingDivider` in front matter splits slides before headings at or above the specified level.
- Directive parsing only recognizes HTML comments that match `<!-- key: value -->`.
- Local directives (no `_` prefix) carry forward to subsequent slides.
- Spot directives (`_` prefix) apply only to the current slide and do not carry forward.

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
