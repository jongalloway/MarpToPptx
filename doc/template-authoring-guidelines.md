# Template Authoring Guidelines

> **Just want to wire a template up to a deck?** Start with
> [`doc/using-templates.md`](using-templates.md). This file is the deep
> technical reference for when things don't work.

## Purpose

This note explains how MarpToPptx consumes a `--template` PowerPoint file
and what a template needs in order to produce usefully distinct output when
you select layouts with `layout:` front matter and `<!-- layout: ... -->` /
`<!-- _layout: ... -->` directives.

Use it when:

- evaluating a conference or corporate template before you start writing
  slides against it
- diagnosing why every generated slide looks the same even though layouts
  are bound correctly
- deciding whether to adjust a received template or to author content
  around its current structure

Related: [#85 Add template diagnostics and recommended layout guidance](https://github.com/jongalloway/MarpToPptx/issues/85)
for the automated version of the manual inspection steps below.

## How MarpToPptx consumes a template

Two things happen when a `layout`/`_layout` directive names a template
layout:

1. The generated slide is bound to that `<p:sldLayout>` part, so PowerPoint
   inherits the layout's background, artwork, and text styles.
2. Text content is written into **placeholder shapes** that reference the
   layout's placeholders by `type` and `idx`. The slide-level shape omits
   geometry and run-level font properties so the layout's own formatting
   cascades. This is what makes templates actually look different.

If the named layout has no title or body placeholder, the renderer falls
back to standalone positioned text boxes for that content. The slide still
opens cleanly, but template text styling will not apply.

There is one separate path for authored template slides: `<!-- _layout: Template[1] -->`
clones template slide 1 itself, preserves its existing slide artwork, and replaces
its standalone text boxes heuristically. Use that when a template's branded title
slide is not actually represented by a reusable PowerPoint layout.

## Checklist for a template you receive

Work through this before committing to a template. The goal is to know which
layouts will pay off when targeted from Markdown.

### 1. Layout names must be populated and distinct

MarpToPptx matches `layout: Foo` against two attributes, case-insensitive,
in this order:

| Where | What |
|---|---|
| `<p:sldLayout matchingName="...">` | optional, often empty in real templates |
| `<p:cSld name="...">` | usually populated; what PowerPoint shows in the layout gallery |

Check that every layout you intend to target has a non-empty `cSld name`.
If two layouts under different masters share the exact same name (common:
"Title and Content" appearing on a primary and a secondary master), the
**first match in iteration order wins**. Rename one in PowerPoint if you
need to target the other.

### 2. Title and body placeholders must exist on targeted layouts

Placeholder-based rendering maps content as follows:

| Markdown content | Target placeholder | `<p:ph>` matched |
|---|---|---|
| First heading (any level `#`–`######`) | Title | `type="title"`, `type="ctrTitle"` |
| Remaining headings, paragraphs, bullet / numbered lists | Body | `type="body"`, `type="subTitle"`, **or typeless with `idx`** (e.g. `<p:ph idx="1"/>`) |
| Images, video, audio, code blocks, tables | *(none — standalone shapes)* | — |

The heading level does **not** affect placeholder selection — `## Topic`
fills the same title slot as `# Topic`. Level only matters for indentation
in standalone (non-placeholder) fallback rendering.

The typeless-`idx` body fallback matters because real "Title and Content"
layouts (`type="obj"`) almost always declare the content slot as
`<p:ph idx="1"/>` with no `type` attribute. That is matched. Footer-like
placeholders (`ftr`, `dt`, `sldNum`) are excluded from body fallback.

Open the layout in PowerPoint's Slide Master view and confirm it carries
a title placeholder and a body/subtitle placeholder. If a placeholder
you need is missing, add it on the layout (see
**View → Slide Master → Insert Placeholder**), then re-save the template.

The `Blank` layout is not expected to carry placeholders; avoid targeting
it with `layout:` unless you actually want standalone-shape fallback.

### 3. Understand where the visual design lives

The most common reason for "every slide looks the same" is that **all the
visible artwork is on the slide master, not on individual layouts**, and
the layouts differ only by placeholder metadata.

The diagnostic split:

| Design element lives on… | Effect in generated output |
|---|---|
| Slide master only | Identical across every slide regardless of layout |
| Individual layout | Visible only on slides bound to that layout |
| Placeholder formatting (font, size, color, alignment) | **Only visible when content is written into that placeholder** — this is what placeholder-based rendering unlocks |

If you want title slides, section dividers, and content slides to look
meaningfully different, the per-layout differences must come from either
layout-level artwork or placeholder text styles. Shared master artwork is
fine for a consistent frame (logos, footers).

### 4. Quick manual inspection

In PowerShell, with the template `.pptx` path in `$t`:

```powershell
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($t)
$ns = @{ p = 'http://schemas.openxmlformats.org/presentationml/2006/main' }
$rows = foreach ($e in $zip.Entries | Where-Object FullName -like 'ppt/slideLayouts/slideLayout*.xml') {
    $xml = [xml](New-Object IO.StreamReader $e.Open()).ReadToEnd()
    $nm = New-Object Xml.XmlNamespaceManager $xml.NameTable
    $nm.AddNamespace('p', $ns.p)
    $shapes = $xml.SelectNodes('//p:sp', $nm)
    $phs    = $xml.SelectNodes('//p:sp//p:ph', $nm)
    [pscustomobject]@{
        Layout   = $e.Name
        Name     = $xml.SelectSingleNode('//p:cSld', $nm).GetAttribute('name')
        Type     = $xml.DocumentElement.GetAttribute('type')
        Shapes   = $shapes.Count
        PhShapes = $phs.Count
        NonPh    = $shapes.Count - $phs.Count
        PhTypes  = ($phs | ForEach-Object {
            $t = $_.GetAttribute('type'); $i = $_.GetAttribute('idx')
            if ($t) { $t } elseif ($i) { "[idx=$i]" } else { '(none)' }
        }) -join ','
    }
}
$zip.Dispose()
$rows | Format-Table -AutoSize
```

Read the output as:

- `NonPh > 0` → the layout carries its own artwork; layout selection alone
  will change the look.
- `NonPh == 0` → all visible difference must come from placeholder
  formatting. Check `PhTypes`.
- **Title:** need `title` or `ctrTitle` in `PhTypes`. Neither present →
  the slide heading renders as a standalone shape.
- **Body:** need `body`, `subTitle`, **or an `[idx=N]` entry** (typeless
  content slot, standard on `obj`-type layouts). None of those → body
  content renders as standalone shapes.

### 5. Recommended layout choices

Based on conventional PowerPoint template structure:

| Use | Directive / front matter | Typical layout name | Layout `type` |
|---|---|---|---|
| Deck-wide default | `layout: Title and Content` (front matter) | "Title and Content" | `obj` or `tx` |
| Title slide | `<!-- _layout: Title Slide -->` | "Title Slide" | `title` |
| Section divider | `<!-- _layout: Section Header -->` | "Section Header" | `secHead` |
| Side-by-side | `<!-- _layout: Two Content -->` | "Two Content" | `twoObj` |
| Captioned picture | `<!-- _layout: Picture with Caption -->` | "Picture with Caption" | `picTx` |

Names vary per template; verify with the inspection snippet above. The
`type` attribute is more stable than the display name but is not what
`layout:` matches against.

## Known limitations

- Only the **first** title-like and the **first** body-like placeholder per
  layout are populated. Layouts with multiple body placeholders
  ("Two Content", "Comparison") receive all text content in the first one.
  Body-placeholder search order: explicit `body`/`subTitle` first, then
  the first typeless `idx`-only placeholder (skipping `ftr`/`dt`/`sldNum`).
- `Template[n]` slide cloning is heuristic, not semantic. It keeps the slide's
  existing artwork, picks the largest upper-half text box as the title box, and
  fills the remaining text boxes top-to-bottom. It is intended for title-slide-like
  cases, not arbitrary multi-slot slide templating.
- Picture placeholders (`type="pic"`) are not yet targeted; images render
  as standalone positioned shapes.
- Slide-number, date, and footer placeholders are inherited from the
  layout but not populated by the renderer; PowerPoint fills them when
  **Insert → Header & Footer** is enabled on the finished deck.
- When `layout:` matches, Marp theme styling (`class`, `backgroundColor`,
  `backgroundImage`, `header`, `footer`) is suppressed for that slide so
  the template is authoritative. `paginate` still emits a slide-number
  field.
