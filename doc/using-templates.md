# Using a PowerPoint Template

You've been handed a conference or corporate `.pptx` template. Here's how to
get your Marp deck to use its layouts.

## 1. Find the layout names

Open the template in PowerPoint. On the **Home** tab, click the **Layout**
dropdown. The names you see there — "Title Slide", "Title and Content",
"Section Header", and so on — are the names you'll use in Markdown.

Matching is case-insensitive, but spelling and spacing must match exactly.

## 2. Set a deck-wide default

In your front matter, `layout:` picks the layout for normal content slides:

```markdown
---
marp: true
layout: Title and Content
---
```

Your opening title slide is still auto-detected — you don't need to set it.

## 3. Override one slide

Use a spot directive (underscore prefix) at the top of that slide:

```markdown
---

<!-- _layout: Section Header -->

## Part Two

Deep dive into the details.
```

The underscore means **this slide only**. The next slide goes back to the
default.

## 4. Override a run of slides

Leave the underscore off and the change **sticks** until you change it again:

```markdown
---

<!-- layout: Two Content -->

## Comparison A

---

## Comparison B

---

<!-- layout: Title and Content -->

## Back to normal
```

## 5. Use a specific slide from the template

Some templates put the fancy title treatment on slide 1 itself instead of on a
reusable layout. In that case, point MarpToPptx at the authored template slide:

```markdown
<!-- _layout: Template[1] -->
```

That clones template slide 1, keeps its background/artwork, and replaces its
text boxes with your content. This is mainly useful for branded title slides.

Use this only when the template's special title slide is not exposed as a real
layout in PowerPoint. If you do see a normal layout name like `Title Slide` in
the **Layout** dropdown, prefer that first.

## A complete minimal example

```markdown
---
marp: true
layout: Title and Content
---

<!-- _layout: Title Slide -->

# My Presentation

Author Name · Conference 2026

---

## First Content Slide

- Point one
- Point two

---

<!-- _layout: Section Header -->

## Part One
```

If your template's fancy title page is slide-authored instead of layout-authored,
swap that first override for:

```markdown
<!-- _layout: Template[1] -->
```

Run it with:

```bash
marp2pptx deck.md --template conference-template.pptx -o deck.pptx
```

## How content fills the layout

On a template-bound slide:

| Your Markdown | Goes into |
|---|---|
| First heading (`#`, `##`, any level) | The layout's title area |
| Paragraphs and bullet lists that follow | The layout's body area |
| Images, code blocks, tables | Placed as standalone shapes |

Font, size, and colour come from the template, not your CSS. That's the
point — you're asking the template to style things.

## It still looks the same?

Nine times out of ten: the layout name in your Markdown doesn't match what
PowerPoint calls it. Open the **Layout** dropdown again and check
character-for-character.

If names match and it *still* looks the same, the template itself may carry
all its artwork on the master rather than per-layout. See
`doc/template-authoring-guidelines.md` for diagnosing that case.

## Diagnosing a template

MarpToPptx ships a standalone diagnostics tool that can report on a template's
layout structure and make recommendations:

```bash
dotnet run --project src/MarpToPptx.TemplateDiagnostics -- diagnose conference.pptx
```

This prints a table of all layouts with their placeholder coverage and semantic
roles, plus recommended `layout:` / `_layout:` values for your Markdown.

## Template doctor: inspecting and repairing templates

Real-world conference and corporate templates are often valid PowerPoint files
but still produce suboptimal results with MarpToPptx.  The **template doctor**
goes beyond diagnostics and can identify — and fix — structural problems.

### What the doctor checks

| Issue code | Severity | What it means |
|---|---|---|
| `DuplicateLayoutName` | Warning | Two or more layouts share the same name; only the first can be targeted by directive. |
| `EmptyLayoutName` | Warning | A layout has no name and will be referenced by position only. |
| `ContentLayoutMissingTitlePlaceholder` | Warning | A content layout lacks a title placeholder; headings won't use the template's title geometry. |
| `ContentLayoutMissingBodyPlaceholder` | Warning | A content layout lacks a body placeholder; body content won't use the template's body geometry. |
| `PlaceholderGeometryInherited` | **Fixable** | A placeholder identity exists on the layout but its position/size is inherited from the slide master instead of being declared on the layout itself.  The renderer recovers via master fallback, but materializing the geometry improves portability. |
| `TypelessIndexedBodyPlaceholder` | Info | Body content is accessible only via a typeless indexed placeholder — this is standard and fully supported. |
| `UnmappableLayoutRole` | Info | A picture-caption or comparison layout can't yet be auto-selected by the renderer. |
| `VisuallyRedundantLayouts` | Info | Multiple layouts share the same role and have no distinct shapes; they appear identical at render time. |

### Dry-run mode

Inspect without modifying anything:

```bash
dotnet run --project src/MarpToPptx.TemplateDiagnostics -- doctor conference.pptx
```

### Write a repaired copy

Apply all safe, automatable fixups to a *new* file (the original is never touched):

```bash
dotnet run --project src/MarpToPptx.TemplateDiagnostics -- \
  doctor conference.pptx \
  --write-fixed-template conference-fixed.pptx
```

The console output lists every fix that was applied.

### JSON output

Pipe the report into other tools:

```bash
dotnet run --project src/MarpToPptx.TemplateDiagnostics -- \
  doctor conference.pptx --json
```
