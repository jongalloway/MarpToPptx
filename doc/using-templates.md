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
