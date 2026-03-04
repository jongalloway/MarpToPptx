---
theme: default
paginate: true
headingDivider: 2
---

# Compatibility Gaps

This deck intentionally uses upstream Marp features that are not fully implemented in `MarpToPptx` yet.

## Heading Divider

Upstream Marpit can split slides automatically with `headingDivider`, but the current parser does not.

---

<!-- _backgroundColor: #EEF2FF -->
# Spot Directive

The `_backgroundColor` syntax is supported upstream for one-slide overrides, but is not implemented here.

---

## Background Image Syntax

![bg](assets/accent-wave.svg)

Upstream Marpit treats `bg` as a background image keyword. The current implementation treats this as a normal image with alt text.

---

## Fragmented List Syntax

- Item one
- Item two

<!--
This deck is a reminder that there is currently no fragment model,
so list reveal behavior should not be expected in PPTX output.
-->

---

## Math And Fit

Inline math like $E = mc^2$ and fit comments such as `# <!-- fit -->` are useful future compatibility checks, but are not currently rendered with Marp semantics.
