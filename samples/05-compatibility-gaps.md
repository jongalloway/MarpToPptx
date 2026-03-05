---
theme: default
paginate: true
headingDivider: 2
---

# Compatibility Gaps

This deck intentionally uses upstream Marp features that are not fully implemented in `MarpToPptx` yet.

## Heading Divider

`headingDivider` is now implemented — this deck uses `headingDivider: 2` in front matter, so `##` headings create automatic slide breaks.

---

<!-- _backgroundColor: #EEF2FF -->
# Implemented Spot Directive

The `_backgroundColor` syntax is now implemented here.

This slide remains in the compatibility deck as a quick visual regression check that spot overrides still apply to a single slide only.

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

---

## Background Size: Contain

<!-- backgroundImage: assets/accent-wave.svg -->
<!-- backgroundSize: contain -->

When `backgroundSize: contain` is set, the background image should be fitted within the slide without cropping.

---

## Background Size: Cover (Default)

<!-- _backgroundSize: cover -->

Without an explicit `backgroundSize` (or with `cover`), the background image should stretch to fill the slide, cropping if needed.
