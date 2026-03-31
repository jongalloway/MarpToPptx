---
# Directive keys set in front matter apply globally to the entire deck.
# All keys listed here are fully supported by MarpToPptx.
theme: gaia
paginate: true
lang: en-US
backgroundColor: "#F7F3E8"
backgroundImage: url(assets/accent-wave.svg)
backgroundSize: contain
header: MarpToPptx Directive Sample
footer: Sample deck footer
# The 'style' key merges inline CSS with any --theme-css file.
style: |
  section.lead {
    font-weight: 600;
  }
  section.contrast {
    font-weight: 600;
  }
---

<!-- class: lead -->
# Directive Coverage

This slide uses front matter plus an inline `class` directive.

- `theme`, `paginate`, `backgroundColor`, `header`, `footer`, `lang`, and `style` are set in front matter and apply globally.
- `backgroundImage: url(...)` and `backgroundSize: contain` are also set in front matter and should normalize and apply globally.
- `<!-- class: lead -->` is a **local directive** — it carries forward to subsequent slides.

<!-- HTML comments that do not match the key: value directive pattern become presenter notes. -->
<!-- presenter note: slide 1 -->

---

# Carry-Forward Check

This slide has no directives of its own.

- `class: lead` from the previous slide **carries forward** here (local directive behavior).
- Front-matter `paginate`, `header`, `footer`, `backgroundColor`, `backgroundImage`, and `backgroundSize` also persist.

---

<!-- backgroundColor: #102A43 -->
<!-- backgroundImage: -->
<!-- class: contrast -->
<!-- color: #FFFFFF -->
# Per-Slide Background Color

These directives are **local** — they apply here and carry forward.

The following directive keys are supported in HTML comments (local and spot):

| Key | Scope | Notes |
| --- | --- | --- |
| `theme` | local | Applies this slide onward |
| `paginate` | local | `true` or `false` |
| `class` | local | CSS class name from theme |
| `backgroundImage` | local | `url(...)` or bare path |
| `backgroundSize` | local | e.g. `cover`, `contain` |
| `backgroundColor` | local | Hex or `rgb()` |
| `color` | local | Default text color for slide text |
| `header` | local | Text string |
| `footer` | local | Text string |

Add a `_` prefix to any key to make it a **spot directive** (current slide only).

---

# Carry-Forward After Local Override

No directives on this slide.

- `class: contrast`, `backgroundColor: #102A43`, and `color: #FFFFFF` from the previous slide carry forward.
- `backgroundImage:` from the previous slide clears the global background image so the dark background remains visible.
- Front-matter `header` and `footer` are still present.

---

<!-- class: -->
<!-- color: -->
<!-- backgroundColor: #F7F3E8 -->
<!-- backgroundImage: url(assets/accent-wave.svg) -->
<!-- backgroundSize: cover -->
# Background Image Directive

`backgroundImage` is a local directive, `backgroundSize: cover` locally overrides the global `contain` behavior, and `class:`, `color:`, plus `backgroundColor: #F7F3E8` return the slide to the deck's lighter styling.

---

# After Background Size Override

No directives on this slide.

- `backgroundImage`, `backgroundSize: cover`, and the cleared `class` and `color` from the previous slide carry forward.
- `backgroundColor: #F7F3E8` from the previous slide also carries forward, restoring a light backdrop for the rest of the deck.
- This verifies local override behavior on top of normalized front-matter defaults.

---

<!-- _paginate: false -->
<!-- _header: Spot Override Header -->
<!-- _footer: Spot Override Footer -->
<!-- _color: #8B0000 -->
## Spot-Directive Override

This slide uses **spot directives** (`_paginate`, `_header`, `_footer`, `_color`).

- Pagination is **off** on this slide only.
- Header and footer text are overridden on this slide only.
- Text color is overridden on this slide only.
- Spot directives use a `_` prefix and do not carry forward.

---

# After Spot Directives

No directives on this slide.

- `_paginate: false` does **not** carry forward — pagination is back on.
- `_header`, `_footer`, and `_color` do **not** carry forward — original inherited values appear again.
- The cleared `class` and `backgroundImage: url(...)` from slide 5 still carry forward (they were local, not spot).

---

<!-- _class: special -->
<!-- _backgroundColor: #FFD700 -->
## Spot Class and Background

This slide uses spot directives for `_class` and `_backgroundColor`.

- Displays with class `special` and gold background on this slide only.
- Neither carries forward to the next slide.

---

# Final Carry-Forward Verification

No directives.

- `class` remains cleared (last **local** directive, from slide 5).
- `backgroundColor` remains `#F7F3E8` (last **local** directive, from slide 5).
- `paginate`, `header`, `footer` reflect their last inherited values.
