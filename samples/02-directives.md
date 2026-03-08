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
  section.contrast {
    color: #FFFFFF;
  }
  section.contrast h1,
  section.contrast h2,
  section.contrast h3,
  section.contrast h4,
  section.contrast h5,
  section.contrast h6 {
    color: #FFFFFF;
  }
  section.contrast code {
    color: #FFFFFF;
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
| `header` | local | Text string |
| `footer` | local | Text string |

Add a `_` prefix to any key to make it a **spot directive** (current slide only).

---

# Carry-Forward After Local Override

No directives on this slide.

- `class: contrast` and `backgroundColor: #102A43` from the previous slide carry forward.
- `backgroundImage:` from the previous slide clears the global background image so the dark background remains visible.
- Front-matter `header` and `footer` are still present.

---

<!-- class: -->
<!-- backgroundImage: url(assets/accent-wave.svg) -->
<!-- backgroundSize: cover -->
# Background Image Directive

`backgroundImage` is a local directive, `backgroundSize: cover` locally overrides the global `contain` behavior, and `class:` clears the inherited `contrast` class.

---

# After Background Size Override

No directives on this slide.

- `backgroundImage`, `backgroundSize: cover`, and the cleared `class` from the previous slide carry forward.
- This verifies local override behavior on top of normalized front-matter defaults.

---

<!-- _paginate: false -->
<!-- _header: Spot Override Header -->
<!-- _footer: Spot Override Footer -->
## Spot-Directive Override

This slide uses **spot directives** (`_paginate`, `_header`, `_footer`).

- Pagination is **off** on this slide only.
- Header and footer text are overridden on this slide only.
- Spot directives use a `_` prefix and do not carry forward.

---

# After Spot Directives

No directives on this slide.

- `_paginate: false` does **not** carry forward — pagination is back on.
- `_header` and `_footer` do **not** carry forward — original front-matter header/footer appear.
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
- `backgroundColor` reverts to `#102A43` (last **local** directive, from slide 3).
- `paginate`, `header`, `footer` reflect their last inherited values.
