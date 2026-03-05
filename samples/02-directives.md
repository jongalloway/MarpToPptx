---
theme: gaia
paginate: true
lang: en-US
backgroundColor: "#F7F3E8"
header: MarpToPptx Directive Sample
footer: Sample deck footer
style: |
  section.contrast {
    color: #FFFFFF;
  }
---

<!-- class: lead -->
# Directive Coverage

This slide uses front matter plus an inline `class` directive.

- Front matter header should appear at the top of every slide.
- Front matter footer should appear at the bottom unless a slide overrides it.

---

# Carry-Forward Check

This slide has no directives of its own.

- `class: lead` from the previous slide should **carry forward** here.
- Front-matter `paginate`, `header`, `footer`, `backgroundColor` should persist.

---

<!-- backgroundColor: #102A43 -->
<!-- class: contrast -->
# Per-Slide Background Color

The current implementation supports these directive keys:

- `theme`
- `paginate`
- `class`
- `backgroundImage`
- `backgroundColor`
- `header`
- `footer`

---

# Carry-Forward After Local Override

No directives on this slide.

- `class: contrast` and `backgroundColor: #102A43` from slide 3 should carry forward.
- Front-matter `header` and `footer` should still be present.

---

<!-- backgroundImage: url(assets/accent-wave.svg) -->
# Background Image Directive

Use this slide to verify that a local background image fills the slide.

---

<!-- _paginate: false -->
<!-- _header: Spot Override Header -->
<!-- _footer: Spot Override Footer -->
## Spot-Directive Override

This slide uses **spot directives** (`_paginate`, `_header`, `_footer`).

- Pagination should be **off** on this slide only.
- Header and footer text should be overridden on this slide only.

---

# After Spot Directives

No directives on this slide.

- `_paginate: false` should **not** carry forward — pagination should be back on.
- `_header` and `_footer` should **not** carry forward — original front-matter header/footer should appear.
- `class: contrast` and `backgroundColor: #102A43` from slide 3 should still carry forward.

---

<!-- _class: special -->
<!-- _backgroundColor: #FFD700 -->
## Spot Class and Background

This slide uses spot directives for `_class` and `_backgroundColor`.

- Should display with class `special` and gold background.
- Neither should carry to the next slide.

---

# Final Carry-Forward Verification

No directives.

- `class` should revert to `contrast` (last local directive, from slide 3).
- `backgroundColor` should revert to `#102A43` (last local directive, from slide 3).
- `paginate`, `header`, `footer` should reflect their last inherited values.
