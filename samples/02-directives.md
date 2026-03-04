---
theme: gaia
paginate: true
backgroundColor: "#F7F3E8"
header: MarpToPptx Directive Sample
footer: Sample deck footer
---

<!-- class: lead -->
# Directive Coverage

This slide uses front matter plus an inline `class` directive.

- Front matter header should appear at the top of every slide.
- Front matter footer should appear at the bottom unless a slide overrides it.

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

<!-- backgroundImage: url(assets/accent-wave.svg) -->
# Background Image Directive

Use this slide to verify that a local background image fills the slide.

---

<!-- paginate: false -->
<!-- header: Directive Override Header -->
<!-- footer: Directive Override Footer -->
## Pagination Override

This slide turns pagination off through a directive comment and overrides the inherited header/footer text.
