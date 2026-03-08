---
theme: community-dracula
paginate: true
lang: en-US
header: Awesome Marp community theme smoke
footer: Dracula-inspired sample
---

<!-- _class: lead -->
# Dracula-Inspired Smoke Deck

This sample takes cues from the community Dracula theme: saturated contrast, neon-accent headings, and code-forward slides.

- Dark stage background
- Bright semantic accents
- Code-heavy presenter workflow
- One compact wrap-up slide

---

## Why It Matters

- Dark themes stress text contrast differently than light themes.
- Accent-heavy headings make hierarchy failures obvious.
- Code blocks need to remain readable without flattening into a single dark rectangle.

---

<!-- _class: codefocus -->
## Terminal Slide

```bash
marp2pptx slides.md \
  --theme-css dracula.css \
  -o slides.pptx
```

The code surface should stay distinct from the slide background while the body copy remains legible.

---

<!-- _class: accent -->
## Accent Slide

Use this slide to verify that a single-slide class can intensify the accent palette without leaking into the next slide.

Inline `code` should still read clearly.

---

<!-- _class: compact -->
## Close

The final slide reduces body size slightly to confirm pagination and footer/header placement remain stable.

1. Theme CSS is loaded.
2. Class variants are isolated.
3. Output remains editable.