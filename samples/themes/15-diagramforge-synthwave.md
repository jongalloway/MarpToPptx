---
theme: diagramforge-synthwave
paginate: true
lang: en-US
header: DiagramForge theme smoke
footer: Synthwave sample
---

<!-- _class: lead -->
# Synthwave Smoke Deck

A retro-future theme shared with DiagramForge: deep purple backgrounds, hot-pink accents, and sunset-orange edge lines.

- Warm dark stage
- Retro neon palette
- Code blocks on surface panels
- Compact closing slide

---

## Why Synthwave?

- Deep-purple backgrounds create a warm dark stage distinct from pure-black themes.
- Sunset-orange and hot-pink accents stress heading hierarchy readability.
- Code blocks must remain readable on the violet surface.

---

<!-- _class: sunset -->
## Sunset Slide

```bash
marp2pptx slides.md \
  --theme-css 15-diagramforge-synthwave.css \
  -o slides.pptx
```

The neon red-pink heading should pop against the deep surface while body text stays legible.

---

<!-- _class: accent -->
## Accent Slide

This slide verifies that the hot-pink accent class applies cleanly to a single slide without bleeding into the next.

Inline `code` should remain readable on the deep-purple surface.

---

## Color Palette

| Role | Color |
|---|---|
| Background | `#1A0030` |
| Foreground | `#F0E0FF` |
| Accent | `#FF6EC7` |
| Surface | `#2A1040` |
| Edge | `#FFB347` |
| Group | `#B24BF3` |

---

<!-- _class: compact -->
## Close

Final compact slide to confirm the layout stays readable at a smaller font size.

- Theme CSS is loaded.
- Single-slide class hints are isolated.
- The rendered PPTX stays editable.
