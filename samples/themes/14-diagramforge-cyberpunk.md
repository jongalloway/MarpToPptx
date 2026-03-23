---
theme: diagramforge-cyberpunk
paginate: true
lang: en-US
header: DiagramForge theme smoke
footer: Cyberpunk sample
---

<!-- _class: lead -->
# Cyberpunk Smoke Deck

A neon-noir theme shared with DiagramForge: near-black backgrounds, hot-pink accents, and cyan edge lines.

- Extreme dark stage
- Neon accent hierarchy
- Code blocks on surface panels
- Compact closing slide

---

## Why Cyberpunk?

- Dark backgrounds stress text contrast and glow readability.
- Multiple neon accents make heading hierarchy failures obvious.
- Code blocks must stay distinct from the near-black stage.

---

<!-- _class: neon -->
## Neon Slide

```bash
marp2pptx slides.md \
  --theme-css 14-diagramforge-cyberpunk.css \
  -o slides.pptx
```

The neon-green heading should contrast sharply with the deep surface while body text stays legible.

---

<!-- _class: accent -->
## Accent Slide

This slide verifies that the hot-pink accent class applies cleanly to a single slide without bleeding into the next.

Inline `code` should remain readable on the dark surface.

---

## Color Palette

| Role | Color |
|---|---|
| Background | `#0A0A1A` |
| Foreground | `#E0E0F0` |
| Accent | `#FF2D95` |
| Surface | `#12122A` |
| Edge | `#00F0FF` |
| Group | `#8B5CF6` |

---

<!-- _class: compact -->
## Close

Final compact slide to confirm the layout stays readable at a smaller font size.

- Theme CSS is loaded.
- Single-slide class hints are isolated.
- The rendered PPTX stays editable.
