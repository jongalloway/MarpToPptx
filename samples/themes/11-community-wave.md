---
theme: community-wave
paginate: true
lang: en-US
header: Awesome Marp community theme smoke
footer: Wave-inspired sample
---

<!-- _class: lead -->
<!-- _backgroundImage: -->
# Wave-Inspired Smoke Deck

This sample is based on the bright conference-talk decks common in community Wave and keynote-style Marp themes.

- Soft background art
- Bold section bands
- Compact closing slide
- High-contrast code sample

---

## Story Arc

Community Wave themes often push a speaker deck toward a more branded, event-ready look.

- open strong
- keep headings oversized
- let color blocks signal transitions
- finish with a compact action slide

---

<!-- _class: band -->
## Transition Slide

Use this slide to verify that a class variant can swap the background and heading treatment without changing the deck structure.

Inline `code` should remain readable while the section color changes.

---

## Presenter Workflow

```bash
pwsh ./scripts/Invoke-PptxSmokeTest.ps1 \
  -InputMarkdown samples/themes/11-community-wave.md \
  -ThemeCss samples/themes/11-community-wave.css \
  -Configuration Release -CiSafe
```

The generated package should validate cleanly before the optional PowerPoint step.

---

<!-- _class: compact -->
## Close

The final slide intentionally reduces body size to make sure compact variants still paginate and lay out correctly.

1. Theme CSS is loaded.
2. Class variants are applied.
3. Open XML validation succeeds.