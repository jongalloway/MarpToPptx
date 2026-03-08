---
theme: community-beam
paginate: true
lang: en-US
header: Awesome Marp community theme smoke
footer: Beam-inspired sample
---

<!-- _class: lead -->
# Beam-Inspired Smoke Deck

This repo-authored sample is based on the academic presentation style common in community Beam and Beamer-like Marp themes.

- Strong title treatment
- Dense agenda slide
- Contrast statement slide
- Appendix-style typography shift

---

## Agenda

1. Frame the problem clearly.
2. Present one crisp claim.
3. Show a code-shaped artifact.
4. Close with a practical takeaway.

This deck is intentionally simple, but the theme should still feel formal and lecture-friendly.

---

<!-- _class: statement -->
## Design Signal

Community Beam themes usually lean on three things:

- restrained color
- assertive headings
- slides that read well from the back of a room

That makes them good smoke fixtures for layout pressure, spacing, and heading contrast.

---

## Implementation Slice

```csharp
var options = new PptxRenderOptions
{
    SourceDirectory = Path.GetDirectoryName(inputPath),
    AllowRemoteAssets = false,
};
```

Block code should stay legible against the darker code surface.

---

<!-- _class: appendix -->
## Appendix Mode

Use this slide to verify that a class variant can tighten the body size without collapsing heading hierarchy.

- Inline `code` should remain readable.
- Pagination should continue to flow.
- Header and footer text should still be visible.