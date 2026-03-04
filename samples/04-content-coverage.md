---
theme: default
paginate: true
---

# Content Coverage

This sample mixes most of the content elements currently modeled by `MarpToPptx`.

---

## Images

![Architecture sketch](assets/stack-diagram.svg)

The image above should be embedded from a local SVG file.

---

## Code Blocks

```json
{
  "title": "MarpToPptx",
  "formats": ["pptx"],
  "mode": "editable"
}
```

---

## Mixed Lists

- Top-level bullet
- Another bullet
  - Nested bullet content should remain associated with the list

1. First ordered item
2. Second ordered item

---

## Table Fallback

| Feature | Expected Behavior |
| --- | --- |
| Table parsing | Create a `TableElement` |
| Current rendering | Editable text fallback |
| Future goal | Native PPTX tables |
