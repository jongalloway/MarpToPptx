---
theme: editorial
paginate: true
---

# Theme CSS Sample

Use this deck together with `samples/03-theme.css`.

- Verifies theme name flow from front matter
- Exercises font, size, color, padding, and background image extraction
- Confirms line height, letter spacing, text transform, and font weight parsing
- Confirms code block styling is picked up from CSS

---

## Body And Heading Hierarchy

This paragraph should inherit the body font and color from the theme CSS.

Body copy should also reflect the theme line height and subtle letter spacing.

### Third-Level Heading

The heading size hierarchy should reflect the CSS file, and the top-level heading should render uppercase.

---

<!-- _class: expansive -->
## Class Variant: Layout Stress

This slide uses a class variant that increases body and heading size enough to change placement.

The layout engine should account for the class-adjusted typography rather than sizing frames from the base theme and rendering with larger text afterward.

This makes the sample useful as a visual smoke check for layout consistency, not just color and font application.

---

```csharp
var compiler = new MarpCompiler();
var deck = compiler.Compile(markdown, sourcePath, themeCss);
```

The code block should use the theme's monospace font, dark background, and configured code typography.

---

<!-- _backgroundImage: -->
<!-- _class: lead -->
# Class Variant: Lead

This slide uses `_class: lead` to select the `section.lead` class variant for this slide only.

- Background should be dark (#102A43).
- Body text should be light (#F0F4F8).
- The heading should be gold (#F7C948).

---

<!-- _backgroundImage: -->
<!-- _class: invert -->
# Class Variant: Invert

This slide uses `_class: invert` for this slide only.

- Background should be dark (#1A1A2E).
- Body text should be light (#E0E0E0).
