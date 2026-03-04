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

```csharp
var compiler = new MarpCompiler();
var deck = compiler.Compile(markdown, sourcePath, themeCss);
```

The code block should use the theme's monospace font, dark background, and configured code typography.
