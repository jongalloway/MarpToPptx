---
theme: editorial
paginate: true
---

# Theme CSS Sample

Use this deck together with `samples/03-theme.css`.

- Verifies theme name flow from front matter
- Exercises font, size, color, and padding extraction
- Confirms code block styling is picked up from CSS

---

## Body And Heading Hierarchy

This paragraph should inherit the body font and color from the theme CSS.

### Third-Level Heading

The heading size hierarchy should reflect the CSS file.

---

```csharp
var compiler = new MarpCompiler();
var deck = compiler.Compile(markdown, sourcePath, themeCss);
```

The code block should use the theme's monospace font and dark background.
