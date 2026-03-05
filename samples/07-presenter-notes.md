# Presenter Notes Smoke

This deck explicitly exercises presenter notes packaging and slide-to-notes attachment.

<!--
Presenter note for slide 1.
- checklist item
- second item with **bold**, *italic*, and `code` markers that should render with formatting

```csharp
var total = items.Count;
Console.WriteLine(total);
```
-->

---

## No Notes On This Slide

This slide is the control case and should not get a notes part.

---

<!-- header: Presenter Notes Header -->
## Directives And Notes

This slide mixes a real directive comment with presenter notes.

<!--
First presenter note paragraph for slide 3.
Formatting markers like **bold** and _italic_ should render as formatted note text.
-->
<!-- Second presenter note line for slide 3. -->

---

## Final Notes Check

The final slide confirms notes still attach correctly later in the deck.

<!--
Presenter note for slide 4.
1. Numbered item
2. Another item

```json
{ "status": "ok" }
```
-->