---
theme: diagram-showcase
paginate: true
header: DiagramForge in MarpToPptx
footer: Mermaid and diagram fences
---

# Diagram Showcase

MarpToPptx can render both `mermaid` and `diagram` fenced blocks through DiagramForge.

---

## Mermaid Sequence

```mermaid
flowchart LR
  Author[Markdown Authoring] --> Parse[Parse Markdown]
  Parse --> Theme[Apply Theme CSS]
  Theme --> Render[Render Open XML Shapes]
  Render --> Deck[Editable PPTX]
  Render -. diagram .-> Svg[DiagramForge SVG]
```

---

## Mermaid Block Diagram

```mermaid
block-beta
  columns 3
  Draft api<[[Review]]>(right) Ship
  space qa<[[QA]]>(down) space
  Archive:3
  Draft --> Ship
  Archive -- "feedback" --> Draft
```

---

## Mermaid State Diagram

```mermaid
stateDiagram-v2
  [*] --> Draft
  Draft --> Review
  Review --> Approved
  Review --> Draft : changes requested
  Approved --> Published
  Published --> [*]
```

---

## Mermaid Mindmap

```mermaid
mindmap
  root(Diagram Support)
    Mermaid
      Flowchart
      Block
      State
    Diagram
      Matrix
      Pyramid
```

---

## Mermaid With Dracula Theme

This Mermaid diagram uses DiagramForge frontmatter to switch to the built-in Dracula theme and apply additional styling overrides.

```mermaid
---
theme: dracula
palette: ["#FFB86C", "#8BE9FD", "#50FA7B"]
borderStyle: rainbow
fillStyle: diagonal-strong
shadowStyle: soft
transparent: true
---
flowchart LR
  A[Plan] --> B[Build]
  B --> C[Ship]
```

---

## Conceptual Matrix

```diagram
diagram: matrix
rows:
  - Important
  - Not Important
columns:
  - Urgent
  - Not Urgent
```

---

## Conceptual Matrix With Prism Theme

This conceptual diagram uses DiagramForge frontmatter to apply the built-in Prism theme inside the fenced block.

```diagram
---
theme: prism
palette: ["#6C5CE7", "#00CEC9", "#FDCB6E", "#FF7675"]
shadowStyle: soft
transparent: true
---
diagram: matrix
rows:
  - High Impact
  - Lower Impact
columns:
  - Quick Wins
  - Strategic Bets
```

---

## Conceptual Pyramid

```diagram
diagram: pyramid
levels:
  - Vision
  - Strategy
  - Delivery
  - Feedback
```

---

## Conceptual Pyramid With Dracula Theme

```diagram
---
theme: dracula
palette: ["#FFB86C", "#8BE9FD", "#BD93F9", "#50FA7B"]
borderStyle: rainbow
shadowStyle: soft
transparent: true
---
diagram: pyramid
levels:
  - Vision
  - Strategy
  - Delivery
  - Feedback
```
