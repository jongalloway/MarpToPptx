---
title: Slide Transitions Demo
transition: fade
---

# Slide Transitions

All slides in this deck inherit the default `fade` transition set in front matter.

---

<!-- transition: push dir:right -->
# Push Right

This slide and all subsequent slides use `push dir:right` (carried forward).

---

# Push Right — Inherited

This slide inherits `push dir:right` from the previous local directive.

---

<!-- _transition: wipe dur:500 -->
# Wipe (Spot, Fast)

This slide uses `_transition: wipe dur:500` — a spot directive that applies only here.

---

# Push Right — Restored

The spot directive does not carry forward. This slide is back to `push dir:right`.

---

<!-- transition: cover dir:down -->
# Cover Down

Local directive: `transition: cover dir:down`. Carries forward from here.

---

# Cover Down — Inherited

Inherits `cover dir:down` from the previous slide.

---

<!-- transition: cut -->
# Instant Cut

`transition: cut` — an instant cut with no animation.

---

<!-- transition: pull dir:up -->
# Pull Up

`transition: pull dir:up`.

---

<!-- transition: random-bar -->
# Random Bar (Horizontal)

`transition: random-bar` — default horizontal bar sweep.

---

<!-- _transition: morph -->
# Morph (Spot, Fade Fallback)

`_transition: morph` — morph requires Office 2019 and an `mc:AlternateContent` wrapper.
A compatible `fade` is emitted as a fallback for all other PowerPoint versions.

---

<!-- transition: push dir:left dur:300 -->
# Push Left, Fast (300 ms)

`transition: push dir:left dur:300` — push left with a 300 ms fast speed.
