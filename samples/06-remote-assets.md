---
theme: default
paginate: true
header: Remote Asset Smoke Test
footer: Commit-pinned raw GitHub URLs
---

# Remote Asset Smoke Test

This deck is an opt-in integration smoke test for actual HTTP(S) image fetches.

- The URLs are commit-pinned so the content stays stable.
- Run this sample with remote assets explicitly enabled.

---

## Remote Inline SVG

![Remote architecture sketch](https://raw.githubusercontent.com/jongalloway/MarpToPptx/378c0623ad8708e51d7bd16092ccd45b48664b69/samples/assets/stack-diagram.svg)

The image above should download and embed from `raw.githubusercontent.com`.

---

<!-- backgroundImage: url(https://raw.githubusercontent.com/jongalloway/MarpToPptx/378c0623ad8708e51d7bd16092ccd45b48664b69/samples/assets/accent-wave.svg) -->
<!-- backgroundSize: contain -->
## Remote Background Image

This slide uses a commit-pinned remote SVG as a background image.

`backgroundSize: contain` should still apply when the image source is remote.
