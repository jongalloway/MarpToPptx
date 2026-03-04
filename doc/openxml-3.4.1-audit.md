# Open XML 3.4.1 Presentation Audit

## Purpose

This note captures what `DocumentFormat.OpenXml` `3.4.1` changes matter to `MarpToPptx` presentation generation, what does not materially change the renderer today, and which follow-up work is worth tracking.

Use it when:

- evaluating whether a renderer change can lean on newer Open XML SDK surface area
- deciding whether package-shaping work can move from raw XML to typed APIs
- revisiting media, metadata, or newer PowerPoint feature support

## Baseline

The repo already references `DocumentFormat.OpenXml` `3.4.1` through central package management in `Directory.Packages.props`.

The renderer in `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs` currently uses a mixed strategy:

- strongly typed SDK objects for most slide, theme, master, and layout content
- raw `XDocument` writes for `docProps/*` and `ppt/viewProps.xml`
- post-save ZIP normalization for `[Content_Types].xml` and relationship target rewriting

That mix matters because the main question for this audit is not only "what new schema classes exist", but also "do they remove any of the manual compatibility work this repo still performs?"

## What 3.4.1 Adds That Matters

### 1. First-class MP4 media part support

The clearest presentation-facing addition in `3.4.1` is `MediaDataPartType.Mp4`.

Relevant upstream changes:

- Open XML SDK `3.4.1` changelog: added `MediaDataPartType.Mp4`
- `MediaDataPartType.Mp4` now maps to `video/mp4` and `.mp4`
- presentation parts such as `SlidePart` and `SlideLayoutPart` expose `AddVideoReferenceRelationship(...)`

Why this matters here:

- prior to `3.4.1`, MP4 embedding required custom content-type handling instead of an obvious enum-backed path
- the SDK now provides a clearer and less error-prone route for embedded video package creation
- this is the only new `3.4.1` change identified in the audit that directly suggests a near-term `MarpToPptx` feature

Follow-up:

- `#32` Support embedded MP4 video assets in PPTX output

### 2. Newer bundled Office schemas are available for advanced presentation features

`3.4.1` updates the bundled schemas to the Q3 2025 Office release. In practice, that means the SDK generator now carries newer strongly typed surface for modern Office presentation features.

Examples visible in the upstream source include:

- newer PowerPoint comment-related extensions
- `Office2019.Drawing.Model3D` and related animation/model types
- presentation/media-adjacent typed surface built around newer Office namespaces

Why this matters here:

- it lowers the barrier if the project ever decides to target richer native PowerPoint capabilities beyond basic editable slides
- it confirms the SDK is not the limiting factor for some advanced presentation constructs

Why it does not materially change the current renderer:

- the current semantic model does not represent comments, 3D content, or advanced Office-only presentation objects
- these features do not map naturally from the current Marp-flavored Markdown input model
- adopting them would require product decisions and new semantic modeling, not just swapping in a newer SDK type

## What 3.4.1 Does Not Change For MarpToPptx Today

### 1. It does not eliminate manual package-shaping work

The current compatibility-sensitive manual work in `OpenXmlPptxRenderer` is still necessary after `3.4.1`:

- `CreateViewPropertiesDocument()` still writes `ppt/viewProps.xml` manually
- `CreateCorePropertiesDocument()` and `CreateExtendedPropertiesDocument()` still write package property XML manually
- `NormalizeContentTypes(...)` still patches `[Content_Types].xml`
- `NormalizeRelationships(...)` still rewrites absolute internal relationship targets to relative ones

Audit conclusion:

- the schema refresh adds feature surface, but it does not remove the package-shape invariants documented in `doc/pptx-compatibility-notes.md`
- no evidence from the `3.4.1` release notes or SDK source suggests that PowerPoint compatibility here can now rely on validator/schema support alone
- the renderer should keep its current normalization steps unless a deliberate package-design change proves they are no longer needed

### 2. It does not simplify the existing editable-content roadmap

The other active fidelity gaps in this repo remain mostly independent of the `3.4.1` schema update:

- `#2` native PPTX tables
- `#3` expanded Marp theme CSS mapping
- `#5` syntax highlighting for code blocks
- `#6` remote assets and additional image formats
- `#27` evaluate richer native PPTX output for current fallback-rendered content

Audit conclusion:

- these items still depend more on renderer design and semantic-model choices than on missing `3.4.1` schema support
- `3.4.1` improves the available SDK surface area around newer Office constructs, but it does not materially change the implementation outlook for the current core rendering backlog

### 3. It does not fix the runtime-host validation issue exposed by the update

The repo already tracks one real operational impact from moving to `3.4.1`:

- `#21` Move PPTX smoke validation into a .NET-based validator path

Audit conclusion:

- this is a tooling/runtime-host issue, not a renderer-feature opportunity
- it should be treated as a consequence of the package update rather than as evidence that the new schema surface changes presentation output behavior

## Candidate Follow-up Changes

Prioritized follow-ups from this audit:

1. `#32` Support embedded MP4 video assets in PPTX output
2. Keep `#21` as the runtime/tooling cleanup required by the package update
3. Continue using existing renderer-fidelity issues for tables, theme coverage, code formatting, and richer native output because `3.4.1` does not materially re-scope them

## No-op Findings To Remember

These are the main no-op conclusions worth documenting so the investigation does not need to be repeated later:

- `3.4.1` is not a general simplification pass for PPTX package generation in this repo.
- The schema refresh does not replace the need for PowerPoint-compatible package wiring and post-save normalization.
- No direct evidence from the `3.4.1` release suggests that current raw XML for view properties or document properties can be removed safely just because the SDK version changed.
- Advanced Office presentation types now present in the SDK are interesting, but they are outside the current Markdown-to-editable-slide scope unless the semantic model expands.

## Bottom Line

For `MarpToPptx`, `DocumentFormat.OpenXml` `3.4.1` is best understood as:

- one concrete new opportunity: embedded MP4 video support
- one already-known operational follow-up: `#21`
- otherwise a mostly neutral schema refresh for the current renderer, not a reason to remove manual package-compatibility logic or to re-scope the existing fidelity backlog
