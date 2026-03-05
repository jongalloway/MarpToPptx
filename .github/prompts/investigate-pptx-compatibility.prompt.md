---
description: "Investigate a PPTX that fails validation, opens with repair, or fails to open in PowerPoint"
name: "Investigate PPTX Compatibility"
argument-hint: "PPTX path, sample deck, failing test, or issue summary"
agent: "agent"
---
Investigate a PPTX compatibility problem for this repository.

Use this prompt when a generated `.pptx`:
- fails `OpenXmlValidator`
- opens with a PowerPoint repair prompt
- passes validation but still fails to open in PowerPoint
- differs from a known-good reference package in suspicious ways

Repository rules:
- PowerPoint compatibility is the real success criterion.
- Passing `OpenXmlValidator` is necessary but not sufficient.
- Prefer package-structure and relationship analysis before changing slide content logic.
- Use `dotnet run --project src/MarpToPptx.Cli -- ...` for local renderer debugging, not `dnx`, unless the published tool path is the thing being tested.

Required references:
- `doc/pptx-compatibility-notes.md`
- `scripts/README.md`
- `doc/openxml-3.4.1-audit.md` when the question involves SDK capabilities or package-shaping assumptions

Expected workflow:
1. Identify the failing artifact, scenario, or repro path.
2. Find the relevant renderer, validator, test, sample, or script entry points.
3. If needed, recommend or run the repo-standard validation flow in this order:
   - targeted tests in `tests/MarpToPptx.Tests`
   - local CLI generation
   - `.NET` Open XML validation through `src/MarpToPptx.OpenXmlValidator`
   - PowerPoint smoke flow on Windows when compatibility risk remains
4. Inspect likely package-shape causes before proposing renderer changes:
   - required parts
   - slide, layout, master, and theme relationships
   - `[Content_Types].xml`
   - relative vs absolute targets
   - part inventory differences against a known-good package
5. If useful, compare against a PowerPoint-authored or other known-good reference package.
6. Summarize the most likely root cause, the evidence, and the smallest safe next change.

Output requirements:
- Start with the most likely findings, ordered by severity or confidence.
- Cite the concrete files, scripts, tests, or package parts that matter.
- Distinguish clearly between:
  - validator/schema failures
  - package-shape or relationship failures
  - unverified hypotheses
- If you recommend a code change, prefer the smallest compatibility-focused change.
- If you recommend more investigation, state the exact next command, script, or comparison to run.

Do not give generic Open XML advice when the repo’s own compatibility notes already define the expected package invariants.
