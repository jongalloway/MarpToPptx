---
description: "Review a proposed implementation against MarpToPptx constraints, using Open XML SDK as primary and ShapeCrawler as a reference when relevant"
name: "Review OpenXml Or ShapeCrawler Approach"
argument-hint: "Feature idea, issue, design sketch, file path, or implementation plan"
agent: "agent"
---
Review a proposed implementation approach for this repository before or during coding.

Use this prompt when evaluating:
- whether a feature should be implemented directly with `DocumentFormat.OpenXml`
- whether ShapeCrawler is a good reference or helper for a specific PPTX task
- whether a proposed refactor is too broad for the compatibility risk involved
- whether newer Open XML SDK surface area actually removes existing manual package-shaping work

Repository rules:
- `DocumentFormat.OpenXml` is the primary implementation surface.
- ShapeCrawler is a reference source or selective helper, not a default abstraction layer.
- Prefer minimal, compatibility-focused changes over broad renderer rewrites.
- PowerPoint-open behavior matters more than validator-only cleanliness.

Required references:
- `doc/pptx-compatibility-notes.md`
- `doc/openxml-3.4.1-audit.md`
- `doc/prd.md` when the question involves intended product direction or optional ShapeCrawler usage
- external references only when needed:
  - https://github.com/dotnet/Open-XML-SDK
  - https://learn.microsoft.com/en-us/office/open-xml/presentation/overview
  - https://github.com/ShapeCrawler/ShapeCrawler

Evaluation criteria:
1. Does the proposal fit the current semantic model and renderer architecture?
2. Does it preserve known package and relationship invariants?
3. Is direct Open XML work clearer and safer than adding or leaning on a helper library?
4. If ShapeCrawler is involved, is it being used for a narrow productivity gain rather than hiding package-critical logic?
5. Does the proposal assume the SDK can now do something that `doc/openxml-3.4.1-audit.md` says still requires manual handling?
6. What is the smallest implementation that proves the behavior safely?
7. What tests, samples, or smoke checks would need to change?

Output requirements:
- Give a clear recommendation first: `use direct Open XML`, `use ShapeCrawler as reference only`, `use ShapeCrawler selectively`, or `do not proceed with this approach`.
- Then explain the reasoning in repo-specific terms.
- Call out concrete risks such as package-shape regressions, hidden abstraction costs, or mismatch with current models.
- If the idea is sound, propose a minimal implementation slice and the validation steps required.
- If the idea is weak, suggest the better alternative instead of stopping at criticism.

Do not answer as if this were a generic PowerPoint library comparison. Ground the recommendation in this repo’s current renderer, validation workflow, and compatibility constraints.