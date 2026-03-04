# PPTX File Structure And Requirements

## Purpose

This note documents the PPTX package structure, relationship requirements, and validation expectations that `MarpToPptx` should preserve.

Use it as a reference when:

- changing `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs`
- debugging a PPTX that opens with repair warnings or fails to open in PowerPoint
- deciding whether a package-level change is safe

The key constraint is that a generated package can pass `OpenXmlValidator` and still fail to open cleanly in PowerPoint. Schema-valid XML is required, but PowerPoint also expects a compatible package structure and relationship graph.

## Requirements

### 1. Validation requirements

The renderer output should satisfy both of these conditions:

- `OpenXmlValidator` should return zero validation errors.
- PowerPoint should open the file without repair prompts.

These checks cover different failure modes:

- `OpenXmlValidator` catches schema and ordering problems in individual XML parts.
- PowerPoint also enforces package-shape and relationship expectations that are not fully captured by validator errors.

### 2. Required package parts

A generated deck should include these canonical parts:

- `docProps/core.xml`
- `docProps/app.xml`
- `ppt/presentation.xml`
- `ppt/presProps.xml`
- `ppt/viewProps.xml`
- `ppt/tableStyles.xml`
- `ppt/theme/theme1.xml`

These are not optional compatibility extras. They are part of the expected presentation scaffold.

### 3. Required layout and master structure

A generated deck should include:

- one slide master
- two slide layouts
- layout relationship parts under `ppt/slideLayouts/_rels/`
- a master-to-theme relationship
- layout-to-master relationships

The working structure is:

- a content-oriented layout
- a blank layout actually used by generated slides

### 4. Required slide relationships

Generated slides should:

- use stable `rIdN` relationship IDs
- point to the intended blank layout
- use relative relationship targets in the final package

Each slide layout should also have its own `.rels` file pointing back to the slide master.

The minimum expected layout relationship files are:

- `ppt/slideLayouts/_rels/slideLayout1.xml.rels`
- `ppt/slideLayouts/_rels/slideLayout2.xml.rels`

### 5. Required content type rules

`[Content_Types].xml` should preserve these rules:

- `.xml` default content type is `application/xml`
- `/ppt/presentation.xml` has an explicit presentation override
- slide, layout, theme, and metadata overrides are present for emitted parts

### 6. XML-level requirements

Some XML-level details are required for validity or for closer PowerPoint compatibility:

- child ordering inside typed elements must be schema-correct
- avoid empty text runs where no visible text is intended
- omit empty `a:pPr` for plain paragraphs unless properties are actually needed
- keep master and layout placeholder trees close to standard PowerPoint structure

One concrete example is run property ordering: `a:solidFill` must not be emitted after `a:latin` under `a:rPr`.

## Troubleshooting Signals

### If `OpenXmlValidator` fails

Start with XML correctness problems such as:

- child ordering within Open XML elements
- invalid element placement
- malformed run, paragraph, or shape structures

### If `OpenXmlValidator` passes but PowerPoint still fails

Check package structure before changing slide content logic. Common package-level suspects include:

- missing `docProps` or `ppt/*Props` parts
- missing `ppt/theme/theme1.xml`
- missing layout relationship parts under `ppt/slideLayouts/_rels/`
- wrong layout wiring between slides, layouts, and master
- absolute internal relationship targets instead of relative ones
- incorrect `[Content_Types].xml` defaults or missing overrides

### If structure looks correct but output is still suspicious

Compare against a known-good reference package, ideally:

- a PowerPoint-saved version of the same deck
- a Marp CLI-generated PPTX for similar content

Use the comparison to inspect:

- part inventory
- relationship targets and IDs
- content types
- layout/master wiring

## Current Renderer Invariants

Future changes to `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs` should preserve these invariants unless there is a deliberate redesign.

### Package parts

The output should continue to include:

- `ppt/presentation.xml`
- `ppt/presProps.xml`
- `ppt/viewProps.xml`
- `ppt/tableStyles.xml`
- `ppt/theme/theme1.xml`
- `docProps/core.xml`
- `docProps/app.xml`

### Layout/master structure

The output should continue to include:

- one slide master
- two slide layouts
- layout relationship parts under `ppt/slideLayouts/_rels/`
- a master-to-theme relationship
- layout-to-master relationships

### Slide relationships

Generated slides should continue to:

- use stable `rIdN` relationship IDs
- point to the intended blank layout
- use relative relationship targets in the final package

### Content types

`[Content_Types].xml` should continue to preserve these rules:

- `.xml` default content type is `application/xml`
- `/ppt/presentation.xml` has an explicit presentation override
- slide, layout, theme, and metadata overrides are present for emitted parts

## Validation Workflow

When changing PPTX generation, use this order of operations.

For reusable local helpers that automate these checks, see `scripts/README.md`.

### 1. Run unit tests

```bash
dotnet test --project tests/MarpToPptx.Tests/MarpToPptx.Tests.csproj -c Release --no-restore
```

The renderer tests cover both semantic content and key package-structure regressions.

When package-shape behavior changes intentionally, regenerate the checked-in golden baselines with:

```bash
UPDATE_GOLDEN_PACKAGES=1 dotnet test --project tests/MarpToPptx.Tests/MarpToPptx.Tests.csproj -c Release --no-restore
```

The fixture files live under `tests/MarpToPptx.Tests/Fixtures/` and are limited to normalized package inventory plus relationship/content-type XML so they stay focused on stable compatibility invariants.

### 2. Generate a sample with the local project

Use the local CLI project when validating renderer changes:

```bash
dotnet run --project src/MarpToPptx.Cli -- samples/01-minimal.md -o artifacts/samples/01-minimal-test.pptx
```

Do not use `dnx MarpToPptx` for local renderer debugging unless you explicitly want to test the published package. `dnx` may execute the NuGet-published tool instead of the current workspace code.

### 3. Check Open XML validation

If a package fails here, fix schema problems first.

If a package passes here but still fails in PowerPoint, move on to package-structure comparison.

### 4. Test with PowerPoint itself

PowerPoint is the source of truth for whether the package opens cleanly.

PowerPoint COM automation can be used to confirm whether a generated file opens successfully and can be re-saved.

If you want one end-to-end local check, `scripts/Invoke-PptxSmokeTest.ps1` combines generation, Open XML validation, and PowerPoint open/save verification. Its `-CiSafe` mode automatically skips the PowerPoint step when COM automation is unavailable.

### 5. Diff against repaired or reference packages

If PowerPoint repairs the file or refuses to open it, compare package structure before modifying rendering logic:

- unzip the generated package
- unzip the repaired package or a Marp CLI reference package
- compare relationships, content types, and part inventory before changing slide content logic

## Practical Rules

- Passing `OpenXmlValidator` does not guarantee PowerPoint acceptance.
- Package-level relationship files can be the real blocker even when slide XML looks valid.
- Diffing the repaired package is a faster path than hand-guessing Open XML structure.
- Marp CLI is a good reference for package conventions, but not every extra part it emits is required.
- Local renderer debugging should use `dotnet run --project src/MarpToPptx.Cli -- ...`, not the published `dnx` tool.

## Relevant Code And Tests

- Renderer: `src/MarpToPptx.Pptx/Rendering/OpenXmlPptxRenderer.cs`
- Regression tests: `tests/MarpToPptx.Tests/PptxRendererTests.cs`
- Sample smoke deck: `samples/01-minimal.md`
- Generated debug artifacts: `artifacts/samples/`
