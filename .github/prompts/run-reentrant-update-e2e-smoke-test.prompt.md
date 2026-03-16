---
description: "Run the local end-to-end re-entrant update smoke test using a real template, simulated PowerPoint edits, and PowerPoint-open validation when available"
name: "Run Re-entrant Update E2E Smoke Test"
argument-hint: "Optional output directory, configuration, or whether to pass the template again on update"
agent: "agent"
---
Run the repository's local end-to-end re-entrant update smoke workflow.

Use this prompt when you want to verify the full update path end to end:
- create a small Marp deck
- generate an initial template-backed PPTX
- simulate manual PowerPoint edits in the generated deck
- update the Marp source
- regenerate with `--update-existing`
- validate the resulting package
- and, when available, confirm the final PPTX opens in PowerPoint

Primary entry point:
- `scripts/Invoke-ReentrantUpdateE2ESmokeTest.ps1`

Default command:

```powershell
pwsh ./scripts/Invoke-ReentrantUpdateE2ESmokeTest.ps1 -Configuration Release
```

Useful variants:

```powershell
pwsh ./scripts/Invoke-ReentrantUpdateE2ESmokeTest.ps1 -Configuration Release -OutputDirectory artifacts/smoke-tests/e2e-my-run
pwsh ./scripts/Invoke-ReentrantUpdateE2ESmokeTest.ps1 -Configuration Release -CiSafe
pwsh ./scripts/Invoke-ReentrantUpdateE2ESmokeTest.ps1 -Configuration Release -PassTemplateOnUpdate
```

Repository rules:
- PowerPoint compatibility is the real success criterion.
- Passing `OpenXmlValidator` is necessary but not sufficient.
- Prefer the local CLI project path via `dotnet run --project src/MarpToPptx.Cli -- ...` rather than `dnx`.
- For update-mode issues, preserve package structure and relationship invariants before changing slide content logic.

Required references:
- `scripts/README.md`
- `doc/pptx-compatibility-notes.md`
- `doc/using-templates.md`

Execution guidance:
1. Prefer running the PowerShell helper rather than reconstructing the steps manually.
2. Use the default path first:
   - initial generation uses the real-world template
   - update uses `--update-existing` against the manually edited, template-origin deck
3. Only use `-PassTemplateOnUpdate` when you explicitly want to exercise the separate-template-on-update path.
4. If PowerPoint COM automation is unavailable, use `-CiSafe` and report that PowerPoint-open validation was skipped.
5. If the script fails after Open XML validation but before PowerPoint open, treat that as a compatibility investigation, not a passing result.

Success criteria:
- the script exits successfully
- initial generation succeeds
- update generation succeeds
- Open XML validation succeeds for the final deck
- when PowerPoint is available, the final PPTX opens successfully in PowerPoint
- the final slide summary shows:
  - updated managed slide content
  - a newly inserted managed slide
  - preservation of the user-added unmanaged slide

Output requirements:
- Report the exact command you ran.
- State whether PowerPoint validation was executed or skipped.
- Summarize the observed behavior of managed-slide replacement and unmanaged-slide preservation.
- If the script fails, identify the failing step and point to the generated artifacts under the chosen output directory.

Do not summarize this as a generic smoke test. Treat it as the repo's canonical local verification flow for re-entrant deck updates.