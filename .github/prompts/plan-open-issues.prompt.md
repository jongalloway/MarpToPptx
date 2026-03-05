---
description: "Plan implementation order for open GitHub issues in this repo"
name: "Plan Open Issues"
argument-hint: "Optional: issue query or label filter (default: all open issues)"
agent: "agent"
---
Look at the remaining open issues for the current repository and produce an implementation plan.

Requirements:
- Use the GitHub CLI (`gh issue list`) or MCP GitHub issue tools to list open issues.
- If authentication is missing or the list cannot be retrieved, ask for help rather than guessing.
- Answer these questions explicitly:
  1) Recommended order to implement issues
  2) Where to run each issue (locally or assign to Copilot in GitHub)
  3) Which issues can be done in parallel vs must be serial to avoid merge conflicts
  4) Provide issue links and S/M/L effort estimates

Process guidance:
- Pull open issues (include number, title, labels, milestone, and short body summary if available).
- Infer dependencies and likely file overlap from labels, titles, and repository structure.
- Prefer assigning docs-only or low-risk documentation changes to Copilot.
- For non-doc issues, choose local vs Copilot based on likelihood of merge conflicts and file overlap.
- Group issues into parallel tracks where file overlap is minimal.
- Flag any issues that should be sequenced first because they unblock others or reduce risk.

Output format:
- **Open issues snapshot**: table with columns `#`, `Title (link)`, `Labels`, `Milestone`, `Effort (S/M/L)`, `Notes`.
- **Recommended order**: numbered list with brief rationale.
- **Execution plan**: table with columns `Issue (link)`, `Run where`, `Parallel group`, `Conflict risk`, `Notes`.

Use concise, action-oriented language. If there are no open issues, say so and stop.