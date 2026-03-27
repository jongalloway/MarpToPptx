# MarpToPptx MCP Server (Experimental)

> **⚠️ Experimental:** The MCP server's tool surface, parameters, and behavior may change between releases without notice.

The MarpToPptx CLI includes a built-in [Model Context Protocol](https://modelcontextprotocol.io/) server that gives AI assistants direct access to the Markdown-to-PowerPoint pipeline — no shell-outs needed.

## Quick Start

Requires [.NET 10 SDK](https://dotnet.microsoft.com/download).

```bash
marp2pptx --mcp
```

Or with `dnx` (no prior install):

```bash
dnx MarpToPptx --mcp --yes
```

## Client Configuration

**VS Code / VS Code Insiders:**

Add to your MCP server configuration (Command Palette → "MCP: Add Server"):

```json
{
  "marp2pptx": {
    "type": "stdio",
    "command": "dnx",
    "args": ["MarpToPptx", "--mcp", "--yes"]
  }
}
```

**Claude Desktop:**

Edit `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "marp2pptx": {
      "command": "dnx",
      "args": ["MarpToPptx", "--mcp", "--yes"]
    }
  }
}
```

**Local development (from source):**

```bash
dotnet run --project src/MarpToPptx.Cli -- --mcp
```

## Tools

| Tool | Description |
| --- | --- |
| `marp_render` | Compile a Marp Markdown file to an editable `.pptx`. Supports templates, CSS themes, remote assets, and atomic slide-ID writing. |
| `marp_render_string` | Compile Marp Markdown provided as a string (no file needed) to `.pptx`. |
| `marp_update_deck` | Re-entrant update: reconcile new Markdown content against an existing deck, preserving manually added slides. |
| `marp_inspect` | Read-only inspection: returns slide count, titles, slide IDs, theme, and per-slide metadata as JSON. |
| `marp_write_slide_ids` | Write stable `<!-- slideId: ... -->` directives into a Markdown file for identity-based updates. |

## Iterative Workflow

The MCP server is designed for content creators who author decks in Markdown and use an AI assistant to manage rendering:

1. **First render** — `marp_render(path, writeSlideIds: true)` compiles to `.pptx` and stamps slide IDs in one pass.
2. **Inspect** — `marp_inspect(path)` returns structured deck metadata for the assistant to reason about.
3. **Edit and re-render** — edit the Markdown, then call `marp_render` again.
4. **Update after PowerPoint edits** — if slides were added or rearranged in PowerPoint, `marp_update_deck` reconciles by slide identity.

## Relationship to the CLI

The `--mcp` flag starts the same binary as a stdio-based MCP server instead of running the normal CLI. Both paths call the same `MarpToPptx.Core` and `MarpToPptx.Pptx` libraries. The CLI remains the primary interface for scripted and CI workflows; the MCP server adds a structured tool layer for AI assistants.

## Limitations

- **Experimental** — the tool surface may change in future releases.
- Contrast auditing (`--contrast-warnings`, `--contrast-report`) is not yet exposed as an MCP tool.
- No MCP Resources or Prompts are provided yet.

## Links

- [MarpToPptx on GitHub](https://github.com/jongalloway/MarpToPptx)
- [MarpToPptx CLI on NuGet](https://www.nuget.org/packages/MarpToPptx/)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [MCP C# SDK](https://github.com/modelcontextprotocol/csharp-sdk)
