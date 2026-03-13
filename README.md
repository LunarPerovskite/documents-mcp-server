# Documents MCP Server

<!-- mcp-name: io.github.LunarPerovskite/document-analyzer -->

An MCP (Model Context Protocol) server that lets AI assistants in VS Code read and visually analyze local documents — **no API keys required**.

Works with **GitHub Copilot**, Claude, and any MCP-compatible host. The AI model itself analyzes charts and images (returned as inline base64), so there are no external LLM calls.

## Supported Formats

| Format | Read text/data | Visual analysis |
|--------|:-:|:-:|
| PDF | ✅ | ✅ (pages rendered as images) |
| Excel (.xlsx/.xls) | ✅ | — |
| CSV / TSV | ✅ | — |
| Word (.docx) | ✅ | — |
| PowerPoint (.pptx) | ✅ | — |
| JSON | ✅ | — |
| TXT / Markdown | ✅ | — |
| Images (.png/.jpg/…) | OCR (optional) | ✅ |

## Installation

### Option A — pip install (recommended)

```bash
pip install documents-mcp-server
```

Or with all optional formats:

```bash
pip install "documents-mcp-server[all]"
```

### Option B — from source

```bash
git clone https://github.com/LunarPerovskite/documents-mcp-server.git
cd documents-mcp-server
pip install -e ".[all]"
```

## VS Code Setup

Add this to your **User** or **Workspace** MCP config:

**Settings → search "mcp" → Edit in mcp.json**, or directly in `~/.vscode/mcp.json` / `.vscode/mcp.json`:

```jsonc
{
  "servers": {
    "document-analyzer": {
      "command": "documents-mcp-server",
      "type": "stdio"
    }
  }
}
```

> **Tip:** Set the `DOCUMENTS_ROOT` env var to restrict file access to a specific folder:
>
> ```jsonc
> {
>   "servers": {
>     "document-analyzer": {
>       "command": "documents-mcp-server",
>       "env": {
>         "DOCUMENTS_ROOT": "C:\\Users\\me\\Projects\\data"
>       },
>       "type": "stdio"
>     }
>   }
> }
> ```

## Tools Provided

| Tool | Description |
|------|-------------|
| `list_documents` | List files under the configured root directory |
| `document_info` | Return metadata (size, sheets, etc.) for a file |
| `read_document` | Read text/data from any supported format |
| `visual_evaluate_document` | Return document pages as inline images for the AI to analyze (charts, diagrams, scanned pages) |

### Example prompts in Copilot Chat

- *"List all PDFs in my project"*
- *"Read the first 10 rows of data.xlsx sheet 'Summary'"*
- *"Visually analyze page 3 of report.pdf — extract the chart data"*

## Configuration

| Environment Variable | Default | Description |
|---|---|---|
| `DOCUMENTS_ROOT` | Server script directory | Root folder for file access |

## License

MIT
