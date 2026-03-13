<!-- mcp-name: io.github.LunarPerovskite/document-analyzer -->

# Document Analyzer MCP Server

An MCP (Model Context Protocol) server that lets AI assistants read and visually analyze local documents — PDFs, Excel spreadsheets, CSV files, Word documents, PowerPoint presentations, and images.

No API keys required. The host AI (GitHub Copilot, Claude, etc.) does all the reasoning directly.

## Supported Formats

| Format | Extensions | Read | Visual |
|--------|-----------|:----:|:------:|
| PDF | `.pdf` | ✅ | ✅ |
| Excel | `.xlsx`, `.xls` | ✅ | ✅ |
| CSV / TSV | `.csv`, `.tsv` | ✅ | — |
| JSON | `.json` | ✅ | — |
| Word | `.docx` | ✅ | — |
| PowerPoint | `.pptx` | ✅ | — |
| Plain text | `.txt`, `.md` | ✅ | — |
| Images | `.png`, `.jpg`, `.jpeg`, `.gif`, `.bmp`, `.tiff`, `.webp` | — | ✅ |

## Tools

| Tool | Description |
|------|-------------|
| `list_documents` | List files under a directory, filtered by glob pattern |
| `document_info` | Get metadata (size, modified date, sheets) for a file |
| `read_document` | Extract text content from a document with pagination |
| `visual_evaluate_document` | Return page images inline so the AI can analyze charts, tables, and diagrams |

## Installation

### From VS Code (recommended)

Search for **document-analyzer** in the MCP server gallery (Extensions sidebar → MCP tab) and click Install.

### From PyPI

```bash
pip install documents-mcp-server
```

### Manual setup

Add to your VS Code `mcp.json` (or `settings.json`):

```jsonc
{
  "servers": {
    "document-analyzer": {
      "type": "stdio",
      "command": "python",
      "args": ["-m", "documents_mcp_server"],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

Or, if you installed via pip and want to use the entry point:

```jsonc
{
  "servers": {
    "document-analyzer": {
      "type": "stdio",
      "command": "documents-mcp-server"
    }
  }
}
```

## Optional Dependencies

The base install handles PDF, Excel, CSV, JSON, and plain text. For additional formats:

```bash
# Word documents
pip install documents-mcp-server[docx]

# PowerPoint
pip install documents-mcp-server[pptx]

# OCR (requires Tesseract installed on your system)
pip install documents-mcp-server[ocr]

# Everything
pip install documents-mcp-server[all]
```

## Configuration

The server reads documents from a configurable root directory. Set the `DOCUMENTS_ROOT` environment variable to change it:

```jsonc
{
  "servers": {
    "document-analyzer": {
      "type": "stdio",
      "command": "documents-mcp-server",
      "env": {
        "DOCUMENTS_ROOT": "/path/to/your/documents"
      }
    }
  }
}
```

If not set, it defaults to the directory containing the server script.

## License

MIT
