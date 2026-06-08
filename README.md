# Office MCP Server

![Python](https://img.shields.io/badge/python-%3E%3D3.10-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

`office-mcp` is a Windows-only MCP server for controlling local Microsoft Office through COM automation. It lets AI agents create, edit, inspect, and export Word, Excel, and PowerPoint documents through the real desktop applications installed on the machine.

This project takes the interactive PowerPoint workflow popularized by tools like `ppt-mcp` and extends the same style of control to Excel and Word.

It supports two deployment shapes:

- a legacy all-in-one `office-mcp` server
- recommended split servers for `Word`, `Excel`, and `PowerPoint`

## Why this project

- Real desktop Office control instead of file-only generation
- One MCP server for Word, Excel, and PowerPoint
- Task-oriented tools plus structured `apply_operations` workflows
- Shared safety boundaries for file access and process cleanup
- Presentation-focused helpers such as templates, theme colors, icons, typography checks, and live navigation

## What it does

- **Word**: create documents, edit paragraphs and tables, insert images, manage styles, headers, footers, TOC, bookmarks, hyperlinks, fields, SmartArt, mail merge, and export to PDF
- **Excel**: create workbooks, write and read ranges, format cells, build charts, manage tables, validation, filtering, names, formulas, pivot tables, and export to PDF
- **PowerPoint**: create presentations, manage slides and layouts, edit shapes and text, apply themes, insert icons, media, tables, charts, control slideshow behavior, and export to PDF

## Quick Start

### Install from the repository

```bash
pip install .
```

Or build and install a wheel:

```bash
python -m build --wheel
pip install dist/office_mcp-*.whl
```

### Configure in an MCP client

Claude Desktop, Cursor, and VS Code can point to either the installed console script or module entry point.

For many agent clients, especially clients that inject MCP tools into the model context dynamically, the recommended setup is to register **three smaller Office MCP servers** instead of one giant combined namespace. In practice this is more robust than exposing hundreds of Word, Excel, and PowerPoint tools through a single MCP namespace.

Recommended split setup:

```json
{
  "mcpServers": {
    "office_word": {
      "command": "python",
      "args": ["-m", "office_mcp.server_word"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop;D:\\work\\client-project;E:\\test"
      }
    },
    "office_excel": {
      "command": "python",
      "args": ["-m", "office_mcp.server_excel"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop;D:\\work\\client-project;E:\\test"
      }
    },
    "office_powerpoint": {
      "command": "python",
      "args": ["-m", "office_mcp.server_powerpoint"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop;D:\\work\\client-project;E:\\test"
      }
    }
  }
}
```

Legacy combined setup:

Installed console script:

```json
{
  "mcpServers": {
    "office": {
      "command": "office-mcp",
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop;D:\\work\\client-project;E:\\test"
      }
    }
  }
}
```

If your MCP client handles large tool surfaces well, the combined server is still available. If your client shows only a small subset of tools, or appears to "lose" Word/Excel/PPT tools after startup, switch to the split configuration above.

Important: on some MCP clients, especially agent clients that lazily inject tools into a fresh conversation, a small visible subset does not necessarily mean the server actually lost tools. The Office MCP server may still be exposing the full Word, Excel, and PowerPoint surface, while the client only injects a smaller subset up front and discovers the rest later. In Codex-style clients, this often shows up as "only a few Office tools are visible at first" even though the remaining tools can still be surfaced through client-side discovery such as `tool_search`.

Python module entry point:

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": ["-m", "office_mcp.server"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop;D:\\work\\client-project;E:\\test"
      }
    }
  }
}
```

### Configure project roots, not just Documents

In practice, most users want to work directly inside their repo, shared drive, or test workspace instead of copying files into `Documents` first. Treat `OFFICE_MCP_ALLOWED_DIRS` as a list of trusted working roots and include every directory where the agent should be allowed to read or write Office files.

Recommended examples:

- `D:\\work\\client-project`
- `D:\\repos\\proposal-kit`
- `E:\\test`
- `D:\\design-assets`

Every Office file path must be absolute and must resolve under one of those roots.

If you do want the server to also trust detected workspace roots automatically, set:

```json
{
  "OFFICE_MCP_AUTO_DISCOVER_DIRS": "true"
}
```

By default, explicit `OFFICE_MCP_ALLOWED_DIRS` entries win and automatic discovery stays off.

## What we learned from `ppt-mcp`

This project borrows the strongest product ideas from a PowerPoint-first MCP workflow and generalizes them:

- Interactive editing matters because users want to see Office update in the real app, not only get a generated file at the end
- Theme-aware operations matter because semantic colors are more reusable than hardcoded RGB values
- Template-first workflows matter because existing decks and document styles are often better starting points than blank files
- Visual assets matter because icon search and SVG insertion raise the quality ceiling for generated slides
- Quality checks matter because typography inspection and layout-aware tools help AI produce cleaner results

The main difference here is breadth: the same server also covers Excel and Word with a shared safety model and a shared automation architecture.

## Agent compatibility notes

This repository is designed to work across different MCP-capable agent clients, not only Codex.

- Some clients handle large MCP namespaces well and can use the combined `office-mcp` entrypoint.
- Some clients degrade when a single MCP server exposes a very large tool surface with long descriptions and schemas.
- Some clients also lazily inject only part of that surface into a new conversation. In that case, the missing tools may still exist server-side and can often be discovered later by the client.
- If you observe missing tools in the client UI or only a tiny subset of Office tools appearing, use the split entrypoints:
  - `office_mcp.server_word`
  - `office_mcp.server_excel`
  - `office_mcp.server_powerpoint`

This split configuration is currently the safest default for general agent compatibility, best upfront tool visibility, and fewer false alarms about "lost" tools.

## Core capabilities

### PowerPoint

- Slide creation, duplication, deletion, reordering, and sections
- Shapes, text boxes, connectors, groups, z-order, alignment, distribution, rotation, crop, gradients, glow, reflection, and soft edges
- Tables, charts, media, SmartArt, hyperlinks, comments, and freeform editing
- Theme presets, theme color operations, master slide helpers, template creation and opening, and slideshow controls
- Icon search and insertion, typography checks, compare and merge helpers, repair, save, and export helpers

### Excel

- Workbook lifecycle, worksheet management, range read and write, number formats, borders, merge, freeze panes, and autofit
- Charts, named ranges, data validation, conditional formatting, tables, filtering, import and export, and formula helpers
- Pivot table and analysis-oriented helpers

### Word

- Document lifecycle, paragraphs, tables, images, replace, page setup, page breaks, styles, and margins
- Headers, footers, page numbers, date and time fields, TOC, bookmarks, hyperlinks, fields, SmartArt, and mail merge
- Icon search, insertion, and PDF export

## Usage examples

Use absolute paths that resolve inside `OFFICE_MCP_ALLOWED_DIRS`. Relative paths are rejected by the path guard.

### Create a quarterly Excel report

```json
{
  "file_path": "C:\\Users\\YourName\\Documents\\sales_report.xlsx",
  "operations": [
    {
      "type": "write_range",
      "sheet": "Sheet1",
      "start_cell": "A1",
      "data": [
        ["Quarter", "Revenue", "Cost"],
        ["Q1", 100000, 60000],
        ["Q2", 120000, 72000],
        ["Q3", 150000, 85000],
        ["Q4", 180000, 95000]
      ]
    },
    {
      "type": "create_chart",
      "sheet": "Sheet1",
      "chart_type": "column",
      "data_range": "A1:C5",
      "title": "Annual Performance"
    },
    {
      "type": "save"
    }
  ]
}
```

### Create a Word document

```json
{
  "file_path": "C:\\Users\\YourName\\Documents\\letter.docx",
  "operations": [
    {"type": "add_paragraph", "text": "Acme Corporation", "style": "Heading 1", "alignment": "center"},
    {"type": "add_paragraph", "text": "123 Business Ave, New York, NY 10001", "style": "Normal", "alignment": "center"},
    {"type": "add_paragraph", "text": ""},
    {"type": "add_paragraph", "text": "Dear Sir or Madam,", "style": "Normal"},
    {"type": "add_paragraph", "text": "We are pleased to inform you...", "style": "Normal"},
    {"type": "save"}
  ]
}
```

### Create a PowerPoint presentation

```json
{
  "file_path": "C:\\Users\\YourName\\Documents\\slides.pptx",
  "operations": [
    {"type": "add_slide", "layout": "title_content"},
    {"type": "set_title", "slide_index": 1, "text": "Project Update"},
    {"type": "add_text", "slide_index": 1, "text": "Progress is on track", "left": 100, "top": 180, "width": 420, "height": 80},
    {"type": "set_theme_preset", "preset": "corporate"},
    {"type": "save"}
  ]
}
```

## Design-oriented prompts

These phrases tend to produce better presentation results:

- `use icons for each point`
- `apply a corporate blue accent`
- `make it minimal with one message per slide`
- `use a two-column layout`
- `create from my template`
- `highlight key numbers`
- `check typography before export`

## Safety model

### Path guard

- Absolute file paths are required and must resolve inside `OFFICE_MCP_ALLOWED_DIRS`
- Configure `OFFICE_MCP_ALLOWED_DIRS` with the real project roots your agent uses, not only profile folders such as `Documents`
- If `OFFICE_MCP_ALLOWED_DIRS` is unset, Office MCP falls back to detected workspace and user-facing directories
- System directories such as `C:\Windows` and `C:\Program Files` are blocked
- If a path is rejected, check that the rejected path is absolute and that the relevant project root is present in `OFFICE_MCP_ALLOWED_DIRS`

### Edit safety

- Existing files can be auto-backed up before editing
- Output paths and referenced asset paths are validated

### Process safety

- Cleanup targets tracked Office instances first
- Force cleanup uses process ancestry checks to avoid killing unrelated user-opened Office sessions where possible

### Network safety

- URL-based helpers apply SSRF checks before downloading remote assets

## Environment variables

| Variable | Description | Default |
|---|---|---|
| `OFFICE_MCP_ALLOWED_DIRS` | Allowed directories, separated by `;` | unset by default; falls back to detected workspace and user-facing directories |
| `OFFICE_MCP_AUTO_DISCOVER_DIRS` | Merge detected workspace roots into the allowlist when explicit roots are configured | `true` when `OFFICE_MCP_ALLOWED_DIRS` is unset, otherwise `false` |
| `OFFICE_MCP_DEFAULT_OVERWRITE` | Allow overwrite by default | `false` |
| `OFFICE_MCP_BACKUP_BEFORE_EDIT` | Auto-backup before editing | `true` |
| `OFFICE_MCP_VISIBLE` | Shared visibility fallback for Office windows | `true` |
| `OFFICE_MCP_WORD_VISIBLE` | Show the Word window | falls back to `OFFICE_MCP_VISIBLE` |
| `OFFICE_MCP_EXCEL_VISIBLE` | Show the Excel window | falls back to `OFFICE_MCP_VISIBLE` |
| `OFFICE_MCP_PPT_VISIBLE` | Show the PowerPoint window | falls back to `OFFICE_MCP_VISIBLE` |

## Architecture

```text
MCP client
  -> office_mcp.server
     or office_mcp.server_word
     or office_mcp.server_excel
     or office_mcp.server_powerpoint
    -> tools/
       -> word.py
       -> excel.py
       -> powerpoint.py
       -> office.py
    -> operations/
       -> word_ops.py
       -> excel_ops.py
       -> ppt_ops.py
    -> core/
       -> office_manager.py
       -> path_guard.py
       -> errors.py
    -> utils/
       -> icons.py
  -> Microsoft Office COM automation
```

## Development Workflow

For non-trivial changes, we recommend using the implementation loop in [docs/implementation-workflow.md](D:/FakeC/MCP/offiiceMCP/docs/implementation-workflow.md).

The short version:

- scope the user-visible outcome first
- implement the smallest complete slice
- review the changed surface for real runtime and release risk
- stress-test assumptions before calling it done
- validate through tests, build, install, and real MCP calls when relevant

If you use Codex globally, the same workflow is also available as the `landable-implementation-loop` skill.

## Verification

### Focused automated checks

```bash
python -m pytest tests
python count_tools.py --min-total 300 --require-prefix excel_ --require-prefix word_ --require-prefix ppt_
python -m build --wheel
```

These checks verify release metadata, path-guard behavior, tool-count contract, and that the project can be built into a wheel.

### Manual Office smoke scripts

```bash
python test_word_advanced.py
python test_excel_advanced.py
python test_excel_supplement.py
python test_ppt_advanced.py
```

These scripts are still useful for interactive COM validation, but they are manual smoke checks rather than CI-grade regression tests. They require Windows and a local Microsoft Office installation.

### Supplemental project docs

- Implementation workflow notes: [docs/implementation-workflow.md](D:/FakeC/MCP/offiiceMCP/docs/implementation-workflow.md)
- Third-version design notes: [docs/third-version-design.md](D:/FakeC/MCP/offiiceMCP/docs/third-version-design.md)
- Handoff reports:
  - [docs/HANDOFF_REPORT_2026-06-06.md](D:/FakeC/MCP/offiiceMCP/docs/HANDOFF_REPORT_2026-06-06.md)
  - [docs/HANDOFF_REPORT_2026-06-06_v2.md](D:/FakeC/MCP/offiiceMCP/docs/HANDOFF_REPORT_2026-06-06_v2.md)

## Requirements

- Windows 10 or Windows 11
- Python 3.10+
- Microsoft Office with Word, Excel, and PowerPoint available through COM

## Repository

- Homepage: [Alex-eng-ux/officeMCP](https://github.com/Alex-eng-ux/officeMCP)
- Issues: [Issue tracker](https://github.com/Alex-eng-ux/officeMCP/issues)

## License

MIT
