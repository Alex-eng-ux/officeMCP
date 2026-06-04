# Office MCP Server

![Python](https://img.shields.io/badge/python-%3E%3D3.10-blue)
![Platform](https://img.shields.io/badge/platform-Windows-lightgrey)
![Tools](https://img.shields.io/badge/tools-366-success)
![PPT](https://img.shields.io/badge/PPT-164-blueviolet)
![Excel](https://img.shields.io/badge/Excel-107-green)
![Word](https://img.shields.io/badge/Word-92-blue)
![License](https://img.shields.io/badge/license-MIT-green)

**366 tools across Word, Excel, and PowerPoint** — Real-time Office control through COM automation. An MCP (Model Context Protocol) server that gives AI agents full control over local Microsoft Office applications.

Unlike file-based libraries, this server interacts directly with running Office applications — changes appear instantly.

## Quick Start

```bash
# Install via pip
pip install office-mcp

# Or via uv (recommended)
uvx office-mcp
```

### Claude Desktop / Cursor / VS Code

```json
{
  "mcpServers": {
    "office": {
      "command": "uvx",
      "args": ["office-mcp"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop"
      }
    }
  }
}
```

## Core Features

- **366 tools across 3 apps** — Word (92), Excel (107), PowerPoint (164), utilities (3)
- **Real-time COM automation** — Directly manipulates running Office; changes appear instantly
- **Batch operations** — `apply_operations` for structured multi-step document creation
- **Security-first** — Path guard (whitelist + system directory block), SSRF protection, process ancestry tracking
- **Theme color awareness** — Use `accent1`, `accent2`, etc. instead of hardcoded RGB
- **Google Material Symbols** — Search 2,500+ icons by keyword and insert as SVG with theme colors

## Tool Categories (vs ppt-mcp)

| Category | Tools | ppt-mcp | Office MCP | Status |
|---|---|---|---|---|
| App | 5 | ✅ | ✅ | ✅ |
| Presentation | 8 | ✅ | ✅ | ✅ |
| Slides | 9 | ✅ | ✅ | ✅ |
| Shapes | 10 | ✅ | ✅ | ✅ |
| Text | 10 | ✅ | ✅ | ✅ |
| Placeholders | 6 | ✅ | ✅ | ✅ |
| Formatting | 3 | ✅ | ✅ | ✅ |
| Tables | 13 | ✅ | ✅ | ✅ |
| Export | 4 | ✅ | ✅ | ✅ |
| Slideshow | 6 | ✅ | ✅ | ✅ |
| Charts | 7 | ✅ | ✅ | ✅ |
| Animation | 6 | ✅ | ✅ | ✅ |
| Themes | 4 | ✅ | ✅ | ✅ |
| Groups | 3 | ✅ | ✅ | ✅ |
| Connectors | 2 | ✅ | ✅ | ✅ |
| Hyperlinks | 3 | ✅ | ✅ | ✅ |
| Sections | 3 | ✅ | ✅ | ✅ |
| Properties | 2 | ✅ | ✅ | ✅ |
| Media | 3 | ✅ | ✅ | ✅ |
| SmartArt | 3 | ✅ | ✅ | ✅ |
| Edit Operations | 6 | ✅ | ✅ | ✅ |
| Layout | 7 | ✅ | ✅ | ✅ |
| Effects | 3 | ✅ | ✅ | ✅ |
| Comments | 3 | ✅ | ✅ | ✅ |
| Advanced | 19 | ✅ | ✅ | ✅ |
| Freeform | 7 | ✅ | ✅ | ✅ |
| **Total PPT** | **164** | **155** | **✅ +9** |
| Excel | **107** | ❌ | **✅** |
| Word | **92** | ❌ | **✅** |
| **Grand Total** | **366** | **155** | **✅ 2.36x** |

## Natural Language Examples

Just describe what you want in plain language:

### "Create a 3-slide intro deck for a productivity app called Flowly."

The server creates:
1. Title slide with "Flowly" in large text
2. Features slide with key capabilities (task management, calendar sync, team collaboration)
3. Call-to-action slide with clean layout

### "Make a quarterly sales report in Excel with a chart."

```json
{
  "file_path": "sales_report.xlsx",
  "operations": [
    {"type": "write_range", "sheet": "Sheet1", "start_cell": "A1", "data": [
      ["Quarter", "Revenue", "Cost"],
      ["Q1", 100000, 60000],
      ["Q2", 120000, 72000],
      ["Q3", 150000, 85000],
      ["Q4", 180000, 95000]
    ]},
    {"type": "create_chart", "sheet": "Sheet1", "chart_type": "column", "data_range": "A1:C5", "title": "Annual Performance"},
    {"type": "add_auto_filter", "sheet": "Sheet1", "range": "A1:C5"},
    {"type": "save"}
  ]
}
```

### "Write a formal letter in Word with letterhead."

```json
{
  "file_path": "letter.docx",
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

## Design Guide

Keywords that elevate your presentations:

| Aspect | Keywords | Effect |
|---|---|---|
| **Icons** | `add icons`, `icon for each point`, `use icons throughout` | Searches Google Material Symbols and places crisp SVG icons |
| **Color scheme** | `dark navy`, `white background`, `monochrome`, `light gray` | Sets the overall color palette |
| **Accent color** | `teal accent`, `blue accent`, `brand color #2563EB` | Applies highlight color to headings, icons, shapes |
| **Style tone** | `modern minimal`, `bold and vibrant`, `clean and professional` | Signals visual personality |
| **Deck type** | `pitch deck`, `investor presentation`, `workshop slides` | Guides layout and content density |
| **Slide structure** | `Slides: title, problem, solution, features, CTA` | Defines narrative arc |
| **Layout** | `card layout`, `two-column`, `centered`, `full-bleed background` | Content arrangement |
| **Text density** | `minimal text`, `one message per slide`, `bullet points` | Text formatting |
| **Backgrounds** | `gradient background`, `solid dark background` | Background treatment |
| **Emphasis** | `highlight key numbers`, `bold headings`, `accent bar` | Visual hierarchy |

## Safety & Security

```
┌──────────────────────────────────────────────┐
│          3-Layer Security Defense            │
├──────────────────────────────────────────────┤
│ Layer 1: Path Guard (validate_path)          │
│   • Absolute path required                   │
│   • C:\Windows\ blocked                      │
│   • Only OFFICE_MCP_ALLOWED_DIRS allowed     │
├──────────────────────────────────────────────┤
│ Layer 2: SSRF Protection (URL tools)         │
│   • http/https only                          │
│   • localhost/127.0.0.1 blocked              │
│   • Private IPs resolved & rejected          │
├──────────────────────────────────────────────┤
│ Layer 3: Process Safety (cleanup)            │
│   • COM parent-child detection (6-level)      │
│   • Only kills MCP-spawned Office instances  │
│   • Never touches user-manual Office windows │
└──────────────────────────────────────────────┘
```

### Environment Variables

| Variable | Description | Default |
|---|---|---|
| `OFFICE_MCP_ALLOWED_DIRS` | Allowed directories (semicolon-separated) | `%USERPROFILE%` |
| `OFFICE_MCP_DEFAULT_OVERWRITE` | Allow overwrite by default | `false` |
| `OFFICE_MCP_BACKUP_BEFORE_EDIT` | Auto-backup before editing | `true` |
| `OFFICE_MCP_VISIBLE` | Show Office windows (debug) | `false` |

## Complete Tool Reference

### Word (92)

- `word_create_document` / `word_open_document` / `word_close_document` — Document lifecycle
- `word_apply_operations` — Batch: paragraphs, tables, find/replace, images, page breaks, styles, headers/footers, TOC, bookmarks, hyperlinks, SmartArt, fields, mail merge, icons
- `word_add_smartart` / `word_add_field` / `word_update_fields` — SmartArt and field operations
- `word_mail_merge` / `word_mail_merge_enhanced` — Mail merge with data sources
- `word_insert_icon` / `word_search_icons` — Google Material Symbols
- `word_check_typography` / `word_export_pdf` — Quality check and export

### Excel (107)

- `excel_create_workbook` / `excel_open_workbook` / `excel_close_workbook` — Workbook lifecycle
- `excel_apply_operations` — Batch: cells, ranges, formulas, charts, formatting, data validation, conditional formatting, named ranges, tables, filtering, protection, objects, view, calculation (7 categories)
- `excel_create_pivot_table` / `excel_add_slicer` / `excel_add_subtotal` — Data analysis
- `excel_define_name` / `excel_list_names` / `excel_set_array_formula` — Formula management
- `excel_import_data` / `excel_export_data` — CSV import/export
- `excel_check_typography` / `excel_export_pdf` — Quality check and export

### PowerPoint (164)

All 26 categories matched and exceeded:
- **Charts (7)**: Add, set data, format, axis, series, change type, get data
- **Slideshow (6)**: Start, stop, next, previous, goto, status
- **Hyperlinks (3)**: Add, get, remove
- **Connectors (2)**: Add, format
- **Groups (3)**: Group, ungroup, get items
- **Effects (3)**: Glow, reflection, soft edge
- **Freeform (7)**: Build, get nodes, set/insert/delete nodes, editing type, segment type
- **Comments (3)**: Add, list, delete
- Plus: SmartArt, animations, transitions, gradients, themes, media (video/audio), master slides, SVG icons, typography check, compare/merge presentations

## Testing

```bash
# Tool count verification
python count_tools.py

# COM integration tests (requires Windows + Office)
python test_excel_supplement.py
python test_word_advanced.py
python test_ppt_advanced.py
```

## Architecture

```
┌──────────────┐     ┌──────────────────────────────────────┐
│   MCP Client  │────▶│        Office MCP Server             │
│ (Claude, etc) │     │                                      │
└──────────────┘     │  tools/          operations/          │
                     │  ├── word.py     ├── word_ops.py      │
                     │  ├── excel.py    ├── excel_ops.py     │
                     │  ├── office.py   └── ppt_ops.py       │
                     │  └── powerpoint.py                     │
                     │         ↕                              │
                     │  core/                                │
                     │  ├── office_manager.py  (COM lifecycle)│
                     │  ├── path_guard.py      (security)    │
                     │  └── errors.py          (exceptions)  │
                     │         ↕                              │
                     │  ┌────────────────────────────┐       │
                     │  │  Windows COM Automation    │       │
                     │  │  (Word/Excel/PowerPoint)   │       │
                     │  └────────────────────────────┘       │
                     └──────────────────────────────────────┘
```

## Requirements

- Windows 10/11
- Python 3.10+
- Microsoft Office (2016+) — Word, Excel, PowerPoint

## License

MIT
