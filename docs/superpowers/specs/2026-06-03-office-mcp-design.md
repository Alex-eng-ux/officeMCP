# Office MCP Server 设计文档

## 1. 项目概述

创建一个 MCP (Model Context Protocol) 服务器，通过 COM 自动化接口调用用户本地安装的 Microsoft Office (Word / Excel / PowerPoint)，实现 AI 对 Office 文档的创建、编辑、导出等操作。

### 核心约束
- **仅支持 Windows** — 依赖 Office COM 自动化
- **需要本地安装 Microsoft Office** — 不支持 WPS / LibreOffice
- **本地 stdio 通信** — 作为本地 MCP 服务器运行

## 2. 技术栈

- **语言**: Python 3.10+
- **MCP 框架**: `mcp` Python SDK (FastMCP)
- **COM 接口**: `pywin32` (win32com.client)
- **配置管理**: Pydantic Settings
- **通信方式**: stdio (MCP 本地服务器标准)

依赖安装：
```bash
pip install mcp pywin32 psutil pydantic-settings
```

## 3. 项目结构

```
office-mcp/
├── src/
│   └── office_mcp/
│       ├── __init__.py
│       ├── server.py              # FastMCP 服务器入口
│       ├── config.py              # 配置管理 (允许目录、安全设置)
│       ├── core/
│       │   ├── __init__.py
│       │   ├── office_manager.py  # Office COM 应用生命周期管理
│       │   ├── path_guard.py      # 路径安全校验
│       │   └── errors.py          # 自定义异常类
│       ├── tools/
│       │   ├── __init__.py
│       │   ├── word.py            # Word 工具注册
│       │   ├── excel.py           # Excel 工具注册
│       │   ├── powerpoint.py      # PowerPoint 工具注册
│       │   └── office.py          # 通用 Office 工具注册
│       └── operations/
│           ├── __init__.py
│           ├── word_ops.py        # Word COM 操作实现
│           ├── excel_ops.py       # Excel COM 操作实现
│           └── ppt_ops.py         # PowerPoint COM 操作实现
├── pyproject.toml
└── README.md
```

## 4. 工具设计

### 4.1 通用 Office 工具

| 工具名 | 功能 | 参数 |
|--------|------|------|
| `office_status` | 查询 Office 应用状态 (运行中/已打开文档) | 无 |
| `office_cleanup` | 关闭本 MCP 管理的 Office 实例；`force=true` 时终止进程 | `force: bool = false` |

### 4.2 Word 工具

| 工具名 | 功能 | 参数 |
|--------|------|------|
| `word_create_document` | 创建新 Word 文档 | `file_path: str, overwrite: bool = false` |
| `word_open_document` | 打开现有 Word 文档 | `file_path: str` |
| `word_apply_operations` | 对文档执行批量操作 | `file_path: str, operations: List[Operation]` |
| `word_export_pdf` | 导出 Word 文档为 PDF | `file_path: str, output_path: str \| None = None` |
| `word_close_document` | 关闭 Word 文档 | `file_path: str, save: bool = true` |
| `word_add_bookmark` | 添加书签 | `file_path: str, name: str` |
| `word_goto_bookmark` | 跳转到书签 | `file_path: str, name: str` |
| `word_delete_bookmark` | 删除书签 | `file_path: str, name: str` |
| `word_insert_at_bookmark` | 在书签位置插入文本 | `file_path: str, name: str, text: str` |
| `word_set_header` | 设置页眉 | `file_path: str, text: str, section: int = 1` |
| `word_set_footer` | 设置页脚 | `file_path: str, text: str, section: int = 1` |
| `word_add_page_number` | 添加页码 | `file_path: str, location: str, format: str` |
| `word_apply_style` | 应用样式 | `file_path: str, style_name: str, range_spec: str` |
| `word_create_style` | 创建自定义样式 | `file_path: str, name: str, ...` |
| `word_list_styles` | 列出所有样式 | `file_path: str` |
| `word_insert_toc` | 插入目录 | `file_path: str, heading_levels: int = 3` |
| `word_update_toc` | 更新目录 | `file_path: str` |

### 4.3 Excel 工具

| 工具名 | 功能 | 参数 |
|--------|------|------|
| `excel_create_workbook` | 创建新 Excel 工作簿 | `file_path: str, overwrite: bool = false` |
| `excel_open_workbook` | 打开现有 Excel 工作簿 | `file_path: str` |
| `excel_apply_operations` | 对工作簿执行批量操作 | `file_path: str, operations: List[Operation]` |
| `excel_export_pdf` | 导出 Excel 工作簿为 PDF | `file_path: str, output_path: str \| None = None` |
| `excel_close_workbook` | 关闭 Excel 工作簿 | `file_path: str, save: bool = true` |
| `excel_add_data_validation` | 添加数据验证 | `file_path, sheet, range, validation_type, formula1` |
| `excel_add_conditional_format` | 添加条件格式 | `file_path, sheet, range, condition_type, operator, formula1, font_color, bg_color` |
| `excel_merge_cells` | 合并单元格 | `file_path, sheet, range` |
| `excel_set_borders` | 设置边框 | `file_path, sheet, range, border_type, style, color` |
| `excel_add_named_range` | 添加命名范围 | `file_path, name, refers_to` |

### 4.4 PowerPoint 工具

| 工具名 | 功能 | 参数 |
|--------|------|------|
| `ppt_create_presentation` | 创建新 PPT 演示文稿 | `file_path: str, overwrite: bool = false` |
| `ppt_open_presentation` | 打开现有 PPT 演示文稿 | `file_path: str` |
| `ppt_apply_operations` | 对演示文稿执行批量操作 | `file_path: str, operations: List[Operation]` |
| `ppt_export_pdf` | 导出 PPT 演示文稿为 PDF | `file_path: str, output_path: str \| None = None` |
| `ppt_close_presentation` | 关闭 PPT 演示文稿 | `file_path: str, save: bool = true` |

## 5. Operation 设计 (结构化批量操作)

每个 `apply_operations` 工具接收 `operations` 参数，是一个 JSON 数组，每个元素是一个操作对象。

### 5.1 Word Operations

```json
[
  { "type": "add_paragraph", "text": "标题", "style": "Heading 1", "alignment": "center" },
  { "type": "add_paragraph", "text": "正文内容...", "style": "Normal" },
  { "type": "insert_table", "rows": 3, "columns": 2, "data": [["A1", "B1"], ["A2", "B2"], ["A3", "B3"]] },
  { "type": "replace_text", "find": "{{name}}", "replace": "张三" },
  { "type": "insert_image", "image_path": "C:\\Users\\me\\Desktop\\logo.png", "width": 200, "height": 100 },
  { "type": "add_page_break" },
  { "type": "set_page_orientation", "orientation": "portrait" },
  { "type": "set_font", "font_name": "Microsoft YaHei", "font_size": 12 },
  { "type": "set_margins", "top": 72, "bottom": 72, "left": 90, "right": 90 },
  { "type": "add_bookmark", "name": "title_bookmark" },
  { "type": "insert_at_bookmark", "name": "title_bookmark", "text": "替换文本" },
  { "type": "delete_bookmark", "name": "title_bookmark" },
  { "type": "set_header", "text": "文档标题" },
  { "type": "set_footer", "text": "第 X 页" },
  { "type": "add_page_number", "location": "footer", "format": "decimal" },
  { "type": "add_date_time", "location": "footer", "format": "yyyy-MM-dd" },
  { "type": "apply_style", "style_name": "Heading 1" },
  { "type": "create_style", "name": "CustomStyle", "font_name": "Arial", "font_size": 12, "bold": true },
  { "type": "list_styles" },
  { "type": "insert_toc", "heading_levels": 3 },
  { "type": "update_toc" },
  { "type": "save" }
]
```

### 5.2 Excel Operations

```json
[
  { "type": "write_cell", "sheet": "Sheet1", "cell": "A1", "value": "标题" },
  { "type": "write_range", "sheet": "Sheet1", "start_cell": "A1", "data": [["姓名", "年龄"], ["张三", 25], ["李四", 30]] },
  { "type": "read_range", "sheet": "Sheet1", "range": "A1:C10" },
  { "type": "add_formula", "sheet": "Sheet1", "cell": "C2", "formula": "=B2*2" },
  { "type": "format_range", "sheet": "Sheet1", "range": "A1:C1", "bold": true, "background_color": "#4472C4", "font_color": "#FFFFFF" },
  { "type": "set_number_format", "sheet": "Sheet1", "range": "B:B", "format": "0.00" },
  { "type": "create_chart", "sheet": "Sheet1", "chart_type": "column", "data_range": "A1:B5", "title": "销售图表", "left": 300, "top": 80, "width": 480, "height": 300 },
  { "type": "add_worksheet", "name": "新工作表" },
  { "type": "rename_worksheet", "old_name": "Sheet1", "new_name": "数据" },
  { "type": "auto_fit_columns", "sheet": "数据", "columns": ["A", "B", "C"] },
  { "type": "freeze_panes", "sheet": "Sheet1", "cell": "A2" },
  { "type": "add_data_validation", "sheet": "Sheet1", "range": "A1:A10", "validation_type": "list", "formula1": "是,否" },
  { "type": "add_conditional_format", "sheet": "Sheet1", "range": "B1:B10", "condition_type": "cell_value", "operator": "greater", "formula1": "60", "bg_color": "#FFCCCC" },
  { "type": "merge_cells", "sheet": "Sheet1", "range": "A1:C1" },
  { "type": "set_borders", "sheet": "Sheet1", "range": "A1:C3", "border_type": "outside", "style": "medium", "color": "#000000" },
  { "type": "add_named_range", "name": "销售数据", "refers_to": "=Sheet1!$A$1:$A$100" },
  { "type": "save" }
]
```

### 5.3 PowerPoint Operations

```json
[
  { "type": "add_slide", "layout": "title_content" },
  { "type": "set_title", "slide_index": 1, "text": "演示标题" },
  { "type": "add_text", "slide_index": 1, "text": "内容文本", "left": 100, "top": 200, "width": 400, "height": 100 },
  { "type": "insert_image", "slide_index": 1, "image_path": "C:\\Users\\me\\Desktop\\chart.png", "left": 100, "top": 300, "width": 400, "height": 300 },
  { "type": "insert_table", "slide_index": 1, "rows": 3, "columns": 2, "data": [["A", "B"], ["1", "2"], ["3", "4"]], "left": 100, "top": 400 },
  { "type": "set_slide_layout", "slide_index": 2, "layout": "blank" },
  { "type": "delete_slide", "slide_index": 3 },
  { "type": "set_background_color", "slide_index": 1, "color": "#FFFFFF" },
  { "type": "add_shape", "slide_index": 1, "shape": "rectangle", "left": 100, "top": 100, "width": 200, "height": 80 },
  { "type": "set_notes", "slide_index": 1, "text": "演讲者备注" },
  { "type": "add_animation", "slide_index": 1, "shape_index": 2, "animation_type": "fade" },
  { "type": "set_transition", "slide_index": 2, "transition_type": "fade", "duration": 1.5 },
  { "type": "add_section", "section_name": "第二章", "after_slide_index": 5 },
  { "type": "format_shape", "slide_index": 1, "shape_index": 1, "fill_color": "#4472C4", "line_color": "#000000", "line_width": 2 },
  { "type": "set_slide_number", "slide_index": 1, "show": true },
  { "type": "save" }
]
```

### 5.4 PPT Layout 映射

| 用户传入值 | COM 常量 |
|-----------|---------|
| `title` | `ppLayoutTitle` |
| `title_content` | `ppLayoutText` |
| `blank` | `ppLayoutBlank` |
| `section_header` | `ppLayoutSectionHeader` |
| `two_content` | `ppLayoutTwoObjects` |
| `comparison` | `ppLayoutComparison` |

## 6. 安全设计

### 6.1 路径安全 (Path Guard)

- 所有文件路径必须是**绝对路径**（包括 `document_path`, `output_path`, `image_path`, `template_path`）
- 禁止访问系统目录 (`C:\Windows`, `C:\Program Files` 等)
- 可通过环境变量 `OFFICE_MCP_ALLOWED_DIRS` 配置允许操作的目录列表（分号分隔）
- 未配置时默认只允许用户主目录 (`%USERPROFILE%`) 及其子目录

### 6.2 配置项

| 环境变量 | 说明 | 默认值 |
|----------|------|--------|
| `OFFICE_MCP_ALLOWED_DIRS` | 允许操作的目录列表，分号分隔 | `%USERPROFILE%` |
| `OFFICE_MCP_DEFAULT_OVERWRITE` | 默认是否允许覆盖文件 | `false` |
| `OFFICE_MCP_BACKUP_BEFORE_EDIT` | 编辑前是否自动备份 | `true` |
| `OFFICE_MCP_VISIBLE` | Office 应用是否可见（调试用） | `false` |

### 6.3 备份策略

当 `OFFICE_MCP_BACKUP_BEFORE_EDIT=true` 时，编辑已有文件前自动创建备份：
```
example.docx -> example.docx.office-mcp-backup-20260603-220000.docx
```
备份文件与原文件同目录，时间戳精确到秒。

## 7. Office COM 生命周期管理

### 7.1 设计原则

- **延迟初始化**: 第一次调用时才启动 Office 应用
- **单例管理**: 每个 Office 应用 (Word/Excel/PowerPoint) 只维护一个 COM 实例
- **文档追踪**: 维护已打开文档的映射表 (`file_path -> document_object`)
- **自动清理**: 服务器退出时自动关闭所有文档和 Office 应用

### 7.2 进程残留处理

- `office_cleanup` **默认**只关闭本 MCP 启动和追踪的 Office 实例
- `force=true` 时才尝试终止进程，并明确提示可能影响用户手动打开的 Office
- OfficeManager 记录自己启动的 Application 实例和打开的文件，优先只清这些
- 每次关闭文档时确保正确调用 COM 的 `Close()` 和 `Quit()` 方法

## 8. 错误处理

定义以下异常类型：

- `OfficeMCPError` — 基类
- `PathNotAllowedError` — 路径不在允许列表中
- `DocumentNotOpenError` — 文档未打开
- `DocumentAlreadyOpenError` — 文档已被打开
- `COMOperationError` — COM 调用失败
- `OfficeNotInstalledError` — 未检测到 Office 安装

所有工具返回的错误信息应包含：
- 错误类型
- 具体描述
- 建议的修复步骤

## 9. 第一版 MVP 范围

### 9.1 包含功能

1. 配置管理和路径安全校验
2. Office COM 应用管理器 (启动/关闭/文档追踪)
3. Word: 创建/打开/关闭/批量操作 (段落、表格、替换、图片、分页、页面设置、字体)
4. Word 高级: 书签、页眉页脚、页码、样式、目录
5. Excel: 创建/打开/关闭/批量操作 (单元格写入、范围读写、公式、格式、图表，工作表、冻结窗格)
6. PowerPoint: 创建/打开/关闭/批量操作 (幻灯片、标题、文本、图片、表格、布局、背景、形状、备注)
7. 通用工具: 状态查询、清理进程
8. 完善的错误处理和日志

### 9.2 不包含 (后续版本)

- 邮件合并
- 宏/VBA 执行
- 复杂图表类型 (仅基础图表)
- 协作/共享功能
- 跨平台支持

## 10. 安装和使用

### 10.1 安装

```bash
pip install office-mcp
```

### 10.2 MCP 配置 (Claude Desktop / Cursor)

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": ["-m", "office_mcp.server"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\Username\\Documents;C:\\Users\\Username\\Desktop"
      }
    }
  }
}
```

### 10.3 使用示例

用户可以对 AI 说：
- "帮我创建一个 Word 文档，写入季度报告大纲，保存到桌面"
- "打开桌面的 sales.xlsx，在 Sheet1 的 A1:C10 写入示例数据，并生成一个柱状图"
- "将当前 PPT 导出为 PDF"
