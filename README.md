# Office MCP Server

通过 MCP (Model Context Protocol) 调用本地 Microsoft Office (Word / Excel / PowerPoint) 进行文档创建、编辑和导出。

## 核心特性

- **本地 Office COM 自动化** — 直接调用用户安装的 Microsoft Office，非模拟生成
- **任务型工具设计** — `apply_operations` 批量结构化操作，适合 AI 调用
- **安全路径控制** — 可配置允许操作的目录，禁止访问系统路径
- **进程安全** — 优先管理本 MCP 启动的 Office 实例，避免误杀用户文档

## 环境要求

- Windows 操作系统
- Python 3.10+
- Microsoft Office (Word, Excel, PowerPoint)

## 安装

```bash
pip install office-mcp
```

## MCP 配置

在 Claude Desktop、Cursor 或其他 MCP 客户端中添加：

```json
{
  "mcpServers": {
    "office": {
      "command": "python",
      "args": ["-m", "office_mcp.server"],
      "env": {
        "OFFICE_MCP_ALLOWED_DIRS": "C:\\Users\\YourName\\Documents;C:\\Users\\YourName\\Desktop"
      }
    }
  }
}
```

### 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| `OFFICE_MCP_ALLOWED_DIRS` | 允许操作的目录，分号分隔 | `%USERPROFILE%` |
| `OFFICE_MCP_DEFAULT_OVERWRITE` | 默认是否允许覆盖 | `false` |
| `OFFICE_MCP_BACKUP_BEFORE_EDIT` | 编辑前是否自动备份 | `true` |
| `OFFICE_MCP_VISIBLE` | Office 是否可见（调试） | `false` |

## 可用工具

### 通用工具

- `office_status` — 查询 Office 应用状态
- `office_cleanup` — 清理 Office 实例

### Word

- `word_create_document` — 创建文档
- `word_open_document` — 打开文档
- `word_apply_operations` — 批量操作（段落、表格、替换、图片、分页等）
- `word_export_pdf` — 导出 PDF
- `word_close_document` — 关闭文档

### Excel

- `excel_create_workbook` — 创建工作簿
- `excel_open_workbook` — 打开工作簿
- `excel_apply_operations` — 批量操作（单元格、范围、公式、图表、格式等）
- `excel_export_pdf` — 导出 PDF
- `excel_close_workbook` — 关闭工作簿

### PowerPoint

- `ppt_create_presentation` — 创建演示文稿
- `ppt_open_presentation` — 打开演示文稿
- `ppt_apply_operations` — 批量操作（幻灯片、文本、图片、表格、布局等）
- `ppt_export_pdf` — 导出 PDF
- `ppt_close_presentation` — 关闭演示文稿

## 使用示例

### 创建 Word 文档

```json
{
  "file_path": "C:\\Users\\Me\\Desktop\\report.docx",
  "operations": [
    {"type": "add_paragraph", "text": "季度报告", "style": "Heading 1", "alignment": "center"},
    {"type": "add_paragraph", "text": "本季度业绩稳步增长...", "style": "Normal"},
    {"type": "insert_table", "rows": 3, "columns": 2, "data": [["项目", "数值"], ["收入", "100万"], ["利润", "20万"]]},
    {"type": "save"}
  ]
}
```

### 创建 Excel 工作簿

```json
{
  "file_path": "C:\\Users\\Me\\Desktop\\data.xlsx",
  "operations": [
    {"type": "write_range", "sheet": "Sheet1", "start_cell": "A1", "data": [["姓名", "年龄"], ["张三", 25], ["李四", 30]]},
    {"type": "create_chart", "sheet": "Sheet1", "chart_type": "column", "data_range": "A1:B3", "title": "年龄分布"},
    {"type": "save"}
  ]
}
```

### 创建 PowerPoint 演示文稿

```json
{
  "file_path": "C:\\Users\\Me\\Desktop\\slides.pptx",
  "operations": [
    {"type": "add_slide", "layout": "title_content"},
    {"type": "set_title", "slide_index": 1, "text": "项目汇报"},
    {"type": "add_text", "slide_index": 1, "text": "项目进展顺利", "left": 100, "top": 200, "width": 400, "height": 100},
    {"type": "save"}
  ]
}
```

## 许可证

MIT
