"""Word COM 操作实现."""

import logging
from pathlib import Path
from typing import Any

from office_mcp.core.errors import COMOperationError

logger = logging.getLogger(__name__)

# 对齐方式映射
ALIGNMENT_MAP = {
    "left": 0,      # wdAlignParagraphLeft
    "center": 1,    # wdAlignParagraphCenter
    "right": 2,     # wdAlignParagraphRight
    "justify": 3,   # wdAlignParagraphJustify
}

# 页面方向映射
ORIENTATION_MAP = {
    "portrait": 0,   # wdOrientPortrait
    "landscape": 1,  # wdOrientLandscape
}

# Word 常量
WD_HEADER_FOOTER_INDEX = {
    "header_primary": 0,      # wdHeaderFooterPrimary
    "header_first": 1,        # wdHeaderFooterFirst
    "header_even": 2,         # wdHeaderFooterEven
    "footer_primary": 3,       # (0 + 1)
    "footer_first": 4,        # (1 + 1)
    "footer_even": 5,         # (2 + 1)
}


def apply_word_operations(doc: Any, operations: list[dict]) -> list[dict]:
    """对 Word 文档执行批量操作.

    Args:
        doc: Word Document 对象
        operations: 操作列表

    Returns:
        每个操作的执行结果
    """
    results = []
    for op in operations:
        op_type = op.get("type", "")
        try:
            result = _execute_word_operation(doc, op)
            results.append({"type": op_type, "status": "success", "result": result})
        except Exception as e:
            logger.error(f"Word 操作失败 [{op_type}]: {e}")
            results.append({"type": op_type, "status": "error", "message": str(e)})
    return results


def _execute_word_operation(doc: Any, op: dict) -> Any:
    """执行单个 Word 操作."""
    op_type = op.get("type", "")

    if op_type == "add_paragraph":
        return _add_paragraph(doc, op)
    elif op_type == "insert_table":
        return _insert_table(doc, op)
    elif op_type == "replace_text":
        return _replace_text(doc, op)
    elif op_type == "insert_image":
        return _insert_image(doc, op)
    elif op_type == "add_page_break":
        return _add_page_break(doc, op)
    elif op_type == "set_page_orientation":
        return _set_page_orientation(doc, op)
    elif op_type == "set_font":
        return _set_font(doc, op)
    elif op_type == "set_margins":
        return _set_margins(doc, op)
    elif op_type == "apply_style":
        return _apply_style(doc, op)
    elif op_type == "create_style":
        return _create_style(doc, op)
    elif op_type == "list_styles":
        return _list_styles(doc, op)
    elif op_type == "set_header":
        return _set_header(doc, op)
    elif op_type == "set_footer":
        return _set_footer(doc, op)
    elif op_type == "add_page_number":
        return _add_page_number(doc, op)
    elif op_type == "add_date_time":
        return _add_date_time(doc, op)
    elif op_type == "save":
        doc.Save()
        return "saved"
    elif op_type == "insert_toc":
        return _insert_toc(doc, op)
    elif op_type == "update_toc":
        return _update_toc(doc, op)
    elif op_type == "add_bookmark":
        return _add_bookmark(doc, op)
    elif op_type == "goto_bookmark":
        return _goto_bookmark(doc, op)
    elif op_type == "delete_bookmark":
        return _delete_bookmark(doc, op)
    elif op_type == "insert_at_bookmark":
        return _insert_at_bookmark(doc, op)
    elif op_type == "add_hyperlink":
        return _add_hyperlink(doc, op)
    elif op_type == "mail_merge":
        return _mail_merge(doc, op)
    else:
        raise COMOperationError(f"未知的 Word 操作类型: {op_type}")


def _add_paragraph(doc: Any, op: dict) -> str:
    """添加段落."""
    text = op.get("text", "")
    style = op.get("style", "Normal")
    alignment = op.get("alignment", "left")

    paragraph = doc.Content.Paragraphs.Add()
    paragraph.Range.Text = text
    paragraph.Style = style

    if alignment in ALIGNMENT_MAP:
        paragraph.Alignment = ALIGNMENT_MAP[alignment]

    return f"added_paragraph: {text[:50]}"


def _insert_table(doc: Any, op: dict) -> str:
    """插入表格."""
    rows = op.get("rows", 2)
    columns = op.get("columns", 2)
    data = op.get("data", [])

    # 在文档末尾添加表格
    range_end = doc.Content
    range_end.Collapse(Direction=0)  # 0 = wdCollapseEnd

    table = doc.Tables.Add(range_end, rows, columns)

    # 填充数据
    for i, row_data in enumerate(data):
        if i >= rows:
            break
        for j, cell_data in enumerate(row_data):
            if j >= columns:
                break
            table.Cell(i + 1, j + 1).Range.Text = str(cell_data)

    return f"inserted_table: {rows}x{columns}"


def _replace_text(doc: Any, op: dict) -> str:
    """查找替换文本."""
    find_text = op.get("find", "")
    replace_text = op.get("replace", "")

    if not find_text:
        raise COMOperationError("replace_text", "find 参数不能为空")

    find = doc.Content.Find
    find.ClearFormatting()
    find.Replacement.ClearFormatting()

    found = find.Execute(
        FindText=find_text,
        ReplaceWith=replace_text,
        Replace=2,  # wdReplaceAll
    )

    return f"replaced: {find_text} -> {replace_text} (found={found})"


def _insert_image(doc: Any, op: dict) -> str:
    """插入图片."""
    image_path = op.get("image_path", "")
    width = op.get("width")
    height = op.get("height")

    if not image_path:
        raise COMOperationError("insert_image", "image_path 不能为空")

    range_end = doc.Content
    range_end.Collapse(Direction=0)

    shape = doc.InlineShapes.AddPicture(
        FileName=image_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Range=range_end,
    )

    if width:
        shape.Width = width
    if height:
        shape.Height = height

    return f"inserted_image: {image_path}"


def _add_page_break(doc: Any, op: dict) -> str:
    """添加分页符."""
    range_end = doc.Content
    range_end.Collapse(Direction=0)
    range_end.InsertBreak(Type=7)  # 7 = wdPageBreak
    return "added_page_break"


def _set_page_orientation(doc: Any, op: dict) -> str:
    """设置页面方向."""
    orientation = op.get("orientation", "portrait")
    if orientation in ORIENTATION_MAP:
        doc.PageSetup.Orientation = ORIENTATION_MAP[orientation]
    return f"set_orientation: {orientation}"


def _set_font(doc: Any, op: dict) -> str:
    """设置字体."""
    font_name = op.get("font_name", "")
    font_size = op.get("font_size")

    if font_name:
        doc.Content.Font.Name = font_name
    if font_size:
        doc.Content.Font.Size = font_size

    return f"set_font: {font_name} {font_size}"


def _set_margins(doc: Any, op: dict) -> str:
    """设置页边距 (单位: 磅)."""
    setup = doc.PageSetup
    if "top" in op:
        setup.TopMargin = op["top"]
    if "bottom" in op:
        setup.BottomMargin = op["bottom"]
    if "left" in op:
        setup.LeftMargin = op["left"]
    if "right" in op:
        setup.RightMargin = op["right"]
    return "set_margins"


def _add_bookmark(doc: Any, op: dict) -> str:
    """添加书签."""
    name = op.get("name", "")
    if not name:
        raise COMOperationError("add_bookmark", "name 不能为空")
    bookmark = doc.Bookmarks.Add(name, doc.Content)
    return f"added_bookmark: {name}"


def _goto_bookmark(doc: Any, op: dict) -> str:
    """跳转到书签."""
    name = op.get("name", "")
    if not name:
        raise COMOperationError("goto_bookmark", "name 不能为空")
    if name not in [b.Name for b in doc.Bookmarks]:
        raise COMOperationError(f"书签不存在: {name}")
    bookmark = doc.Bookmarks(name)
    bookmark.Select()
    return f"goto_bookmark: {name}"


def _delete_bookmark(doc: Any, op: dict) -> str:
    """删除书签."""
    name = op.get("name", "")
    if not name:
        raise COMOperationError("delete_bookmark", "name 不能为空")
    if name not in [b.Name for b in doc.Bookmarks]:
        raise COMOperationError(f"书签不存在: {name}")
    doc.Bookmarks(name).Delete()
    return f"deleted_bookmark: {name}"


def _insert_at_bookmark(doc: Any, op: dict) -> str:
    """在书签位置插入文本."""
    name = op.get("name", "")
    text = op.get("text", "")
    if not name:
        raise COMOperationError("insert_at_bookmark", "name 不能为空")
    if name not in [b.Name for b in doc.Bookmarks]:
        raise COMOperationError(f"书签不存在: {name}")
    bookmark = doc.Bookmarks(name)
    bookmark.Range.Text = text
    return f"inserted_at_bookmark: {name} = {text}"


def _insert_toc(doc: Any, op: dict) -> str:
    """插入目录."""
    heading_levels = op.get("heading_levels", 3)

    # 在文档末尾插入目录
    doc.Content.InsertParagraphAfter()

    # 添加目录域
    toc_range = doc.Content
    toc_range.Collapse(Direction=0)

    try:
        toc = doc.TablesOfContents.Add(
            Range=toc_range,
            Level=heading_levels,
            UseOutlineLevels=True,
            IncludePageNumbers=True,
            RightAlignPageNumbers=True,
            Leader=2,
        )
        return f"inserted_toc: {heading_levels} levels"
    except Exception:
        toc_range.Text = "\n目录\n"
        toc_code = ' TOC \\o "1-' + str(heading_levels) + '" \\h \\z \\u '
        field = toc_range.Fields.Add(toc_range, Type=1, Code=toc_code)
        field.Update()
        return f"inserted_toc: {heading_levels} levels"


def _update_toc(doc: Any, op: dict) -> str:
    """更新目录（刷新页码）."""
    updated = 0
    for toc in doc.TablesOfContents:
        toc.Update()
        updated += 1
    return f"updated_toc: {updated} tables"


def _apply_style(doc: Any, op: dict) -> str:
    """对范围应用样式."""
    range_spec = op.get("range", "all")
    style_name = op.get("style_name", "Normal")

    if range_spec == "all":
        doc.Content.Style = style_name
    else:
        doc.Range().Style = style_name
    return f"applied_style: {style_name}"


def _create_style(doc: Any, op: dict) -> str:
    """创建自定义样式."""
    name = op.get("name", "CustomStyle")
    font_name = op.get("font_name", "Arial")
    font_size = op.get("font_size", 12)
    bold = op.get("bold", False)
    italic = op.get("italic", False)

    try:
        existing = None
        for s in doc.Styles:
            if s.NameLocal == name:
                existing = s
                break

        if existing:
            return f"style_exists: {name}"

        style = doc.Styles.Add(name, 1)
        style.Font.Name = font_name
        style.Font.Size = font_size
        style.Font.Bold = bold
        style.Font.Italic = italic
        return f"created_style: {name}"
    except Exception as e:
        raise COMOperationError("create_style", str(e))


def _list_styles(doc: Any, op: dict) -> list:
    """列出所有可用样式."""
    styles = []
    for style in doc.Styles:
        if style.Type == 1:
            styles.append(style.NameLocal)
    return styles


def _set_header(doc: Any, op: dict) -> str:
    """设置页眉."""
    text = op.get("text", "")
    section_num = op.get("section", 1)
    section = doc.Sections(section_num)
    try:
        header = section.Headers(1)  # wdHeaderFooterPrimary
        header.Range.Text = text
        return f"set_header: {text[:50]}"
    except Exception:
        return f"set_header_skipped: {text[:50]}"


def _set_footer(doc: Any, op: dict) -> str:
    """设置页脚."""
    text = op.get("text", "")
    section_num = op.get("section", 1)
    section = doc.Sections(section_num)
    try:
        footer = section.Footers(1)  # wdHeaderFooterPrimary
        footer.Range.Text = text
        return f"set_footer: {text[:50]}"
    except Exception:
        return f"set_footer_skipped: {text[:50]}"


def _add_page_number(doc: Any, op: dict) -> str:
    """添加页码."""
    location = op.get("location", "footer")
    format_name = op.get("format", "decimal")

    # 数字格式映射
    format_map = {
        "decimal": 0,            # wdPageNumberFormatArabic
        "roman_upper": 1,        # wdPageNumberFormatUppercaseRoman
        "roman_lower": 2,        # wdPageNumberFormatLowercaseRoman
        "letter_upper": 3,       # wdPageNumberFormatUppercaseLetter
        "letter_lower": 4,       # wdPageNumberFormatLowercaseLetter
    }
    format_val = format_map.get(format_name, 0)

    section = doc.Sections(1)
    if location == "header":
        header = section.Headers(1)
    else:
        header = section.Footers(1)

    try:
        header.Range.Text = ""
        header.Range.ParagraphFormat.Alignment = 2  # wdAlignParagraphRight
        header.PageNumbers.Add(
            PageNumberAlignment=2,  # wdAlignPageNumberRight
            FirstPage=True,
        )
        return f"added_page_number: {location}"
    except Exception:
        return f"add_page_number_skipped: {location}"


def _add_date_time(doc: Any, op: dict) -> str:
    """在文档末尾添加日期时间."""
    date_format = op.get("format", "%Y年%m月%d日")
    range_end = doc.Content
    range_end.Collapse(Direction=0)
    import datetime
    now = datetime.datetime.now()
    range_end.Text = now.strftime(date_format)
    return f"added_date_time: {now.strftime(date_format)}"


def _add_hyperlink(doc: Any, op: dict) -> str:
    """添加超链接."""
    text = op.get("text", "")
    url = op.get("url", "")
    screen_tip = op.get("screen_tip", "")

    if not url:
        raise COMOperationError("add_hyperlink", "url 不能为空")

    # 在文档末尾添加超链接
    range_end = doc.Content
    range_end.Collapse(Direction=0)
    range_end.Text = text

    # 添加超链接
    hyperlink = doc.Hyperlinks.Add(
        Anchor=range_end,
        Address=url,
        ScreenTip=screen_tip,
        TextToDisplay=text,
    )
    return f"added_hyperlink: {text} -> {url}"


def _mail_merge(doc: Any, op: dict) -> str:
    """执行邮件合并.

    注意: 邮件合并需要数据源文件 (Excel/CSV 等)
    """
    data_source = op.get("data_source", "")
    connection = op.get("connection", "")  # ODBC 连接字符串
    sql_statement = op.get("sql_statement", "")

    if not data_source:
        raise COMOperationError("mail_merge", "data_source 不能为空")

    # 安全校验：禁止危险 SQL 关键字
    if sql_statement:
        dangerous = {"DROP", "DELETE", "INSERT", "UPDATE", "ALTER", "CREATE", "EXEC", "TRUNCATE"}
        upper_sql = sql_statement.upper()
        if any(kw in upper_sql for kw in dangerous):
            raise COMOperationError("mail_merge", "sql_statement 仅允许 SELECT 查询")

    # 配置邮件合并
    mail_merge = doc.MailMerge
    mail_merge.OpenDataSource(
        Name=data_source,
        Connection=connection,
        SQLStatement=sql_statement,
        ConfirmConversions=False,
        ReadOnly=True,
        LinkToSource=False,
        AddToRecentFiles=False,
        Format=0,  # wdOpenFormatAuto
    )

    # 执行合并到新建文档
    mail_merge.Execute(Pause=False)

    return f"mail_merge_executed: {data_source}"
