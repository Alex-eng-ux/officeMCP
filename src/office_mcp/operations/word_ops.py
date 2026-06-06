"""Word COM 操作实现."""

import logging
import os
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from office_mcp.core.errors import COMOperationError
from office_mcp.core.path_guard import validate_path
from office_mcp.utils.icons import get_icon_url, search_icons

logger = logging.getLogger(__name__)


# 操作 op 中疑似路径字段的 key 集合
_PATH_FIELDS = (
    "image_path", "source_path", "target_path", "template_path",
    "file_path", "new_path", "output_path", "output_dir", "from_file",
    "to_file", "src_path", "dst_path", "data_source",
)


def _validate_op_paths(op: dict) -> None:
    """校验 op dict 中所有疑似路径字段."""
    for key, value in op.items():
        if not isinstance(value, str) or not value:
            continue
        if key.lower() in _PATH_FIELDS or key.lower().endswith(("_path", "_file", "path")):
            try:
                validate_path(value)
            except COMOperationError:
                raise
            except Exception as e:
                raise COMOperationError(op.get("type", "?"), f"路径校验失败 {key}={value}: {e}")

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

# SmartArt 布局类型映射 (基于 Office 2010+ 的 LayoutID 或 Type)
SMARTART_MAP = {
    "org_chart": 92,          # Organization Chart
    "cycle": 88,              # Basic Cycle
    "pyramid": 95,            # Pyramid List
    "process": 85,            # Basic Process
    "list": 83,               # Basic List
}

# 域类型映射
FIELD_TYPE_MAP = {
    "DATE": 31,               # wdFieldDate
    "TIME": 32,               # wdFieldTime
    "PAGE": 33,               # wdFieldPage
    "NUMPAGES": 36,           # wdFieldNumPages
    "REF": 3,                 # wdFieldRef
    "AUTHOR": 17,             # wdFieldAuthor
    "TITLE": 25,              # wdFieldTitle
    "FILENAME": 29,           # wdFieldFileName
    "HYPERLINK": 26,          # wdFieldHyperlink
}

# 分节符类型映射
SECTION_BREAK_MAP = {
    "next_page": 7,           # wdSectionBreakNextPage
    "continuous": 8,          # wdSectionBreakContinuous
    "even_page": 9,           # wdSectionBreakEvenPage
    "odd_page": 10,           # wdSectionBreakOddPage
    "column": 11,             # wdSectionBreakColumn
}

# 图片文字环绕类型映射
WRAP_TYPE_MAP = {
    "none": 0,                # wdWrapNone
    "square": 1,              # wdWrapSquare
    "tight": 2,               # wdWrapTight
    "through": 3,             # wdWrapThrough
    "top_bottom": 4,          # wdWrapTopBottom
    "behind": 5,              # wdWrapBehind
    "in_front": 6,            # wdWrapFront
    "inline": 7,              # wdWrapInline (特殊)
}

# 保护类型映射
PROTECTION_TYPE_MAP = {
    "none": 0,                # wdNoProtection
    "read_only": 1,           # wdAllowOnlyRevisions
    "comments": 2,            # wdAllowOnlyComments
    "tracked_changes": 3,     # wdAllowOnlyRevisions
    "forms": 4,               # wdAllowOnlyFormFields
    "all": 5,                 # wdAllowOnlyReading
}

MAIL_MERGE_RETRYABLE_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED
    -2147418110,  # RPC_E_SERVERCALL_RETRYLATER
    -2147417848,  # RPC_E_DISCONNECTED
    -2147417846,  # busy/system call failed variants
    -2147023170,  # RPC_S_CALL_FAILED
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
}


def _retry_word_mail_merge_call(stage: str, callable_obj, *args, **kwargs):  # noqa: ANN001, ANN002, ANN003
    """Retry Word COM calls that commonly fail while Office is busy."""
    last_error = None
    for _ in range(15):
        try:
            return callable_obj(*args, **kwargs)
        except Exception as error:  # noqa: BLE001
            last_error = error
            error_text = str(error).lower()
            hresult = getattr(error, "hresult", None)
            if hresult not in MAIL_MERGE_RETRYABLE_HRESULTS and not any(
                marker in error_text
                for marker in (
                    "call was rejected by callee",
                    "rpc server is unavailable",
                    "object invoked has disconnected",
                    "remote procedure call failed",
                    "server unavailable",
                    "application is busy",
                )
            ):
                raise COMOperationError("mail_merge", f"{stage} failed: {error}") from error
            time.sleep(1.0)
    raise COMOperationError("mail_merge", f"{stage} failed after retries: {last_error}")


def _list_excel_sheet_names(data_source: Path) -> list[str]:
    """Read visible worksheet names from an xlsx/xlsm workbook without Excel COM."""
    try:
        with zipfile.ZipFile(data_source) as archive:
            with archive.open("xl/workbook.xml") as workbook_xml:
                root = ElementTree.parse(workbook_xml).getroot()
    except (FileNotFoundError, zipfile.BadZipFile, KeyError, ElementTree.ParseError):
        return []

    namespace = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    sheets: list[str] = []
    for sheet in root.findall("main:sheets/main:sheet", namespace):
        state = (sheet.attrib.get("state") or "").lower()
        if state == "hidden":
            continue
        name = (sheet.attrib.get("name") or "").strip()
        if name:
            sheets.append(name)
    return sheets


def _build_excel_mail_merge_defaults(data_source: Path) -> dict[str, str | int]:
    """Build deterministic defaults for Excel-backed mail merge sources."""
    sheet_names = _list_excel_sheet_names(data_source)
    if not sheet_names:
        return {}

    first_sheet = sheet_names[0].replace("]", "]]")
    return {
        "SubType": 16,  # wdMergeSubTypeOther
        "Connection": (
            f"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source={data_source};"
            'Mode=Read;Extended Properties="HDR=YES;IMEX=1";'
        ),
        "SQLStatement": f"SELECT * FROM [{first_sheet}$]",
    }


def _build_mail_merge_open_kwargs(
    data_source: str,
    connection: str,
    sql_statement: str,
) -> dict[str, str | bool | int]:
    """Build OpenDataSource kwargs with stable defaults for Excel sources."""
    data_path = Path(data_source)
    kwargs: dict[str, str | bool | int] = {
        "Name": data_source,
        "ConfirmConversions": False,
        "ReadOnly": True,
        "LinkToSource": False,
        "AddToRecentFiles": False,
        "Format": 0,  # wdOpenFormatAuto
    }

    if data_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        kwargs.update(_build_excel_mail_merge_defaults(data_path))

    if connection:
        kwargs["Connection"] = connection
    if sql_statement:
        kwargs["SQLStatement"] = sql_statement

    return kwargs


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
            _validate_op_paths(op)
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
    elif op_type == "insert_icon":
        return _insert_icon_to_word(doc, op)
    elif op_type == "add_smartart":
        return _add_smartart(doc, op)
    elif op_type == "add_field":
        return _add_field(doc, op)
    elif op_type == "update_fields":
        return _update_fields(doc, op)
    elif op_type == "check_typography":
        return _check_typography(doc, op)
    # Sections 类
    elif op_type == "list_sections":
        return _list_sections(doc, op)
    elif op_type == "add_section_break":
        return _add_section_break(doc, op)
    elif op_type == "delete_section":
        return _delete_section(doc, op)
    elif op_type == "set_section_orientation":
        return _set_section_orientation(doc, op)
    elif op_type == "set_section_margins":
        return _set_section_margins(doc, op)
    elif op_type == "set_section_columns":
        return _set_section_columns(doc, op)
    # Fields 类
    elif op_type == "list_fields":
        return _list_fields(doc, op)
    elif op_type == "update_field":
        return _update_field(doc, op)
    elif op_type == "delete_field":
        return _delete_field(doc, op)
    elif op_type == "get_field_result":
        return _get_field_result(doc, op)
    # Images 类
    elif op_type == "list_images":
        return _list_images(doc, op)
    elif op_type == "get_image_info":
        return _get_image_info(doc, op)
    elif op_type == "resize_image":
        return _resize_image(doc, op)
    elif op_type == "crop_image":
        return _crop_image(doc, op)
    elif op_type == "set_image_position":
        return _set_image_position(doc, op)
    elif op_type == "set_image_wrap":
        return _set_image_wrap(doc, op)
    elif op_type == "replace_image":
        return _replace_image(doc, op)
    elif op_type == "delete_image":
        return _delete_image(doc, op)
    # Hyperlinks 类
    elif op_type == "list_hyperlinks":
        return _list_hyperlinks(doc, op)
    elif op_type == "get_hyperlink":
        return _get_hyperlink(doc, op)
    elif op_type == "remove_hyperlink":
        return _remove_hyperlink(doc, op)
    # Protection 类
    elif op_type == "set_document_protection":
        return _set_document_protection(doc, op)
    elif op_type == "unprotect_document":
        return _unprotect_document(doc, op)
    elif op_type == "set_read_only":
        return _set_read_only(doc, op)
    elif op_type == "set_password":
        return _set_password(doc, op)
    elif op_type == "accept_all_changes":
        return _accept_all_changes(doc, op)
    elif op_type == "reject_all_changes":
        return _reject_all_changes(doc, op)
    # Document 类
    elif op_type == "get_document_info":
        return _get_document_info(doc, op)
    elif op_type == "set_document_properties":
        return _set_document_properties(doc, op)
    elif op_type == "get_document_properties":
        return _get_document_properties(doc, op)
    elif op_type == "add_comment":
        return _add_comment(doc, op)
    elif op_type == "list_comments":
        return _list_comments(doc, op)
    elif op_type == "delete_comment":
        return _delete_comment(doc, op)
    elif op_type == "track_changes":
        return _track_changes(doc, op)
    elif op_type == "compare_documents":
        return _compare_documents(doc, op)
    # Paragraphs 类
    elif op_type == "list_paragraphs":
        return _list_paragraphs(doc, op)
    elif op_type == "get_paragraph":
        return _get_paragraph(doc, op)
    elif op_type == "set_paragraph_alignment":
        return _set_paragraph_alignment(doc, op)
    elif op_type == "set_paragraph_spacing":
        return _set_paragraph_spacing(doc, op)
    elif op_type == "set_line_spacing":
        return _set_line_spacing(doc, op)
    elif op_type == "set_indent":
        return _set_indent(doc, op)
    elif op_type == "add_bullet":
        return _add_bullet(doc, op)
    elif op_type == "remove_bullet":
        return _remove_bullet(doc, op)
    elif op_type == "merge_paragraphs":
        return _merge_paragraphs(doc, op)
    elif op_type == "split_paragraph":
        return _split_paragraph(doc, op)
    # Tables 类
    elif op_type == "list_tables":
        return _list_tables(doc, op)
    elif op_type == "get_table_info":
        return _get_table_info(doc, op)
    elif op_type == "set_cell":
        return _set_cell(doc, op)
    elif op_type == "get_cell":
        return _get_cell(doc, op)
    elif op_type == "add_row":
        return _add_row(doc, op)
    elif op_type == "add_column":
        return _add_column(doc, op)
    elif op_type == "delete_row":
        return _delete_row(doc, op)
    elif op_type == "delete_column":
        return _delete_column(doc, op)
    elif op_type == "merge_cells":
        return _merge_cells(doc, op)
    elif op_type == "split_cell":
        return _split_cell(doc, op)
    elif op_type == "set_table_borders":
        return _set_table_borders(doc, op)
    elif op_type == "set_table_style":
        return _set_table_style(doc, op)
    # Styles 类
    elif op_type == "apply_bold":
        return _apply_bold(doc, op)
    elif op_type == "apply_italic":
        return _apply_italic(doc, op)
    elif op_type == "apply_underline":
        return _apply_underline(doc, op)
    elif op_type == "set_font_color":
        return _set_font_color(doc, op)
    elif op_type == "set_font_size":
        return _set_font_size(doc, op)
    elif op_type == "set_highlight_color":
        return _set_highlight_color(doc, op)
    elif op_type == "set_strikethrough":
        return _set_strikethrough(doc, op)
    elif op_type == "set_subscript_superscript":
        return _set_subscript_superscript(doc, op)
    else:
        raise COMOperationError(f"未知的 Word 操作类型: {op_type}")


def _check_typography(doc: Any, op: dict) -> list[dict]:
    """检查 Word 文档排版问题.

    Args:
        doc: Word 文档对象
        op: 操作配置

    Returns:
        问题列表，每个问题包含 type, description, location
    """
    issues = []

    try:
        # 1. 检查孤立行和短行
        issues.extend(_check_widows_and_short_lines(doc))

        # 2. 检查标点符号前后空格
        issues.extend(_check_punctuation_spacing(doc))

        # 3. 检查标题层级一致性
        issues.extend(_check_heading_hierarchy(doc))

    except Exception as e:
        logger.error(f"排版检查出错: {e}")
        issues.append({
            "type": "error",
            "description": f"排版检查过程中发生错误: {str(e)}",
            "location": "整个文档"
        })

    return issues


def _check_widows_and_short_lines(doc: Any) -> list[dict]:
    """检查孤立行和短行."""
    issues = []
    try:
        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            # 检查是否是段落的最后一行只有少量字符
            if len(text) > 0:
                # 简单检查：段落长度小于10个字符可能是短行
                if len(text) < 10 and para.Previous() and para.Next():
                    issues.append({
                        "type": "short_line",
                        "description": f"可能的短行，长度: {text}",
                        "location": f"段落 {para.Range.Start}"
                    })
    except Exception as e:
        logger.warning(f"检查孤立行和短行出错: {e}")
    return issues


def _check_punctuation_spacing(doc: Any) -> list[dict]:
    """检查标点符号前后空格."""
    issues = []
    try:
        # 定义需要检查的标点符号
        punctuation_pairs = [
            (',', '后应无空格'),
            ('.', '后应无空格'),
            ('!', '后应无空格'),
            ('?', '后应无空格'),
            ('；', '后应无空格'),
            ('，', '后应无空格'),
            ('。', '后应无空格'),
            ('！', '后应无空格'),
            ('？', '后应无空格'),
        ]

        # 遍历文档内容
        content = doc.Content.Text
        for char, description in punctuation_pairs:
            # 检查标点符号后是否有不必要的空格
            import re
            pattern = char + r'\s+'
            matches = re.finditer(pattern, content)
            for match in matches:
                issues.append({
                    "type": "punctuation_spacing",
                    "description": f"{char}{description}",
                    "location": f"位置 {match.start()}"
                })
    except Exception as e:
        logger.warning(f"检查标点符号空格出错: {e}")
    return issues


def _check_heading_hierarchy(doc: Any) -> list[dict]:
    """检查标题层级一致性."""
    issues = []
    try:
        heading_levels = []
        # 收集所有标题样式
        for para in doc.Paragraphs:
            style_name = para.Style.NameLocal if hasattr(para.Style, 'NameLocal') else ''
            # 查找 "Heading 1", "Heading 2" 等样式
            if style_name.startswith('Heading '):
                try:
                    level = int(style_name.split(' ')[1])
                    heading_levels.append((level, para.Range.Start))
                except (IndexError, ValueError):
                    continue

        # 检查层级跳跃
        prev_level = 0
        for i, (level, location) in enumerate(heading_levels):
            if prev_level == 0:
                prev_level = level
                continue
            if level > prev_level + 1:
                issues.append({
                    "type": "heading_hierarchy",
                    "description": f"标题层级跳跃: 从 Heading {prev_level} 跳到 Heading {level}",
                    "location": f"位置 {location}"
                })
            prev_level = level
    except Exception as e:
        logger.warning(f"检查标题层级出错: {e}")
    return issues


def _add_paragraph(doc: Any, op: dict) -> str:
    """添加段落."""
    text = op.get("text", "")
    style = op.get("style", "")
    alignment = op.get("alignment", "left")

    paragraph = doc.Content.Paragraphs.Add()
    paragraph.Range.Text = text
    if style:
        try:
            paragraph.Style = style
        except Exception:  # noqa: BLE001
            # Style name may differ by locale (e.g. "Normal" vs "正文")
            logger.debug("Style '%s' not found, using default", style)

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
    heading_levels = int(op.get("heading_levels", 3))

    # 在文档末尾插入目录
    doc.Content.InsertParagraphAfter()

    # 添加目录域
    toc_range = doc.Content
    toc_range.Collapse(Direction=0)

    try:
        toc = doc.TablesOfContents.Add(
            Range=toc_range,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=heading_levels,
            UseFields=True,
        )
        return f"inserted_toc: {heading_levels} levels"
    except Exception:
        try:
            toc = doc.TablesOfContents.Add(
                toc_range,
                True,
                1,
                heading_levels,
                True,
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

    try:
        if range_spec == "all":
            doc.Content.Style = style_name
        else:
            doc.Range().Style = style_name
    except Exception as e:
        raise COMOperationError("apply_style", f"样式 '{style_name}' 不存在或无法应用: {e}") from e
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


def _insert_icon_to_word(doc: Any, op: dict) -> str:
    """搜索并插入 Material Icon 到 Word 文档光标位置.

    Args:
        icon_name: 图标名称
        query: 搜索关键词
        width: 宽度 (pt)
        height: 高度 (pt)
        fill_color: 填充颜色
        size: 图标尺寸 (px)
    """
    icon_name = op.get("icon_name", "")
    query = op.get("query", "")
    width = op.get("width", 24)
    height = op.get("height", 24)
    fill_color = op.get("fill_color", "#000000")
    size = op.get("size", 24)

    if not icon_name and query:
        results = search_icons(query, limit=1)
        if results:
            icon_name = results[0]["name"]

    if not icon_name:
        raise COMOperationError("insert_icon", "需要提供 icon_name 或 query")

    # 下载 PNG
    png_url = get_icon_url(icon_name, size)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            temp_path = f.name
        urllib.request.urlretrieve(png_url, temp_path)
        # 插入到文档末尾
        range_end = doc.Content
        range_end.Collapse(Direction=0)
        shape = range_end.InlineShapes.AddPicture(temp_path, False, True)

        # 调整尺寸
        if width:
            shape.Width = width
        if height:
            shape.Height = height
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass

    return f"inserted_icon: {icon_name}"


def _add_smartart(doc: Any, op: dict) -> str:
    """添加 SmartArt 图形.

    Args:
        doc: Word Document 对象
        op: 操作配置
            - type: SmartArt 类型 (org_chart, cycle, pyramid, process, list)
            - text_list: 文本列表，用于填充图形内容
    """
    smartart_type = op.get("type", "list")
    text_list = op.get("text_list", [])

    if smartart_type not in SMARTART_MAP:
        raise COMOperationError(
            "add_smartart",
            f"不支持的 SmartArt 类型: {smartart_type}, 支持: {list(SMARTART_MAP.keys())}"
        )

    try:
        # 在文档末尾添加位置
        range_end = doc.Content
        range_end.Collapse(Direction=0)

        # 添加 SmartArt 图形
        layout_id = SMARTART_MAP[smartart_type]
        # 注意: Word COM 中 SmartArt 的添加通常通过 Application.SmartArtLayouts
        # 这里我们尝试通用方法
        try:
            shape = range_end.InlineShapes.AddSmartArt(
                Layout=doc.Application.SmartArtLayouts(layout_id),
                Range=range_end
            )
        except Exception:
            # 备选方法：尝试通过 Shapes 添加
            shape = doc.Shapes.AddSmartArt(
                Layout=doc.Application.SmartArtLayouts(layout_id),
                Range=range_end
            )
            shape = shape.ConvertToInlineShape()

        # 填充文本内容 (如果有的话)
        if text_list:
            try:
                # 尝试访问 SmartArt 节点
                smartart = shape.SmartArt
                nodes = smartart.AllNodes
                for i, text in enumerate(text_list):
                    if i < nodes.Count:
                        node = nodes(i + 1)
                        node.TextFrame2.TextRange.Text = text
                    else:
                        # 添加新节点
                        new_node = nodes.Add()
                        new_node.TextFrame2.TextRange.Text = text
            except Exception:
                logger.warning("填充 SmartArt 文本内容失败")

        return f"added_smartart: {smartart_type}"
    except Exception as e:
        raise COMOperationError("add_smartart", str(e))


def _add_field(doc: Any, op: dict) -> str:
    """添加域.

    Args:
        doc: Word Document 对象
        op: 操作配置
            - field_type: 域类型 (DATE, TIME, PAGE, NUMPAGES, REF)
            - code: 域代码 (可选)
            - preserve_formatting: 是否保留格式
    """
    field_type_name = op.get("field_type", "DATE")
    code = op.get("code", "")
    preserve_formatting = op.get("preserve_formatting", True)

    if field_type_name not in FIELD_TYPE_MAP:
        raise COMOperationError(
            "add_field",
            f"不支持的域类型: {field_type_name}, 支持: {list(FIELD_TYPE_MAP.keys())}"
        )

    try:
        field_type = FIELD_TYPE_MAP[field_type_name]
        range_end = doc.Content
        range_end.Collapse(Direction=0)

        # 添加域
        field = range_end.Fields.Add(
            Range=range_end,
            Type=field_type,
            Text=code,
            PreserveFormatting=preserve_formatting
        )

        return f"added_field: {field_type_name}"
    except Exception as e:
        raise COMOperationError("add_field", str(e))


def _update_fields(doc: Any, op: dict) -> str:
    """更新文档中的所有域.

    Args:
        doc: Word Document 对象
        op: 操作配置
            - update_links: 是否更新链接 (可选)
    """
    update_links = op.get("update_links", False)
    count = 0

    try:
        # 更新文档正文中的域
        for field in doc.Fields:
            try:
                field.Update()
                count += 1
            except Exception:
                pass

        # 更新页眉页脚中的域
        for section in doc.Sections:
            for header in section.Headers:
                try:
                    for field in header.Range.Fields:
                        field.Update()
                        count += 1
                except Exception:
                    pass
            for footer in section.Footers:
                try:
                    for field in footer.Range.Fields:
                        field.Update()
                        count += 1
                except Exception:
                    pass

        return f"updated_fields: {count} fields"
    except Exception as e:
        raise COMOperationError("update_fields", str(e))


def _mail_merge(doc: Any, op: dict) -> str:
    """执行邮件合并.

    注意: 邮件合并需要数据源文件 (Excel/CSV 等)

    Args:
        doc: Word Document 对象
        op: 操作配置
            - data_source: 数据源文件路径
            - connection: ODBC 连接字符串 (可选)
            - sql_statement: SQL 查询语句 (可选)
            - output_path: 输出到新文档的路径 (可选，不提供则合并到当前)
    """
    data_source = op.get("data_source", "")
    connection = op.get("connection", "").strip()
    sql_statement = op.get("sql_statement", "").strip()
    output_path = op.get("output_path", "").strip()
    send_to_new_document = bool(op.get("send_to_new_document", False) or output_path)

    if not data_source:
        raise COMOperationError("mail_merge", "data_source 不能为空")

    # 安全校验：禁止危险 SQL 关键字
    if sql_statement:
        dangerous = {"DROP", "DELETE", "INSERT", "UPDATE", "ALTER", "CREATE", "EXEC", "TRUNCATE"}
        upper_sql = sql_statement.upper()
        if any(kw in upper_sql for kw in dangerous):
            raise COMOperationError("mail_merge", "sql_statement 仅允许 SELECT 查询")

    app = None
    previous_alerts = None
    previous_screen_updating = None
    previous_confirm_conversions = None
    merged_doc = None

    try:
        app = doc.Application
        previous_alerts = getattr(app, "DisplayAlerts", None)
        previous_screen_updating = getattr(app, "ScreenUpdating", None)
        options = getattr(app, "Options", None)
        if options is not None:
            previous_confirm_conversions = getattr(options, "ConfirmConversions", None)

        if previous_alerts is not None:
            app.DisplayAlerts = 0
        if previous_screen_updating is not None:
            app.ScreenUpdating = False
        if previous_confirm_conversions is not None:
            options.ConfirmConversions = False

        # ── Phase 1: Snapshot pre-merge state ──
        template_full_name = getattr(doc, "FullName", "")
        template_name = getattr(doc, "Name", "")
        pre_merge_doc_count = getattr(app.Documents, "Count", 0)
        pre_merge_doc_names = set()
        try:
            for i in range(1, pre_merge_doc_count + 1):
                d = app.Documents.Item(i)
                pre_merge_doc_names.add(getattr(d, "FullName", ""))
        except Exception:
            logger.warning("Phase 1: could not enumerate pre-merge documents, falling back to count-only detection")
        logger.info("Phase 1 complete: template=%s, pre_merge_doc_count=%d", template_full_name, pre_merge_doc_count)

        # ── Phase 2: Bind data source ──
        mail_merge = doc.MailMerge
        open_data_source_kwargs = _build_mail_merge_open_kwargs(data_source, connection, sql_statement)

        if send_to_new_document and hasattr(mail_merge, "Destination"):
            mail_merge.Destination = 0  # wdSendToNewDocument

        logger.info("Phase 2: binding data source: %s", open_data_source_kwargs.get("Name"))
        _retry_word_mail_merge_call("OpenDataSource", mail_merge.OpenDataSource, **open_data_source_kwargs)
        logger.info("Phase 2 complete: data source bound")

        # ── Phase 3: Execute merge ──
        logger.info("Phase 3: executing mail merge")
        try:
            _retry_word_mail_merge_call("Execute", mail_merge.Execute, Pause=False)
        except Exception as exec_err:
            logger.error("Phase 3: Execute failed: %s", exec_err)
            try:
                post_fail_count = getattr(app.Documents, "Count", 0)
                if post_fail_count > pre_merge_doc_count:
                    for i in range(1, post_fail_count + 1):
                        d = app.Documents.Item(i)
                        d_full = getattr(d, "FullName", "")
                        if d_full not in pre_merge_doc_names and d_full != template_full_name:
                            logger.info("Phase 3 error recovery: closing new document %s without saving", d_full)
                            d.Close(SaveChanges=False)
            except Exception:
                pass
            raise

        # Identify the result document
        # Strategy 1: ActiveDocument differs from template
        active_doc = _retry_word_mail_merge_call("ActiveDocument", getattr, app, "ActiveDocument", None)
        if active_doc is not None:
            active_full_name = getattr(active_doc, "FullName", "")
            if active_doc is not doc and active_full_name != template_full_name:
                merged_doc = active_doc
                logger.info("Phase 3: identified result via ActiveDocument: %s", active_full_name)

        # Strategy 2: enumerate Documents and find the new one
        if merged_doc is None:
            try:
                post_merge_count = getattr(app.Documents, "Count", 0)
                for i in range(1, post_merge_count + 1):
                    d = app.Documents.Item(i)
                    d_full = getattr(d, "FullName", "")
                    if d_full not in pre_merge_doc_names and d_full != template_full_name:
                        merged_doc = d
                        logger.info("Phase 3: identified result via enumeration: %s", d_full)
                        break
            except Exception as enum_err:
                logger.warning("Phase 3: document enumeration failed: %s", enum_err)

        # Strategy 3: last resort — only template remains, merge may have failed
        if merged_doc is None:
            logger.error("Phase 3: could not identify a result document; template=%s", template_full_name)
            raise COMOperationError("mail_merge", "mail merge did not produce a result document")

        logger.info("Phase 3 complete: result document identified")

        # ── Phase 4: Save result ──
        if output_path:
            output_file = Path(output_path)
            output_file.parent.mkdir(parents=True, exist_ok=True)
            try:
                _retry_word_mail_merge_call("SaveAs", merged_doc.SaveAs, str(output_file))
                logger.info("Phase 4: saved result to %s", output_file)
            except Exception as save_err:
                logger.error("Phase 4: SaveAs failed: %s", save_err)
                try:
                    _retry_word_mail_merge_call("CloseMergedDocument", merged_doc.Close, SaveChanges=False)
                except Exception:
                    pass
                merged_doc = None
                raise
            try:
                _retry_word_mail_merge_call("CloseMergedDocument", merged_doc.Close, SaveChanges=False)
            except Exception:
                pass
            merged_doc = None
            return f"mail_merge_executed_to_file: {output_file}"

        if send_to_new_document:
            return f"mail_merge_executed_to_new_document: {data_source}"

        return f"mail_merge_executed: {data_source}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("mail_merge", str(e)) from e
    finally:
        try:
            if app is not None and previous_alerts is not None:
                app.DisplayAlerts = previous_alerts
            if app is not None and previous_screen_updating is not None:
                app.ScreenUpdating = previous_screen_updating
            if app is not None and previous_confirm_conversions is not None:
                app.Options.ConfirmConversions = previous_confirm_conversions
        except Exception:
            pass


# ===================== Document 类工具 =====================

# Word 文档属性名映射
WD_PROPERTY_MAP = {
    "title": "Title",                    # wdPropertyTitle
    "subject": "Subject",                # wdPropertySubject
    "author": "Author",                  # wdPropertyAuthor
    "keywords": "Keywords",              # wdPropertyKeywords
    "comments": "Comments",              # wdPropertyComments
    "category": "Category",              # wdPropertyCategory
    "manager": "Manager",                # wdPropertyManager
    "company": "Company",                # wdPropertyCompany
}


def _get_document_info(doc: Any, op: dict) -> dict:
    """获取文档信息（页数、字数、作者等）.

    Args:
        doc: Word 文档对象
        op: 操作配置（无参数）
    """
    try:
        info = {
            "page_count": doc.ComputeStatistics(2),       # wdStatisticPages
            "word_count": doc.ComputeStatistics(0),       # wdStatisticWords
            "character_count": doc.ComputeStatistics(3),  # wdStatisticCharacters
            "paragraph_count": doc.ComputeStatistics(4),  # wdStatisticParagraphs
            "line_count": doc.ComputeStatistics(1),       # wdStatisticLines
            "author": doc.BuiltInDocumentProperties("Author").Value or "",
            "title": doc.BuiltInDocumentProperties("Title").Value or "",
            "subject": doc.BuiltInDocumentProperties("Subject").Value or "",
            "creation_date": str(doc.BuiltInDocumentProperties("Creation Date").Value or ""),
            "last_saved": str(doc.BuiltInDocumentProperties("Last Save Time").Value or ""),
            "revision_number": doc.BuiltInDocumentProperties("Revision Number").Value or 0,
            "is_protected": doc.ProtectionType != 0,     # wdNoProtection = 0
            "track_changes": doc.TrackRevisions,
        }
        return info
    except Exception as e:
        raise COMOperationError("get_document_info", str(e))


def _set_document_properties(doc: Any, op: dict) -> str:
    """设置文档属性.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - properties: dict, 属性名 -> 值
              支持的键: title, subject, author, keywords, comments, category, manager, company
    """
    properties = op.get("properties", {})
    if not properties:
        raise COMOperationError("set_document_properties", "properties 不能为空")

    updated = []
    for key, value in properties.items():
        if key in WD_PROPERTY_MAP:
            try:
                doc.BuiltInDocumentProperties(WD_PROPERTY_MAP[key]).Value = value
                updated.append(key)
            except Exception as e:
                logger.warning(f"设置属性 {key} 失败: {e}")
    return f"set_properties: {', '.join(updated)}"


def _get_document_properties(doc: Any, op: dict) -> dict:
    """获取文档属性.

    Args:
        doc: Word 文档对象
        op: 操作配置（无参数）
    """
    properties = {}
    for key, prop_name in WD_PROPERTY_MAP.items():
        try:
            value = doc.BuiltInDocumentProperties(prop_name).Value
            properties[key] = str(value) if value is not None else ""
        except Exception:
            properties[key] = ""
    return properties


def _add_comment(doc: Any, op: dict) -> str:
    """添加批注.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - text: 批注内容
            - range_start: 选区起始位置 (可选, 默认光标位置)
            - range_end: 选区结束位置 (可选, 默认光标位置)
    """
    text = op.get("text", "")
    if not text:
        raise COMOperationError("add_comment", "text 不能为空")

    try:
        if "range_start" in op and "range_end" in op:
            target_range = doc.Range(op["range_start"], op["range_end"])
            comment = doc.Comments.Add(Range=target_range, Text=text)
        else:
            # 在文档末尾添加
            range_end = doc.Content
            range_end.Collapse(Direction=0)  # wdCollapseEnd
            comment = doc.Comments.Add(Range=range_end, Text=text)
        return f"added_comment: {comment.Index}"
    except Exception as e:
        raise COMOperationError("add_comment", str(e))


def _list_comments(doc: Any, op: dict) -> list:
    """列出所有批注.

    Args:
        doc: Word 文档对象
        op: 操作配置（无参数）
    """
    comments = []
    try:
        for comment in doc.Comments:
            comments.append({
                "index": comment.Index,
                "author": comment.Author,
                "text": comment.Range.Text,
                "date": str(comment.Date) if comment.Date else "",
                "initial": comment.Initial,
            })
    except Exception as e:
        logger.warning(f"列出批注失败: {e}")
    return comments


def _delete_comment(doc: Any, op: dict) -> str:
    """删除批注.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 批注索引 (1-based)
    """
    index = op.get("index")
    if index is None:
        raise COMOperationError("delete_comment", "index 不能为空")

    try:
        if index < 1 or index > doc.Comments.Count:
            raise COMOperationError("delete_comment", f"批注索引超出范围: {index}")
        doc.Comments(index).Delete()
        return f"deleted_comment: {index}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("delete_comment", str(e))


def _track_changes(doc: Any, op: dict) -> str:
    """启用/禁用修订跟踪.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - enable: bool, 是否启用
    """
    enable = op.get("enable", True)
    doc.TrackRevisions = bool(enable)
    return f"track_changes: {enable}"


def _compare_documents(doc: Any, op: dict) -> str:
    """比较两个文档 (将结果写入当前文档).

    Args:
        doc: Word 文档对象 (作为比较结果的目标)
        op: 操作配置
            - compare_path: 要比较的文档路径 (绝对路径)
            - output_path: 输出文档路径 (绝对路径, 可选)
    """
    compare_path = op.get("compare_path", "")
    output_path = op.get("output_path", "")

    if not compare_path:
        raise COMOperationError("compare_documents", "compare_path 不能为空")

    try:
        app = doc.Application
        if output_path:
            # 打开修订文档以获取 Document COM 对象
            revised_doc = app.Documents.Open(compare_path)
            try:
                result = app.CompareDocuments(
                    OriginalDocument=doc,
                    RevisedDocument=revised_doc,
                    Destination=2,        # wdCompareDestinationNew
                    Granularity=1,         # wdGranularityWordLevel
                    CompareFormatting=True,
                    CompareCaseChanges=True,
                    CompareWhitespace=False,
                    CompareTables=True,
                    CompareHeaders=True,
                    CompareFootnotes=True,
                    CompareTextboxes=True,
                    CompareFields=True,
                    CompareComments=True,
                    CompareMoves=True,
                    RevisedAuthor="",
                    IgnoreAllComparisonWarnings=True,
                )
                try:
                    result.SaveAs(output_path)
                    result.Close()
                    return f"compared_and_saved: {output_path}"
                except Exception:
                    return f"compared_to_new_document (output_path 可能需要手动保存)"
            finally:
                revised_doc.Close(SaveChanges=False)
        else:
            # 使用 doc.Compare 方法，直接传入文件路径即可
            # Word COM 的 Compare 方法只接受 Name 和可选的 AuthorTarget/CompareTarget 参数
            try:
                doc.Compare(Name=compare_path)
            except Exception:
                # 某些 Word 版本不接受关键字参数，尝试位置参数
                doc.Compare(compare_path)
            return f"compared: {compare_path}"
    except Exception as e:
        raise COMOperationError("compare_documents", str(e))


# ===================== Paragraphs 类工具 =====================

# 行距规则映射
WD_LINE_SPACING_RULE = {
    "single": 0,            # wdLineSpaceSingle
    "1.5": 1,               # wdLineSpace1pt5
    "double": 2,            # wdLineSpaceDouble
    "multiple": 4,          # wdLineSpaceMultiple (需要 LineSpacingRule = 4)
    "at_least": 3,          # wdLineSpaceAtLeast (单位: 磅)
    "exactly": 4,           # wdLineSpaceExactly (单位: 磅)
}


def _get_paragraph_by_index(doc: Any, index: int) -> Any:
    """根据索引获取段落（1-based，超出范围抛错）."""
    if index < 1 or index > doc.Paragraphs.Count:
        raise COMOperationError(
            "paragraph_index",
            f"段落索引超出范围: {index}, 总数: {doc.Paragraphs.Count}",
        )
    return doc.Paragraphs(index)


def _list_paragraphs(doc: Any, op: dict) -> list:
    """列出所有段落.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - max_count: 最大返回数量（可选, 默认返回全部）
    """
    max_count = op.get("max_count")
    paragraphs = []
    total = doc.Paragraphs.Count
    limit = total if max_count is None else min(int(max_count), total)
    for i in range(1, limit + 1):
        try:
            p = doc.Paragraphs(i)
            style_name = ""
            try:
                style_name = p.Style.NameLocal if hasattr(p.Style, "NameLocal") else ""
            except Exception:
                pass
            paragraphs.append({
                "index": i,
                "text": p.Range.Text,
                "style": style_name,
                "alignment": p.Alignment,
            })
        except Exception as e:
            logger.warning(f"获取段落 {i} 失败: {e}")
    return {"total": total, "paragraphs": paragraphs}


def _get_paragraph(doc: Any, op: dict) -> dict:
    """获取段落内容.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
    """
    index = op.get("index", 1)
    p = _get_paragraph_by_index(doc, index)
    style_name = ""
    try:
        style_name = p.Style.NameLocal if hasattr(p.Style, "NameLocal") else ""
    except Exception:
        pass
    return {
        "index": index,
        "text": p.Range.Text,
        "style": style_name,
        "alignment": p.Alignment,
        "line_spacing": p.LineSpacing,
        "left_indent": p.LeftIndent,
        "right_indent": p.RightIndent,
        "first_line_indent": p.FirstLineIndent,
    }


def _set_paragraph_alignment(doc: Any, op: dict) -> str:
    """设置段落对齐.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
            - alignment: 对齐方式 (left/center/right/justify)
    """
    index = op.get("index", 1)
    alignment = op.get("alignment", "left")
    if alignment not in ALIGNMENT_MAP:
        raise COMOperationError(
            "set_paragraph_alignment",
            f"不支持的对齐方式: {alignment}, 支持: {list(ALIGNMENT_MAP.keys())}",
        )
    p = _get_paragraph_by_index(doc, index)
    p.Alignment = ALIGNMENT_MAP[alignment]
    return f"set_paragraph_alignment: {index} -> {alignment}"


def _set_paragraph_spacing(doc: Any, op: dict) -> str:
    """设置段落间距 (单位: 磅).

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
            - space_before: 段前间距 (可选)
            - space_after: 段后间距 (可选)
    """
    index = op.get("index", 1)
    p = _get_paragraph_by_index(doc, index)
    pf = p.Format
    if "space_before" in op:
        pf.SpaceBefore = op["space_before"]
    if "space_after" in op:
        pf.SpaceAfter = op["space_after"]
    return f"set_paragraph_spacing: {index}"


def _set_line_spacing(doc: Any, op: dict) -> str:
    """设置行距.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
            - rule: 行距规则 (single, 1.5, double, multiple, at_least, exactly)
            - value: 数值（multiple 时为倍数, at_least/exactly 时为磅数）
    """
    index = op.get("index", 1)
    rule = op.get("rule", "single")
    value = op.get("value", 0)

    if rule not in WD_LINE_SPACING_RULE:
        raise COMOperationError(
            "set_line_spacing",
            f"不支持的行距规则: {rule}, 支持: {list(WD_LINE_SPACING_RULE.keys())}",
        )

    p = _get_paragraph_by_index(doc, index)
    p.LineSpacingRule = WD_LINE_SPACING_RULE[rule]
    if rule in ("single", "1.5", "double"):
        # 这三种规则由 LineSpacingRule 自动控制
        pass
    else:
        p.LineSpacing = float(value)
    return f"set_line_spacing: {index} -> {rule} ({value})"


def _set_indent(doc: Any, op: dict) -> str:
    """设置缩进 (单位: 磅).

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
            - left: 左缩进 (可选)
            - right: 右缩进 (可选)
            - first_line: 首行缩进 (可选, 负数表示悬挂缩进)
    """
    index = op.get("index", 1)
    p = _get_paragraph_by_index(doc, index)
    if "left" in op:
        p.LeftIndent = op["left"]
    if "right" in op:
        p.RightIndent = op["right"]
    if "first_line" in op:
        p.FirstLineIndent = op["first_line"]
    return f"set_indent: {index}"


def _add_bullet(doc: Any, op: dict) -> str:
    """添加项目符号.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
            - character: 项目符号字符 (可选, 默认 "*")
            - font_name: 字体名称 (可选, 默认 "Symbol")
    """
    index = op.get("index", 1)
    character = op.get("character", "*")
    font_name = op.get("font_name", "Symbol")

    p = _get_paragraph_by_index(doc, index)
    p.Range.ListFormat.ApplyBulletDefault()
    # 自定义项目符号字符
    if character != "*" or font_name != "Symbol":
        try:
            p.Range.ListFormat.ListLevelNumber = 1
            list_template = p.Range.ListFormat.ListTemplate
            list_level = list_template.ListLevels(1)
            list_level.NumberFormat = character
            list_level.NumberStyle = 23  # wdListNumberStyleBullet
            list_level.Font.Name = font_name
        except Exception as e:
            logger.warning(f"自定义项目符号失败: {e}")
    return f"added_bullet: {index}"


def _remove_bullet(doc: Any, op: dict) -> str:
    """移除项目符号.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
    """
    index = op.get("index", 1)
    p = _get_paragraph_by_index(doc, index)
    try:
        p.Range.ListFormat.RemoveNumbers(NumberType=1)  # wdNumberListNum = 1
    except Exception:
        try:
            p.Range.ListFormat.RemoveNumbers()
        except Exception as e:
            raise COMOperationError("remove_bullet", str(e))
    return f"removed_bullet: {index}"


def _merge_paragraphs(doc: Any, op: dict) -> str:
    """合并段落 (将后一个段落合并到前一个段落).

    Args:
        doc: Word 文档对象
        op: 操作配置
            - start_index: 起始段落索引 (1-based)
            - end_index: 结束段落索引 (1-based, 包含)
    """
    start_index = op.get("start_index", 1)
    end_index = op.get("end_index", 1)
    if end_index < start_index:
        raise COMOperationError("merge_paragraphs", "end_index 必须 >= start_index")

    _get_paragraph_by_index(doc, start_index)
    _get_paragraph_by_index(doc, end_index)

    # 起始段落的结束位置
    start_para = doc.Paragraphs(start_index)
    end_para = doc.Paragraphs(end_index)
    merge_range = doc.Range(start_para.Range.End, end_para.Range.End)
    merge_range.Delete()
    return f"merged_paragraphs: {start_index}-{end_index}"


def _split_paragraph(doc: Any, op: dict) -> str:
    """拆分段落 (在指定位置插入段落符).

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 段落索引 (1-based)
            - position: 段内字符位置 (可选, 默认末尾)
    """
    index = op.get("index", 1)
    position = op.get("position")

    p = _get_paragraph_by_index(doc, index)
    if position is None:
        split_pos = p.Range.End - 1  # 段尾的换行符位置
    else:
        split_pos = p.Range.Start + int(position)
    doc.Range(split_pos, split_pos).InsertParagraphAfter()
    return f"split_paragraph: {index}"


# ===================== Tables 类工具 =====================

# 表格边框样式映射
WD_LINE_STYLE_MAP = {
    "none": 0,         # wdLineStyleNone
    "single": 1,       # wdLineStyleSingle
    "double": 7,       # wdLineStyleDouble
    "dotted": 4,       # wdLineStyleDot
    "dashed": 8,       # wdLineStyleDash
    "thick": 5,        # wdLineStyleThick
}


def _get_table_by_index(doc: Any, index: int) -> Any:
    """根据索引获取表格（1-based）."""
    if index < 1 or index > doc.Tables.Count:
        raise COMOperationError(
            "table_index",
            f"表格索引超出范围: {index}, 总数: {doc.Tables.Count}",
        )
    return doc.Tables(index)


def _list_tables(doc: Any, op: dict) -> list:
    """列出所有表格.

    Args:
        doc: Word 文档对象
        op: 操作配置（无参数）
    """
    tables = []
    for i in range(1, doc.Tables.Count + 1):
        try:
            t = doc.Tables(i)
            tables.append({
                "index": i,
                "rows": t.Rows.Count,
                "columns": t.Columns.Count,
                "style": t.Style.NameLocal if t.Style and hasattr(t.Style, "NameLocal") else "",
            })
        except Exception as e:
            logger.warning(f"获取表格 {i} 失败: {e}")
    return tables


def _get_table_info(doc: Any, op: dict) -> dict:
    """获取表格信息.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - index: 表格索引 (1-based)
    """
    index = op.get("index", 1)
    t = _get_table_by_index(doc, index)
    style_name = ""
    try:
        style_name = t.Style.NameLocal if t.Style and hasattr(t.Style, "NameLocal") else ""
    except Exception:
        pass
    return {
        "index": index,
        "rows": t.Rows.Count,
        "columns": t.Columns.Count,
        "style": style_name,
    }


def _set_cell(doc: Any, op: dict) -> str:
    """设置单元格内容.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - row: 行号 (1-based)
            - column: 列号 (1-based)
            - text: 单元格文本
    """
    table_index = op.get("table_index", 1)
    row = op.get("row", 1)
    column = op.get("column", 1)
    text = op.get("text", "")

    t = _get_table_by_index(doc, table_index)
    if row < 1 or row > t.Rows.Count:
        raise COMOperationError("set_cell", f"行号超出范围: {row}")
    if column < 1 or column > t.Columns.Count:
        raise COMOperationError("set_cell", f"列号超出范围: {column}")
    t.Cell(row, column).Range.Text = str(text)
    return f"set_cell: table={table_index} ({row},{column})"


def _get_cell(doc: Any, op: dict) -> str:
    """获取单元格内容.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - row: 行号 (1-based)
            - column: 列号 (1-based)
    """
    table_index = op.get("table_index", 1)
    row = op.get("row", 1)
    column = op.get("column", 1)

    t = _get_table_by_index(doc, table_index)
    if row < 1 or row > t.Rows.Count:
        raise COMOperationError("get_cell", f"行号超出范围: {row}")
    if column < 1 or column > t.Columns.Count:
        raise COMOperationError("get_cell", f"列号超出范围: {column}")
    return t.Cell(row, column).Range.Text


def _add_row(doc: Any, op: dict) -> str:
    """添加表格行.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - count: 添加行数 (可选, 默认 1)
    """
    table_index = op.get("table_index", 1)
    count = op.get("count", 1)
    t = _get_table_by_index(doc, table_index)
    for _ in range(count):
        t.Rows.Add()
    return f"added_row: table={table_index} count={count}"


def _add_column(doc: Any, op: dict) -> str:
    """添加表格列.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - count: 添加列数 (可选, 默认 1)
    """
    table_index = op.get("table_index", 1)
    count = op.get("count", 1)
    t = _get_table_by_index(doc, table_index)
    for _ in range(count):
        t.Columns.Add()
    return f"added_column: table={table_index} count={count}"


def _delete_row(doc: Any, op: dict) -> str:
    """删除表格行.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - row: 行号 (1-based)
            - count: 删除行数 (可选, 默认 1)
    """
    table_index = op.get("table_index", 1)
    row = op.get("row", 1)
    count = op.get("count", 1)
    t = _get_table_by_index(doc, table_index)
    if row < 1 or row > t.Rows.Count:
        raise COMOperationError("delete_row", f"行号超出范围: {row}")
    if row + count - 1 > t.Rows.Count:
        raise COMOperationError("delete_row", f"删除范围超出表格")
    for i in range(count):
        t.Rows(row).Delete()
    return f"deleted_row: table={table_index} row={row}"


def _delete_column(doc: Any, op: dict) -> str:
    """删除表格列.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - column: 列号 (1-based)
            - count: 删除列数 (可选, 默认 1)
    """
    table_index = op.get("table_index", 1)
    column = op.get("column", 1)
    count = op.get("count", 1)
    t = _get_table_by_index(doc, table_index)
    if column < 1 or column > t.Columns.Count:
        raise COMOperationError("delete_column", f"列号超出范围: {column}")
    if column + count - 1 > t.Columns.Count:
        raise COMOperationError("delete_column", f"删除范围超出表格")
    for i in range(count):
        t.Columns(column).Delete()
    return f"deleted_column: table={table_index} column={column}"


def _merge_cells(doc: Any, op: dict) -> str:
    """合并单元格.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - start_row: 起始行 (1-based)
            - start_column: 起始列 (1-based)
            - end_row: 结束行 (1-based, 包含)
            - end_column: 结束列 (1-based, 包含)
    """
    table_index = op.get("table_index", 1)
    start_row = op.get("start_row", 1)
    start_column = op.get("start_column", 1)
    end_row = op.get("end_row", 1)
    end_column = op.get("end_column", 1)

    t = _get_table_by_index(doc, table_index)
    t.Cell(start_row, start_column).Merge(
        MergeTo=t.Cell(end_row, end_column)
    )
    return f"merged_cells: table={table_index} ({start_row},{start_column})-({end_row},{end_column})"


def _split_cell(doc: Any, op: dict) -> str:
    """拆分单元格.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - row: 行号 (1-based)
            - column: 列号 (1-based)
            - rows: 拆分行数 (可选, 默认 1)
            - columns: 拆分列数 (可选, 默认 1)
    """
    table_index = op.get("table_index", 1)
    row = op.get("row", 1)
    column = op.get("column", 1)
    rows = op.get("rows", 1)
    columns = op.get("columns", 1)

    t = _get_table_by_index(doc, table_index)
    t.Cell(row, column).Split(
        NumRows=int(rows),
        NumColumns=int(columns),
    )
    return f"split_cell: table={table_index} ({row},{column}) -> {rows}x{columns}"


def _set_table_borders(doc: Any, op: dict) -> str:
    """设置表格边框.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - style: 边框样式 (none, single, double, dotted, dashed, thick)
            - line_width: 线条宽度磅数 (可选, 默认 0.5)
    """
    table_index = op.get("table_index", 1)
    style = op.get("style", "single")
    line_width = op.get("line_width", 0.5)

    if style not in WD_LINE_STYLE_MAP:
        raise COMOperationError(
            "set_table_borders",
            f"不支持的边框样式: {style}, 支持: {list(WD_LINE_STYLE_MAP.keys())}",
        )

    t = _get_table_by_index(doc, table_index)
    t.Borders.Enable = True
    line_style = WD_LINE_STYLE_MAP[style]
    for border in t.Borders:
        try:
            border.LineStyle = line_style
            border.LineWidth = float(line_width)
        except Exception as e:
            logger.warning(f"设置边框失败: {e}")
    return f"set_table_borders: table={table_index} style={style}"


def _set_table_style(doc: Any, op: dict) -> str:
    """设置表格样式.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - table_index: 表格索引 (1-based)
            - style_name: 样式名称 (如 "Table Grid", "Light List")
    """
    table_index = op.get("table_index", 1)
    style_name = op.get("style_name", "")

    if not style_name:
        raise COMOperationError("set_table_style", "style_name 不能为空")

    t = _get_table_by_index(doc, table_index)

    # 样式名称在不同语言的 Word 中可能不同，尝试多个备选名称
    style_alternatives = [style_name]
    _TABLE_STYLE_ALIASES = {
        "Table Grid": ["网格型", "Table Grid"],
        "Light List": ["浅色列表", "Light List"],
        "Light Shading": ["浅色底纹", "Light Shading"],
        "Medium Shading 1": ["中等底纹 1", "Medium Shading 1"],
        "Medium Shading 2": ["中等底纹 2", "Medium Shading 2"],
        "Medium List 1": ["中等列表 1", "Medium List 1"],
    }
    if style_name in _TABLE_STYLE_ALIASES:
        style_alternatives = _TABLE_STYLE_ALIASES[style_name]

    last_error = None
    for alt_name in style_alternatives:
        try:
            t.Style = alt_name
            return f"set_table_style: table={table_index} style={alt_name}"
        except Exception as e:
            last_error = e
            continue

    raise COMOperationError(
        "set_table_style",
        f"设置表格样式失败，已尝试 {style_alternatives}，均不成功: {last_error}",
    )


# ===================== Styles 类工具 =====================

# 颜色名称到 RGB 映射
WD_COLOR_MAP = {
    "black": 0,        # wdColorBlack
    "blue": 16711680,
    "red": 255,
    "green": 65280,
    "yellow": 65535,
    "white": 16777215,
    "gray": 8421504,
    "orange": 42495,
    "purple": 10498160,
}


def _resolve_color(color: Any) -> int:
    """将颜色字符串/整数解析为 RGB 整数."""
    if isinstance(color, int):
        return color
    if isinstance(color, str):
        hex_str = color.lstrip("#")
        try:
            if len(hex_str) == 6:
                r = int(hex_str[0:2], 16)
                g = int(hex_str[2:4], 16)
                b = int(hex_str[4:6], 16)
                return r + (g << 8) + (b << 16)  # BGR
        except ValueError:
            pass
        key = color.lower().strip()
        if key in WD_COLOR_MAP:
            return WD_COLOR_MAP[key]
    raise COMOperationError("color", f"无法解析颜色: {color}")


def _resolve_range(doc: Any, op: dict) -> Any:
    """根据 op 解析目标 Range.

    - "all"/None: 整个文档
    - start/end: 字符位置范围
    - index: 单个段落
    """
    start = op.get("range_start")
    end = op.get("range_end")
    if start is not None and end is not None:
        return doc.Range(int(start), int(end))
    if "index" in op:
        return _get_paragraph_by_index(doc, op["index"]).Range
    return doc.Content


def _apply_bold(doc: Any, op: dict) -> str:
    """设置粗体.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - value: bool (可选, 默认 True)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    value = op.get("value", True)
    target = _resolve_range(doc, op)
    target.Font.Bold = bool(value)
    return f"apply_bold: {value}"


def _apply_italic(doc: Any, op: dict) -> str:
    """设置斜体.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - value: bool (可选, 默认 True)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    value = op.get("value", True)
    target = _resolve_range(doc, op)
    target.Font.Italic = bool(value)
    return f"apply_italic: {value}"


def _apply_underline(doc: Any, op: dict) -> str:
    """设置下划线.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - value: bool (可选, 默认 True)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    value = op.get("value", True)
    target = _resolve_range(doc, op)
    target.Font.Underline = 1 if value else 0  # wdUnderlineSingle=1, wdUnderlineNone=0
    return f"apply_underline: {value}"


def _set_font_color(doc: Any, op: dict) -> str:
    """设置字体颜色.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - color: 颜色 (名称 black/blue/red/green 等 或 #RRGGBB)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    color = op.get("color", "black")
    rgb = _resolve_color(color)
    target = _resolve_range(doc, op)
    target.Font.Color = rgb
    return f"set_font_color: {color}"


def _set_font_size(doc: Any, op: dict) -> str:
    """设置字体大小.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - size: 字号(磅)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    size = op.get("size", 12)
    target = _resolve_range(doc, op)
    target.Font.Size = float(size)
    return f"set_font_size: {size}"


def _set_highlight_color(doc: Any, op: dict) -> str:
    """设置高亮颜色.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - color: 颜色名称 (yellow, green, cyan, magenta, red, blue, dark_blue, teal, gray, black, white)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    color = op.get("color", "yellow")
    # 高亮颜色索引映射 (wdColorIndex)
    highlight_map = {
        "yellow": 7,        # wdYellow
        "green": 11,        # wdBrightGreen
        "cyan": 9,          # wdTurquoise
        "magenta": 13,      # wdPink
        "red": 6,           # wdRed
        "blue": 8,          # wdBlue
        "dark_blue": 12,    # wdDarkBlue
        "teal": 10,         # wdTeal
        "gray": 16,         # wdGray25
        "black": 1,         # wdBlack
        "white": 2,         # wdWhite
        "none": 0,          # wdNoHighlight
    }
    color_index = highlight_map.get(color.lower())
    if color_index is None:
        raise COMOperationError(
            "set_highlight_color",
            f"不支持的高亮颜色: {color}, 支持: {list(highlight_map.keys())}",
        )
    target = _resolve_range(doc, op)
    target.HighlightColorIndex = color_index
    return f"set_highlight_color: {color}"


def _set_strikethrough(doc: Any, op: dict) -> str:
    """设置删除线.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - value: bool (可选, 默认 True)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    value = op.get("value", True)
    target = _resolve_range(doc, op)
    target.Font.StrikeThrough = bool(value)
    return f"set_strikethrough: {value}"


def _set_subscript_superscript(doc: Any, op: dict) -> str:
    """设置下标/上标.

    Args:
        doc: Word 文档对象
        op: 操作配置
            - mode: subscript(下标) 或 superscript(上标) 或 normal(正常)
            - index: 段落索引 (可选)
            - range_start, range_end: 范围 (可选)
    """
    mode = op.get("mode", "normal")
    mode_map = {
        "normal": 0,         # wdNormal
        "superscript": 1,    # wdSuperscript
        "subscript": 2,      # wdSubscript
    }
    if mode not in mode_map:
        raise COMOperationError(
            "set_subscript_superscript",
            f"不支持的模式: {mode}, 支持: {list(mode_map.keys())}",
        )
    target = _resolve_range(doc, op)
    target.Font.Subscript = (mode == "subscript")
    target.Font.Superscript = (mode == "superscript")
    return f"set_subscript_superscript: {mode}"


# ====================================================================
# Sections 类操作 (6 个)
# ====================================================================

def _list_sections(doc: Any, op: dict) -> list[dict]:
    """列出文档中的所有分节.

    Returns:
        包含每个分节信息的字典列表
    """
    sections_info = []
    try:
        for i, section in enumerate(doc.Sections, start=1):
            section_info = {
                "index": i,
                "orientation": "portrait" if section.PageSetup.Orientation == 0 else "landscape",
                "top_margin": section.PageSetup.TopMargin,
                "bottom_margin": section.PageSetup.BottomMargin,
                "left_margin": section.PageSetup.LeftMargin,
                "right_margin": section.PageSetup.RightMargin,
                "page_width": section.PageSetup.PageWidth,
                "page_height": section.PageSetup.PageHeight,
            }
            sections_info.append(section_info)
    except Exception as e:
        raise COMOperationError("list_sections", str(e))
    return sections_info


def _add_section_break(doc: Any, op: dict) -> str:
    """添加分节符.

    Args:
        op: 操作配置
            - break_type: 分节符类型 (next_page, continuous, even_page, odd_page, column)
            - position: 位置 ("end" 或 "start", 默认 end)
    """
    break_type = op.get("break_type", "next_page")
    if break_type not in SECTION_BREAK_MAP:
        raise COMOperationError(
            "add_section_break",
            f"不支持的分节符类型: {break_type}, 支持: {list(SECTION_BREAK_MAP.keys())}"
        )

    try:
        position = op.get("position", "end")
        if position == "start":
            # 在文档开头插入分节符
            rng = doc.Range(0, 0)
            rng.InsertBreak(Type=SECTION_BREAK_MAP[break_type])
        elif position == "middle":
            # 在第一个段落处插入分节符（不折叠，确保创建新 Section）
            doc.Paragraphs(1).Range.InsertBreak(Type=SECTION_BREAK_MAP[break_type])
        else:
            # 默认在文档末尾插入分节符
            # 注意：在文档末尾插入分节符可能不会增加 Section 数量
            # 改为在最后一个段落之前插入，确保创建新 Section
            if doc.Paragraphs.Count > 1:
                rng = doc.Paragraphs(doc.Paragraphs.Count - 1).Range
                rng.Collapse(Direction=0)  # wdCollapseEnd
                rng.InsertBreak(Type=SECTION_BREAK_MAP[break_type])
            else:
                rng = doc.Content
                rng.Collapse(Direction=0)  # wdCollapseEnd
                rng.InsertBreak(Type=SECTION_BREAK_MAP[break_type])
        return f"added_section_break: {break_type} (sections={doc.Sections.Count})"
    except Exception as e:
        raise COMOperationError("add_section_break", str(e))


def _delete_section(doc: Any, op: dict) -> str:
    """删除指定分节.

    Args:
        op: 操作配置
            - section: 分节索引 (从1开始)
    """
    section_num = op.get("section", 1)
    try:
        if section_num < 1 or section_num > doc.Sections.Count:
            raise COMOperationError(
                "delete_section",
                f"分节索引超出范围: {section_num}, 共 {doc.Sections.Count} 个分节"
            )
        if doc.Sections.Count <= 1:
            raise COMOperationError("delete_section", "文档至少需要一个分节")

        # 取消分节符与其前一节合并
        section = doc.Sections(section_num)
        # 获取下一节(若有)的起始范围,然后取消分节
        if section_num < doc.Sections.Count:
            # 将当前分节的起始处内容合并到上一分节
            # 取消分节符
            rng = section.Range
            rng.Delete()
        return f"deleted_section: {section_num}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("delete_section", str(e))


def _set_section_orientation(doc: Any, op: dict) -> str:
    """设置指定分节的页面方向.

    Args:
        op: 操作配置
            - section: 分节索引 (默认1)
            - orientation: portrait 或 landscape
    """
    section_num = op.get("section", 1)
    orientation = op.get("orientation", "portrait")

    if orientation not in ORIENTATION_MAP:
        raise COMOperationError(
            "set_section_orientation",
            f"不支持的方向: {orientation}, 支持: {list(ORIENTATION_MAP.keys())}"
        )

    try:
        if section_num < 1 or section_num > doc.Sections.Count:
            raise COMOperationError(
                "set_section_orientation",
                f"分节索引超出范围: {section_num}"
            )
        section = doc.Sections(section_num)
        section.PageSetup.Orientation = ORIENTATION_MAP[orientation]
        return f"set_section_orientation: section={section_num}, orientation={orientation}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("set_section_orientation", str(e))


def _set_section_margins(doc: Any, op: dict) -> str:
    """设置指定分节的页边距 (单位: 磅).

    Args:
        op: 操作配置
            - section: 分节索引 (默认1)
            - top, bottom, left, right: 边距值 (磅)
    """
    section_num = op.get("section", 1)
    try:
        if section_num < 1 or section_num > doc.Sections.Count:
            raise COMOperationError(
                "set_section_margins",
                f"分节索引超出范围: {section_num}"
            )
        section = doc.Sections(section_num)
        setup = section.PageSetup
        if "top" in op:
            setup.TopMargin = op["top"]
        if "bottom" in op:
            setup.BottomMargin = op["bottom"]
        if "left" in op:
            setup.LeftMargin = op["left"]
        if "right" in op:
            setup.RightMargin = op["right"]
        return f"set_section_margins: section={section_num}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("set_section_margins", str(e))


def _set_section_columns(doc: Any, op: dict) -> str:
    """设置指定分节的列数.

    Args:
        op: 操作配置
            - section: 分节索引 (默认1)
            - columns: 列数
            - space_between: 列间距 (磅)
            - equal_width: 是否等宽 (默认 True)
    """
    section_num = op.get("section", 1)
    columns = op.get("columns", 1)
    space_between = op.get("space_between", 36)  # 默认 0.5 英寸
    equal_width = op.get("equal_width", True)

    try:
        if section_num < 1 or section_num > doc.Sections.Count:
            raise COMOperationError(
                "set_section_columns",
                f"分节索引超出范围: {section_num}"
            )
        section = doc.Sections(section_num)
        setup = section.PageSetup.TextColumns
        setup.SetCount(NumColumns=columns)
        if equal_width:
            setup.EvenlySpaced = True
            setup.Spacing = space_between
        return f"set_section_columns: section={section_num}, columns={columns}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("set_section_columns", str(e))


# ====================================================================
# Fields 类操作 (4 个新: list_fields, update_field, delete_field, get_field_result)
# ====================================================================

def _list_fields(doc: Any, op: dict) -> list[dict]:
    """列出文档中的所有域.

    Returns:
        包含每个域信息的字典列表
    """
    fields_info = []
    try:
        for i, field in enumerate(doc.Fields, start=1):
            try:
                fields_info.append({
                    "index": i,
                    "type": field.Type,
                    "code": field.Code.Text if field.Code else "",
                    "result": field.Result.Text[:100] if field.Result else "",
                })
            except Exception:
                fields_info.append({
                    "index": i,
                    "type": None,
                    "code": "",
                    "result": "",
                })
    except Exception as e:
        raise COMOperationError("list_fields", str(e))
    return fields_info


def _update_field(doc: Any, op: dict) -> str:
    """更新单个域.

    Args:
        op: 操作配置
            - field_index: 域索引 (从1开始)
    """
    field_index = op.get("field_index", 1)
    try:
        if field_index < 1 or field_index > doc.Fields.Count:
            raise COMOperationError(
                "update_field",
                f"域索引超出范围: {field_index}, 共 {doc.Fields.Count} 个域"
            )
        field = doc.Fields(field_index)
        field.Update()
        return f"updated_field: {field_index}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("update_field", str(e))


def _delete_field(doc: Any, op: dict) -> str:
    """删除单个域.

    Args:
        op: 操作配置
            - field_index: 域索引 (从1开始)
            - keep_result: 是否保留域结果文本 (默认 False)
    """
    field_index = op.get("field_index", 1)
    keep_result = op.get("keep_result", False)
    try:
        if field_index < 1 or field_index > doc.Fields.Count:
            raise COMOperationError(
                "delete_field",
                f"域索引超出范围: {field_index}, 共 {doc.Fields.Count} 个域"
            )
        field = doc.Fields(field_index)
        if keep_result:
            # 取消域绑定,保留结果文本
            field.Unlink()
        else:
            # 完全删除域
            field.Delete()
        return f"deleted_field: {field_index}, keep_result={keep_result}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("delete_field", str(e))


def _get_field_result(doc: Any, op: dict) -> dict:
    """获取域结果.

    Args:
        op: 操作配置
            - field_index: 域索引 (从1开始)
    """
    field_index = op.get("field_index", 1)
    try:
        if field_index < 1 or field_index > doc.Fields.Count:
            raise COMOperationError(
                "get_field_result",
                f"域索引超出范围: {field_index}, 共 {doc.Fields.Count} 个域"
            )
        field = doc.Fields(field_index)
        result_text = field.Result.Text if field.Result else ""
        code_text = field.Code.Text if field.Code else ""
        return {
            "field_index": field_index,
            "type": field.Type,
            "code": code_text,
            "result": result_text,
        }
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("get_field_result", str(e))


# ====================================================================
# Images 类操作 (8 个)
# ====================================================================

def _list_images(doc: Any, op: dict) -> list[dict]:
    """列出文档中的所有图片 (内联形状与浮动形状).

    Returns:
        包含每个图片信息的字典列表
    """
    images_info = []
    try:
        # 列出内联图片
        for i, shape in enumerate(doc.InlineShapes, start=1):
            try:
                images_info.append({
                    "index": i,
                    "type": "inline",
                    "width": shape.Width,
                    "height": shape.Height,
                })
            except Exception:
                pass

        # 列出浮动图片 (来自 doc.Shapes)
        offset = len(images_info)
        for j, shape in enumerate(doc.Shapes, start=1):
            try:
                if shape.Type in (13, 17, 18):  # msoPicture, msoLinkedPicture, msoPlaceholderPicture
                    images_info.append({
                        "index": offset + j,
                        "type": "floating",
                        "shape_type": int(shape.Type),
                        "width": shape.Width,
                        "height": shape.Height,
                    })
            except Exception:
                pass
    except Exception as e:
        raise COMOperationError("list_images", str(e))
    return images_info


def _get_image_info(doc: Any, op: dict) -> dict:
    """获取指定图片的详细信息.

    Args:
        op: 操作配置
            - image_index: 图片索引 (从1开始, 按 InlineShapes + Shapes 顺序)
    """
    image_index = op.get("image_index", 1)
    try:
        # 先在 InlineShapes 中查找
        if 1 <= image_index <= doc.InlineShapes.Count:
            shape = doc.InlineShapes(image_index)
            return {
                "image_index": image_index,
                "location": "inline",
                "type": int(shape.Type),
                "width": shape.Width,
                "height": shape.Height,
                "lock_aspect_ratio": shape.LockAspectRatio,
            }
        # 再在 Shapes 中查找
        offset = doc.InlineShapes.Count
        floating_index = image_index - offset
        if 1 <= floating_index <= doc.Shapes.Count:
            shape = doc.Shapes(floating_index)
            return {
                "image_index": image_index,
                "location": "floating",
                "type": int(shape.Type),
                "width": shape.Width,
                "height": shape.Height,
                "left": shape.Left,
                "top": shape.Top,
                "rotation": shape.Rotation,
            }
        raise COMOperationError(
            "get_image_info",
            f"图片索引超出范围: {image_index}"
        )
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("get_image_info", str(e))


def _resize_image(doc: Any, op: dict) -> str:
    """调整图片大小.

    Args:
        op: 操作配置
            - image_index: 图片索引 (从1开始)
            - width: 宽度 (磅)
            - height: 高度 (磅)
            - lock_aspect_ratio: 锁定纵横比 (默认 True)
    """
    image_index = op.get("image_index", 1)
    width = op.get("width")
    height = op.get("height")
    lock_aspect_ratio = op.get("lock_aspect_ratio", True)

    try:
        if 1 <= image_index <= doc.InlineShapes.Count:
            shape = doc.InlineShapes(image_index)
            shape.LockAspectRatio = lock_aspect_ratio
            if width is not None:
                shape.Width = width
            if height is not None:
                shape.Height = height
        else:
            offset = doc.InlineShapes.Count
            floating_index = image_index - offset
            if 1 <= floating_index <= doc.Shapes.Count:
                shape = doc.Shapes(floating_index)
                if lock_aspect_ratio:
                    shape.LockAspectRatio = -1  # msoTrue
                if width is not None:
                    shape.Width = width
                if height is not None:
                    shape.Height = height
            else:
                raise COMOperationError(
                    "resize_image",
                    f"图片索引超出范围: {image_index}"
                )
        return f"resized_image: {image_index}, width={width}, height={height}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("resize_image", str(e))


def _crop_image(doc: Any, op: dict) -> str:
    """裁剪图片.

    Args:
        op: 操作配置
            - image_index: 图片索引 (从1开始)
            - crop_left: 左侧裁剪量 (磅)
            - crop_right: 右侧裁剪量 (磅)
            - crop_top: 顶部裁剪量 (磅)
            - crop_bottom: 底部裁剪量 (磅)
    """
    image_index = op.get("image_index", 1)
    try:
        # 裁剪仅适用于浮动形状
        offset = doc.InlineShapes.Count
        if image_index <= offset:
            raise COMOperationError(
                "crop_image",
                "内联图片不支持裁剪, 请使用浮动图片"
            )
        floating_index = image_index - offset
        if floating_index < 1 or floating_index > doc.Shapes.Count:
            raise COMOperationError(
                "crop_image",
                f"图片索引超出范围: {image_index}"
            )
        shape = doc.Shapes(floating_index)
        if "crop_left" in op:
            shape.PictureFormat.CropLeft = op["crop_left"]
        if "crop_right" in op:
            shape.PictureFormat.CropRight = op["crop_right"]
        if "crop_top" in op:
            shape.PictureFormat.CropTop = op["crop_top"]
        if "crop_bottom" in op:
            shape.PictureFormat.CropBottom = op["crop_bottom"]
        return f"cropped_image: {image_index}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("crop_image", str(e))


def _set_image_position(doc: Any, op: dict) -> str:
    """设置图片位置 (仅对浮动图片有效).

    Args:
        op: 操作配置
            - image_index: 图片索引
            - left: 左边距 (磅)
            - top: 上边距 (磅)
    """
    image_index = op.get("image_index", 1)
    try:
        offset = doc.InlineShapes.Count
        if image_index <= offset:
            raise COMOperationError(
                "set_image_position",
                "内联图片不支持位置设置, 请使用浮动图片"
            )
        floating_index = image_index - offset
        if floating_index < 1 or floating_index > doc.Shapes.Count:
            raise COMOperationError(
                "set_image_position",
                f"图片索引超出范围: {image_index}"
            )
        shape = doc.Shapes(floating_index)
        if "left" in op:
            shape.Left = op["left"]
        if "top" in op:
            shape.Top = op["top"]
        return f"set_image_position: {image_index}, left={op.get('left')}, top={op.get('top')}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("set_image_position", str(e))


def _set_image_wrap(doc: Any, op: dict) -> str:
    """设置图片文字环绕类型.

    Args:
        op: 操作配置
            - image_index: 图片索引
            - wrap_type: 环绕类型 (none, square, tight, through, top_bottom, behind, in_front, inline)
    """
    image_index = op.get("image_index", 1)
    wrap_type = op.get("wrap_type", "inline")
    if wrap_type not in WRAP_TYPE_MAP:
        raise COMOperationError(
            "set_image_wrap",
            f"不支持的环绕类型: {wrap_type}, 支持: {list(WRAP_TYPE_MAP.keys())}"
        )

    try:
        if wrap_type == "inline":
            # 转为内联形状
            offset = doc.InlineShapes.Count
            if image_index > offset:
                floating_index = image_index - offset
                if 1 <= floating_index <= doc.Shapes.Count:
                    shape = doc.Shapes(floating_index)
                    shape.ConvertToInlineShape()
            return f"set_image_wrap: {image_index} -> inline"

        offset = doc.InlineShapes.Count
        if image_index <= offset:
            raise COMOperationError(
                "set_image_wrap",
                "内联图片需要先转换为浮动图片才能设置环绕"
            )
        floating_index = image_index - offset
        if floating_index < 1 or floating_index > doc.Shapes.Count:
            raise COMOperationError(
                "set_image_wrap",
                f"图片索引超出范围: {image_index}"
            )
        shape = doc.Shapes(floating_index)
        shape.WrapFormat.Type = WRAP_TYPE_MAP[wrap_type]
        return f"set_image_wrap: {image_index} -> {wrap_type}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("set_image_wrap", str(e))


def _replace_image(doc: Any, op: dict) -> str:
    """替换图片.

    Args:
        op: 操作配置
            - image_index: 图片索引
            - new_image_path: 新图片路径
    """
    image_index = op.get("image_index", 1)
    new_image_path = op.get("new_image_path", "")
    if not new_image_path:
        raise COMOperationError("replace_image", "new_image_path 不能为空")
    if not os.path.exists(new_image_path):
        raise COMOperationError("replace_image", f"新图片文件不存在: {new_image_path}")

    try:
        if 1 <= image_index <= doc.InlineShapes.Count:
            shape = doc.InlineShapes(image_index)
            shape.Range.InlineShapes.AddPicture(
                FileName=new_image_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Range=shape.Range,
            )
            shape.Delete()
        else:
            offset = doc.InlineShapes.Count
            floating_index = image_index - offset
            if 1 <= floating_index <= doc.Shapes.Count:
                shape = doc.Shapes(floating_index)
                shape.Fill.UserPicture(new_image_path)
            else:
                raise COMOperationError(
                    "replace_image",
                    f"图片索引超出范围: {image_index}"
                )
        return f"replaced_image: {image_index} -> {new_image_path}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("replace_image", str(e))


def _delete_image(doc: Any, op: dict) -> str:
    """删除图片.

    Args:
        op: 操作配置
            - image_index: 图片索引
    """
    image_index = op.get("image_index", 1)
    try:
        if 1 <= image_index <= doc.InlineShapes.Count:
            doc.InlineShapes(image_index).Delete()
        else:
            offset = doc.InlineShapes.Count
            floating_index = image_index - offset
            if 1 <= floating_index <= doc.Shapes.Count:
                doc.Shapes(floating_index).Delete()
            else:
                raise COMOperationError(
                    "delete_image",
                    f"图片索引超出范围: {image_index}"
                )
        return f"deleted_image: {image_index}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("delete_image", str(e))


# ====================================================================
# Hyperlinks 类操作 (3 个新: list_hyperlinks, get_hyperlink, remove_hyperlink)
# ====================================================================

def _list_hyperlinks(doc: Any, op: dict) -> list[dict]:
    """列出文档中的所有超链接.

    Returns:
        包含每个超链接信息的字典列表
    """
    links_info = []
    try:
        for i, link in enumerate(doc.Hyperlinks, start=1):
            try:
                links_info.append({
                    "index": i,
                    "text": link.TextToDisplay,
                    "address": link.Address,
                    "screen_tip": link.ScreenTip,
                    "anchor": link.Anchor,
                })
            except Exception:
                links_info.append({
                    "index": i,
                    "text": "",
                    "address": "",
                    "screen_tip": "",
                    "anchor": "",
                })
    except Exception as e:
        raise COMOperationError("list_hyperlinks", str(e))
    return links_info


def _get_hyperlink(doc: Any, op: dict) -> dict:
    """获取指定超链接的详细信息.

    Args:
        op: 操作配置
            - hyperlink_index: 超链接索引 (从1开始)
    """
    hyperlink_index = op.get("hyperlink_index", 1)
    try:
        if hyperlink_index < 1 or hyperlink_index > doc.Hyperlinks.Count:
            raise COMOperationError(
                "get_hyperlink",
                f"超链接索引超出范围: {hyperlink_index}, 共 {doc.Hyperlinks.Count} 个超链接"
            )
        link = doc.Hyperlinks(hyperlink_index)
        return {
            "hyperlink_index": hyperlink_index,
            "text": link.TextToDisplay,
            "address": link.Address,
            "screen_tip": link.ScreenTip,
            "anchor": link.Anchor,
        }
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("get_hyperlink", str(e))


def _remove_hyperlink(doc: Any, op: dict) -> str:
    """移除超链接 (保留文本).

    Args:
        op: 操作配置
            - hyperlink_index: 超链接索引 (从1开始)
    """
    hyperlink_index = op.get("hyperlink_index", 1)
    try:
        if hyperlink_index < 1 or hyperlink_index > doc.Hyperlinks.Count:
            raise COMOperationError(
                "remove_hyperlink",
                f"超链接索引超出范围: {hyperlink_index}, 共 {doc.Hyperlinks.Count} 个超链接"
            )
        doc.Hyperlinks(hyperlink_index).Delete()
        return f"removed_hyperlink: {hyperlink_index}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("remove_hyperlink", str(e))


# ====================================================================
# Protection 类操作 (6 个)
# ====================================================================

def _set_document_protection(doc: Any, op: dict) -> str:
    """设置文档保护.

    Args:
        op: 操作配置
            - protection_type: 保护类型 (none, read_only, comments, tracked_changes, forms, all)
            - password: 密码 (可选)
            - enforce: 是否强制 (默认 True)
    """
    protection_type = op.get("protection_type", "read_only")
    password = op.get("password", "")
    enforce = op.get("enforce", True)

    if protection_type not in PROTECTION_TYPE_MAP:
        raise COMOperationError(
            "set_document_protection",
            f"不支持的保护类型: {protection_type}, 支持: {list(PROTECTION_TYPE_MAP.keys())}"
        )

    try:
        if password:
            doc.Protect(
                Type=PROTECTION_TYPE_MAP[protection_type],
                NoReset=True,
                Password=password,
            )
        else:
            doc.Protect(
                Type=PROTECTION_TYPE_MAP[protection_type],
                NoReset=True,
            )
        return f"set_document_protection: type={protection_type}, enforce={enforce}"
    except Exception as e:
        raise COMOperationError("set_document_protection", str(e))


def _unprotect_document(doc: Any, op: dict) -> str:
    """取消文档保护.

    Args:
        op: 操作配置
            - password: 密码 (可选, 如果设置时提供了密码)
    """
    password = op.get("password", "")
    try:
        if password:
            doc.Unprotect(Password=password)
        else:
            doc.Unprotect()
        return "unprotected_document"
    except Exception as e:
        raise COMOperationError("unprotect_document", str(e))


def _set_read_only(doc: Any, op: dict) -> str:
    """设置文档为只读.

    Args:
        op: 操作配置
            - read_only: True/False (默认 True)
    """
    read_only = op.get("read_only", True)
    try:
        if read_only:
            doc.ReadOnlyRecommended = True
            doc.Protect(Type=PROTECTION_TYPE_MAP["read_only"], NoReset=True)
        else:
            doc.ReadOnlyRecommended = False
        return f"set_read_only: {read_only}"
    except Exception as e:
        raise COMOperationError("set_read_only", str(e))


def _set_password(doc: Any, op: dict) -> str:
    """设置文档打开密码.

    Args:
        op: 操作配置
            - password: 密码
    """
    password = op.get("password", "")
    if not password:
        raise COMOperationError("set_password", "password 不能为空")
    try:
        doc.Password = password
        return "password_set"
    except Exception as e:
        raise COMOperationError("set_password", str(e))


def _accept_all_changes(doc: Any, op: dict) -> str:
    """接受文档中的所有修订."""
    try:
        if hasattr(doc, 'AcceptAllRevisions'):
            doc.AcceptAllRevisions()
        # 兼容旧版本
        if hasattr(doc, 'Revisions'):
            count = doc.Revisions.Count
        else:
            count = 0
        return f"accepted_all_changes: {count} revisions"
    except Exception as e:
        raise COMOperationError("accept_all_changes", str(e))


def _reject_all_changes(doc: Any, op: dict) -> str:
    """拒绝文档中的所有修订."""
    try:
        if hasattr(doc, 'RejectAllRevisions'):
            doc.RejectAllRevisions()
        if hasattr(doc, 'Revisions'):
            count = doc.Revisions.Count
        else:
            count = 0
        return f"rejected_all_changes: {count} revisions"
    except Exception as e:
        raise COMOperationError("reject_all_changes", str(e))
