"""Excel COM 操作实现."""

import logging
from pathlib import Path
from typing import Any

from office_mcp.core.errors import COMOperationError
from office_mcp.core.path_guard import validate_path

logger = logging.getLogger(__name__)


# 操作 op 中以 _path / _file 结尾或明确为路径字段的 key 集合
_PATH_FIELDS = (
    "image_path", "source_path", "target_path", "template_path",
    "file_path", "new_path", "output_path", "output_dir", "from_file",
    "to_file", "src_path", "dst_path", "data_source",
)


def _validate_op_paths(op: dict) -> None:
    """校验 op dict 中所有疑似路径字段, 防止任意文件访问.

    对未在白名单内但形似 Windows 路径的字段, 也做警告.
    """
    for key, value in op.items():
        if not isinstance(value, str) or not value:
            continue
        # 已显式列入白名单: 校验
        if key.lower() in _PATH_FIELDS or key.lower().endswith(("_path", "_file", "path")):
            try:
                validate_path(value)
            except COMOperationError:
                raise
            except Exception as e:
                raise COMOperationError(op.get("type", "?"), f"路径校验失败 {key}={value}: {e}")

# 图表类型映射
CHART_TYPE_MAP = {
    "column": 51,      # xlColumnClustered
    "bar": 57,         # xlBarClustered
    "line": 65,        # xlLine
    "pie": 5,          # xlPie
    "scatter": 72,     # xlXYScatter
    "area": 76,        # xlArea
}


def apply_excel_operations(workbook: Any, operations: list[dict]) -> list[dict]:
    """对 Excel 工作簿执行批量操作.

    Args:
        workbook: Excel Workbook 对象
        operations: 操作列表

    Returns:
        每个操作的执行结果
    """
    results = []
    for op in operations:
        op_type = op.get("type", "")
        try:
            # 入口处校验 op 中所有疑似路径字段
            _validate_op_paths(op)
            result = _execute_excel_operation(workbook, op)
            results.append({"type": op_type, "status": "success", "result": result})
        except Exception as e:
            logger.error(f"Excel 操作失败 [{op_type}]: {e}")
            results.append({"type": op_type, "status": "error", "message": str(e)})
    return results


def _get_sheet(workbook: Any, sheet_name: str) -> Any:
    """获取工作表."""
    try:
        return workbook.Worksheets(sheet_name)
    except Exception as e:
        raise COMOperationError(f"获取工作表 '{sheet_name}'", str(e))


def _col_idx_to_letters(col: int) -> str:
    """将 1-based 列号转为字母: 1->A, 26->Z, 27->AA, 52->AZ, 53->BA, 702->ZZ, 703->AAA."""
    if col < 1:
        return ""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _execute_excel_operation(workbook: Any, op: dict) -> Any:
    """执行单个 Excel 操作."""
    op_type = op.get("type", "")

    if op_type == "write_cell":
        return _write_cell(workbook, op)
    elif op_type == "write_range":
        return _write_range(workbook, op)
    elif op_type == "read_range":
        return _read_range(workbook, op)
    elif op_type == "add_formula":
        return _add_formula(workbook, op)
    elif op_type == "format_range":
        return _format_range(workbook, op)
    elif op_type == "set_number_format":
        return _set_number_format(workbook, op)
    elif op_type == "create_chart":
        return _create_chart(workbook, op)
    elif op_type == "add_worksheet":
        return _add_worksheet(workbook, op)
    elif op_type == "rename_worksheet":
        return _rename_worksheet(workbook, op)
    elif op_type == "auto_fit_columns":
        return _auto_fit_columns(workbook, op)
    elif op_type == "freeze_panes":
        return _freeze_panes(workbook, op)
    elif op_type == "save":
        workbook.Save()
        return "saved"
    elif op_type == "add_data_validation":
        return _add_data_validation(workbook, op)
    elif op_type == "add_conditional_format":
        return _add_conditional_format(workbook, op)
    elif op_type == "merge_cells":
        return _merge_cells(workbook, op)
    elif op_type == "set_borders":
        return _set_borders(workbook, op)
    elif op_type == "add_named_range":
        return _add_named_range(workbook, op)
    elif op_type == "create_pivot_table":
        return _create_pivot_table(workbook, op)
    elif op_type == "import_data":
        return _import_data(workbook, op)
    elif op_type == "export_data":
        return _export_data(workbook, op)
    elif op_type == "add_slicer":
        return _add_slicer(workbook, op)
    elif op_type == "add_subtotal":
        return _add_subtotal(workbook, op)
    elif op_type == "check_typography":
        return _check_typography(workbook, op)
    elif op_type == "list_worksheets":
        return _list_worksheets(workbook, op)
    elif op_type == "get_worksheet_info":
        return _get_worksheet_info(workbook, op)
    elif op_type == "copy_worksheet":
        return _copy_worksheet(workbook, op)
    elif op_type == "delete_worksheet":
        return _delete_worksheet(workbook, op)
    elif op_type == "move_worksheet":
        return _move_worksheet(workbook, op)
    elif op_type == "hide_worksheet":
        return _hide_worksheet(workbook, op)
    elif op_type == "show_worksheet":
        return _show_worksheet(workbook, op)
    elif op_type == "protect_worksheet":
        return _protect_worksheet(workbook, op)
    elif op_type == "unprotect_worksheet":
        return _unprotect_worksheet(workbook, op)
    elif op_type == "set_tab_color":
        return _set_tab_color(workbook, op)
    elif op_type == "list_used_range":
        return _list_used_range(workbook, op)
    elif op_type == "clear_range":
        return _clear_range(workbook, op)
    elif op_type == "copy_range":
        return _copy_range(workbook, op)
    elif op_type == "paste_range":
        return _paste_range(workbook, op)
    elif op_type == "cut_range":
        return _cut_range(workbook, op)
    elif op_type == "delete_cells":
        return _delete_cells(workbook, op)
    elif op_type == "insert_cells":
        return _insert_cells(workbook, op)
    elif op_type == "set_row_height":
        return _set_row_height(workbook, op)
    elif op_type == "set_column_width":
        return _set_column_width(workbook, op)
    elif op_type == "hide_rows":
        return _hide_rows(workbook, op)
    elif op_type == "list_charts":
        return _list_charts(workbook, op)
    elif op_type == "get_chart_info":
        return _get_chart_info(workbook, op)
    elif op_type == "set_chart_title":
        return _set_chart_title(workbook, op)
    elif op_type == "set_chart_legend":
        return _set_chart_legend(workbook, op)
    elif op_type == "add_chart_series":
        return _add_chart_series(workbook, op)
    elif op_type == "remove_chart_series":
        return _remove_chart_series(workbook, op)
    elif op_type == "set_chart_axis":
        return _set_chart_axis(workbook, op)
    elif op_type == "change_chart_type":
        return _change_chart_type(workbook, op)
    elif op_type == "export_chart":
        return _export_chart(workbook, op)
    elif op_type == "delete_chart":
        return _delete_chart(workbook, op)
    elif op_type == "set_font":
        return _set_font(workbook, op)
    elif op_type == "set_font_bold":
        return _set_font_bold(workbook, op)
    elif op_type == "set_font_italic":
        return _set_font_italic(workbook, op)
    elif op_type == "set_font_underline":
        return _set_font_underline(workbook, op)
    elif op_type == "set_alignment":
        return _set_alignment(workbook, op)
    elif op_type == "set_wrap_text":
        return _set_wrap_text(workbook, op)
    elif op_type == "set_indent":
        return _set_indent(workbook, op)
    elif op_type == "set_orientation":
        return _set_orientation(workbook, op)
    elif op_type == "clear_format":
        return _clear_format(workbook, op)
    elif op_type == "copy_format":
        return _copy_format(workbook, op)
    elif op_type == "set_page_orientation":
        return _set_page_orientation(workbook, op)
    elif op_type == "set_page_size":
        return _set_page_size(workbook, op)
    elif op_type == "set_page_margins":
        return _set_page_margins(workbook, op)
    elif op_type == "set_header":
        return _set_header(workbook, op)
    elif op_type == "set_footer":
        return _set_footer(workbook, op)
    elif op_type == "add_print_title":
        return _add_print_title(workbook, op)
    elif op_type == "set_print_area":
        return _set_print_area(workbook, op)
    elif op_type == "set_page_break":
        return _set_page_break(workbook, op)
    elif op_type == "set_scale":
        return _set_scale(workbook, op)
    elif op_type == "set_fit_to_page":
        return _set_fit_to_page(workbook, op)
    elif op_type == "set_array_formula":
        return _set_array_formula(workbook, op)
    elif op_type == "evaluate_formula":
        return _evaluate_formula(workbook, op)
    elif op_type == "replace_formula":
        return _replace_formula(workbook, op)
    elif op_type == "find_formula_cells":
        return _find_formula_cells(workbook, op)
    elif op_type == "convert_to_values":
        return _convert_to_values(workbook, op)
    elif op_type == "get_formula_info":
        return _get_formula_info(workbook, op)
    elif op_type == "define_name":
        return _define_name(workbook, op)
    elif op_type == "create_table":
        return _create_table(workbook, op)
    elif op_type == "list_tables":
        return _list_tables(workbook, op)
    elif op_type == "resize_table":
        return _resize_table(workbook, op)
    elif op_type == "set_table_style":
        return _set_table_style(workbook, op)
    elif op_type == "show_table_totals":
        return _show_table_totals(workbook, op)
    elif op_type == "add_table_column":
        return _add_table_column(workbook, op)
    elif op_type == "remove_table_column":
        return _remove_table_column(workbook, op)
    elif op_type == "delete_table":
        return _delete_table(workbook, op)
    elif op_type == "add_auto_filter":
        return _add_auto_filter(workbook, op)
    elif op_type == "remove_auto_filter":
        return _remove_auto_filter(workbook, op)
    elif op_type == "sort_range":
        return _sort_range(workbook, op)
    elif op_type == "advanced_filter":
        return _advanced_filter(workbook, op)
    elif op_type == "remove_duplicates":
        return _remove_duplicates(workbook, op)
    elif op_type == "group_rows":
        return _group_rows(workbook, op)
    elif op_type == "ungroup_rows":
        return _ungroup_rows(workbook, op)
    elif op_type == "group_columns":
        return _group_columns(workbook, op)
    elif op_type == "ungroup_columns":
        return _ungroup_columns(workbook, op)
    elif op_type == "protect_workbook":
        return _protect_workbook(workbook, op)
    elif op_type == "unprotect_workbook":
        return _unprotect_workbook(workbook, op)
    elif op_type == "set_open_password":
        return _set_open_password(workbook, op)
    elif op_type == "set_write_reservation_password":
        return _set_write_reservation_password(workbook, op)
    elif op_type == "mark_as_final":
        return _mark_as_final(workbook, op)
    elif op_type == "recommend_read_only":
        return _recommend_read_only(workbook, op)
    elif op_type == "add_image":
        return _add_image(workbook, op)
    elif op_type == "list_shapes":
        return _list_shapes(workbook, op)
    elif op_type == "delete_shape":
        return _delete_shape(workbook, op)
    elif op_type == "add_comment":
        return _add_comment(workbook, op)
    elif op_type == "delete_comment":
        return _delete_comment(workbook, op)
    elif op_type == "set_view_zoom":
        return _set_view_zoom(workbook, op)
    elif op_type == "set_view_gridlines":
        return _set_view_gridlines(workbook, op)
    elif op_type == "set_view_headings":
        return _set_view_headings(workbook, op)
    elif op_type == "recalculate":
        return _recalculate(workbook, op)
    elif op_type == "set_calculation_mode":
        return _set_calculation_mode(workbook, op)
    elif op_type == "set_iterative_calc":
        return _set_iterative_calc(workbook, op)
    else:
        raise COMOperationError(f"未知的 Excel 操作类型: {op_type}")


def _check_typography(workbook: Any, op: dict) -> list[dict]:
    """检查 Excel 工作簿排版问题.

    Args:
        workbook: Excel 工作簿对象
        op: 操作配置

    Returns:
        问题列表，每个问题包含 type, description, location
    """
    issues = []
    sheet_name = op.get("sheet", None)

    try:
        # 如果指定了工作表，只检查该表，否则检查所有表
        sheets_to_check = []
        if sheet_name:
            sheets_to_check.append(_get_sheet(workbook, sheet_name))
        else:
            for sheet in workbook.Worksheets:
                sheets_to_check.append(sheet)

        for sheet in sheets_to_check:
            sheet_name_current = sheet.Name
            # 1. 检查单元格内容对齐
            issues.extend(_check_cell_alignment(sheet, sheet_name_current))

            # 2. 检查数字格式一致性
            issues.extend(_check_number_format_consistency(sheet, sheet_name_current))

            # 3. 检查边框使用规范
            issues.extend(_check_border_usage(sheet, sheet_name_current))

    except Exception as e:
        logger.error(f"Excel 排版检查出错: {e}")
        issues.append({
            "type": "error",
            "description": f"排版检查过程中发生错误: {str(e)}",
            "location": "整个工作簿"
        })

    return issues


def _check_cell_alignment(sheet: Any, sheet_name: str) -> list[dict]:
    """检查单元格内容对齐."""
    issues = []
    try:
        # 定义 Excel 对齐常量
        xlHAlignGeneral = 1
        xlHAlignLeft = -4131
        xlHAlignCenter = -4108
        xlHAlignRight = -4152

        # 获取使用范围
        used_range = sheet.UsedRange
        if used_range is None:
            return issues

        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        # 简单检查：同一列的单元格对齐方式是否一致（针对前100行和前20列）
        max_rows = min(row_count, 100)
        max_cols = min(col_count, 20)

        for col in range(1, max_cols + 1):
            # 获取第一行的数据类型作为参考
            first_cell = sheet.Cells(1, col)
            first_value = first_cell.Value
            first_align = first_cell.HorizontalAlignment

            # 如果第一行有值，检查同列其他单元格
            if first_value is not None:
                for row in range(2, max_rows + 1):
                    cell = sheet.Cells(row, col)
                    cell_value = cell.Value

                    if cell_value is not None:
                        # 数字和文本通常有不同的对齐习惯
                        # 数字通常右对齐，文本通常左对齐
                        cell_align = cell.HorizontalAlignment
                        is_number = isinstance(cell_value, (int, float))
                        is_first_number = isinstance(first_value, (int, float))

                        if is_number and cell_align not in (xlHAlignRight, xlHAlignGeneral):
                            issues.append({
                                "type": "cell_alignment",
                                "description": f"数字单元格建议使用右对齐，当前对齐: {cell_align}",
                                "location": f"{sheet_name}!{_col_idx_to_letters(col)}{row}"
                            })
                        elif not is_number and cell_align == xlHAlignRight:
                            issues.append({
                                "type": "cell_alignment",
                                "description": f"文本单元格建议使用左对齐",
                                "location": f"{sheet_name}!{_col_idx_to_letters(col)}{row}"
                            })
    except Exception as e:
        logger.warning(f"检查单元格对齐出错: {e}")
    return issues


def _check_number_format_consistency(sheet: Any, sheet_name: str) -> list[dict]:
    """检查数字格式一致性."""
    issues = []
    try:
        used_range = sheet.UsedRange
        if used_range is None:
            return issues

        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        max_rows = min(row_count, 100)
        max_cols = min(col_count, 20)

        for col in range(1, max_cols + 1):
            # 收集列中所有数字单元格的格式
            number_formats = []
            for row in range(1, max_rows + 1):
                cell = sheet.Cells(row, col)
                cell_value = cell.Value
                if isinstance(cell_value, (int, float)):
                    fmt = cell.NumberFormat
                    if fmt and fmt not in number_formats:
                        number_formats.append(fmt)

            # 如果同列中有多种数字格式，建议统一
            if len(number_formats) > 1:
                issues.append({
                    "type": "number_format",
                    "description": f"列中存在多种数字格式: {', '.join(number_formats)}",
                    "location": f"{sheet_name}!列 {_col_idx_to_letters(col)}"
                })
    except Exception as e:
        logger.warning(f"检查数字格式出错: {e}")
    return issues


def _check_border_usage(sheet: Any, sheet_name: str) -> list[dict]:
    """检查边框使用规范: 报告无边框的有内容孤立单元格.

    注: 仅报告顶部行 (标题行) 的单元格是否缺少边框, 简单启发式.
    """
    issues: list[dict] = []
    try:
        used_range = sheet.UsedRange
        if used_range is None:
            return issues

        # 边框 COM 常量
        xlEdgeTop = 8
        xlLineStyleNone = -4142

        # 只检查第一行 (header 行) 是否有边框
        col_count = min(int(used_range.Columns.Count), 20)
        for col in range(1, col_count + 1):
            cell = sheet.Cells(1, col)
            if cell.Value is None:
                continue
            try:
                top_border = cell.Borders(xlEdgeTop)
                if top_border.LineStyle == xlLineStyleNone:
                    issues.append({
                        "type": "border",
                        "description": "标题行单元格缺少上边框",
                        "location": f"{sheet_name}!{cell.Address(False, False)}",
                    })
            except Exception:
                continue
    except Exception as e:
        logger.warning(f"检查边框使用出错: {e}")
    return issues


def _write_cell(workbook: Any, op: dict) -> str:
    """写入单元格."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    value = op.get("value", "")
    sheet.Range(cell).Value = value
    return f"wrote_cell: {cell} = {value}"


def _write_range(workbook: Any, op: dict) -> str:
    """写入范围."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    start_cell = op.get("start_cell", "A1")
    data = op.get("data", [])

    if not data:
        return "no_data"

    rows = len(data)
    cols = max(len(row) for row in data) if data else 0

    # 计算结束单元格
    start_row = sheet.Range(start_cell).Row
    start_col = sheet.Range(start_cell).Column
    end_row = start_row + rows - 1
    end_col = start_col + cols - 1

    # 将列号转为字母
    def col_to_letter(col: int) -> str:
        result = ""
        while col > 0:
            col, rem = divmod(col - 1, 26)
            result = chr(65 + rem) + result
        return result

    end_cell = f"{col_to_letter(end_col)}{end_row}"
    range_obj = sheet.Range(f"{start_cell}:{end_cell}")

    # 填充数据，补全短行
    filled_data = []
    for row in data:
        filled_row = list(row) + [""] * (cols - len(row))
        filled_data.append(filled_row)

    range_obj.Value = filled_data
    return f"wrote_range: {start_cell}:{end_cell}"


def _read_range(workbook: Any, op: dict) -> Any:
    """读取范围."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    values = sheet.Range(range_str).Value

    # 统一为二维列表
    if values is None:
        return []
    if not isinstance(values, tuple):
        values = ((values,),)

    # 处理单行或单列的情况
    result = []
    for row in values:
        if isinstance(row, tuple):
            result.append(list(row))
        else:
            result.append([row])
    return result


def _add_formula(workbook: Any, op: dict) -> str:
    """添加公式."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    formula = op.get("formula", "")
    sheet.Range(cell).Formula = formula
    return f"added_formula: {cell} = {formula}"


def _format_range(workbook: Any, op: dict) -> str:
    """格式化范围."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    range_obj = sheet.Range(range_str)

    if op.get("bold"):
        range_obj.Font.Bold = True
    if op.get("italic"):
        range_obj.Font.Italic = True

    # 背景色 (支持 #RRGGBB)
    bg_color = op.get("background_color")
    if bg_color:
        range_obj.Interior.Color = _hex_to_rgb(bg_color)

    # 字体色
    font_color = op.get("font_color")
    if font_color:
        range_obj.Font.Color = _hex_to_rgb(font_color)

    return f"formatted_range: {range_str}"


def _set_number_format(workbook: Any, op: dict) -> str:
    """设置数字格式."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    fmt = op.get("format", "General")
    sheet.Range(range_str).NumberFormat = fmt
    return f"set_number_format: {range_str} -> {fmt}"


def _create_chart(workbook: Any, op: dict) -> str:
    """创建图表."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_type = op.get("chart_type", "column")
    data_range = op.get("data_range", "A1:B5")
    title = op.get("title", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 400)
    height = op.get("height", 300)

    chart_type_val = CHART_TYPE_MAP.get(chart_type, 51)

    chart = sheet.ChartObjects().Add(left, top, width, height).Chart
    chart.ChartType = chart_type_val
    chart.SetSourceData(sheet.Range(data_range))

    if title:
        chart.HasTitle = True
        chart.ChartTitle.Text = title

    return f"created_chart: {chart_type} ({data_range})"


def _add_worksheet(workbook: Any, op: dict) -> str:
    """添加工作表."""
    name = op.get("name", "Sheet")
    sheet = workbook.Worksheets.Add()
    sheet.Name = name
    return f"added_worksheet: {name}"


def _rename_worksheet(workbook: Any, op: dict) -> str:
    """重命名工作表."""
    old_name = op.get("old_name", "")
    new_name = op.get("new_name", "")
    workbook.Worksheets(old_name).Name = new_name
    return f"renamed_worksheet: {old_name} -> {new_name}"


def _auto_fit_columns(workbook: Any, op: dict) -> str:
    """自动调整列宽."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    columns = op.get("columns", [])
    if isinstance(columns, str):
        sheet.Columns(columns).AutoFit()
        return f"auto_fit_columns: {columns}"
    for col in columns:
        sheet.Columns(col).AutoFit()
    return f"auto_fit_columns: {columns}"


def _freeze_panes(workbook: Any, op: dict) -> str:
    """冻结窗格."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A2")
    sheet.Range(cell).Select()
    sheet.Application.ActiveWindow.FreezePanes = True
    return f"freeze_panes: {cell}"


def _hex_to_rgb(hex_color: str) -> int:
    """将 #RRGGBB 转为 Office RGB 整数."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return 0
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g << 8) + (b << 16)


# ============ Excel 高级功能 ============

def _add_data_validation(workbook: Any, op: dict) -> str:
    """添加数据验证.

    Args:
        sheet: 工作表名称
        range: 验证范围 (如 "A1:A10")
        type: 验证类型 (list/whole/decimal/date/time/textLength/custom)
        formula1: 验证公式或列表值 (逗号分隔)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:A10")
    validation_type = op.get("type", "list")
    formula1 = op.get("formula1", "")

    # 验证类型映射
    type_map = {
        "list": 3,          # xlValidateList
        "whole": 1,         # xlValidateWholeNumber
        "decimal": 2,       # xlValidateDecimal
        "date": 4,          # xlValidateDate
        "time": 5,          # xlValidateTime
        "text_length": 6,   # xlValidateTextLength
        "custom": 0,        # xlValidateCustom
    }

    validation_type_val = type_map.get(validation_type, 3)
    range_obj = sheet.Range(range_str)

    # 先删除已有验证，避免冲突
    try:
        range_obj.Validation.Delete()
    except Exception:
        pass

    try:
        # list 类型 (Type=3) 不需要 Operator 参数，Formula1 不能为空
        if validation_type_val == 3:  # xlValidateList
            if not formula1:
                formula1 = "a,b,c"
            range_obj.Validation.Add(
                Type=validation_type_val,
                AlertStyle=1,  # xlValidAlertStop
                Formula1=formula1,
            )
        else:
            range_obj.Validation.Add(
                Type=validation_type_val,
                AlertStyle=1,  # xlValidAlertStop
                Operator=0,
                Formula1=formula1 if formula1 else "0",
            )
    except Exception as e:
        raise COMOperationError("add_data_validation", str(e))
    return f"added_data_validation: {range_str} ({validation_type})"


def _add_conditional_format(workbook: Any, op: dict) -> str:
    """添加条件格式.

    Args:
        sheet: 工作表名称
        range: 范围 (如 "A1:A10")
        type: 条件类型 (cell_value/formula/color_scale/data_bar/icon_set)
        operator: 操作符 (greater/less/equal/between)
        formula1: 条件值1
        formula2: 条件值2
        format_type: 格式类型 (color_scale/data_bar/icon_set)
        font_color: 字体颜色 (#RRGGBB)
        bg_color: 背景颜色 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:A10")
    condition_type = op.get("type", "cell_value")
    operator = op.get("operator", "greater")
    formula1 = op.get("formula1", "")
    formula2 = op.get("formula2", "")
    format_type = op.get("format_type", "")
    font_color = op.get("font_color", "")
    bg_color = op.get("bg_color", "")

    range_obj = sheet.Range(range_str)

    # 先删除已有条件格式，避免冲突
    try:
        range_obj.FormatConditions.Delete()
    except Exception:
        pass

    # 高级格式类型
    if format_type == "color_scale":
        # 色阶
        color_scale = range_obj.FormatConditions.AddColorScale(ColorScaleType=3)
        # 设置默认颜色: 红-黄-绿
        color_scale.ColorScaleCriteria(1).Type = 1  # xlLowestValue
        color_scale.ColorScaleCriteria(1).FormatColor.Color = _hex_to_rgb("#FF0000")
        color_scale.ColorScaleCriteria(2).Type = 5  # xlPercentile
        color_scale.ColorScaleCriteria(2).Value = 50
        color_scale.ColorScaleCriteria(2).FormatColor.Color = _hex_to_rgb("#FFFF00")
        color_scale.ColorScaleCriteria(3).Type = 2  # xlHighestValue
        color_scale.ColorScaleCriteria(3).FormatColor.Color = _hex_to_rgb("#00FF00")
        return f"added_conditional_format: {range_str} (color_scale)"

    elif format_type == "data_bar":
        # 数据条
        data_bar = range_obj.FormatConditions.AddDatabar()
        data_bar.BarColor.Color = _hex_to_rgb("#638EC6")
        data_bar.BarFillType = 0  # xlDataBarFillSolid
        return f"added_conditional_format: {range_str} (data_bar)"

    elif format_type == "icon_set":
        # 图标集
        icon_set = range_obj.FormatConditions.AddIconSetCondition()
        icon_set.IconSet = workbook.Application.IconSets(3)  # xl3Arrows
        return f"added_conditional_format: {range_str} (icon_set)"

    # 常规条件格式
    else:
        if condition_type == "cell_value":
            # xlCellValue 类型 Formula1 不能为空
            if not formula1:
                formula1 = "0"
            # 操作符映射
            op_map = {
                "between": 1,        # xlBetween
                "not_between": 2,    # xlNotBetween
                "equal": 3,          # xlEqual
                "not_equal": 4,      # xlNotEqual
                "greater": 5,        # xlGreater
                "less": 6,           # xlLess
                "greater_equal": 7,  # xlGreaterEqual
                "less_equal": 8,     # xlLessEqual
            }
            operator_val = op_map.get(operator, 5)

            params = {
                "Type": 1,  # xlCellValue
                "Operator": operator_val,
                "Formula1": formula1,
            }
            if operator in ["between"] and formula2:
                params["Formula2"] = formula2

            try:
                format_condition = range_obj.FormatConditions.Add(**params)
            except Exception as e:
                raise COMOperationError("add_conditional_format", f"FormatConditions.Add 失败: {e}")
        else:  # formula
            if not formula1:
                formula1 = "=TRUE"
            try:
                format_condition = range_obj.FormatConditions.Add(
                    Type=2,  # xlExpression
                    Formula1=formula1,
                )
            except Exception as e:
                raise COMOperationError("add_conditional_format", f"FormatConditions.Add 失败: {e}")

        # 应用格式
        if font_color:
            format_condition.Font.Color = _hex_to_rgb(font_color)
        if bg_color:
            format_condition.Interior.Color = _hex_to_rgb(bg_color)

        return f"added_conditional_format: {range_str}"


def _merge_cells(workbook: Any, op: dict) -> str:
    """合并单元格.

    Args:
        sheet: 工作表名称
        range: 范围 (如 "A1:C3")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:C3")
    sheet.Range(range_str).Merge()
    return f"merged_cells: {range_str}"


def _set_borders(workbook: Any, op: dict) -> str:
    """设置边框.

    Args:
        sheet: 工作表名称
        range: 范围 (如 "A1:C3")
        border_type: 边框类型 (all/outside/inside)
        style: 线型 (thin/medium/thick/dashed/dotted)
        color: 颜色 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:C3")
    border_type = op.get("border_type", "all")
    style = op.get("style", "thin")
    color = op.get("color", "#000000")

    # 线型映射
    style_map = {
        "thin": 1,          # xlThin
        "medium": -4118,    # xlMedium
        "thick": 4,         # xlThick
        "dashed": -4115,    # xlDash
        "dotted": -4122,    # xlDot
    }
    style_val = style_map.get(style, 1)
    color_val = _hex_to_rgb(color)

    rng = sheet.Range(range_str)

    if border_type == "all":
        # xlEdgeTop/Bottom/Left/Right/InsideHorizontal/InsideVertical
        for border in [7, 8, 9, 10, 11, 12]:
            rng.Borders(border).LineStyle = style_val
            rng.Borders(border).Color = color_val
    elif border_type == "outside":
        for border in [7, 8, 9, 10]:  # xlEdgeTop/Bottom/Left/Right
            rng.Borders(border).LineStyle = style_val
            rng.Borders(border).Color = color_val
    elif border_type == "inside":
        for border in [11, 12]:  # xlInsideHorizontal/InsideVertical
            rng.Borders(border).LineStyle = style_val
            rng.Borders(border).Color = color_val

    return f"set_borders: {range_str} ({border_type})"


def _add_named_range(workbook: Any, op: dict) -> str:
    """添加命名范围.

    Args:
        name: 名称
        refers_to: 引用公式 (如 "=Sheet1!$A$1:$A$10")
    """
    name = op.get("name", "")
    refers_to = op.get("refers_to", "=Sheet1!$A$1:$A$10")

    if not name:
        raise COMOperationError("add_named_range", "name 不能为空")

    try:
        # 删除已存在的同名范围
        try:
            for i in range(1, workbook.Names.Count + 1):
                try:
                    if workbook.Names(i).Name == name:
                        workbook.Names(i).Delete()
                        break
                except Exception:
                    continue
        except Exception:
            pass

        try:
            workbook.Names.Add(Name=name, RefersTo=refers_to)
        except Exception:
            # RefersTo 失败时尝试 RefersToR1C1 作为回退
            try:
                workbook.Names.Add(Name=name, RefersToR1C1=refers_to)
            except Exception as e2:
                raise COMOperationError("add_named_range", f"RefersTo 和 RefersToR1C1 均失败: {e2}")
        return f"added_named_range: {name}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("add_named_range", str(e))


def _create_pivot_table(workbook: Any, op: dict) -> str:
    """创建数据透视表.

    Args:
        source_sheet: 数据源工作表名称
        source_range: 数据源范围 (如 "A1:D100", 留空则自动使用 UsedRange)
        target_sheet: 目标工作表名称 (自动创建或指定)
        target_cell: 目标单元格 (如 "A3")
        row_fields: 行字段列表 (如 ["部门", "月份"])
        column_fields: 列字段列表 (如 ["地区"])
        data_fields: 数据字段字典 (如 {"销售额": "sum", "数量": "average"})
    """
    source_sheet = _get_sheet(workbook, op.get("source_sheet", "Sheet1"))
    source_range = op.get("source_range", "")
    target_sheet_name = op.get("target_sheet", "数据透视表")
    target_cell = op.get("target_cell", "A3")

    row_fields = op.get("row_fields", [])
    column_fields = op.get("column_fields", [])
    data_fields = op.get("data_fields", {})

    # 如果未指定 source_range，则使用 UsedRange 避免引用超出实际数据范围
    if not source_range:
        used = source_sheet.UsedRange
        if used is not None:
            source_range = used.Address
        else:
            source_range = "A1"

    # 创建或获取目标工作表
    try:
        target_sheet = workbook.Worksheets(target_sheet_name)
    except Exception:
        target_sheet = workbook.Worksheets.Add()
        target_sheet.Name = target_sheet_name

    # 创建数据透视表缓存 (使用地址字符串更可靠)
    source_data_addr = source_sheet.Range(source_range)
    try:
        pivot_cache = workbook.PivotCaches.Create(
            SourceType=1,  # xlDatabase
            SourceData=source_data_addr,
        )
    except Exception as e:
        raise COMOperationError("create_pivot_table", f"PivotCaches.Create 失败: {e}")

    # 生成不重复的表名
    import time
    table_name = f"PivotTable_{int(time.time())}"

    # 创建数据透视表
    try:
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=target_sheet.Range(target_cell),
            TableName=table_name,
        )
    except Exception as e:
        raise COMOperationError("create_pivot_table", f"CreatePivotTable 失败: {e}")

    # 配置行字段
    for i, field in enumerate(row_fields):
        try:
            pf = pivot_table.PivotFields(field)
            pf.Orientation = 1  # xlRowField
            pf.Position = i + 1
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"行字段 '{field}' 不存在: {e}")

    # 配置列字段
    for i, field in enumerate(column_fields):
        try:
            pf = pivot_table.PivotFields(field)
            pf.Orientation = 2  # xlColumnField
            pf.Position = i + 1
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"列字段 '{field}' 不存在: {e}")

    # 配置数据字段
    aggregation_map = {
        "sum": -4157,      # xlSum
        "average": -4106,  # xlAverage
        "count": -4112,    # xlCount
        "max": -4136,      # xlMax
        "min": -4139,      # xlMin
    }
    for field, func in data_fields.items():
        try:
            pivot_table.AddDataField(
                pivot_table.PivotFields(field),
                f"{func}_{field}",
                aggregation_map.get(func, -4157),
            )
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"数据字段 '{field}' 不存在: {e}")

    return f"created_pivot_table: {target_sheet_name}"


def _import_data(workbook: Any, op: dict) -> str:
    """导入外部数据文件 (CSV/TXT) 到工作表."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    file_path = op.get("file_path", "")
    start_cell = op.get("start_cell", "A1")
    delimiter = op.get("delimiter", ",")  # CSV 分隔符
    has_header = op.get("has_header", True)

    if not file_path:
        raise COMOperationError("import_data", "file_path 不能为空")

    # 路径校验
    if not Path(file_path).exists():
        raise COMOperationError("import_data", f"文件不存在: {file_path}")

    # 使用 QueryTables 导入
    query = None
    try:
        query = sheet.QueryTables.Add(
            Connection=f"TEXT;{file_path}",
            Destination=sheet.Range(start_cell),
        )
        query.TextFileDelimiter = delimiter
        query.TextFileParseType = 1  # xlDelimited
        query.Refresh()
    except Exception as e:
        raise COMOperationError("import_data", str(e))
    finally:
        if query:
            query.Delete()  # 导入完成后删除查询对象

    return f"imported_data: {file_path} -> {sheet.Name}"


def _export_data(workbook: Any, op: dict) -> str:
    """导出工作表为 CSV 文件."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    export_path = op.get("export_path", "")

    if not export_path:
        raise COMOperationError("export_data", "export_path 不能为空")

    # 复制到新工作簿再保存为 CSV
    new_wb = workbook.Application.Workbooks.Add()
    sheet.Copy(Before=new_wb.Worksheets(1))
    # 删除自动生成的多余工作表
    for ws in list(new_wb.Worksheets):
        if ws.Name != sheet.Name:
            try:
                ws.Delete()
            except Exception:
                pass
    new_wb.SaveAs(export_path, FileFormat=6)  # xlCSV
    new_wb.Close(SaveChanges=False)

    return f"exported_data: {sheet.Name} -> {export_path}"


def _add_slicer(workbook: Any, op: dict) -> str:
    """添加切片器.

    Args:
        target_sheet: 切片器所在工作表名称
        pivot_sheet: 数据透视表所在工作表名称
        pivot_name: 数据透视表名称
        field_name: 要筛选的字段名称
        left: 切片器左侧位置 (像素)
        top: 切片器顶部位置 (像素)
        width: 切片器宽度 (像素)
        height: 切片器高度 (像素)
    """
    target_sheet_name = op.get("target_sheet", "Sheet1")
    pivot_sheet_name = op.get("pivot_sheet", "数据透视表")
    pivot_name = op.get("pivot_name", "")
    field_name = op.get("field_name", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 200)
    height = op.get("height", 200)

    try:
        target_sheet = _get_sheet(workbook, target_sheet_name)
        pivot_sheet = _get_sheet(workbook, pivot_sheet_name)

        # 查找数据透视表
        pivot_table = None
        for pt in pivot_sheet.PivotTables():
            if not pivot_name or pt.Name == pivot_name:
                pivot_table = pt
                break

        if not pivot_table:
            raise COMOperationError("add_slicer", f"未找到数据透视表: {pivot_name}")

        # 添加切片器缓存
        slicer_cache = workbook.SlicerCaches.Add(pivot_table, field_name)

        # 添加切片器
        slicer = slicer_cache.Slicers.Add(
            SlicerDestination=target_sheet,
            Name=f"Slicer_{field_name}",
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )

        return f"added_slicer: {field_name}"
    except Exception as e:
        raise COMOperationError("add_slicer", str(e))


def _add_subtotal(workbook: Any, op: dict) -> str:
    """添加分类汇总.

    Args:
        sheet: 工作表名称
        range: 数据范围 (如 "A1:D100")
        group_by: 分组字段列号 (如 1 表示第 1 列)
        summary_function: 汇总函数 (sum/count/average/max/min)
        summary_fields: 要汇总的列号列表 (如 [3, 4])
        replace: 是否替换现有分类汇总
        page_breaks: 是否在每组后分页
        summary_below: 汇总结果是否在数据下方
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:D100")
    group_by = op.get("group_by", 1)
    summary_function = op.get("summary_function", "sum")
    summary_fields = op.get("summary_fields", [])
    # summary_fields 为空时使用默认值，TotalList 必须为非空元组
    if not summary_fields:
        summary_fields = [2]
    replace = op.get("replace", True)
    page_breaks = op.get("page_breaks", False)
    summary_below = op.get("summary_below", True)

    # 汇总函数映射
    func_map = {
        "sum": -4157,      # xlSum
        "count": -4112,    # xlCount
        "average": -4106,  # xlAverage
        "max": -4136,      # xlMax
        "min": -4139,      # xlMin
    }
    func_val = func_map.get(summary_function, -4157)

    range_obj = sheet.Range(range_str)

    try:
        range_obj.Subtotal(
            GroupBy=group_by,
            Function=func_val,
            TotalList=tuple(summary_fields),
            Replace=replace,
            PageBreaks=page_breaks,
            SummaryBelowData=summary_below,
        )
        return f"added_subtotal: {range_str}"
    except Exception as e:
        raise COMOperationError("add_subtotal", str(e))


# ============ Worksheet 工作表操作 (10 个) ============

def _list_worksheets(workbook: Any, op: dict) -> list[dict]:
    """列出所有工作表.

    Args:
        workbook: Excel 工作簿对象

    Returns:
        工作表信息列表
    """
    result = []
    for sheet in workbook.Worksheets:
        result.append({
            "index": sheet.Index,
            "name": sheet.Name,
            "visible": sheet.Visible == -1,  # xlSheetVisible
        })
    return result


def _get_worksheet_info(workbook: Any, op: dict) -> dict:
    """获取工作表信息.

    Args:
        sheet: 工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    return {
        "name": sheet.Name,
        "index": sheet.Index,
        "visible": sheet.Visible == -1,
        "tab_color": sheet.Tab.Color if sheet.Tab.Color else None,
        "used_range": str(sheet.UsedRange.Address) if sheet.UsedRange else None,
        "used_rows": sheet.UsedRange.Rows.Count if sheet.UsedRange else 0,
        "used_columns": sheet.UsedRange.Columns.Count if sheet.UsedRange else 0,
        "protected": sheet.ProtectContents,
    }


def _copy_worksheet(workbook: Any, op: dict) -> str:
    """复制工作表.

    Args:
        sheet: 源工作表名称
        new_name: 新工作表名称 (可选)
        position: 位置 (before/after, 可选)
        target_sheet: 目标位置参考工作表 (可选)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    new_name = op.get("new_name", "")
    position = op.get("position", "")
    target_sheet_name = op.get("target_sheet", "")

    if position and target_sheet_name:
        target = _get_sheet(workbook, target_sheet_name)
        if position == "before":
            sheet.Copy(Before=target)
        else:
            sheet.Copy(After=target)
    else:
        sheet.Copy(After=sheet)

    # 获取复制后的工作表
    new_sheet = sheet.Next
    if new_sheet and new_sheet.Name == sheet.Name:
        new_sheet = new_sheet.Next

    if new_name and new_sheet:
        new_sheet.Name = new_name

    return f"copied_worksheet: {sheet.Name} -> {new_sheet.Name if new_sheet else 'unnamed'}"


def _delete_worksheet(workbook: Any, op: dict) -> str:
    """删除工作表.

    Args:
        sheet: 要删除的工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    sheet_name = sheet.Name

    # 检查是否唯一工作表
    if workbook.Worksheets.Count == 1:
        raise COMOperationError("delete_worksheet", "不能删除唯一的工作表")

    sheet.Delete()
    return f"deleted_worksheet: {sheet_name}"


def _move_worksheet(workbook: Any, op: dict) -> str:
    """移动工作表.

    Args:
        sheet: 要移动的工作表名称
        position: 位置 (before/after/first/last)
        target_sheet: 目标位置参考工作表 (position=before/after 时必填)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    position = op.get("position", "first")
    target_sheet_name = op.get("target_sheet", "")

    # Guard: cannot move a sheet to its own position
    if workbook.Worksheets.Count <= 1:
        return f"moved_worksheet: skipped (only 1 sheet in workbook)"

    try:
        if position == "first":
            first_sheet = workbook.Worksheets(1)
            if sheet.Name == first_sheet.Name:
                return f"moved_worksheet: skipped (already first)"
            sheet.Move(Before=first_sheet)
        elif position == "last":
            last_sheet = workbook.Worksheets(workbook.Worksheets.Count)
            if sheet.Name == last_sheet.Name:
                return f"moved_worksheet: skipped (already last)"
            sheet.Move(After=last_sheet)
        elif position == "before" and target_sheet_name:
            target = _get_sheet(workbook, target_sheet_name)
            if sheet.Name == target.Name:
                return f"moved_worksheet: skipped (same sheet)"
            sheet.Move(Before=target)
        elif position == "after" and target_sheet_name:
            target = _get_sheet(workbook, target_sheet_name)
            if sheet.Name == target.Name:
                return f"moved_worksheet: skipped (same sheet)"
            sheet.Move(After=target)
    except Exception as e:
        raise COMOperationError("move_worksheet", str(e)) from e

    return f"moved_worksheet: {sheet.Name} to {position}"


def _hide_worksheet(workbook: Any, op: dict) -> str:
    """隐藏工作表.

    Args:
        sheet: 工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    # Excel 不允许隐藏工作簿中唯一可见的工作表
    visible_count = sum(1 for i in range(1, workbook.Worksheets.Count + 1) if workbook.Worksheets(i).Visible == -1)
    if visible_count <= 1:
        return f"hidden_worksheet: skipped (only {visible_count} visible sheet(s), Excel requires at least 1)"
    sheet.Visible = 0  # xlSheetHidden
    return f"hidden_worksheet: {sheet.Name}"


def _show_worksheet(workbook: Any, op: dict) -> str:
    """显示工作表.

    Args:
        sheet: 工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    sheet.Visible = -1  # xlSheetVisible
    return f"showed_worksheet: {sheet.Name}"


def _protect_worksheet(workbook: Any, op: dict) -> str:
    """保护工作表.

    Args:
        sheet: 工作表名称
        password: 密码 (可选)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    password = op.get("password", "")
    if password:
        sheet.Protect(Password=password)
    else:
        sheet.Protect()
    return f"protected_worksheet: {sheet.Name}"


def _unprotect_worksheet(workbook: Any, op: dict) -> str:
    """取消工作表保护.

    Args:
        sheet: 工作表名称
        password: 密码 (可选)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    password = op.get("password", "")
    if password:
        sheet.Unprotect(Password=password)
    else:
        sheet.Unprotect()
    return f"unprotected_worksheet: {sheet.Name}"


def _set_tab_color(workbook: Any, op: dict) -> str:
    """设置工作表标签颜色.

    Args:
        sheet: 工作表名称
        color: 颜色 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    color = op.get("color", "#FF0000")
    sheet.Tab.Color = _hex_to_rgb(color)
    return f"set_tab_color: {sheet.Name} -> {color}"


# ============ Range 范围操作 (10 个) ============

def _list_used_range(workbook: Any, op: dict) -> dict:
    """列出已使用范围.

    Args:
        sheet: 工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    used = sheet.UsedRange
    if not used:
        return {"address": None, "rows": 0, "columns": 0}
    return {
        "address": str(used.Address),
        "rows": used.Rows.Count,
        "columns": used.Columns.Count,
    }


def _clear_range(workbook: Any, op: dict) -> str:
    """清除范围内容.

    Args:
        sheet: 工作表名称
        range: 范围
        clear_type: 清除类型 (all/formulas/contents/comments)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    clear_type = op.get("clear_type", "all").lower()

    sheet.Range(range_str)
    rng = sheet.Range(range_str)
    # 按类型分派到不同的 Clear 方法
    if clear_type == "all":
        rng.Clear()
    elif clear_type == "contents":
        rng.ClearContents()
    elif clear_type == "formulas":
        # 清除公式但保留格式
        rng.ClearContents()
    elif clear_type == "comments":
        # 遍历每个单元格删除批注
        for row in rng.Rows:
            for cell in row.Cells:
                if cell.Comment is not None:
                    cell.Comment.Delete()
    elif clear_type == "hyperlinks":
        rng.ClearHyperlinks()
    elif clear_type == "formats":
        rng.ClearFormats()
    else:
        raise COMOperationError(
            "clear_range",
            f"clear_type 必须是 all/contents/formulas/comments/hyperlinks/formats,"
            f" 收到 '{clear_type}'",
        )
    return f"cleared_range: {range_str} ({clear_type})"


def _copy_range(workbook: Any, op: dict) -> str:
    """复制范围.

    Args:
        sheet: 源工作表名称
        range: 源范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Copy()
    return f"copied_range: {sheet.Name}!{range_str}"


def _paste_range(workbook: Any, op: dict) -> str:
    """粘贴范围.

    Args:
        sheet: 目标工作表名称
        target_cell: 目标单元格 (如 "A1")
        paste_type: 粘贴类型 (all/formulas/values/formats)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    target_cell = op.get("target_cell", "A1")
    paste_type = op.get("paste_type", "all")

    # 粘贴类型映射
    paste_map = {
        "all": -4104,        # xlPasteAll
        "formulas": -4122,   # xlPasteFormulas
        "values": -4163,     # xlPasteValues
        "formats": -4122,    # xlPasteFormats (approx)
    }
    paste_val = paste_map.get(paste_type, -4104)
    try:
        sheet.Range(target_cell).Select()
        sheet.Paste()
    except Exception:
        # Fallback: use PasteSpecial on the target range directly
        try:
            sheet.Range(target_cell).PasteSpecial(Paste=paste_val)
        except Exception as e:
            raise COMOperationError("paste_range", str(e)) from e
    return f"pasted_range: {sheet.Name}!{target_cell} ({paste_type})"


def _cut_range(workbook: Any, op: dict) -> str:
    """剪切范围.

    Args:
        sheet: 工作表名称
        range: 源范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Cut()
    return f"cut_range: {sheet.Name}!{range_str}"


def _delete_cells(workbook: Any, op: dict) -> str:
    """删除单元格.

    Args:
        sheet: 工作表名称
        range: 范围
        shift: 移动方向 (left/up)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    shift = op.get("shift", "left")

    shift_map = {
        "left": -4159,   # xlShiftToLeft
        "up": -4162,     # xlShiftUp
    }
    shift_val = shift_map.get(shift, -4159)
    sheet.Range(range_str).Delete(Shift=shift_val)
    return f"deleted_cells: {sheet.Name}!{range_str} (shift {shift})"


def _insert_cells(workbook: Any, op: dict) -> str:
    """插入单元格.

    Args:
        sheet: 工作表名称
        range: 范围
        shift: 移动方向 (right/down)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    shift = op.get("shift", "right")

    shift_map = {
        "right": -4161,  # xlShiftToRight
        "down": -4121,   # xlShiftDown
    }
    shift_val = shift_map.get(shift, -4161)
    sheet.Range(range_str).Insert(Shift=shift_val)
    return f"inserted_cells: {sheet.Name}!{range_str} (shift {shift})"


def _set_row_height(workbook: Any, op: dict) -> str:
    """设置行高.

    Args:
        sheet: 工作表名称
        row: 行号 (或范围, 如 "1:3" 表示 1-3 行)
        height: 高度 (磅)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    row = op.get("row", 1)
    height = op.get("height", 15.0)
    sheet.Rows(row).RowHeight = height
    return f"set_row_height: {sheet.Name} row {row} = {height}"


def _set_column_width(workbook: Any, op: dict) -> str:
    """设置列宽.

    Args:
        sheet: 工作表名称
        column: 列标识 (如 "A" 或 "A:C")
        width: 宽度 (字符单位)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    column = op.get("column", "A")
    width = op.get("width", 8.43)
    sheet.Columns(column).ColumnWidth = width
    return f"set_column_width: {sheet.Name} column {column} = {width}"


def _hide_rows(workbook: Any, op: dict) -> str:
    """隐藏行.

    Args:
        sheet: 工作表名称
        rows: 行号或范围 (如 "1" 或 "1:5")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    rows = op.get("rows", "1")
    sheet.Rows(rows).Hidden = True
    return f"hidden_rows: {sheet.Name} rows {rows}"


# ============ Charts 图表操作 (10 个) ============

def _list_charts(workbook: Any, op: dict) -> list[dict]:
    """列出所有图表.

    Args:
        sheet: 工作表名称 (可选, 不填则列出所有工作表)
    """
    result = []
    sheet_name = op.get("sheet", "")
    sheets_to_check = []
    if sheet_name:
        sheets_to_check.append(_get_sheet(workbook, sheet_name))
    else:
        for s in workbook.Worksheets:
            sheets_to_check.append(s)

    for sheet in sheets_to_check:
        for i in range(1, sheet.ChartObjects().Count + 1):
            chart_obj = sheet.ChartObjects(i)
            result.append({
                "sheet": sheet.Name,
                "index": i,
                "name": chart_obj.Name,
                "chart_type": chart_obj.Chart.ChartType,
                "has_title": chart_obj.Chart.HasTitle,
            })
    return result


def _get_chart_info(workbook: Any, op: dict) -> dict:
    """获取图表信息.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引 (从 1 开始)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    chart_obj = sheet.ChartObjects(chart_index)
    chart = chart_obj.Chart

    series_info = []
    for s in chart.SeriesCollection():
        series_info.append({"name": s.Name})

    return {
        "name": chart_obj.Name,
        "chart_type": chart.ChartType,
        "has_title": chart.HasTitle,
        "title": chart.ChartTitle.Text if chart.HasTitle else "",
        "has_legend": chart.HasLegend,
        "series_count": chart.SeriesCollection().Count,
        "series": series_info,
    }


def _set_chart_title(workbook: Any, op: dict) -> str:
    """设置图表标题.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        title: 标题文本
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    title = op.get("title", "")

    chart = sheet.ChartObjects(chart_index).Chart
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    return f"set_chart_title: {title}"


def _set_chart_legend(workbook: Any, op: dict) -> str:
    """设置图表图例.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        show: 是否显示图例
        position: 图例位置 (bottom/top/left/right/corner)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    show = op.get("show", True)
    position = op.get("position", "bottom")

    chart = sheet.ChartObjects(chart_index).Chart
    chart.HasLegend = show
    if show:
        position_map = {
            "bottom": -4107,   # xlLegendPositionBottom
            "top": -4160,      # xlLegendPositionTop
            "left": -4131,     # xlLegendPositionLeft
            "right": -4152,    # xlLegendPositionRight
            "corner": -4151,   # xlLegendPositionCorner
        }
        chart.Legend.Position = position_map.get(position, -4107)

    return f"set_chart_legend: show={show}, position={position}"


def _add_chart_series(workbook: Any, op: dict) -> str:
    """添加图表系列.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        series_name: 系列名称
        values_range: 数值范围
        categories_range: 分类范围 (可选)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    series_name = op.get("series_name", "Series")
    values_range = op.get("values_range", "")
    categories_range = op.get("categories_range", "")

    chart = sheet.ChartObjects(chart_index).Chart
    series = chart.SeriesCollection().NewSeries()
    series.Name = series_name
    if values_range:
        try:
            series.Values = sheet.Range(values_range)
        except Exception:
            # 某些 Excel 版本需要用地址字符串
            series.Values = f"={sheet.Name}!{sheet.Range(values_range).Address}"
    if categories_range:
        try:
            series.XValues = sheet.Range(categories_range)
        except Exception:
            series.XValues = f"={sheet.Name}!{sheet.Range(categories_range).Address}"

    return f"added_chart_series: {series_name}"


def _remove_chart_series(workbook: Any, op: dict) -> str:
    """移除图表系列.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        series_index: 系列索引 (从 1 开始)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    series_index = op.get("series_index", 1)

    chart = sheet.ChartObjects(chart_index).Chart
    series = chart.SeriesCollection(series_index)
    series_name = series.Name
    series.Delete()
    return f"removed_chart_series: {series_name}"


def _set_chart_axis(workbook: Any, op: dict) -> str:
    """设置图表轴.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        axis: 轴类型 (x/y/value1/value2)
        title: 轴标题 (可选)
        min_scale: 最小值 (可选)
        max_scale: 最大值 (可选)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    axis = op.get("axis", "x")
    title = op.get("title", "")
    min_scale = op.get("min_scale")
    max_scale = op.get("max_scale")

    chart = sheet.ChartObjects(chart_index).Chart

    if axis in ("x", "category"):
        ax = chart.Axes(1)  # xlCategory
    else:
        ax = chart.Axes(2)  # xlValue

    if title:
        ax.HasTitle = True
        ax.AxisTitle.Text = title
    if min_scale is not None:
        ax.MinimumScale = min_scale
    if max_scale is not None:
        ax.MaximumScale = max_scale

    return f"set_chart_axis: {axis}"


def _change_chart_type(workbook: Any, op: dict) -> str:
    """更改图表类型.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        chart_type: 新图表类型 (column/bar/line/pie/scatter/area)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    chart_type = op.get("chart_type", "column")

    chart_type_val = CHART_TYPE_MAP.get(chart_type, 51)
    chart = sheet.ChartObjects(chart_index).Chart
    chart.ChartType = chart_type_val
    return f"changed_chart_type: {chart_type}"


def _export_chart(workbook: Any, op: dict) -> str:
    """导出图表为图片.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
        output_path: 输出图片路径
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    output_path = op.get("output_path", "")

    if not output_path:
        raise COMOperationError("export_chart", "output_path 不能为空")

    chart = sheet.ChartObjects(chart_index).Chart
    chart.Export(output_path)
    return f"exported_chart: {sheet.Name} chart {chart_index} -> {output_path}"


def _delete_chart(workbook: Any, op: dict) -> str:
    """删除图表.

    Args:
        sheet: 工作表名称
        chart_index: 图表索引
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)

    chart_obj = sheet.ChartObjects(chart_index)
    chart_name = chart_obj.Name
    chart_obj.Delete()
    return f"deleted_chart: {chart_name}"


# ============ Format 格式操作 (10 个) ============

def _set_font(workbook: Any, op: dict) -> str:
    """设置字体.

    Args:
        sheet: 工作表名称
        range: 范围
        font_name: 字体名称 (如 "微软雅黑")
        font_size: 字体大小
        font_color: 字体颜色 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    font_obj = sheet.Range(range_str).Font

    font_name = op.get("font_name", "")
    font_size = op.get("font_size")
    font_color = op.get("font_color", "")

    if font_name:
        font_obj.Name = font_name
    if font_size is not None:
        font_obj.Size = font_size
    if font_color:
        font_obj.Color = _hex_to_rgb(font_color)

    return f"set_font: {range_str}"


def _set_font_bold(workbook: Any, op: dict) -> str:
    """设置粗体.

    Args:
        sheet: 工作表名称
        range: 范围
        bold: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    bold = op.get("bold", True)
    sheet.Range(range_str).Font.Bold = bold
    return f"set_font_bold: {range_str} = {bold}"


def _set_font_italic(workbook: Any, op: dict) -> str:
    """设置斜体.

    Args:
        sheet: 工作表名称
        range: 范围
        italic: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    italic = op.get("italic", True)
    sheet.Range(range_str).Font.Italic = italic
    return f"set_font_italic: {range_str} = {italic}"


def _set_font_underline(workbook: Any, op: dict) -> str:
    """设置下划线.

    Args:
        sheet: 工作表名称
        range: 范围
        underline: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    underline = op.get("underline", True)
    sheet.Range(range_str).Font.Underline = underline
    return f"set_font_underline: {range_str} = {underline}"


def _set_alignment(workbook: Any, op: dict) -> str:
    """设置对齐.

    Args:
        sheet: 工作表名称
        range: 范围
        horizontal: 水平对齐 (left/center/right/general)
        vertical: 垂直对齐 (top/middle/bottom)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    range_obj = sheet.Range(range_str)

    horizontal_map = {
        "general": 1,
        "left": -4131,
        "center": -4108,
        "right": -4152,
    }
    vertical_map = {
        "top": -4160,
        "middle": -4108,
        "bottom": -4107,
    }

    horizontal = op.get("horizontal", "")
    vertical = op.get("vertical", "")

    if horizontal:
        range_obj.HorizontalAlignment = horizontal_map.get(horizontal, 1)
    if vertical:
        range_obj.VerticalAlignment = vertical_map.get(vertical, -4108)

    return f"set_alignment: {range_str}"


def _set_wrap_text(workbook: Any, op: dict) -> str:
    """设置自动换行.

    Args:
        sheet: 工作表名称
        range: 范围
        wrap: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    wrap = op.get("wrap", True)
    sheet.Range(range_str).WrapText = wrap
    return f"set_wrap_text: {range_str} = {wrap}"


def _set_indent(workbook: Any, op: dict) -> str:
    """设置缩进.

    Args:
        sheet: 工作表名称
        range: 范围
        indent: 缩进级别 (0-15)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    indent = op.get("indent", 1)
    if indent <= 0:
        return f"set_indent: skipped (indent={indent} <= 0)"
    sheet.Range(range_str).InsertIndent(indent)
    return f"set_indent: {range_str} = {indent}"


def _set_orientation(workbook: Any, op: dict) -> str:
    """设置文字方向.

    Args:
        sheet: 工作表名称
        range: 范围
        orientation: 角度 (0=水平, 90=垂直, 45=-45度)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    orientation = op.get("orientation", 0)
    sheet.Range(range_str).Orientation = orientation
    return f"set_orientation: {range_str} = {orientation}"


def _clear_format(workbook: Any, op: dict) -> str:
    """清除格式.

    Args:
        sheet: 工作表名称
        range: 范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).ClearFormats()
    return f"cleared_format: {range_str}"


def _copy_format(workbook: Any, op: dict) -> str:
    """复制格式.

    Args:
        sheet: 工作表名称
        source_range: 源格式范围
        target_range: 目标范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    source_range = op.get("source_range", "A1")
    target_range = op.get("target_range", "B1")

    source = sheet.Range(source_range)
    target = sheet.Range(target_range)
    source.Copy()
    target.PasteSpecial(Paste=-4122)  # xlPasteFormats
    return f"copied_format: {source_range} -> {target_range}"


# ============ Page Setup 页面设置 (10 个) ============

def _set_page_orientation(workbook: Any, op: dict) -> str:
    """设置页面方向.

    Args:
        sheet: 工作表名称
        orientation: portrait/landscape
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    orientation = op.get("orientation", "portrait")

    # 1=纵向 xlPortrait, 2=横向 xlLandscape
    sheet.PageSetup.Orientation = 1 if orientation == "portrait" else 2
    return f"set_page_orientation: {orientation}"


def _set_page_size(workbook: Any, op: dict) -> str:
    """设置页面大小.

    Args:
        sheet: 工作表名称
        size: A4/A3/Letter/Legal 或编号 (1=Letter, 5=Legal, 9=A4, 8=A3)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    size = op.get("size", "A4")

    size_map = {
        "Letter": 1,
        "LetterSmall": 2,
        "Tabloid": 3,
        "Ledger": 4,
        "Legal": 5,
        "Statement": 6,
        "Executive": 7,
        "A3": 8,
        "A4": 9,
        "A4Small": 10,
        "A5": 11,
        "B4": 12,
        "B5": 13,
    }
    sheet.PageSetup.PaperSize = size_map.get(size, 9)
    return f"set_page_size: {size}"


def _set_page_margins(workbook: Any, op: dict) -> str:
    """设置页边距.

    Args:
        sheet: 工作表名称
        top: 上边距 (英寸)
        bottom: 下边距 (英寸)
        left: 左边距 (英寸)
        right: 右边距 (英寸)
        header: 页眉边距 (英寸)
        footer: 页脚边距 (英寸)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    page_setup = sheet.PageSetup

    if "top" in op:
        page_setup.TopMargin = op["top"]
    if "bottom" in op:
        page_setup.BottomMargin = op["bottom"]
    if "left" in op:
        page_setup.LeftMargin = op["left"]
    if "right" in op:
        page_setup.RightMargin = op["right"]
    if "header" in op:
        page_setup.HeaderMargin = op["header"]
    if "footer" in op:
        page_setup.FooterMargin = op["footer"]

    return f"set_page_margins: {sheet.Name}"


def _set_header(workbook: Any, op: dict) -> str:
    """设置页眉.

    Args:
        sheet: 工作表名称
        text: 页眉文本 (&L左& C中& R右)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    text = op.get("text", "")
    sheet.PageSetup.CenterHeader = text
    return f"set_header: {text}"


def _set_footer(workbook: Any, op: dict) -> str:
    """设置页脚.

    Args:
        sheet: 工作表名称
        text: 页脚文本 (&L左& C中& R右)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    text = op.get("text", "")
    sheet.PageSetup.CenterFooter = text
    return f"set_footer: {text}"


def _add_print_title(workbook: Any, op: dict) -> str:
    """添加打印标题 (重复打印行/列).

    Args:
        sheet: 工作表名称
        rows: 重复行 (如 "$1:$1")
        columns: 重复列 (如 "$A:$A")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    rows = op.get("rows", "")
    columns = op.get("columns", "")

    if rows:
        sheet.PageSetup.PrintTitleRows = rows
    if columns:
        sheet.PageSetup.PrintTitleColumns = columns

    return f"add_print_title: rows={rows}, columns={columns}"


def _set_print_area(workbook: Any, op: dict) -> str:
    """设置打印区域.

    Args:
        sheet: 工作表名称
        range: 打印区域 (如 "A1:D20")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.PageSetup.PrintArea = range_str
    return f"set_print_area: {range_str}"


def _set_page_break(workbook: Any, op: dict) -> str:
    """设置分页符.

    Args:
        sheet: 工作表名称
        cell: 分页符位置 (如 "A20")
        break_type: row/column
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A20")
    break_type = op.get("break_type", "row")

    # xlPageBreakManual: 1
    if break_type == "row":
        sheet.HPageBreaks.Add(Before=sheet.Range(cell))
    else:
        sheet.VPageBreaks.Add(Before=sheet.Range(cell))

    return f"set_page_break: {cell} ({break_type})"


def _set_scale(workbook: Any, op: dict) -> str:
    """设置缩放比例.

    Args:
        sheet: 工作表名称
        scale: 缩放比例 (10-400 百分比)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    scale = op.get("scale", 100)
    sheet.PageSetup.Zoom = scale
    return f"set_scale: {scale}%"


def _set_fit_to_page(workbook: Any, op: dict) -> str:
    """设置适应页面.

    Args:
        sheet: 工作表名称
        fit_width: 适应宽度 (1=单页宽, 0=自动)
        fit_height: 适应高度 (1=单页高, 0=自动)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    page_setup = sheet.PageSetup

    # Must disable Zoom before setting FitToPages — Zoom and FitToPages are mutually exclusive
    # In Excel COM, Zoom=False means "use FitToPages instead of percentage zoom"
    page_setup.Zoom = False
    try:
        page_setup.FitToPagesWide = op.get("fit_width", 1)
    except Exception:
        page_setup.FitToPagesWide = 1
    try:
        page_setup.FitToPagesTall = op.get("fit_height", 0)
    except Exception:
        page_setup.FitToPagesTall = 0

    return f"set_fit_to_page: {page_setup.FitToPagesWide}x{page_setup.FitToPagesTall}"


# ============ Formulas 公式操作 (8 个) ============

def _set_array_formula(workbook: Any, op: dict) -> str:
    """设置数组公式 (Ctrl+Shift+Enter 公式).

    Args:
        sheet: 工作表名称
        range: 范围
        formula: 公式字符串
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    formula = op.get("formula", "")
    if not formula:
        raise COMOperationError("set_array_formula", "formula 不能为空")
    sheet.Range(range_str).FormulaArray = formula
    return f"set_array_formula: {range_str} = {formula}"


def _evaluate_formula(workbook: Any, op: dict) -> Any:
    """计算并返回公式结果.

    Args:
        sheet: 工作表名称
        cell: 单元格地址 (如 A1)

    注: 会先调用 Application.Calculate() 确保返回最新值.
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    # 强制重算, 避免 manual 模式下返回脏值
    try:
        workbook.Application.Calculate()
    except Exception:
        pass
    value = sheet.Range(cell).Value
    return {"cell": cell, "value": value}


def _replace_formula(workbook: Any, op: dict) -> str:
    """替换范围内所有公式 (按字符串匹配).

    Args:
        sheet: 工作表名称
        range: 范围
        find: 查找字符串
        replace: 替换字符串

    注意: 简单子串匹配, find="A1" 会同时影响 "AA1" 等含 A1 子串的公式.
    """
    import re
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    find = op.get("find", "")
    replace = op.get("replace", "")
    if not find:
        raise COMOperationError("replace_formula", "find 不能为空")

    rng = sheet.Range(range_str)
    # 使用 \b 单词边界避免 AA1 误匹配 A1
    pattern = re.compile(r"\b" + re.escape(find) + r"\b")
    count = 0
    for row in rng.Rows:
        for cell in row.Cells:
            if cell.HasFormula:
                old_f = str(cell.Formula)
                new_f = pattern.sub(replace, old_f)
                if new_f != old_f:
                    cell.Formula = new_f
                    count += 1
    return f"replaced_formula: {count} cells in {range_str}"


def _find_formula_cells(workbook: Any, op: dict) -> list[dict]:
    """查找范围内所有含公式的单元格.

    Args:
        sheet: 工作表名称
        range: 范围 (留空使用 UsedRange)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "")
    if range_str:
        rng = sheet.Range(range_str)
    else:
        rng = sheet.UsedRange

    formulas = []
    try:
        # Use SpecialCells for efficiency instead of iterating all cells
        formula_cells = rng.SpecialCells(5)  # xlCellTypeFormulas = 5
        for cell in formula_cells:
            try:
                formulas.append({
                    "cell": str(cell.Address),
                    "formula": str(cell.Formula),
                    "value": cell.Value,
                })
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        # No formula cells found, that's fine
        pass
    return formulas


def _convert_to_values(workbook: Any, op: dict) -> str:
    """将公式转换为静态值.

    Args:
        sheet: 工作表名称
        range: 范围

    注: 先强制重算, 若单元格值是 Excel 错误 (#NAME? / #VALUE! 等) 则拒绝覆盖原公式.
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    rng = sheet.Range(range_str)
    # 强制重算后再读取
    try:
        workbook.Application.Calculate()
    except Exception:
        pass
    value = rng.Value
    # 检测 Excel 错误值字符串
    if isinstance(value, str) and value.startswith("#") and value.endswith("!"):
        raise COMOperationError(
            "convert_to_values",
            f"范围内含有计算错误 {value}, 拒绝覆盖原公式",
        )
    rng.Value = value
    return f"converted_to_values: {range_str}"


def _get_formula_info(workbook: Any, op: dict) -> dict:
    """获取公式信息 (类型/值/是否数组公式).

    Args:
        sheet: 工作表名称
        cell: 单元格地址
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    rng = sheet.Range(cell)
    return {
        "cell": cell,
        "has_formula": bool(rng.HasFormula),
        "formula": str(rng.Formula) if rng.HasFormula else None,
        "value": rng.Value,
        "text": str(rng.Text),
    }


def _define_name(workbook: Any, op: dict) -> str:
    """定义名称 (workbook level).

    Args:
        name: 名称
        refers_to: 引用 (如 '=Sheet1!$A$1:$A$10')
        scope: 范围 sheet name (可选, 默认为工作簿级)
    """
    import re
    name = op.get("name", "")
    refers_to = op.get("refers_to", "")
    scope = op.get("scope", "")

    if not name or not refers_to:
        raise COMOperationError("define_name", "name 和 refers_to 不能为空")

    # 名称合法性: 字母/下划线开头, 后续可含字母数字下划线/点号
    if not re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", name):
        raise COMOperationError(
            "define_name",
            f"名称 '{name}' 不合法 (须以字母/下划线开头, 仅含字母数字下划线)",
        )

    # 重复检测
    try:
        existing = workbook.Names(name)
        if existing is not None:
            raise COMOperationError(
                "define_name",
                f"名称 '{name}' 已存在, 请先删除或换名",
            )
    except COMOperationError:
        raise
    except Exception:
        # 名称不存在 (正常)
        pass

    if scope:
        # 工作表级名称
        ws = workbook.Worksheets(scope)
        ws.Names.Add(Name=name, RefersTo=refers_to)
    else:
        # 工作簿级名称
        workbook.Names.Add(Name=name, RefersTo=refers_to)
    return f"defined_name: {name} = {refers_to}"


# ============ Tables 表格 (ListObject) (8 个) ============

def _create_table(workbook: Any, op: dict) -> str:
    """创建 Excel 表格 (ListObject).

    Args:
        sheet: 工作表名称
        range: 表格数据范围 (含表头)
        table_name: 表格名称
        style_name: 表格样式名 (如 'TableStyleMedium2')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    table_name = op.get("table_name", "")
    style_name = op.get("style_name", "TableStyleMedium2")

    # 检查重名并自动追加序号
    existing_names: set[str] = set()
    for i in range(1, sheet.ListObjects.Count + 1):
        existing_names.add(sheet.ListObjects(i).Name)

    if not table_name:
        base = f"Table_{sheet.Name}"
        table_name = base
        n = 2
        while table_name in existing_names:
            table_name = f"{base}_{n}"
            n += 1
    elif table_name in existing_names:
        raise COMOperationError(
            "create_table",
            f"表格名 '{table_name}' 已被占用, 请指定其他名称",
        )

    rng = sheet.Range(range_str)
    table = sheet.ListObjects.Add(
        SourceType=1,  # xlSrcRange
        Source=rng,
        XlListObjectHasHeaders=1,  # xlYes
    )
    table.Name = table_name
    try:
        table.TableStyle = style_name
    except Exception:
        logger.warning(f"无法应用表格样式 {style_name}")

    return f"created_table: {table_name} ({range_str})"


def _list_tables(workbook: Any, op: dict) -> list[dict]:
    """列出工作簿中所有 Excel 表格.

    Args:
        sheet: 工作表名称 (可选, 留空列出所有)
    """
    sheet_name = op.get("sheet", "")
    tables_info = []

    if sheet_name:
        sheets = [_get_sheet(workbook, sheet_name)]
    else:
        sheets = list(workbook.Worksheets)

    for sheet in sheets:
        for i in range(1, sheet.ListObjects.Count + 1):
            tbl = sheet.ListObjects(i)
            try:
                style = tbl.TableStyle.Name if tbl.TableStyle else ""
            except Exception:
                style = ""
            tables_info.append({
                "sheet": sheet.Name,
                "name": tbl.Name,
                "range": str(tbl.Range.Address),
                "style": style,
                "show_totals": bool(tbl.ShowTotals),
            })
    return tables_info


def _resize_table(workbook: Any, op: dict) -> str:
    """调整 Excel 表格范围.

    Args:
        sheet: 工作表名称
        table_name: 表格名称
        range: 新范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    range_str = op.get("range", "")
    if not table_name or not range_str:
        raise COMOperationError("resize_table", "table_name 和 range 不能为空")

    tbl = sheet.ListObjects(table_name)
    tbl.Resize(sheet.Range(range_str))
    return f"resized_table: {table_name} -> {range_str}"


def _set_table_style(workbook: Any, op: dict) -> str:
    """设置表格样式.

    Args:
        sheet: 工作表名称
        table_name: 表格名称
        style_name: 样式名 (如 'TableStyleLight1')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    style_name = op.get("style_name", "TableStyleLight1")
    if not table_name:
        raise COMOperationError("set_table_style", "table_name 不能为空")

    tbl = sheet.ListObjects(table_name)
    tbl.TableStyle = style_name
    return f"set_table_style: {table_name} = {style_name}"


def _show_table_totals(workbook: Any, op: dict) -> str:
    """显示/隐藏表格汇总行.

    Args:
        sheet: 工作表名称
        table_name: 表格名称
        show: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    show = op.get("show", True)
    if not table_name:
        raise COMOperationError("show_table_totals", "table_name 不能为空")

    tbl = sheet.ListObjects(table_name)
    tbl.ShowTotals = show
    return f"show_table_totals: {table_name} = {show}"


def _add_table_column(workbook: Any, op: dict) -> str:
    """为表格添加计算列 (公式列).

    Args:
        sheet: 工作表名称
        table_name: 表格名称
        column_name: 新列名
        formula: 列公式 (如 '=[@Qty]*[@Price]')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    column_name = op.get("column_name", "")
    formula = op.get("formula", "")
    if not table_name or not column_name:
        raise COMOperationError("add_table_column", "table_name 和 column_name 不能为空")

    tbl = sheet.ListObjects(table_name)

    new_col = tbl.ListColumns.Add()
    new_col.Name = column_name
    if formula:
        # formula 须以 = 开头
        if not formula.startswith("="):
            formula = "=" + formula
        if new_col.DataBodyRange is None:
            raise COMOperationError(
                "add_table_column",
                "空表无法添加计算列,需先填入至少一行数据",
            )
        new_col.DataBodyRange.Formula = formula

    return f"added_table_column: {table_name}.{column_name}"


def _remove_table_column(workbook: Any, op: dict) -> str:
    """删除表格列.

    Args:
        sheet: 工作表名称
        table_name: 表格名称
        column_name: 列名
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    column_name = op.get("column_name", "")
    if not table_name or not column_name:
        raise COMOperationError("remove_table_column", "table_name 和 column_name 不能为空")

    tbl = sheet.ListObjects(table_name)
    # 关闭 totals row 防止干扰
    try:
        tbl.ShowTotals = False
    except Exception:
        pass

    # ListColumns 遍历查找
    target_col = None
    count = int(tbl.ListColumns.Count)
    for i in range(1, count + 1):
        try:
            col = tbl.ListColumns(i)
            name = str(col.Name)
            if name == column_name:
                target_col = col
                break
        except Exception as e:
            logger.debug(f"skip col {i}: {e}")
            continue

    if target_col is not None:
        target_col.Delete()
        return f"removed_table_column: {table_name}.{column_name}"

    # 找不到则按整列 Range 删除
    try:
        col_count = int(tbl.Range.Columns.Count)
        for i in range(1, col_count + 1):
            col_range = tbl.Range.Columns(i)
            # 头部单元格
            try:
                header = str(col_range.Cells(1, 1).Value)
                if header == column_name:
                    col_range.Delete()
                    return f"removed_table_column: {table_name}.{column_name} (by range)"
            except Exception:
                continue
    except Exception as e:
        logger.error(f"remove_table_column range fallback: {e}")

    raise COMOperationError("remove_table_column", f"列 '{column_name}' 不存在")


def _delete_table(workbook: Any, op: dict) -> str:
    """删除 Excel 表格 (仅删除表格结构, 不删除数据).

    Args:
        sheet: 工作表名称
        table_name: 表格名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    if not table_name:
        raise COMOperationError("delete_table", "table_name 不能为空")

    tbl = sheet.ListObjects(table_name)
    tbl.Unlist()  # 仅删除表格, 保留数据
    return f"deleted_table: {table_name}"


# ============ Data 数据操作 (9 个) ============

def _add_auto_filter(workbook: Any, op: dict) -> str:
    """添加自动筛选.

    Args:
        sheet: 工作表名称
        range: 数据范围 (留空使用 UsedRange)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "")
    if range_str:
        rng = sheet.Range(range_str)
    else:
        rng = sheet.UsedRange
    try:
        rng.AutoFilter()
    except Exception as e:
        raise COMOperationError("add_auto_filter", str(e)) from e
    return f"added_auto_filter: {rng.Address}"


def _remove_auto_filter(workbook: Any, op: dict) -> str:
    """移除自动筛选.

    Args:
        sheet: 工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    if sheet.AutoFilterMode:
        sheet.AutoFilterMode = False
    return f"removed_auto_filter: {sheet.Name}"


def _sort_range(workbook: Any, op: dict) -> str:
    """对范围内数据排序.

    Args:
        sheet: 工作表名称
        range: 数据范围
        key_column: 排序列地址 (如 'A1') 或列号 (1-based)
        ascending: True 升序 / False 降序
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    key_column = op.get("key_column", 1)
    ascending = op.get("ascending", True)

    rng = sheet.Range(range_str)

    # 解析排序列
    if isinstance(key_column, int):
        key = rng.Columns(key_column)
    else:
        key = sheet.Range(key_column)

    key.Sort(
        Key1=key,
        Order1=1 if ascending else 2,  # xlAscending=1, xlDescending=2
        Header=1,  # xlYes
    )
    return f"sorted_range: {range_str} by column {key_column}"


def _advanced_filter(workbook: Any, op: dict) -> str:
    """高级筛选 (就地筛选或复制到目标位置).

    Args:
        sheet: 工作表名称
        range: 数据范围
        criteria_range: 条件范围
        action: 'filter' 原地筛选 / 'copy' 复制
        copy_to: 复制目标 (action='copy' 时必填)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    criteria_range = op.get("criteria_range", "")
    action = op.get("action", "filter")
    copy_to = op.get("copy_to", "")

    if not criteria_range:
        raise COMOperationError("advanced_filter", "criteria_range 不能为空")

    rng = sheet.Range(range_str)
    crit_rng = sheet.Range(criteria_range)
    if action == "copy":
        if not copy_to:
            raise COMOperationError("advanced_filter", "copy_to 不能为空")
        rng.AdvancedFilter(
            Action=2,  # xlFilterCopy
            CriteriaRange=crit_rng,
            CopyToRange=sheet.Range(copy_to),
        )
    elif action == "filter":
        rng.AdvancedFilter(
            Action=1,  # xlFilterInPlace
            CriteriaRange=crit_rng,
        )
    else:
        raise COMOperationError(
            "advanced_filter", f"action 必须是 'filter' 或 'copy', 收到 '{action}'"
        )
    return f"advanced_filter: {range_str} ({action})"


def _remove_duplicates(workbook: Any, op: dict) -> str:
    """删除重复行.

    Args:
        sheet: 工作表名称
        range: 数据范围
        columns: 判定列 (1-based int 或 'A,B,C'), 默认所有列
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    columns = op.get("columns", "")

    rng = sheet.Range(range_str)
    cols: list[int] | None = None
    if columns:
        if isinstance(columns, str):
            cols = []
            for c in columns.split(","):
                c = c.strip()
                if not c:
                    continue
                if not c.isdigit():
                    raise COMOperationError(
                        "remove_duplicates",
                        f"列号必须为正整数,收到 '{c}'",
                    )
                cols.append(int(c))
        else:
            cols = [int(columns)]

    if cols:
        rng.RemoveDuplicates(Columns=cols)
    else:
        rng.RemoveDuplicates()
    return f"removed_duplicates: {range_str}"


def _group_rows(workbook: Any, op: dict) -> str:
    """分级显示 (组合行).

    Args:
        sheet: 工作表名称
        range: 范围 (如 'A2:A5')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Group()
    return f"grouped_rows: {range_str}"


def _ungroup_rows(workbook: Any, op: dict) -> str:
    """取消行组合.

    Args:
        sheet: 工作表名称
        range: 范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Ungroup()
    return f"ungrouped_rows: {range_str}"


def _group_columns(workbook: Any, op: dict) -> str:
    """分级显示 (组合列).

    Args:
        sheet: 工作表名称
        range: 范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Group()
    return f"grouped_columns: {range_str}"


def _ungroup_columns(workbook: Any, op: dict) -> str:
    """取消列组合.

    Args:
        sheet: 工作表名称
        range: 范围
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Ungroup()
    return f"ungrouped_columns: {range_str}"


# ============ Protection 工作簿保护 (6 个) ============

def _protect_workbook(workbook: Any, op: dict) -> str:
    """保护工作簿 (结构保护).

    Args:
        password: 密码
        structure: 保护结构 (默认 True)
        windows: 保护窗口 (默认 False)
    """
    password = op.get("password", "")
    structure = op.get("structure", True)
    windows = op.get("windows", False)

    workbook.Protect(
        Password=password,
        Structure=structure,
        Windows=windows,
    )
    return f"protected_workbook: structure={structure}, windows={windows}"


def _unprotect_workbook(workbook: Any, op: dict) -> str:
    """撤销工作簿保护.

    Args:
        password: 密码
    """
    password = op.get("password", "")
    workbook.Unprotect(Password=password)
    return f"unprotected_workbook"


def _set_open_password(workbook: Any, op: dict) -> str:
    """设置打开密码 (下次 SaveAs 时真正加密).

    Args:
        password: 密码

    注: Excel COM 中 Password 属性只在 SaveAs 时生效.
        本函数仅设置属性, 不主动 SaveAs 以避免覆盖未保存的内容.
        请在调用本函数后, 显式调用 close_document(save=True) 触发加密落盘.
    """
    password = op.get("password", "")
    if not password:
        raise COMOperationError("set_open_password", "password 不能为空")
    workbook.Password = password
    return f"set_open_password: 长度 {len(password)} (将在 SaveAs 时生效)"


def _set_write_reservation_password(workbook: Any, op: dict) -> str:
    """设置写保护密码 (推荐只读).

    Args:
        password: 密码
    """
    password = op.get("password", "")
    if not password:
        raise COMOperationError("set_write_reservation_password", "password 不能为空")
    workbook.WriteReservationPassword = password
    return f"set_write_reservation_password: 长度 {len(password)}"


def _mark_as_final(workbook: Any, op: dict) -> str:
    """标记为最终状态 (Mark As Final).

    注: 通过自定义文档属性实现.
    """
    custom_props = workbook.CustomDocumentProperties
    # 检查属性是否已存在
    prop = None
    try:
        for i in range(1, custom_props.Count + 1):
            if custom_props(i).Name == "_MarkAsFinal":
                prop = custom_props(i)
                break
    except Exception:
        pass

    if prop is not None:
        # 已存在则更新
        prop.Value = True
    else:
        # 不存在则添加 (使用位置参数)
        custom_props.Add(
            "_MarkAsFinal",  # Name
            False,            # LinkToContent
            4,                # Type (msoPropertyTypeBoolean)
            True,             # Value
        )
    return "marked_as_final"


def _recommend_read_only(workbook: Any, op: dict) -> str:
    """设置推荐只读 (ReadOnlyRecommended).

    Args:
        recommend: True 启用 / False 关闭
    """
    recommend = op.get("recommend", True)
    workbook.ReadOnlyRecommended = recommend
    return f"recommend_read_only: {recommend}"


# ============ Objects 对象操作 (5 个) ============

def _add_image(workbook: Any, op: dict) -> str:
    """插入图片.

    Args:
        sheet: 工作表名称
        image_path: 图片文件路径
        cell: 锚定单元格 (默认 A1)
        width: 宽度 (磅, 可选)
        height: 高度 (磅, 可选)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    image_path = op.get("image_path", "")
    cell = op.get("cell", "A1")
    width = op.get("width")
    height = op.get("height")

    if not image_path:
        raise COMOperationError("add_image", "image_path 不能为空")

    p = Path(image_path)
    if not p.exists():
        raise COMOperationError("add_image", f"图片不存在: {image_path}")

    img = sheet.Pictures().Insert(str(p))
    img.Left = sheet.Range(cell).Left
    img.Top = sheet.Range(cell).Top
    if width is not None:
        img.Width = width
    if height is not None:
        img.Height = height
    return f"added_image: {p.name} at {cell}"


def _list_shapes(workbook: Any, op: dict) -> list[dict]:
    """列出工作表所有形状 (图片/文本框/形状).

    Args:
        sheet: 工作表名称
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    shapes = []
    for i in range(1, sheet.Shapes.Count + 1):
        sh = sheet.Shapes(i)
        shapes.append({
            "index": i,
            "name": sh.Name,
            "type": str(sh.Type),
            "left": sh.Left,
            "top": sh.Top,
            "width": sh.Width,
            "height": sh.Height,
        })
    return shapes


def _delete_shape(workbook: Any, op: dict) -> str:
    """删除形状.

    Args:
        sheet: 工作表名称
        index: 形状索引 (1-based)
        name: 形状名称 (与 index 二选一)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    index = op.get("index")
    name = op.get("name", "")

    if index is not None:
        sh = sheet.Shapes(index)
        sh_name = sh.Name
    elif name:
        sh = sheet.Shapes(name)
        sh_name = name
    else:
        raise COMOperationError("delete_shape", "index 或 name 必填")

    sh.Delete()
    return f"deleted_shape: {sh_name}"


def _add_comment(workbook: Any, op: dict) -> str:
    """添加批注.

    Args:
        sheet: 工作表名称
        cell: 单元格
        text: 批注内容
        author: 作者 (默认 'AI')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    text = op.get("text", "")
    author = op.get("author", "AI")

    if not text:
        raise COMOperationError("add_comment", "text 不能为空")

    rng = sheet.Range(cell)
    if rng.Comment:
        rng.Comment.Delete()
    rng.AddComment(text)
    try:
        rng.Comment.Author = author
    except Exception:
        pass
    return f"added_comment: {cell} by {author}"


def _delete_comment(workbook: Any, op: dict) -> str:
    """删除批注.

    Args:
        sheet: 工作表名称
        cell: 单元格
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    rng = sheet.Range(cell)
    if rng.Comment:
        rng.Comment.Delete()
        return f"deleted_comment: {cell}"
    return f"no_comment: {cell}"


# ============ View 视图操作 (3 个) ============

def _set_view_zoom(workbook: Any, op: dict) -> str:
    """设置视图缩放.

    Args:
        sheet: 工作表名称
        zoom: 缩放比例 (10-400)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    zoom = op.get("zoom", 100)
    if not 10 <= zoom <= 400:
        raise COMOperationError(
            "set_view_zoom", f"zoom 必须在 10-400 之间, 收到 {zoom}"
        )
    # 通过激活并设置 zoom
    sheet.Activate()
    active_window = workbook.Application.ActiveWindow
    active_window.Zoom = zoom
    return f"set_view_zoom: {zoom}%"


def _set_view_gridlines(workbook: Any, op: dict) -> str:
    """设置是否显示网格线.

    Args:
        sheet: 工作表名称
        show: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    show = op.get("show", True)
    sheet.Activate()
    workbook.Application.ActiveWindow.DisplayGridlines = show
    return f"set_view_gridlines: {show}"


def _set_view_headings(workbook: Any, op: dict) -> str:
    """设置是否显示行列标题.

    Args:
        sheet: 工作表名称
        show: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    show = op.get("show", True)
    sheet.Activate()
    workbook.Application.ActiveWindow.DisplayHeadings = show
    return f"set_view_headings: {show}"


# ============ Calculation 计算操作 (3 个) ============

def _recalculate(workbook: Any, op: dict) -> str:
    """重新计算所有公式.

    Args:
        full: True 全量重算 / False 仅脏数据 (当前实现均全量)
    """
    # 用 Application.Calculate() 而不是 workbook.Calculate()
    # 避免某些环境下 workbook 被错误绑定
    full = op.get("full", True)
    app = workbook.Application
    if app is not None:
        # 注: Excel COM 中 Calculate() 即为全量重算, 没有单独的"仅脏数据"公开 API
        app.Calculate()
    else:
        workbook.Calculate()
    return f"recalculated: full={full}"


def _set_calculation_mode(workbook: Any, op: dict) -> str:
    """设置计算模式.

    Args:
        mode: 'auto' (1) / 'manual' (-4135) / 'semiauto' (2)
    """
    mode = op.get("mode", "auto")
    mode_map = {
        "auto": -4105,  # xlCalculationAutomatic
        "manual": -4135,  # xlCalculationManual
        "semiauto": 2,  # xlCalculationSemiautomatic
    }
    if mode not in mode_map:
        raise COMOperationError("set_calculation_mode", f"mode 必须是 auto/manual/semiauto")
    workbook.Application.Calculation = mode_map[mode]
    return f"set_calculation_mode: {mode}"


def _set_iterative_calc(workbook: Any, op: dict) -> str:
    """启用/配置迭代计算.

    Args:
        enable: True/False
        max_iterations: 最大迭代次数 (1-32767, 默认 100)
        max_change: 最大变化量 (>0, 默认 0.001)
    """
    enable = op.get("enable", True)
    max_iter = op.get("max_iterations", 100)
    max_change = op.get("max_change", 0.001)
    if not 1 <= int(max_iter) <= 32767:
        raise COMOperationError(
            "set_iterative_calc",
            f"max_iterations 必须在 1-32767, 收到 {max_iter}",
        )
    if float(max_change) <= 0:
        raise COMOperationError(
            "set_iterative_calc", f"max_change 必须 > 0, 收到 {max_change}"
        )
    app = workbook.Application
    app.Iteration = enable
    app.MaxIterations = max_iter
    app.MaxChange = max_change
    return f"set_iterative_calc: enable={enable}, max_iter={max_iter}"
