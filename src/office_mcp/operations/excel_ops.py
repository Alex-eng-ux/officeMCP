"""Excel COM 鎿嶄綔瀹炵幇."""

import logging
from pathlib import Path
from typing import Any

from office_mcp.core.errors import COMOperationError
from office_mcp.core.path_guard import validate_path

logger = logging.getLogger(__name__)


# 鎿嶄綔 op 涓互 _path / _file 缁撳熬鎴栨槑纭负璺緞瀛楁鐨?key 闆嗗悎
_PATH_FIELDS = (
    "image_path", "source_path", "target_path", "template_path",
    "file_path", "new_path", "output_path", "output_dir", "from_file",
    "to_file", "src_path", "dst_path", "data_source",
)


def _validate_op_paths(op: dict) -> None:
    """鏍￠獙 op dict 涓墍鏈夌枒浼艰矾寰勫瓧娈? 闃叉浠绘剰鏂囦欢璁块棶.

    瀵规湭鍦ㄧ櫧鍚嶅崟鍐呬絾褰技 Windows 璺緞鐨勫瓧娈? 涔熷仛璀﹀憡.
    """
    for key, value in op.items():
        if not isinstance(value, str) or not value:
            continue
        # 宸叉樉寮忓垪鍏ョ櫧鍚嶅崟: 鏍￠獙
        if key.lower() in _PATH_FIELDS or key.lower().endswith(("_path", "_file", "path")):
            try:
                validate_path(value)
            except COMOperationError:
                raise
            except Exception as e:
                raise COMOperationError(op.get("type", "?"), f"璺緞鏍￠獙澶辫触 {key}={value}: {e}")

# 鍥捐〃绫诲瀷鏄犲皠
CHART_TYPE_MAP = {
    "column": 51,      # xlColumnClustered
    "bar": 57,         # xlBarClustered
    "line": 65,        # xlLine
    "pie": 5,          # xlPie
    "scatter": 72,     # xlXYScatter
    "area": 76,        # xlArea
}


def apply_excel_operations(workbook: Any, operations: list[dict]) -> list[dict]:
    """瀵?Excel 宸ヤ綔绨挎墽琛屾壒閲忔搷浣?

    Args:
        workbook: Excel Workbook 瀵硅薄
        operations: 鎿嶄綔鍒楄〃

    Returns:
        姣忎釜鎿嶄綔鐨勬墽琛岀粨鏋?
    """
    results = []
    for op in operations:
        op_type = op.get("type", "")
        try:
            # 鍏ュ彛澶勬牎楠?op 涓墍鏈夌枒浼艰矾寰勫瓧娈?
            _validate_op_paths(op)
            result = _execute_excel_operation(workbook, op)
            results.append({"type": op_type, "status": "success", "result": result})
        except Exception as e:
            logger.error(f"Excel 鎿嶄綔澶辫触 [{op_type}]: {e}")
            results.append({"type": op_type, "status": "error", "message": str(e)})
    return results


def _get_sheet(workbook: Any, sheet_name: str) -> Any:
    """鑾峰彇宸ヤ綔琛?"""
    try:
        return workbook.Worksheets(sheet_name)
    except Exception as e:
        raise COMOperationError(f"鑾峰彇宸ヤ綔琛?'{sheet_name}'", str(e))


def _col_idx_to_letters(col: int) -> str:
    """灏?1-based 鍒楀彿杞负瀛楁瘝: 1->A, 26->Z, 27->AA, 52->AZ, 53->BA, 702->ZZ, 703->AAA."""
    if col < 1:
        return ""
    result = ""
    while col > 0:
        col, rem = divmod(col - 1, 26)
        result = chr(65 + rem) + result
    return result


def _excel_require_active_window(workbook: Any, operation: str) -> Any:
    """Return Excel's active window or raise a clear error."""
    app = getattr(workbook, "Application", None)
    active_window = getattr(app, "ActiveWindow", None) if app is not None else None
    if active_window is None:
        raise COMOperationError(operation, "Excel active window is unavailable; activate the workbook before changing the view")
    return active_window


def _excel_external_range_address(range_obj: Any) -> str:
    """Return a pivot-safe external address string for a range."""
    try:
        return str(range_obj.Address(True, True, 1, True))
    except Exception:
        worksheet = getattr(range_obj, "Worksheet", None)
        sheet_name = getattr(worksheet, "Name", "Sheet1")
        return f"'{sheet_name}'!{range_obj.Address}"


def _excel_find_pivot_table(sheet: Any, pivot_name: str = "") -> Any | None:
    """Return the first matching PivotTable on a sheet."""
    pivot_tables = sheet.PivotTables()
    for index in range(1, pivot_tables.Count + 1):
        candidate = pivot_tables(index)
        if not pivot_name or candidate.Name == pivot_name:
            return candidate
    return None


def _execute_excel_operation(workbook: Any, op: dict) -> Any:
    """鎵ц鍗曚釜 Excel 鎿嶄綔."""
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
    elif op_type == "goal_seek":
        return _goal_seek(workbook, op)
    else:
        raise COMOperationError(f"鏈煡鐨?Excel 鎿嶄綔绫诲瀷: {op_type}")


def _check_typography(workbook: Any, op: dict) -> list[dict]:
    """妫€鏌?Excel 宸ヤ綔绨挎帓鐗堥棶棰?

    Args:
        workbook: Excel 宸ヤ綔绨垮璞?
        op: 鎿嶄綔閰嶇疆

    Returns:
        闂鍒楄〃锛屾瘡涓棶棰樺寘鍚?type, description, location
    """
    issues = []
    sheet_name = op.get("sheet", None)

    try:
        # 濡傛灉鎸囧畾浜嗗伐浣滆〃锛屽彧妫€鏌ヨ琛紝鍚﹀垯妫€鏌ユ墍鏈夎〃
        sheets_to_check = []
        if sheet_name:
            sheets_to_check.append(_get_sheet(workbook, sheet_name))
        else:
            for sheet in workbook.Worksheets:
                sheets_to_check.append(sheet)

        for sheet in sheets_to_check:
            sheet_name_current = sheet.Name
            # 1. 妫€鏌ュ崟鍏冩牸鍐呭瀵归綈
            issues.extend(_check_cell_alignment(sheet, sheet_name_current))

            # 2. 妫€鏌ユ暟瀛楁牸寮忎竴鑷存€?
            issues.extend(_check_number_format_consistency(sheet, sheet_name_current))

            # 3. 妫€鏌ヨ竟妗嗕娇鐢ㄨ鑼?
            issues.extend(_check_border_usage(sheet, sheet_name_current))

    except Exception as e:
        logger.error(f"Excel 鎺掔増妫€鏌ュ嚭閿? {e}")
        issues.append({
            "type": "error",
            "description": f"鎺掔増妫€鏌ヨ繃绋嬩腑鍙戠敓閿欒: {str(e)}",
            "location": "entire_workbook",
        })

    return issues


def _check_cell_alignment(sheet: Any, sheet_name: str) -> list[dict]:
    """妫€鏌ュ崟鍏冩牸鍐呭瀵归綈."""
    issues = []
    try:
        # 瀹氫箟 Excel 瀵归綈甯搁噺
        xlHAlignGeneral = 1
        xlHAlignLeft = -4131
        xlHAlignCenter = -4108
        xlHAlignRight = -4152

        # 鑾峰彇浣跨敤鑼冨洿
        used_range = sheet.UsedRange
        if used_range is None:
            return issues

        row_count = used_range.Rows.Count
        col_count = used_range.Columns.Count

        # 绠€鍗曟鏌ワ細鍚屼竴鍒楃殑鍗曞厓鏍煎榻愭柟寮忔槸鍚︿竴鑷达紙閽堝鍓?00琛屽拰鍓?0鍒楋級
        max_rows = min(row_count, 100)
        max_cols = min(col_count, 20)

        for col in range(1, max_cols + 1):
            # 鑾峰彇绗竴琛岀殑鏁版嵁绫诲瀷浣滀负鍙傝€?
            first_cell = sheet.Cells(1, col)
            first_value = first_cell.Value
            first_align = first_cell.HorizontalAlignment

            # 濡傛灉绗竴琛屾湁鍊硷紝妫€鏌ュ悓鍒楀叾浠栧崟鍏冩牸
            if first_value is not None:
                for row in range(2, max_rows + 1):
                    cell = sheet.Cells(row, col)
                    cell_value = cell.Value

                    if cell_value is not None:
                        # 鏁板瓧鍜屾枃鏈€氬父鏈変笉鍚岀殑瀵归綈涔犳儻
                        # 鏁板瓧閫氬父鍙冲榻愶紝鏂囨湰閫氬父宸﹀榻?
                        cell_align = cell.HorizontalAlignment
                        is_number = isinstance(cell_value, (int, float))
                        is_first_number = isinstance(first_value, (int, float))

                        if is_number and cell_align not in (xlHAlignRight, xlHAlignGeneral):
                            issues.append({
                                "type": "cell_alignment",
                                "description": f"鏁板瓧鍗曞厓鏍煎缓璁娇鐢ㄥ彸瀵归綈锛屽綋鍓嶅榻? {cell_align}",
                                "location": f"{sheet_name}!{_col_idx_to_letters(col)}{row}"
                            })
                        elif not is_number and cell_align == xlHAlignRight:
                            issues.append({
                                "type": "cell_alignment",
                                "description": f"鏂囨湰鍗曞厓鏍煎缓璁娇鐢ㄥ乏瀵归綈",
                                "location": f"{sheet_name}!{_col_idx_to_letters(col)}{row}"
                            })
    except Exception as e:
        logger.warning(f"妫€鏌ュ崟鍏冩牸瀵归綈鍑洪敊: {e}")
    return issues


def _check_number_format_consistency(sheet: Any, sheet_name: str) -> list[dict]:
    """妫€鏌ユ暟瀛楁牸寮忎竴鑷存€?"""
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
            # 鏀堕泦鍒椾腑鎵€鏈夋暟瀛楀崟鍏冩牸鐨勬牸寮?
            number_formats = []
            for row in range(1, max_rows + 1):
                cell = sheet.Cells(row, col)
                cell_value = cell.Value
                if isinstance(cell_value, (int, float)):
                    fmt = cell.NumberFormat
                    if fmt and fmt not in number_formats:
                        number_formats.append(fmt)

            # 濡傛灉鍚屽垪涓湁澶氱鏁板瓧鏍煎紡锛屽缓璁粺涓€
            if len(number_formats) > 1:
                issues.append({
                    "type": "number_format",
                    "description": f"鍒椾腑瀛樺湪澶氱鏁板瓧鏍煎紡: {', '.join(number_formats)}",
                    "location": f"{sheet_name}!鍒?{_col_idx_to_letters(col)}"
                })
    except Exception as e:
        logger.warning(f"妫€鏌ユ暟瀛楁牸寮忓嚭閿? {e}")
    return issues


def _check_border_usage(sheet: Any, sheet_name: str) -> list[dict]:
    """妫€鏌ヨ竟妗嗕娇鐢ㄨ鑼? 鎶ュ憡鏃犺竟妗嗙殑鏈夊唴瀹瑰绔嬪崟鍏冩牸.

    娉? 浠呮姤鍛婇《閮ㄨ (鏍囬琛? 鐨勫崟鍏冩牸鏄惁缂哄皯杈规, 绠€鍗曞惎鍙戝紡.
    """
    issues: list[dict] = []
    try:
        used_range = sheet.UsedRange
        if used_range is None:
            return issues

        # 杈规 COM 甯搁噺
        xlEdgeTop = 8
        xlLineStyleNone = -4142

        # 鍙鏌ョ涓€琛?(header 琛? 鏄惁鏈夎竟妗?
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
        logger.warning(f"妫€鏌ヨ竟妗嗕娇鐢ㄥ嚭閿? {e}")
    return issues


def _write_cell(workbook: Any, op: dict) -> str:
    """鍐欏叆鍗曞厓鏍?"""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    value = op.get("value", "")
    sheet.Range(cell).Value = value
    return f"wrote_cell: {cell} = {value}"


def _write_range(workbook: Any, op: dict) -> str:
    """鍐欏叆鑼冨洿."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    start_cell = op.get("start_cell", "A1")
    data = op.get("data", [])

    if not data:
        return "no_data"

    rows = len(data)
    cols = max(len(row) for row in data) if data else 0

    # 璁＄畻缁撴潫鍗曞厓鏍?
    start_row = sheet.Range(start_cell).Row
    start_col = sheet.Range(start_cell).Column
    end_row = start_row + rows - 1
    end_col = start_col + cols - 1

    # 灏嗗垪鍙疯浆涓哄瓧姣?
    def col_to_letter(col: int) -> str:
        result = ""
        while col > 0:
            col, rem = divmod(col - 1, 26)
            result = chr(65 + rem) + result
        return result

    end_cell = f"{col_to_letter(end_col)}{end_row}"
    range_obj = sheet.Range(f"{start_cell}:{end_cell}")

    # 濉厖鏁版嵁锛岃ˉ鍏ㄧ煭琛?
    filled_data = []
    for row in data:
        filled_row = list(row) + [""] * (cols - len(row))
        filled_data.append(filled_row)

    range_obj.Value = filled_data
    return f"wrote_range: {start_cell}:{end_cell}"


def _read_range(workbook: Any, op: dict) -> Any:
    """璇诲彇鑼冨洿."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    values = sheet.Range(range_str).Value

    # 缁熶竴涓轰簩缁村垪琛?
    if values is None:
        return []
    if not isinstance(values, tuple):
        values = ((values,),)

    # 澶勭悊鍗曡鎴栧崟鍒楃殑鎯呭喌
    result = []
    for row in values:
        if isinstance(row, tuple):
            result.append(list(row))
        else:
            result.append([row])
    return result


def _add_formula(workbook: Any, op: dict) -> str:
    """娣诲姞鍏紡."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    formula = op.get("formula", "")
    sheet.Range(cell).Formula = formula
    return f"added_formula: {cell} = {formula}"


def _format_range(workbook: Any, op: dict) -> str:
    """鏍煎紡鍖栬寖鍥?"""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    range_obj = sheet.Range(range_str)

    if op.get("bold"):
        range_obj.Font.Bold = True
    if op.get("italic"):
        range_obj.Font.Italic = True

    # 鑳屾櫙鑹?(鏀寔 #RRGGBB)
    bg_color = op.get("background_color")
    if bg_color:
        range_obj.Interior.Color = _hex_to_rgb(bg_color)

    # 瀛椾綋鑹?
    font_color = op.get("font_color")
    if font_color:
        range_obj.Font.Color = _hex_to_rgb(font_color)

    return f"formatted_range: {range_str}"


def _set_number_format(workbook: Any, op: dict) -> str:
    """璁剧疆鏁板瓧鏍煎紡."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    fmt = op.get("format", "General")
    sheet.Range(range_str).NumberFormat = fmt
    return f"set_number_format: {range_str} -> {fmt}"


def _create_chart(workbook: Any, op: dict) -> str:
    """鍒涘缓鍥捐〃."""
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
    """娣诲姞宸ヤ綔琛?"""
    name = op.get("name", "Sheet")
    sheet = workbook.Worksheets.Add()
    sheet.Name = name
    return f"added_worksheet: {name}"


def _rename_worksheet(workbook: Any, op: dict) -> str:
    """閲嶅懡鍚嶅伐浣滆〃."""
    old_name = op.get("old_name", "")
    new_name = op.get("new_name", "")
    workbook.Worksheets(old_name).Name = new_name
    return f"renamed_worksheet: {old_name} -> {new_name}"


def _auto_fit_columns(workbook: Any, op: dict) -> str:
    """鑷姩璋冩暣鍒楀."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    columns = op.get("columns", [])
    if isinstance(columns, str):
        sheet.Columns(columns).AutoFit()
        return f"auto_fit_columns: {columns}"
    for col in columns:
        sheet.Columns(col).AutoFit()
    return f"auto_fit_columns: {columns}"


def _freeze_panes(workbook: Any, op: dict) -> str:
    """鍐荤粨绐楁牸."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A2")
    sheet.Range(cell).Select()
    sheet.Application.ActiveWindow.FreezePanes = True
    return f"freeze_panes: {cell}"


def _hex_to_rgb(hex_color: str) -> int:
    """灏?#RRGGBB 杞负 Office RGB 鏁存暟."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return 0
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g << 8) + (b << 16)


# ============ Excel 楂樼骇鍔熻兘 ============

def _add_data_validation(workbook: Any, op: dict) -> str:
    """娣诲姞鏁版嵁楠岃瘉.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 楠岃瘉鑼冨洿 (濡?"A1:A10")
        type: 楠岃瘉绫诲瀷 (list/whole/decimal/date/time/textLength/custom)
        formula1: 楠岃瘉鍏紡鎴栧垪琛ㄥ€?(閫楀彿鍒嗛殧)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:A10")
    validation_type = op.get("type", "list")
    formula1 = op.get("formula1", "")
    formula2 = op.get("formula2", "")
    operator_name = op.get("operator", "")
    ignore_blank = op.get("ignore_blank", True)
    in_cell_dropdown = op.get("in_cell_dropdown", True)
    show_input = op.get("show_input", True)
    input_title = op.get("input_title", "")
    input_message = op.get("input_message", "")
    show_error = op.get("show_error", True)
    error_title = op.get("error_title", "")
    error_message = op.get("error_message", "")
    error_style_name = op.get("error_style", "stop")

    # 楠岃瘉绫诲瀷鏄犲皠
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
    if not formula1:
        raise COMOperationError("add_data_validation", "formula1 is required")
    if operator_name in {"between", "not_between"} and not formula2:
        raise COMOperationError("add_data_validation", "formula2 is required for between/not_between validations")

    # 鍏堝垹闄ゅ凡鏈夐獙璇侊紝閬垮厤鍐茬獊
    try:
        range_obj.Validation.Delete()
    except Exception:
        pass

    try:
        # list 绫诲瀷 (Type=3) 涓嶉渶瑕?Operator 鍙傛暟锛孎ormula1 涓嶈兘涓虹┖
        if validation_type_val == 3:  # xlValidateList
            range_obj.Validation.Add(
                Type=validation_type_val,
                AlertStyle={"stop": 1, "warning": 2, "information": 3}.get(error_style_name, 1),
                Formula1=formula1,
            )
        else:
            operator_map = {
                "between": 1,
                "not_between": 2,
                "equal": 3,
                "not_equal": 4,
                "greater": 5,
                "less": 6,
                "greater_equal": 7,
                "less_equal": 8,
            }
            validation_kwargs = {
                "Type": validation_type_val,
                "AlertStyle": {"stop": 1, "warning": 2, "information": 3}.get(error_style_name, 1),
                "Formula1": formula1,
            }
            if operator_name in operator_map:
                validation_kwargs["Operator"] = operator_map[operator_name]
            if formula2:
                validation_kwargs["Formula2"] = formula2
            range_obj.Validation.Add(**validation_kwargs)
        validation = range_obj.Validation
        validation.IgnoreBlank = bool(ignore_blank)
        try:
            validation.InCellDropdown = bool(in_cell_dropdown)
        except Exception:
            logger.debug("Validation dropdown flag unsupported for this validation type")
        validation.ShowInput = bool(show_input)
        validation.InputTitle = input_title
        validation.InputMessage = input_message
        validation.ShowError = bool(show_error)
        validation.ErrorTitle = error_title
        validation.ErrorMessage = error_message
    except Exception as e:
        raise COMOperationError("add_data_validation", str(e))
    return f"added_data_validation: {range_str} ({validation_type})"


def _add_conditional_format(workbook: Any, op: dict) -> str:
    """娣诲姞鏉′欢鏍煎紡.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿 (濡?"A1:A10")
        type: 鏉′欢绫诲瀷 (cell_value/formula/color_scale/data_bar/icon_set)
        operator: 鎿嶄綔绗?(greater/less/equal/between)
        formula1: 鏉′欢鍊?
        formula2: 鏉′欢鍊?
        format_type: 鏍煎紡绫诲瀷 (color_scale/data_bar/icon_set)
        font_color: 瀛椾綋棰滆壊 (#RRGGBB)
        bg_color: 鑳屾櫙棰滆壊 (#RRGGBB)
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

    # 鍏堝垹闄ゅ凡鏈夋潯浠舵牸寮忥紝閬垮厤鍐茬獊
    try:
        range_obj.FormatConditions.Delete()
    except Exception:
        pass

    # 楂樼骇鏍煎紡绫诲瀷
    if format_type == "color_scale":
        # 鑹查樁
        color_scale = range_obj.FormatConditions.AddColorScale(ColorScaleType=3)
        # 璁剧疆榛樿棰滆壊: 绾?榛?缁?
        color_scale.ColorScaleCriteria(1).Type = 1  # xlLowestValue
        color_scale.ColorScaleCriteria(1).FormatColor.Color = _hex_to_rgb("#FF0000")
        color_scale.ColorScaleCriteria(2).Type = 5  # xlPercentile
        color_scale.ColorScaleCriteria(2).Value = 50
        color_scale.ColorScaleCriteria(2).FormatColor.Color = _hex_to_rgb("#FFFF00")
        color_scale.ColorScaleCriteria(3).Type = 2  # xlHighestValue
        color_scale.ColorScaleCriteria(3).FormatColor.Color = _hex_to_rgb("#00FF00")
        return f"added_conditional_format: {range_str} (color_scale)"

    elif format_type == "data_bar":
        # 鏁版嵁鏉?
        data_bar = range_obj.FormatConditions.AddDatabar()
        data_bar.BarColor.Color = _hex_to_rgb("#638EC6")
        data_bar.BarFillType = 0  # xlDataBarFillSolid
        return f"added_conditional_format: {range_str} (data_bar)"

    elif format_type == "icon_set":
        # 鍥炬爣闆?
        icon_set = range_obj.FormatConditions.AddIconSetCondition()
        icon_set.IconSet = workbook.Application.IconSets(3)  # xl3Arrows
        return f"added_conditional_format: {range_str} (icon_set)"

    # 甯歌鏉′欢鏍煎紡
    else:
        if condition_type == "cell_value":
            # xlCellValue 绫诲瀷 Formula1 涓嶈兘涓虹┖
            if not formula1:
                formula1 = "0"
            # 鎿嶄綔绗︽槧灏?
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
                raise COMOperationError("add_conditional_format", f"FormatConditions.Add 澶辫触: {e}")
        else:  # formula
            if not formula1:
                formula1 = "=TRUE"
            try:
                format_condition = range_obj.FormatConditions.Add(
                    Type=2,  # xlExpression
                    Formula1=formula1,
                )
            except Exception as e:
                raise COMOperationError("add_conditional_format", f"FormatConditions.Add 澶辫触: {e}")

        # 搴旂敤鏍煎紡
        if font_color:
            format_condition.Font.Color = _hex_to_rgb(font_color)
        if bg_color:
            format_condition.Interior.Color = _hex_to_rgb(bg_color)

        return f"added_conditional_format: {range_str}"


def _merge_cells(workbook: Any, op: dict) -> str:
    """鍚堝苟鍗曞厓鏍?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿 (濡?"A1:C3")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:C3")
    sheet.Range(range_str).Merge()
    return f"merged_cells: {range_str}"


def _set_borders(workbook: Any, op: dict) -> str:
    """璁剧疆杈规.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿 (濡?"A1:C3")
        border_type: 杈规绫诲瀷 (all/outside/inside)
        style: 绾垮瀷 (thin/medium/thick/dashed/dotted)
        color: 棰滆壊 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:C3")
    border_type = op.get("border_type", "all")
    style = op.get("style", "thin")
    color = op.get("color", "#000000")

    # 绾垮瀷鏄犲皠
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
    """娣诲姞鍛藉悕鑼冨洿.

    Args:
        name: 鍚嶇О
        refers_to: 寮曠敤鍏紡 (濡?"=Sheet1!$A$1:$A$10")
    """
    name = op.get("name", "")
    refers_to = op.get("refers_to", "=Sheet1!$A$1:$A$10")

    if not name:
        raise COMOperationError("add_named_range", "name 涓嶈兘涓虹┖")

    try:
        # 鍒犻櫎宸插瓨鍦ㄧ殑鍚屽悕鑼冨洿
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
            # RefersTo 澶辫触鏃跺皾璇?RefersToR1C1 浣滀负鍥為€€
            try:
                workbook.Names.Add(Name=name, RefersToR1C1=refers_to)
            except Exception as e2:
                raise COMOperationError("add_named_range", f"RefersTo 鍜?RefersToR1C1 鍧囧け璐? {e2}")
        return f"added_named_range: {name}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("add_named_range", str(e))


def _create_pivot_table(workbook: Any, op: dict) -> str:
    """鍒涘缓鏁版嵁閫忚琛?

    Args:
        source_sheet: 鏁版嵁婧愬伐浣滆〃鍚嶇О
        source_range: 鏁版嵁婧愯寖鍥?(濡?"A1:D100", 鐣欑┖鍒欒嚜鍔ㄤ娇鐢?UsedRange)
        target_sheet: 鐩爣宸ヤ綔琛ㄥ悕绉?(鑷姩鍒涘缓鎴栨寚瀹?
        target_cell: 鐩爣鍗曞厓鏍?(濡?"A3")
        row_fields: 琛屽瓧娈靛垪琛?(濡?["閮ㄩ棬", "鏈堜唤"])
        column_fields: 鍒楀瓧娈靛垪琛?(濡?["鍦板尯"])
        data_fields: 鏁版嵁瀛楁瀛楀吀 (濡?{"閿€鍞": "sum", "鏁伴噺": "average"})
    """
    source_sheet = _get_sheet(workbook, op.get("source_sheet", "Sheet1"))
    source_range = op.get("source_range", "")
    target_sheet_name = op.get("target_sheet", "数据透视表")
    target_cell = op.get("target_cell", "A3")
    pivot_name = op.get("pivot_name", "").strip()

    row_fields = op.get("row_fields", [])
    column_fields = op.get("column_fields", [])
    filter_fields = op.get("filter_fields", [])
    data_fields = op.get("data_fields", {})
    style_name = op.get("style_name", "").strip()

    # 濡傛灉鏈寚瀹?source_range锛屽垯浣跨敤 UsedRange 閬垮厤寮曠敤瓒呭嚭瀹為檯鏁版嵁鑼冨洿
    if not source_range:
        used = source_sheet.UsedRange
        if used is not None:
            source_range = used.Address
        else:
            source_range = "A1"

    # 鍒涘缓鎴栬幏鍙栫洰鏍囧伐浣滆〃
    try:
        target_sheet = workbook.Worksheets(target_sheet_name)
    except Exception:
        target_sheet = workbook.Worksheets.Add()
        target_sheet.Name = target_sheet_name

    # 鍒涘缓鏁版嵁閫忚琛ㄧ紦瀛?(浣跨敤鍦板潃瀛楃涓叉洿鍙潬)
    source_data_addr = source_sheet.Range(source_range)
    source_data_ref = _excel_external_range_address(source_data_addr)
    try:
        pivot_cache = workbook.PivotCaches.Create(
            SourceType=1,  # xlDatabase
            SourceData=source_data_ref,
        )
    except Exception as e:
        raise COMOperationError("create_pivot_table", f"PivotCaches.Create 澶辫触: {e}")

    # 鐢熸垚涓嶉噸澶嶇殑琛ㄥ悕
    import time
    table_name = pivot_name or f"PivotTable_{int(time.time())}"

    # 鍒涘缓鏁版嵁閫忚琛?
    try:
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=target_sheet.Range(target_cell),
            TableName=table_name,
        )
    except Exception as e:
        raise COMOperationError("create_pivot_table", f"CreatePivotTable 澶辫触: {e}")

    # 閰嶇疆琛屽瓧娈?
    for i, field in enumerate(row_fields):
        try:
            pf = pivot_table.PivotFields(field)
            pf.Orientation = 1  # xlRowField
            pf.Position = i + 1
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"琛屽瓧娈?'{field}' 涓嶅瓨鍦? {e}")

    # 閰嶇疆鍒楀瓧娈?
    for i, field in enumerate(column_fields):
        try:
            pf = pivot_table.PivotFields(field)
            pf.Orientation = 2  # xlColumnField
            pf.Position = i + 1
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"鍒楀瓧娈?'{field}' 涓嶅瓨鍦? {e}")

    for i, field in enumerate(filter_fields):
        try:
            pf = pivot_table.PivotFields(field)
            pf.Orientation = 3  # xlPageField
            pf.Position = i + 1
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"筛选字段 '{field}' 不存在: {e}")

    # 閰嶇疆鏁版嵁瀛楁
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
            raise COMOperationError("create_pivot_table", f"鏁版嵁瀛楁 '{field}' 涓嶅瓨鍦? {e}")

    if style_name:
        try:
            pivot_table.TableStyle2 = style_name
        except Exception as e:
            raise COMOperationError("create_pivot_table", f"无法应用透视表样式 '{style_name}': {e}")

    return f"created_pivot_table: {target_sheet_name}!{target_cell} ({table_name})"


def _import_data(workbook: Any, op: dict) -> str:
    """瀵煎叆澶栭儴鏁版嵁鏂囦欢 (CSV/TXT) 鍒板伐浣滆〃."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    file_path = op.get("file_path", "")
    start_cell = op.get("start_cell", "A1")
    delimiter = op.get("delimiter", ",")  # CSV 鍒嗛殧绗?
    has_header = op.get("has_header", True)

    if not file_path:
        raise COMOperationError("import_data", "file_path 涓嶈兘涓虹┖")

    # 璺緞鏍￠獙
    if not Path(file_path).exists():
        raise COMOperationError("import_data", f"鏂囦欢涓嶅瓨鍦? {file_path}")

    # 浣跨敤 QueryTables 瀵煎叆
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
            query.Delete()  # 瀵煎叆瀹屾垚鍚庡垹闄ゆ煡璇㈠璞?

    return f"imported_data: {file_path} -> {sheet.Name}"


def _export_data(workbook: Any, op: dict) -> str:
    """瀵煎嚭宸ヤ綔琛ㄤ负 CSV 鏂囦欢."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    export_path = op.get("export_path", "")

    if not export_path:
        raise COMOperationError("export_data", "export_path 涓嶈兘涓虹┖")

    # 澶嶅埗鍒版柊宸ヤ綔绨垮啀淇濆瓨涓?CSV
    new_wb = workbook.Application.Workbooks.Add()
    sheet.Copy(Before=new_wb.Worksheets(1))
    # 鍒犻櫎鑷姩鐢熸垚鐨勫浣欏伐浣滆〃
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
    """娣诲姞鍒囩墖鍣?

    Args:
        target_sheet: 鍒囩墖鍣ㄦ墍鍦ㄥ伐浣滆〃鍚嶇О
        pivot_sheet: 鏁版嵁閫忚琛ㄦ墍鍦ㄥ伐浣滆〃鍚嶇О
        pivot_name: 鏁版嵁閫忚琛ㄥ悕绉?
        field_name: 瑕佺瓫閫夌殑瀛楁鍚嶇О
        left: 鍒囩墖鍣ㄥ乏渚т綅缃?(鍍忕礌)
        top: 鍒囩墖鍣ㄩ《閮ㄤ綅缃?(鍍忕礌)
        width: 鍒囩墖鍣ㄥ搴?(鍍忕礌)
        height: 鍒囩墖鍣ㄩ珮搴?(鍍忕礌)
    """
    target_sheet_name = op.get("target_sheet", "Sheet1")
    pivot_sheet_name = op.get("pivot_sheet", "数据透视表")
    pivot_name = op.get("pivot_name", "")
    field_name = op.get("field_name", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 200)
    height = op.get("height", 200)

    if not field_name:
        raise COMOperationError("add_slicer", "field_name is required")

    try:
        target_sheet = _get_sheet(workbook, target_sheet_name)
        pivot_sheet = _get_sheet(workbook, pivot_sheet_name)

        # 鏌ユ壘鏁版嵁閫忚琛?
        pivot_table = _excel_find_pivot_table(pivot_sheet, pivot_name)

        if not pivot_table:
            raise COMOperationError("add_slicer", f"鏈壘鍒版暟鎹€忚琛? {pivot_name}")

        # 娣诲姞鍒囩墖鍣ㄧ紦瀛?
        slicer_caches = workbook.SlicerCaches
        slicer_cache = None
        try:
            for index in range(1, slicer_caches.Count + 1):
                candidate = slicer_caches(index)
                candidate_name = str(getattr(candidate, "Name", "") or "")
                source_name = str(getattr(candidate, "SourceName", "") or "")
                if field_name in {candidate_name, source_name}:
                    slicer_cache = candidate
                    break
        except Exception:
            logger.debug("Could not enumerate existing slicer caches", exc_info=True)

        try:
            if slicer_cache is None:
                slicer_cache = slicer_caches.Add2(pivot_table, field_name)
        except Exception:
            try:
                if slicer_cache is None:
                    slicer_cache = slicer_caches.Add(pivot_table, field_name)
            except Exception as e:
                raise COMOperationError("add_slicer", f"Slicer API unavailable for field '{field_name}': {e}")

        # 娣诲姞鍒囩墖鍣?
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
    """娣诲姞鍒嗙被姹囨€?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鏁版嵁鑼冨洿 (濡?"A1:D100")
        group_by: 鍒嗙粍瀛楁鍒楀彿 (濡?1 琛ㄧず绗?1 鍒?
        summary_function: 姹囨€诲嚱鏁?(sum/count/average/max/min)
        summary_fields: 瑕佹眹鎬荤殑鍒楀彿鍒楄〃 (濡?[3, 4])
        replace: 鏄惁鏇挎崲鐜版湁鍒嗙被姹囨€?
        page_breaks: 鏄惁鍦ㄦ瘡缁勫悗鍒嗛〉
        summary_below: 姹囨€荤粨鏋滄槸鍚﹀湪鏁版嵁涓嬫柟
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:D100")
    group_by = op.get("group_by", 1)
    summary_function = op.get("summary_function", "sum")
    summary_fields = op.get("summary_fields", [])
    # summary_fields 涓虹┖鏃朵娇鐢ㄩ粯璁ゅ€硷紝TotalList 蹇呴』涓洪潪绌哄厓缁?
    if not summary_fields:
        summary_fields = [2]
    replace = op.get("replace", True)
    page_breaks = op.get("page_breaks", False)
    summary_below = op.get("summary_below", True)

    # 姹囨€诲嚱鏁版槧灏?
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


# ============ Worksheet 宸ヤ綔琛ㄦ搷浣?(10 涓? ============

def _list_worksheets(workbook: Any, op: dict) -> list[dict]:
    """鍒楀嚭鎵€鏈夊伐浣滆〃.

    Args:
        workbook: Excel 宸ヤ綔绨垮璞?

    Returns:
        宸ヤ綔琛ㄤ俊鎭垪琛?
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
    """鑾峰彇宸ヤ綔琛ㄤ俊鎭?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
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
        "protect_drawing_objects": getattr(sheet, "ProtectDrawingObjects", False),
        "protect_scenarios": getattr(sheet, "ProtectScenarios", False),
        "protection_mode": getattr(sheet, "ProtectionMode", False),
    }


def _copy_worksheet(workbook: Any, op: dict) -> str:
    """澶嶅埗宸ヤ綔琛?

    Args:
        sheet: 婧愬伐浣滆〃鍚嶇О
        new_name: 鏂板伐浣滆〃鍚嶇О (鍙€?
        position: 浣嶇疆 (before/after, 鍙€?
        target_sheet: 鐩爣浣嶇疆鍙傝€冨伐浣滆〃 (鍙€?
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

    # 鑾峰彇澶嶅埗鍚庣殑宸ヤ綔琛?
    new_sheet = sheet.Next
    if new_sheet and new_sheet.Name == sheet.Name:
        new_sheet = new_sheet.Next

    if new_name and new_sheet:
        new_sheet.Name = new_name

    return f"copied_worksheet: {sheet.Name} -> {new_sheet.Name if new_sheet else 'unnamed'}"


def _delete_worksheet(workbook: Any, op: dict) -> str:
    """鍒犻櫎宸ヤ綔琛?

    Args:
        sheet: 瑕佸垹闄ょ殑宸ヤ綔琛ㄥ悕绉?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    sheet_name = sheet.Name

    # 妫€鏌ユ槸鍚﹀敮涓€宸ヤ綔琛?
    if workbook.Worksheets.Count == 1:
        raise COMOperationError("delete_worksheet", "涓嶈兘鍒犻櫎鍞竴鐨勫伐浣滆〃")

    sheet.Delete()
    return f"deleted_worksheet: {sheet_name}"


def _move_worksheet(workbook: Any, op: dict) -> str:
    """绉诲姩宸ヤ綔琛?

    Args:
        sheet: 瑕佺Щ鍔ㄧ殑宸ヤ綔琛ㄥ悕绉?
        position: 浣嶇疆 (before/after/first/last)
        target_sheet: 鐩爣浣嶇疆鍙傝€冨伐浣滆〃 (position=before/after 鏃跺繀濉?
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
    """闅愯棌宸ヤ綔琛?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    # Excel 涓嶅厑璁搁殣钘忓伐浣滅翱涓敮涓€鍙鐨勫伐浣滆〃
    visible_count = sum(1 for i in range(1, workbook.Worksheets.Count + 1) if workbook.Worksheets(i).Visible == -1)
    if visible_count <= 1:
        return f"hidden_worksheet: skipped (only {visible_count} visible sheet(s), Excel requires at least 1)"
    sheet.Visible = 0  # xlSheetHidden
    return f"hidden_worksheet: {sheet.Name}"


def _show_worksheet(workbook: Any, op: dict) -> str:
    """鏄剧ず宸ヤ綔琛?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    sheet.Visible = -1  # xlSheetVisible
    return f"showed_worksheet: {sheet.Name}"


def _protect_worksheet(workbook: Any, op: dict) -> str:
    """淇濇姢宸ヤ綔琛?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        password: 瀵嗙爜 (鍙€?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    password = op.get("password", "")
    protect_kwargs = {
        "DrawingObjects": op.get("drawing_objects", True),
        "Contents": op.get("contents", True),
        "Scenarios": op.get("scenarios", True),
        "UserInterfaceOnly": op.get("user_interface_only", False),
        "AllowFormattingCells": op.get("allow_formatting_cells", False),
        "AllowFormattingColumns": op.get("allow_formatting_columns", False),
        "AllowFormattingRows": op.get("allow_formatting_rows", False),
        "AllowInsertingColumns": op.get("allow_inserting_columns", False),
        "AllowInsertingRows": op.get("allow_inserting_rows", False),
        "AllowInsertingHyperlinks": op.get("allow_inserting_hyperlinks", False),
        "AllowDeletingColumns": op.get("allow_deleting_columns", False),
        "AllowDeletingRows": op.get("allow_deleting_rows", False),
        "AllowSorting": op.get("allow_sorting", False),
        "AllowFiltering": op.get("allow_filtering", False),
        "AllowUsingPivotTables": op.get("allow_using_pivot_tables", False),
    }
    if password:
        sheet.Protect(Password=password, **protect_kwargs)
    else:
        sheet.Protect(**protect_kwargs)
    return f"protected_worksheet: {sheet.Name}"


def _unprotect_worksheet(workbook: Any, op: dict) -> str:
    """鍙栨秷宸ヤ綔琛ㄤ繚鎶?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        password: 瀵嗙爜 (鍙€?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    password = op.get("password", "")
    if password:
        sheet.Unprotect(Password=password)
    else:
        sheet.Unprotect()
    return f"unprotected_worksheet: {sheet.Name}"


def _set_tab_color(workbook: Any, op: dict) -> str:
    """璁剧疆宸ヤ綔琛ㄦ爣绛鹃鑹?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        color: 棰滆壊 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    color = op.get("color", "#FF0000")
    sheet.Tab.Color = _hex_to_rgb(color)
    return f"set_tab_color: {sheet.Name} -> {color}"


# ============ Range 鑼冨洿鎿嶄綔 (10 涓? ============

def _list_used_range(workbook: Any, op: dict) -> dict:
    """鍒楀嚭宸蹭娇鐢ㄨ寖鍥?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
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
    """娓呴櫎鑼冨洿鍐呭.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        clear_type: 娓呴櫎绫诲瀷 (all/formulas/contents/comments)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    clear_type = op.get("clear_type", "all").lower()

    sheet.Range(range_str)
    rng = sheet.Range(range_str)
    # 鎸夌被鍨嬪垎娲惧埌涓嶅悓鐨?Clear 鏂规硶
    if clear_type == "all":
        rng.Clear()
    elif clear_type == "contents":
        rng.ClearContents()
    elif clear_type == "formulas":
        # 娓呴櫎鍏紡浣嗕繚鐣欐牸寮?
        rng.ClearContents()
    elif clear_type == "comments":
        # 閬嶅巻姣忎釜鍗曞厓鏍煎垹闄ゆ壒娉?
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
            f"clear_type 蹇呴』鏄?all/contents/formulas/comments/hyperlinks/formats,"
            f" 鏀跺埌 '{clear_type}'",
        )
    return f"cleared_range: {range_str} ({clear_type})"


def _copy_range(workbook: Any, op: dict) -> str:
    """澶嶅埗鑼冨洿.

    Args:
        sheet: 婧愬伐浣滆〃鍚嶇О
        range: 婧愯寖鍥?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Copy()
    return f"copied_range: {sheet.Name}!{range_str}"


def _paste_range(workbook: Any, op: dict) -> str:
    """绮樿创鑼冨洿.

    Args:
        sheet: 鐩爣宸ヤ綔琛ㄥ悕绉?
        target_cell: 鐩爣鍗曞厓鏍?(濡?"A1")
        paste_type: 绮樿创绫诲瀷 (all/formulas/values/formats)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    target_cell = op.get("target_cell", "A1")
    paste_type = op.get("paste_type", "all")

    # 绮樿创绫诲瀷鏄犲皠
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
    """鍓垏鑼冨洿.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 婧愯寖鍥?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Cut()
    return f"cut_range: {sheet.Name}!{range_str}"


def _delete_cells(workbook: Any, op: dict) -> str:
    """鍒犻櫎鍗曞厓鏍?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        shift: 绉诲姩鏂瑰悜 (left/up)
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
    """鎻掑叆鍗曞厓鏍?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        shift: 绉诲姩鏂瑰悜 (right/down)
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
    """璁剧疆琛岄珮.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        row: 琛屽彿 (鎴栬寖鍥? 濡?"1:3" 琛ㄧず 1-3 琛?
        height: 楂樺害 (纾?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    row = op.get("row", 1)
    height = op.get("height", 15.0)
    sheet.Rows(row).RowHeight = height
    return f"set_row_height: {sheet.Name} row {row} = {height}"


def _set_column_width(workbook: Any, op: dict) -> str:
    """璁剧疆鍒楀.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        column: 鍒楁爣璇?(濡?"A" 鎴?"A:C")
        width: 瀹藉害 (瀛楃鍗曚綅)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    column = op.get("column", "A")
    width = op.get("width", 8.43)
    sheet.Columns(column).ColumnWidth = width
    return f"set_column_width: {sheet.Name} column {column} = {width}"


def _hide_rows(workbook: Any, op: dict) -> str:
    """闅愯棌琛?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        rows: 琛屽彿鎴栬寖鍥?(濡?"1" 鎴?"1:5")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    rows = op.get("rows", "1")
    sheet.Rows(rows).Hidden = True
    return f"hidden_rows: {sheet.Name} rows {rows}"


# ============ Charts 鍥捐〃鎿嶄綔 (10 涓? ============

def _list_charts(workbook: Any, op: dict) -> list[dict]:
    """鍒楀嚭鎵€鏈夊浘琛?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?(鍙€? 涓嶅～鍒欏垪鍑烘墍鏈夊伐浣滆〃)
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
    """鑾峰彇鍥捐〃淇℃伅.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩 (浠?1 寮€濮?
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
    """璁剧疆鍥捐〃鏍囬.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        title: 鏍囬鏂囨湰
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    title = op.get("title", "")

    chart = sheet.ChartObjects(chart_index).Chart
    chart.HasTitle = True
    chart.ChartTitle.Text = title
    return f"set_chart_title: {title}"


def _set_chart_legend(workbook: Any, op: dict) -> str:
    """璁剧疆鍥捐〃鍥句緥.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        show: 鏄惁鏄剧ず鍥句緥
        position: 鍥句緥浣嶇疆 (bottom/top/left/right/corner)
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
    """娣诲姞鍥捐〃绯诲垪.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        series_name: 绯诲垪鍚嶇О
        values_range: 鏁板€艰寖鍥?
        categories_range: 鍒嗙被鑼冨洿 (鍙€?
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
            # 鏌愪簺 Excel 鐗堟湰闇€瑕佺敤鍦板潃瀛楃涓?
            series.Values = f"={sheet.Name}!{sheet.Range(values_range).Address}"
    if categories_range:
        try:
            series.XValues = sheet.Range(categories_range)
        except Exception:
            series.XValues = f"={sheet.Name}!{sheet.Range(categories_range).Address}"

    return f"added_chart_series: {series_name}"


def _remove_chart_series(workbook: Any, op: dict) -> str:
    """绉婚櫎鍥捐〃绯诲垪.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        series_index: 绯诲垪绱㈠紩 (浠?1 寮€濮?
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
    """璁剧疆鍥捐〃杞?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        axis: 杞寸被鍨?(x/y/value1/value2)
        title: 杞存爣棰?(鍙€?
        min_scale: 鏈€灏忓€?(鍙€?
        max_scale: 鏈€澶у€?(鍙€?
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
    """鏇存敼鍥捐〃绫诲瀷.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        chart_type: 鏂板浘琛ㄧ被鍨?(column/bar/line/pie/scatter/area)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    chart_type = op.get("chart_type", "column")

    chart_type_val = CHART_TYPE_MAP.get(chart_type, 51)
    chart = sheet.ChartObjects(chart_index).Chart
    chart.ChartType = chart_type_val
    return f"changed_chart_type: {chart_type}"


def _export_chart(workbook: Any, op: dict) -> str:
    """瀵煎嚭鍥捐〃涓哄浘鐗?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
        output_path: 杈撳嚭鍥剧墖璺緞
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)
    output_path = op.get("output_path", "")

    if not output_path:
        raise COMOperationError("export_chart", "output_path 涓嶈兘涓虹┖")

    chart = sheet.ChartObjects(chart_index).Chart
    chart.Export(output_path)
    return f"exported_chart: {sheet.Name} chart {chart_index} -> {output_path}"


def _delete_chart(workbook: Any, op: dict) -> str:
    """鍒犻櫎鍥捐〃.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        chart_index: 鍥捐〃绱㈠紩
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    chart_index = op.get("chart_index", 1)

    chart_obj = sheet.ChartObjects(chart_index)
    chart_name = chart_obj.Name
    chart_obj.Delete()
    return f"deleted_chart: {chart_name}"


# ============ Format 鏍煎紡鎿嶄綔 (10 涓? ============

def _set_font(workbook: Any, op: dict) -> str:
    """璁剧疆瀛椾綋.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        font_name: 瀛椾綋鍚嶇О (濡?"寰蒋闆呴粦")
        font_size: 瀛椾綋澶у皬
        font_color: 瀛椾綋棰滆壊 (#RRGGBB)
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
    """璁剧疆绮椾綋.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        bold: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    bold = op.get("bold", True)
    sheet.Range(range_str).Font.Bold = bold
    return f"set_font_bold: {range_str} = {bold}"


def _set_font_italic(workbook: Any, op: dict) -> str:
    """璁剧疆鏂滀綋.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        italic: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    italic = op.get("italic", True)
    sheet.Range(range_str).Font.Italic = italic
    return f"set_font_italic: {range_str} = {italic}"


def _set_font_underline(workbook: Any, op: dict) -> str:
    """璁剧疆涓嬪垝绾?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        underline: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    underline = op.get("underline", True)
    sheet.Range(range_str).Font.Underline = underline
    return f"set_font_underline: {range_str} = {underline}"


def _set_alignment(workbook: Any, op: dict) -> str:
    """璁剧疆瀵归綈.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        horizontal: 姘村钩瀵归綈 (left/center/right/general)
        vertical: 鍨傜洿瀵归綈 (top/middle/bottom)
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
    """璁剧疆鑷姩鎹㈣.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        wrap: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    wrap = op.get("wrap", True)
    sheet.Range(range_str).WrapText = wrap
    return f"set_wrap_text: {range_str} = {wrap}"


def _set_indent(workbook: Any, op: dict) -> str:
    """璁剧疆缂╄繘.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        indent: 缂╄繘绾у埆 (0-15)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    indent = op.get("indent", 1)
    if indent <= 0:
        return f"set_indent: skipped (indent={indent} <= 0)"
    sheet.Range(range_str).InsertIndent(indent)
    return f"set_indent: {range_str} = {indent}"


def _set_orientation(workbook: Any, op: dict) -> str:
    """璁剧疆鏂囧瓧鏂瑰悜.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        orientation: 瑙掑害 (0=姘村钩, 90=鍨傜洿, 45=-45搴?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    orientation = op.get("orientation", 0)
    sheet.Range(range_str).Orientation = orientation
    return f"set_orientation: {range_str} = {orientation}"


def _clear_format(workbook: Any, op: dict) -> str:
    """娓呴櫎鏍煎紡.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).ClearFormats()
    return f"cleared_format: {range_str}"


def _copy_format(workbook: Any, op: dict) -> str:
    """澶嶅埗鏍煎紡.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        source_range: 婧愭牸寮忚寖鍥?
        target_range: 鐩爣鑼冨洿
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    source_range = op.get("source_range", "A1")
    target_range = op.get("target_range", "B1")

    source = sheet.Range(source_range)
    target = sheet.Range(target_range)
    source.Copy()
    target.PasteSpecial(Paste=-4122)  # xlPasteFormats
    return f"copied_format: {source_range} -> {target_range}"


# ============ Page Setup 椤甸潰璁剧疆 (10 涓? ============

def _set_page_orientation(workbook: Any, op: dict) -> str:
    """璁剧疆椤甸潰鏂瑰悜.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        orientation: portrait/landscape
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    orientation = op.get("orientation", "portrait")

    # 1=绾靛悜 xlPortrait, 2=妯悜 xlLandscape
    sheet.PageSetup.Orientation = 1 if orientation == "portrait" else 2
    return f"set_page_orientation: {orientation}"


def _set_page_size(workbook: Any, op: dict) -> str:
    """璁剧疆椤甸潰澶у皬.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        size: A4/A3/Letter/Legal 鎴栫紪鍙?(1=Letter, 5=Legal, 9=A4, 8=A3)
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
    """璁剧疆椤佃竟璺?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        top: 涓婅竟璺?(鑻卞)
        bottom: 涓嬭竟璺?(鑻卞)
        left: 宸﹁竟璺?(鑻卞)
        right: 鍙宠竟璺?(鑻卞)
        header: 椤电湁杈硅窛 (鑻卞)
        footer: 椤佃剼杈硅窛 (鑻卞)
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
    """璁剧疆椤电湁.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        text: 椤电湁鏂囨湰 (&L宸? C涓? R鍙?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    text = op.get("text", "")
    sheet.PageSetup.CenterHeader = text
    return f"set_header: {text}"


def _set_footer(workbook: Any, op: dict) -> str:
    """璁剧疆椤佃剼.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        text: 椤佃剼鏂囨湰 (&L宸? C涓? R鍙?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    text = op.get("text", "")
    sheet.PageSetup.CenterFooter = text
    return f"set_footer: {text}"


def _add_print_title(workbook: Any, op: dict) -> str:
    """娣诲姞鎵撳嵃鏍囬 (閲嶅鎵撳嵃琛?鍒?.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        rows: 閲嶅琛?(濡?"$1:$1")
        columns: 閲嶅鍒?(濡?"$A:$A")
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
    """璁剧疆鎵撳嵃鍖哄煙.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鎵撳嵃鍖哄煙 (濡?"A1:D20")
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.PageSetup.PrintArea = range_str
    return f"set_print_area: {range_str}"


def _set_page_break(workbook: Any, op: dict) -> str:
    """璁剧疆鍒嗛〉绗?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        cell: 鍒嗛〉绗︿綅缃?(濡?"A20")
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
    """璁剧疆缂╂斁姣斾緥.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        scale: 缂╂斁姣斾緥 (10-400 鐧惧垎姣?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    scale = op.get("scale", 100)
    sheet.PageSetup.Zoom = scale
    return f"set_scale: {scale}%"


def _set_fit_to_page(workbook: Any, op: dict) -> str:
    """璁剧疆閫傚簲椤甸潰.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        fit_width: 閫傚簲瀹藉害 (1=鍗曢〉瀹? 0=鑷姩)
        fit_height: 閫傚簲楂樺害 (1=鍗曢〉楂? 0=鑷姩)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    page_setup = sheet.PageSetup

    # Must disable Zoom before setting FitToPages 鈥?Zoom and FitToPages are mutually exclusive
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


# ============ Formulas 鍏紡鎿嶄綔 (8 涓? ============

def _set_array_formula(workbook: Any, op: dict) -> str:
    """璁剧疆鏁扮粍鍏紡 (Ctrl+Shift+Enter 鍏紡).

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        formula: 鍏紡瀛楃涓?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    formula = op.get("formula", "")
    if not formula:
        raise COMOperationError("set_array_formula", "formula 涓嶈兘涓虹┖")
    sheet.Range(range_str).FormulaArray = formula
    return f"set_array_formula: {range_str} = {formula}"


def _evaluate_formula(workbook: Any, op: dict) -> Any:
    """璁＄畻骞惰繑鍥炲叕寮忕粨鏋?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        cell: 鍗曞厓鏍煎湴鍧€ (濡?A1)

    娉? 浼氬厛璋冪敤 Application.Calculate() 纭繚杩斿洖鏈€鏂板€?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    # 寮哄埗閲嶇畻, 閬垮厤 manual 妯″紡涓嬭繑鍥炶剰鍊?
    try:
        workbook.Application.Calculate()
    except Exception:
        pass
    value = sheet.Range(cell).Value
    return {"cell": cell, "value": value}


def _replace_formula(workbook: Any, op: dict) -> str:
    """鏇挎崲鑼冨洿鍐呮墍鏈夊叕寮?(鎸夊瓧绗︿覆鍖归厤).

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
        find: 鏌ユ壘瀛楃涓?
        replace: 鏇挎崲瀛楃涓?

    娉ㄦ剰: 绠€鍗曞瓙涓插尮閰? find="A1" 浼氬悓鏃跺奖鍝?"AA1" 绛夊惈 A1 瀛愪覆鐨勫叕寮?
    """
    import re
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    find = op.get("find", "")
    replace = op.get("replace", "")
    if not find:
        raise COMOperationError("replace_formula", "find 涓嶈兘涓虹┖")

    rng = sheet.Range(range_str)
    # 浣跨敤 \b 鍗曡瘝杈圭晫閬垮厤 AA1 璇尮閰?A1
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
    """鏌ユ壘鑼冨洿鍐呮墍鏈夊惈鍏紡鐨勫崟鍏冩牸.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿 (鐣欑┖浣跨敤 UsedRange)
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
    """灏嗗叕寮忚浆鎹负闈欐€佸€?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿

    娉? 鍏堝己鍒堕噸绠? 鑻ュ崟鍏冩牸鍊兼槸 Excel 閿欒 (#NAME? / #VALUE! 绛? 鍒欐嫆缁濊鐩栧師鍏紡.
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    rng = sheet.Range(range_str)
    # 寮哄埗閲嶇畻鍚庡啀璇诲彇
    try:
        workbook.Application.Calculate()
    except Exception:
        pass
    value = rng.Value
    # 妫€娴?Excel 閿欒鍊煎瓧绗︿覆
    if isinstance(value, str) and value.startswith("#") and value.endswith("!"):
        raise COMOperationError(
            "convert_to_values",
            f"范围内含有计算错误 {value}，拒绝覆盖原公式",
        )
    rng.Value = value
    return f"converted_to_values: {range_str}"


def _get_formula_info(workbook: Any, op: dict) -> dict:
    """鑾峰彇鍏紡淇℃伅 (绫诲瀷/鍊?鏄惁鏁扮粍鍏紡).

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        cell: 鍗曞厓鏍煎湴鍧€
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
    """瀹氫箟鍚嶇О (workbook level).

    Args:
        name: 鍚嶇О
        refers_to: 寮曠敤 (濡?'=Sheet1!$A$1:$A$10')
        scope: 鑼冨洿 sheet name (鍙€? 榛樿涓哄伐浣滅翱绾?
    """
    import re
    name = op.get("name", "")
    refers_to = op.get("refers_to", "")
    scope = op.get("scope", "")

    if not name or not refers_to:
        raise COMOperationError("define_name", "name 鍜?refers_to 涓嶈兘涓虹┖")

    # 鍚嶇О鍚堟硶鎬? 瀛楁瘝/涓嬪垝绾垮紑澶? 鍚庣画鍙惈瀛楁瘝鏁板瓧涓嬪垝绾?鐐瑰彿
    if not re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", name):
        raise COMOperationError(
            "define_name",
            f"鍚嶇О '{name}' 涓嶅悎娉?(椤讳互瀛楁瘝/涓嬪垝绾垮紑澶? 浠呭惈瀛楁瘝鏁板瓧涓嬪垝绾?",
        )

    # 閲嶅妫€娴?
    try:
        existing = workbook.Names(name)
        if existing is not None:
            raise COMOperationError(
                "define_name",
            f"名称 '{name}' 已存在，请先删除或更换名称",
            )
    except COMOperationError:
        raise
    except Exception:
        # 鍚嶇О涓嶅瓨鍦?(姝ｅ父)
        pass

    if scope:
        # 宸ヤ綔琛ㄧ骇鍚嶇О
        ws = workbook.Worksheets(scope)
        ws.Names.Add(Name=name, RefersTo=refers_to)
    else:
        # 宸ヤ綔绨跨骇鍚嶇О
        workbook.Names.Add(Name=name, RefersTo=refers_to)
    return f"defined_name: {name} = {refers_to}"


# ============ Tables 琛ㄦ牸 (ListObject) (8 涓? ============

def _create_table(workbook: Any, op: dict) -> str:
    """鍒涘缓 Excel 琛ㄦ牸 (ListObject).

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 琛ㄦ牸鏁版嵁鑼冨洿 (鍚〃澶?
        table_name: 琛ㄦ牸鍚嶇О
        style_name: 琛ㄦ牸鏍峰紡鍚?(濡?'TableStyleMedium2')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    table_name = op.get("table_name", "")
    style_name = op.get("style_name", "TableStyleMedium2")

    # 妫€鏌ラ噸鍚嶅苟鑷姩杩藉姞搴忓彿
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
            f"表格名 '{table_name}' 已被占用，请指定其他名称",
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
        logger.warning(f"鏃犳硶搴旂敤琛ㄦ牸鏍峰紡 {style_name}")

    return f"created_table: {table_name} ({range_str})"


def _list_tables(workbook: Any, op: dict) -> list[dict]:
    """鍒楀嚭宸ヤ綔绨夸腑鎵€鏈?Excel 琛ㄦ牸.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?(鍙€? 鐣欑┖鍒楀嚭鎵€鏈?
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
    """璋冩暣 Excel 琛ㄦ牸鑼冨洿.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        table_name: 琛ㄦ牸鍚嶇О
        range: 鏂拌寖鍥?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    range_str = op.get("range", "")
    if not table_name or not range_str:
        raise COMOperationError("resize_table", "table_name 鍜?range 涓嶈兘涓虹┖")

    tbl = sheet.ListObjects(table_name)
    tbl.Resize(sheet.Range(range_str))
    return f"resized_table: {table_name} -> {range_str}"


def _set_table_style(workbook: Any, op: dict) -> str:
    """璁剧疆琛ㄦ牸鏍峰紡.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        table_name: 琛ㄦ牸鍚嶇О
        style_name: 鏍峰紡鍚?(濡?'TableStyleLight1')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    style_name = op.get("style_name", "TableStyleLight1")
    if not table_name:
        raise COMOperationError("set_table_style", "table_name 涓嶈兘涓虹┖")

    tbl = sheet.ListObjects(table_name)
    tbl.TableStyle = style_name
    return f"set_table_style: {table_name} = {style_name}"


def _show_table_totals(workbook: Any, op: dict) -> str:
    """鏄剧ず/闅愯棌琛ㄦ牸姹囨€昏.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        table_name: 琛ㄦ牸鍚嶇О
        show: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    show = op.get("show", True)
    if not table_name:
        raise COMOperationError("show_table_totals", "table_name 涓嶈兘涓虹┖")

    tbl = sheet.ListObjects(table_name)
    tbl.ShowTotals = show
    return f"show_table_totals: {table_name} = {show}"


def _add_table_column(workbook: Any, op: dict) -> str:
    """涓鸿〃鏍兼坊鍔犺绠楀垪 (鍏紡鍒?.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        table_name: 琛ㄦ牸鍚嶇О
        column_name: 鏂板垪鍚?
        formula: 鍒楀叕寮?(濡?'=[@Qty]*[@Price]')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    column_name = op.get("column_name", "")
    formula = op.get("formula", "")
    if not table_name or not column_name:
        raise COMOperationError("add_table_column", "table_name 鍜?column_name 涓嶈兘涓虹┖")

    tbl = sheet.ListObjects(table_name)

    new_col = tbl.ListColumns.Add()
    new_col.Name = column_name
    if formula:
        # formula 椤讳互 = 寮€澶?
        if not formula.startswith("="):
            formula = "=" + formula
        if new_col.DataBodyRange is None:
            raise COMOperationError(
                "add_table_column",
                "空表无法添加计算列，需要先填入至少一行数据",
            )
        new_col.DataBodyRange.Formula = formula

    return f"added_table_column: {table_name}.{column_name}"


def _remove_table_column(workbook: Any, op: dict) -> str:
    """鍒犻櫎琛ㄦ牸鍒?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        table_name: 琛ㄦ牸鍚嶇О
        column_name: 鍒楀悕
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    column_name = op.get("column_name", "")
    if not table_name or not column_name:
        raise COMOperationError("remove_table_column", "table_name 鍜?column_name 涓嶈兘涓虹┖")

    tbl = sheet.ListObjects(table_name)
    # 鍏抽棴 totals row 闃叉骞叉壈
    try:
        tbl.ShowTotals = False
    except Exception:
        pass

    # ListColumns 閬嶅巻鏌ユ壘
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

    # 鎵句笉鍒板垯鎸夋暣鍒?Range 鍒犻櫎
    try:
        col_count = int(tbl.Range.Columns.Count)
        for i in range(1, col_count + 1):
            col_range = tbl.Range.Columns(i)
            # 澶撮儴鍗曞厓鏍?
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
    """鍒犻櫎 Excel 琛ㄦ牸 (浠呭垹闄よ〃鏍肩粨鏋? 涓嶅垹闄ゆ暟鎹?.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        table_name: 琛ㄦ牸鍚嶇О
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    table_name = op.get("table_name", "")
    if not table_name:
        raise COMOperationError("delete_table", "table_name 涓嶈兘涓虹┖")

    tbl = sheet.ListObjects(table_name)
    tbl.Unlist()  # 浠呭垹闄よ〃鏍? 淇濈暀鏁版嵁
    return f"deleted_table: {table_name}"


# ============ Data 鏁版嵁鎿嶄綔 (9 涓? ============

def _add_auto_filter(workbook: Any, op: dict) -> str:
    """娣诲姞鑷姩绛涢€?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鏁版嵁鑼冨洿 (鐣欑┖浣跨敤 UsedRange)
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
    """绉婚櫎鑷姩绛涢€?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    if sheet.AutoFilterMode:
        sheet.AutoFilterMode = False
    return f"removed_auto_filter: {sheet.Name}"


def _sort_range(workbook: Any, op: dict) -> str:
    """瀵硅寖鍥村唴鏁版嵁鎺掑簭.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鏁版嵁鑼冨洿
        key_column: 鎺掑簭鍒楀湴鍧€ (濡?'A1') 鎴栧垪鍙?(1-based)
        ascending: True 鍗囧簭 / False 闄嶅簭
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    key_column = op.get("key_column", 1)
    ascending = op.get("ascending", True)

    rng = sheet.Range(range_str)

    # 瑙ｆ瀽鎺掑簭鍒?
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
    """楂樼骇绛涢€?(灏卞湴绛涢€夋垨澶嶅埗鍒扮洰鏍囦綅缃?.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鏁版嵁鑼冨洿
        criteria_range: 鏉′欢鑼冨洿
        action: 'filter' 鍘熷湴绛涢€?/ 'copy' 澶嶅埗
        copy_to: 澶嶅埗鐩爣 (action='copy' 鏃跺繀濉?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    criteria_range = op.get("criteria_range", "")
    action = op.get("action", "filter")
    copy_to = op.get("copy_to", "")

    if not criteria_range:
        raise COMOperationError("advanced_filter", "criteria_range 涓嶈兘涓虹┖")

    rng = sheet.Range(range_str)
    crit_rng = sheet.Range(criteria_range)
    if action == "copy":
        if not copy_to:
            raise COMOperationError("advanced_filter", "copy_to 涓嶈兘涓虹┖")
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
            "advanced_filter", f"action 蹇呴』鏄?'filter' 鎴?'copy', 鏀跺埌 '{action}'"
        )
    return f"advanced_filter: {range_str} ({action})"


def _remove_duplicates(workbook: Any, op: dict) -> str:
    """鍒犻櫎閲嶅琛?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鏁版嵁鑼冨洿
        columns: 鍒ゅ畾鍒?(1-based int 鎴?'A,B,C'), 榛樿鎵€鏈夊垪
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
                        f"鍒楀彿蹇呴』涓烘鏁存暟,鏀跺埌 '{c}'",
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
    """鍒嗙骇鏄剧ず (缁勫悎琛?.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿 (濡?'A2:A5')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Group()
    return f"grouped_rows: {range_str}"


def _ungroup_rows(workbook: Any, op: dict) -> str:
    """鍙栨秷琛岀粍鍚?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Ungroup()
    return f"ungrouped_rows: {range_str}"


def _group_columns(workbook: Any, op: dict) -> str:
    """鍒嗙骇鏄剧ず (缁勫悎鍒?.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Group()
    return f"grouped_columns: {range_str}"


def _ungroup_columns(workbook: Any, op: dict) -> str:
    """鍙栨秷鍒楃粍鍚?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        range: 鑼冨洿
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1")
    sheet.Range(range_str).Ungroup()
    return f"ungrouped_columns: {range_str}"


# ============ Protection 宸ヤ綔绨夸繚鎶?(6 涓? ============

def _protect_workbook(workbook: Any, op: dict) -> str:
    """淇濇姢宸ヤ綔绨?(缁撴瀯淇濇姢).

    Args:
        password: 瀵嗙爜
        structure: 淇濇姢缁撴瀯 (榛樿 True)
        windows: 淇濇姢绐楀彛 (榛樿 False)
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
    """鎾ら攢宸ヤ綔绨夸繚鎶?

    Args:
        password: 瀵嗙爜
    """
    password = op.get("password", "")
    workbook.Unprotect(Password=password)
    return f"unprotected_workbook"


def _set_open_password(workbook: Any, op: dict) -> str:
    """璁剧疆鎵撳紑瀵嗙爜 (涓嬫 SaveAs 鏃剁湡姝ｅ姞瀵?.

    Args:
        password: 瀵嗙爜

    娉? Excel COM 涓?Password 灞炴€у彧鍦?SaveAs 鏃剁敓鏁?
        鏈嚱鏁颁粎璁剧疆灞炴€? 涓嶄富鍔?SaveAs 浠ラ伩鍏嶈鐩栨湭淇濆瓨鐨勫唴瀹?
        璇峰湪璋冪敤鏈嚱鏁板悗, 鏄惧紡璋冪敤 close_document(save=True) 瑙﹀彂鍔犲瘑钀界洏.
    """
    password = op.get("password", "")
    if not password:
        raise COMOperationError("set_open_password", "password 涓嶈兘涓虹┖")
    workbook.Password = password
    return f"set_open_password: 闀垮害 {len(password)} (灏嗗湪 SaveAs 鏃剁敓鏁?"


def _set_write_reservation_password(workbook: Any, op: dict) -> str:
    """璁剧疆鍐欎繚鎶ゅ瘑鐮?(鎺ㄨ崘鍙).

    Args:
        password: 瀵嗙爜
    """
    password = op.get("password", "")
    if not password:
        raise COMOperationError("set_write_reservation_password", "password 涓嶈兘涓虹┖")
    workbook.WriteReservationPassword = password
    return f"set_write_reservation_password: 闀垮害 {len(password)}"


def _mark_as_final(workbook: Any, op: dict) -> str:
    """鏍囪涓烘渶缁堢姸鎬?(Mark As Final).

    娉? 閫氳繃鑷畾涔夋枃妗ｅ睘鎬у疄鐜?
    """
    custom_props = workbook.CustomDocumentProperties
    # 妫€鏌ュ睘鎬ф槸鍚﹀凡瀛樺湪
    prop = None
    try:
        for i in range(1, custom_props.Count + 1):
            if custom_props(i).Name == "_MarkAsFinal":
                prop = custom_props(i)
                break
    except Exception:
        pass

    if prop is not None:
        # 宸插瓨鍦ㄥ垯鏇存柊
        prop.Value = True
    else:
        # 涓嶅瓨鍦ㄥ垯娣诲姞 (浣跨敤浣嶇疆鍙傛暟)
        custom_props.Add(
            "_MarkAsFinal",  # Name
            False,            # LinkToContent
            4,                # Type (msoPropertyTypeBoolean)
            True,             # Value
        )
    return "marked_as_final"


def _recommend_read_only(workbook: Any, op: dict) -> str:
    """璁剧疆鎺ㄨ崘鍙 (ReadOnlyRecommended).

    Args:
        recommend: True 鍚敤 / False 鍏抽棴
    """
    recommend = op.get("recommend", True)
    workbook.ReadOnlyRecommended = recommend
    return f"recommend_read_only: {recommend}"


# ============ Objects 瀵硅薄鎿嶄綔 (5 涓? ============

def _add_image(workbook: Any, op: dict) -> str:
    """鎻掑叆鍥剧墖.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        image_path: 鍥剧墖鏂囦欢璺緞
        cell: 閿氬畾鍗曞厓鏍?(榛樿 A1)
        width: 瀹藉害 (纾? 鍙€?
        height: 楂樺害 (纾? 鍙€?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    image_path = op.get("image_path", "")
    cell = op.get("cell", "A1")
    width = op.get("width")
    height = op.get("height")

    if not image_path:
        raise COMOperationError("add_image", "image_path 涓嶈兘涓虹┖")

    p = Path(image_path)
    if not p.exists():
        raise COMOperationError("add_image", f"鍥剧墖涓嶅瓨鍦? {image_path}")

    img = sheet.Pictures().Insert(str(p))
    img.Left = sheet.Range(cell).Left
    img.Top = sheet.Range(cell).Top
    if width is not None:
        img.Width = width
    if height is not None:
        img.Height = height
    return f"added_image: {p.name} at {cell}"


def _list_shapes(workbook: Any, op: dict) -> list[dict]:
    """鍒楀嚭宸ヤ綔琛ㄦ墍鏈夊舰鐘?(鍥剧墖/鏂囨湰妗?褰㈢姸).

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
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
    """鍒犻櫎褰㈢姸.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        index: 褰㈢姸绱㈠紩 (1-based)
        name: 褰㈢姸鍚嶇О (涓?index 浜岄€変竴)
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
        raise COMOperationError("delete_shape", "index 鎴?name 蹇呭～")

    sh.Delete()
    return f"deleted_shape: {sh_name}"


def _add_comment(workbook: Any, op: dict) -> str:
    """娣诲姞鎵规敞.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        cell: 鍗曞厓鏍?
        text: 鎵规敞鍐呭
        author: 浣滆€?(榛樿 'AI')
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    text = op.get("text", "")
    author = op.get("author", "AI")

    if not text:
        raise COMOperationError("add_comment", "text 涓嶈兘涓虹┖")

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
    """鍒犻櫎鎵规敞.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        cell: 鍗曞厓鏍?
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    cell = op.get("cell", "A1")
    rng = sheet.Range(cell)
    if rng.Comment:
        rng.Comment.Delete()
        return f"deleted_comment: {cell}"
    return f"no_comment: {cell}"


# ============ View 瑙嗗浘鎿嶄綔 (3 涓? ============

def _set_view_zoom(workbook: Any, op: dict) -> str:
    """璁剧疆瑙嗗浘缂╂斁.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        zoom: 缂╂斁姣斾緥 (10-400)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    zoom = op.get("zoom", 100)
    if not 10 <= zoom <= 400:
        raise COMOperationError(
            "set_view_zoom", f"zoom 蹇呴』鍦?10-400 涔嬮棿, 鏀跺埌 {zoom}"
        )
    # 閫氳繃婵€娲诲苟璁剧疆 zoom
    sheet.Activate()
    active_window = _excel_require_active_window(workbook, "set_view_zoom")
    active_window.Zoom = zoom
    return f"set_view_zoom: {zoom}%"


def _set_view_gridlines(workbook: Any, op: dict) -> str:
    """璁剧疆鏄惁鏄剧ず缃戞牸绾?

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        show: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    show = op.get("show", True)
    sheet.Activate()
    _excel_require_active_window(workbook, "set_view_gridlines").DisplayGridlines = show
    return f"set_view_gridlines: {show}"


def _set_view_headings(workbook: Any, op: dict) -> str:
    """璁剧疆鏄惁鏄剧ず琛屽垪鏍囬.

    Args:
        sheet: 宸ヤ綔琛ㄥ悕绉?
        show: True/False
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    show = op.get("show", True)
    sheet.Activate()
    _excel_require_active_window(workbook, "set_view_headings").DisplayHeadings = show
    return f"set_view_headings: {show}"


# ============ Calculation 璁＄畻鎿嶄綔 (3 涓? ============

def _recalculate(workbook: Any, op: dict) -> str:
    """閲嶆柊璁＄畻鎵€鏈夊叕寮?

    Args:
        full: True 鍏ㄩ噺閲嶇畻 / False 浠呰剰鏁版嵁 (褰撳墠瀹炵幇鍧囧叏閲?
    """
    # 鐢?Application.Calculate() 鑰屼笉鏄?workbook.Calculate()
    # 閬垮厤鏌愪簺鐜涓?workbook 琚敊璇粦瀹?
    full = op.get("full", True)
    app = workbook.Application
    if app is not None:
        # 娉? Excel COM 涓?Calculate() 鍗充负鍏ㄩ噺閲嶇畻, 娌℃湁鍗曠嫭鐨?浠呰剰鏁版嵁"鍏紑 API
        app.Calculate()
    else:
        workbook.Calculate()
    return f"recalculated: full={full}"


def _set_calculation_mode(workbook: Any, op: dict) -> str:
    """璁剧疆璁＄畻妯″紡.

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
        raise COMOperationError("set_calculation_mode", f"mode 蹇呴』鏄?auto/manual/semiauto")
    workbook.Application.Calculation = mode_map[mode]
    return f"set_calculation_mode: {mode}"


def _set_iterative_calc(workbook: Any, op: dict) -> str:
    """鍚敤/閰嶇疆杩唬璁＄畻.

    Args:
        enable: True/False
        max_iterations: 鏈€澶ц凯浠ｆ鏁?(1-32767, 榛樿 100)
        max_change: 鏈€澶у彉鍖栭噺 (>0, 榛樿 0.001)
    """
    enable = op.get("enable", True)
    max_iter = op.get("max_iterations", 100)
    max_change = op.get("max_change", 0.001)
    if not 1 <= int(max_iter) <= 32767:
        raise COMOperationError(
            "set_iterative_calc",
            f"max_iterations 蹇呴』鍦?1-32767, 鏀跺埌 {max_iter}",
        )
    if float(max_change) <= 0:
        raise COMOperationError(
            "set_iterative_calc", f"max_change 蹇呴』 > 0, 鏀跺埌 {max_change}"
        )
    app = workbook.Application
    app.Iteration = enable
    app.MaxIterations = max_iter
    app.MaxChange = max_change
    return f"set_iterative_calc: enable={enable}, max_iter={max_iter}"


def _goal_seek(workbook: Any, op: dict) -> str:
    """Run Excel Goal Seek on a target cell."""
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    set_cell = op.get("set_cell", "")
    changing_cell = op.get("changing_cell", "")
    goal = op.get("goal")

    if not set_cell or not changing_cell:
        raise COMOperationError("goal_seek", "set_cell and changing_cell are required")
    if goal is None:
        raise COMOperationError("goal_seek", "goal is required")

    try:
        result = sheet.Range(set_cell).GoalSeek(Goal=goal, ChangingCell=sheet.Range(changing_cell))
    except Exception as e:
        raise COMOperationError("goal_seek", str(e))

    return f"goal_seek: set_cell={set_cell}, changing_cell={changing_cell}, converged={bool(result)}"

