"""Excel COM 操作实现."""

import logging
from pathlib import Path
from typing import Any

from office_mcp.core.errors import COMOperationError

logger = logging.getLogger(__name__)

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
    else:
        raise COMOperationError(f"未知的 Excel 操作类型: {op_type}")


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

    validation = range_obj.Validation.Add(
        Type=validation_type_val,
        AlertStyle=1,  # xlValidAlertStop
        Operator=0,
        Formula1=formula1,
    )
    return f"added_data_validation: {range_str} ({validation_type})"


def _add_conditional_format(workbook: Any, op: dict) -> str:
    """添加条件格式.

    Args:
        sheet: 工作表名称
        range: 范围 (如 "A1:A10")
        type: 条件类型 (cell_value/formula)
        operator: 操作符 (greater/less/equal/between)
        formula1: 条件值1
        formula2: 条件值2
        format_type: 格式类型 (color_scale/data_bar)
        font_color: 字体颜色 (#RRGGBB)
        bg_color: 背景颜色 (#RRGGBB)
    """
    sheet = _get_sheet(workbook, op.get("sheet", "Sheet1"))
    range_str = op.get("range", "A1:A10")
    condition_type = op.get("type", "cell_value")
    operator = op.get("operator", "greater")
    formula1 = op.get("formula1", "")
    font_color = op.get("font_color", "")
    bg_color = op.get("bg_color", "")

    range_obj = sheet.Range(range_str)

    if condition_type == "cell_value":
        # 操作符映射
        op_map = {
            "greater": 5,      # xlGreater
            "less": 4,         # xlLess
            "equal": 3,        # xlEqual
            "between": 1,      # xlBetween
            "greater_equal": 7,
            "less_equal": 6,
        }
        operator_val = op_map.get(operator, 5)

        format_condition = range_obj.FormatConditions.Add(
            Type=1,  # xlCellValue
            Operator=operator_val,
            Formula1=formula1,
        )
    else:  # formula
        format_condition = range_obj.FormatConditions.Add(
            Type=2,  # xlExpression
            Formula1=formula1,
        )

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
        for border in [7, 8, 9, 10, 11, 12]:  # xlEdgeTop/Bottom/Left/Right/InsideHorizontal/InsideVertical
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
        for named_range in workbook.Names:
            if named_range.Name == name:
                named_range.Delete()
                break

        workbook.Names.Add(Name=name, RefersToR1C1=refers_to)
        return f"added_named_range: {name}"
    except Exception as e:
        raise COMOperationError("add_named_range", str(e))


def _create_pivot_table(workbook: Any, op: dict) -> str:
    """创建数据透视表.

    Args:
        source_sheet: 数据源工作表名称
        source_range: 数据源范围 (如 "A1:D100")
        target_sheet: 目标工作表名称 (自动创建或指定)
        target_cell: 目标单元格 (如 "A3")
        row_fields: 行字段列表 (如 ["部门", "月份"])
        column_fields: 列字段列表 (如 ["地区"])
        data_fields: 数据字段字典 (如 {"销售额": "sum", "数量": "average"})
    """
    source_sheet = _get_sheet(workbook, op.get("source_sheet", "Sheet1"))
    source_range = op.get("source_range", "A1:D100")
    target_sheet_name = op.get("target_sheet", "数据透视表")
    target_cell = op.get("target_cell", "A3")

    row_fields = op.get("row_fields", [])
    column_fields = op.get("column_fields", [])
    data_fields = op.get("data_fields", {})

    # 创建或获取目标工作表
    try:
        target_sheet = workbook.Worksheets(target_sheet_name)
    except Exception:
        target_sheet = workbook.Worksheets.Add()
        target_sheet.Name = target_sheet_name

    # 创建数据透视表缓存 (使用地址字符串更可靠)
    source_data_addr = source_sheet.Range(source_range).Address(External=True)
    pivot_cache = workbook.PivotCaches.Create(
        SourceType=1,  # xlDatabase
        SourceData=source_data_addr,
    )

    # 生成不重复的表名
    import time
    table_name = f"PivotTable_{int(time.time())}"

    # 创建数据透视表
    pivot_table = pivot_cache.CreatePivotTable(
        TableDestination=target_sheet.Range(target_cell),
        TableName=table_name,
    )

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
