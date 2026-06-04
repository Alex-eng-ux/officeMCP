"""Excel MCP 工具."""

from mcp.server.fastmcp import FastMCP

from office_mcp.config import settings
from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.operations.excel_ops import (
    _add_auto_filter,
    _add_chart_series,
    _add_comment,
    _add_conditional_format,
    _add_data_validation,
    _add_image,
    _add_named_range,
    _add_print_title,
    _add_slicer,
    _add_subtotal,
    _add_table_column,
    _advanced_filter,
    _change_chart_type,
    _check_typography,
    _clear_format,
    _clear_range,
    _convert_to_values,
    _copy_format,
    _copy_range,
    _copy_worksheet,
    _create_pivot_table,
    _create_table,
    _cut_range,
    _define_name,
    _delete_cells,
    _delete_chart,
    _delete_comment,
    _delete_shape,
    _delete_table,
    _delete_worksheet,
    _evaluate_formula,
    _export_chart,
    _export_data,
    _find_formula_cells,
    _get_chart_info,
    _get_formula_info,
    _get_worksheet_info,
    _group_columns,
    _group_rows,
    _hide_rows,
    _hide_worksheet,
    _import_data,
    _insert_cells,
    _list_charts,
    _list_shapes,
    _list_tables,
    _list_used_range,
    _list_worksheets,
    _mark_as_final,
    _merge_cells,
    _move_worksheet,
    _paste_range,
    _protect_workbook,
    _protect_worksheet,
    _recalculate,
    _recommend_read_only,
    _remove_auto_filter,
    _remove_chart_series,
    _remove_duplicates,
    _remove_table_column,
    _replace_formula,
    _resize_table,
    _set_alignment,
    _set_array_formula,
    _set_borders,
    _set_calculation_mode,
    _set_chart_axis,
    _set_chart_legend,
    _set_chart_title,
    _set_column_width,
    _set_fit_to_page,
    _set_font,
    _set_font_bold,
    _set_font_italic,
    _set_font_underline,
    _set_footer,
    _set_header,
    _set_indent,
    _set_iterative_calc,
    _set_open_password,
    _set_orientation,
    _set_page_break,
    _set_page_margins,
    _set_page_orientation,
    _set_page_size,
    _set_print_area,
    _set_row_height,
    _set_scale,
    _set_tab_color,
    _set_table_style,
    _set_view_gridlines,
    _set_view_headings,
    _set_view_zoom,
    _set_wrap_text,
    _set_write_reservation_password,
    _show_table_totals,
    _show_worksheet,
    _sort_range,
    _ungroup_columns,
    _ungroup_rows,
    _unprotect_workbook,
    _unprotect_worksheet,
    apply_excel_operations,
)


def register_excel_tools(mcp: FastMCP) -> None:
    """注册 Excel 相关工具."""

    @mcp.tool()
    def excel_create_workbook(file_path: str, overwrite: bool = False) -> str:
        """创建新的 Excel 工作簿.

        Args:
            file_path: 工作簿保存路径 (绝对路径)
            overwrite: 是否覆盖已存在的文件
        """
        try:
            path = validate_path(file_path)
            if path.exists() and not overwrite and not settings.default_overwrite:
                return f"错误: 文件已存在，请设置 overwrite=true 覆盖: {file_path}"
            office_manager.create_document(path, overwrite=overwrite)
            return f"已创建 Excel 工作簿: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_open_workbook(file_path: str) -> str:
        """打开现有 Excel 工作簿.

        Args:
            file_path: 工作簿路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            office_manager.open_document(path)
            return f"已打开 Excel 工作簿: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_apply_operations(file_path: str, operations: list[dict]) -> dict:
        """对 Excel 工作簿执行批量操作.

        Args:
            file_path: 工作簿路径 (绝对路径)
            operations: 操作列表，每个操作是一个字典，包含 type 和其他参数

        支持的操作类型:
        - write_cell: 写入单元格 (sheet, cell, value)
        - write_range: 写入范围 (sheet, start_cell, data)
        - read_range: 读取范围 (sheet, range) -> 返回数据
        - add_formula: 添加公式 (sheet, cell, formula)
        - format_range: 格式化范围 (sheet, range, bold, italic, background_color, font_color)
        - set_number_format: 设置数字格式 (sheet, range, format)
        - create_chart: 创建图表 (sheet, chart_type, data_range, title, left, top, width, height)
        - add_worksheet: 添加工作表 (name)
        - rename_worksheet: 重命名工作表 (old_name, new_name)
        - auto_fit_columns: 自动调整列宽 (sheet, columns)
        - freeze_panes: 冻结窗格 (sheet, cell)
        - save: 保存工作簿
        - create_pivot_table: 创建数据透视表
        - add_slicer: 添加切片器
        - add_subtotal: 添加分类汇总
        """
        try:
            path = validate_path(file_path)
            workbook = office_manager.get_document(path)
            results = apply_excel_operations(workbook, operations)
            return {"file_path": file_path, "results": results}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def excel_export_pdf(file_path: str, output_path: str | None = None) -> str:
        """将 Excel 工作簿导出为 PDF.

        Args:
            file_path: 工作簿路径 (绝对路径)
            output_path: PDF 输出路径，默认为同目录同名 .pdf
        """
        try:
            path = validate_path(file_path)
            out_path = None
            if output_path:
                out_path = validate_path(output_path)
            result = office_manager.export_pdf(path, out_path)
            return f"已导出 PDF: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_close_workbook(file_path: str, save: bool = True) -> str:
        """关闭 Excel 工作簿.

        Args:
            file_path: 工作簿路径 (绝对路径)
            save: 是否保存更改
        """
        try:
            path = validate_path(file_path)
            office_manager.close_document(path, save=save)
            action = "保存并关闭" if save else "关闭(未保存)"
            return f"{action} Excel 工作簿: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_data_validation(file_path: str, sheet: str, range: str, validation_type: str = "list", formula1: str = "") -> str:
        """添加数据验证.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 验证范围
            validation_type: 验证类型 (list/whole/decimal/date/time/text_length/custom)
            formula1: 验证公式或逗号分隔的列表值
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_data_validation(wb, {"sheet": sheet, "range": range, "type": validation_type, "formula1": formula1})
            return f"已添加数据验证: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_conditional_format(
        file_path: str,
        sheet: str,
        range: str,
        condition_type: str = "cell_value",
        operator: str = "greater",
        formula1: str = "",
        formula2: str = "",
        format_type: str = "",
        font_color: str = "",
        bg_color: str = ""
    ) -> str:
        """添加条件格式.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            condition_type: 条件类型 (cell_value/formula)
            operator: 操作符 (greater/less/equal/between)
            formula1: 条件值1
            formula2: 条件值2
            format_type: 高级格式类型 (color_scale/data_bar/icon_set)
            font_color: 字体颜色 (#RRGGBB)
            bg_color: 背景颜色 (#RRGGBB)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_conditional_format(wb, {
                "sheet": sheet,
                "range": range,
                "type": condition_type,
                "operator": operator,
                "formula1": formula1,
                "formula2": formula2,
                "format_type": format_type,
                "font_color": font_color,
                "bg_color": bg_color
            })
            return f"已添加条件格式: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_slicer(
        file_path: str,
        target_sheet: str,
        pivot_sheet: str,
        field_name: str,
        pivot_name: str = "",
        left: int = 100,
        top: int = 100,
        width: int = 200,
        height: int = 200
    ) -> str:
        """添加切片器.

        Args:
            file_path: 工作簿路径
            target_sheet: 切片器所在工作表名称
            pivot_sheet: 数据透视表所在工作表名称
            pivot_name: 数据透视表名称 (可选)
            field_name: 要筛选的字段名称
            left: 切片器左侧位置 (像素)
            top: 切片器顶部位置 (像素)
            width: 切片器宽度 (像素)
            height: 切片器高度 (像素)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_slicer(wb, {
                "target_sheet": target_sheet,
                "pivot_sheet": pivot_sheet,
                "pivot_name": pivot_name,
                "field_name": field_name,
                "left": left,
                "top": top,
                "width": width,
                "height": height
            })
            return f"已添加切片器: {field_name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_subtotal(
        file_path: str,
        sheet: str,
        range: str,
        group_by: int = 1,
        summary_function: str = "sum",
        summary_fields: list[int] | None = None,
        replace: bool = True,
        page_breaks: bool = False,
        summary_below: bool = True
    ) -> str:
        """添加分类汇总.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 数据范围 (如 "A1:D100")
            group_by: 分组字段列号 (如 1 表示第 1 列)
            summary_function: 汇总函数 (sum/count/average/max/min)
            summary_fields: 要汇总的列号列表 (如 [3, 4])
            replace: 是否替换现有分类汇总
            page_breaks: 是否在每组后分页
            summary_below: 汇总结果是否在数据下方
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_subtotal(wb, {
                "sheet": sheet,
                "range": range,
                "group_by": group_by,
                "summary_function": summary_function,
                "summary_fields": summary_fields or [],
                "replace": replace,
                "page_breaks": page_breaks,
                "summary_below": summary_below
            })
            return f"已添加分类汇总: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_merge_cells(file_path: str, sheet: str, range: str) -> str:
        """合并单元格.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _merge_cells(wb, {"sheet": sheet, "range": range})
            return f"已合并单元格: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_borders(file_path: str, sheet: str, range: str, border_type: str = "all", style: str = "thin", color: str = "#000000") -> str:
        """设置边框.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            border_type: 边框类型 (all/outside/inside)
            style: 线型 (thin/medium/thick/dashed/dotted)
            color: 颜色 (#RRGGBB)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_borders(wb, {"sheet": sheet, "range": range, "border_type": border_type, "style": style, "color": color})
            return f"已设置边框: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_named_range(file_path: str, name: str, refers_to: str) -> str:
        """添加命名范围.

        Args:
            file_path: 工作簿路径
            name: 名称
            refers_to: 引用公式 (如 "=Sheet1!$A$1:$A$10")
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_named_range(wb, {"name": name, "refers_to": refers_to})
            return f"已添加命名范围: {name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_create_pivot_table(
        file_path: str,
        source_sheet: str = "Sheet1",
        source_range: str = "A1:D100",
        target_sheet: str = "数据透视表",
        target_cell: str = "A3",
        row_fields: list[str] | None = None,
        column_fields: list[str] | None = None,
        data_fields: dict[str, str] | None = None,
    ) -> str:
        """创建 Excel 数据透视表.

        Args:
            file_path: 工作簿路径
            source_sheet: 数据源工作表名称
            source_range: 数据源范围 (如 "A1:D100")
            target_sheet: 目标工作表名称 (自动创建或指定)
            target_cell: 目标单元格 (如 "A3")
            row_fields: 行字段列表 (如 ["部门", "月份"])
            column_fields: 列字段列表 (如 ["地区"])
            data_fields: 数据字段字典 (如 {"销售额": "sum", "数量": "average"})
                         支持的聚合函数: sum, average, count, max, min
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _create_pivot_table(wb, {
                "source_sheet": source_sheet,
                "source_range": source_range,
                "target_sheet": target_sheet,
                "target_cell": target_cell,
                "row_fields": row_fields or [],
                "column_fields": column_fields or [],
                "data_fields": data_fields or {},
            })
            return f"已创建数据透视表: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_import_data(file_path: str, csv_file: str, sheet: str = "Sheet1", start_cell: str = "A1", delimiter: str = ",") -> str:
        """导入外部数据文件 (CSV/TXT) 到工作表.

        Args:
            file_path: 工作簿路径
            csv_file: CSV/TXT 文件路径
            sheet: 目标工作表名称
            start_cell: 起始单元格
            delimiter: CSV 分隔符
        """
        try:
            path = validate_path(file_path)
            validate_path(csv_file)
            wb = office_manager.get_document(path)
            result = _import_data(wb, {
                "sheet": sheet,
                "file_path": csv_file,
                "start_cell": start_cell,
                "delimiter": delimiter,
            })
            return f"已导入数据: {csv_file}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_export_data(file_path: str, sheet: str = "Sheet1", export_path: str = "") -> str:
        """导出工作表为 CSV 文件.

        Args:
            file_path: 工作簿路径 (绝对路径)
            sheet: 源工作表名称
            export_path: 导出 CSV 文件路径
        """
        try:
            path = validate_path(file_path)
            validate_path(export_path)
            wb = office_manager.get_document(path)
            result = _export_data(wb, {
                "sheet": sheet,
                "export_path": export_path,
            })
            return f"已导出数据: {sheet} -> {export_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_check_typography(file_path: str, sheet: str = None) -> dict:
        """检查 Excel 工作簿排版问题.

        检查内容：
        - 单元格内容对齐
        - 数字格式一致性
        - 边框使用规范

        Args:
            file_path: Excel 工作簿路径 (绝对路径)
            sheet: 要检查的工作表名称 (可选，不提供则检查所有表)

        Returns:
            包含问题列表的字典
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            issues = _check_typography(wb, {"sheet": sheet})
            return {
                "file_path": file_path,
                "sheet": sheet,
                "issue_count": len(issues),
                "issues": issues
            }
        except OfficeMCPError as e:
            return {
                "file_path": file_path,
                "sheet": sheet,
                "error": str(e),
                "issue_count": 0,
                "issues": []
            }

    # ============ Worksheet 类工具 (10 个) ============

    @mcp.tool()
    def excel_list_worksheets(file_path: str) -> dict:
        """列出工作簿所有工作表.

        Args:
            file_path: 工作簿路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            sheets = _list_worksheets(wb, {})
            return {"file_path": file_path, "count": len(sheets), "worksheets": sheets}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def excel_get_worksheet_info(file_path: str, sheet: str) -> dict:
        """获取工作表信息.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            info = _get_worksheet_info(wb, {"sheet": sheet})
            return {"file_path": file_path, **info}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def excel_copy_worksheet(
        file_path: str,
        sheet: str,
        new_name: str = "",
        position: str = "",
        target_sheet: str = "",
    ) -> str:
        """复制工作表.

        Args:
            file_path: 工作簿路径
            sheet: 源工作表名称
            new_name: 新工作表名称 (可选)
            position: 位置 (before/after, 可选)
            target_sheet: 目标位置参考工作表 (可选)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _copy_worksheet(wb, {
                "sheet": sheet,
                "new_name": new_name,
                "position": position,
                "target_sheet": target_sheet,
            })
            return f"已复制工作表: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_delete_worksheet(file_path: str, sheet: str) -> str:
        """删除工作表.

        Args:
            file_path: 工作簿路径
            sheet: 要删除的工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _delete_worksheet(wb, {"sheet": sheet})
            return f"已删除工作表: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_move_worksheet(
        file_path: str, sheet: str, position: str = "last", target_sheet: str = ""
    ) -> str:
        """移动工作表.

        Args:
            file_path: 工作簿路径
            sheet: 要移动的工作表名称
            position: 位置 (first/last/before/after)
            target_sheet: 目标位置参考工作表 (position=before/after 时必填)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _move_worksheet(wb, {
                "sheet": sheet,
                "position": position,
                "target_sheet": target_sheet,
            })
            return f"已移动工作表: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_hide_worksheet(file_path: str, sheet: str) -> str:
        """隐藏工作表.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _hide_worksheet(wb, {"sheet": sheet})
            return f"已隐藏工作表: {sheet}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_show_worksheet(file_path: str, sheet: str) -> str:
        """显示工作表.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _show_worksheet(wb, {"sheet": sheet})
            return f"已显示工作表: {sheet}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_protect_worksheet(file_path: str, sheet: str, password: str = "") -> str:
        """保护工作表.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            password: 密码 (可选)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _protect_worksheet(wb, {"sheet": sheet, "password": password})
            return f"已保护工作表: {sheet}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_unprotect_worksheet(file_path: str, sheet: str, password: str = "") -> str:
        """取消工作表保护.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            password: 密码 (可选)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _unprotect_worksheet(wb, {"sheet": sheet, "password": password})
            return f"已取消保护: {sheet}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_tab_color(file_path: str, sheet: str, color: str) -> str:
        """设置工作表标签颜色.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            color: 颜色 (#RRGGBB)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_tab_color(wb, {"sheet": sheet, "color": color})
            return f"已设置标签颜色: {sheet} -> {color}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Range 类工具 (10 个) ============

    @mcp.tool()
    def excel_list_used_range(file_path: str, sheet: str) -> dict:
        """列出工作表已使用范围.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _list_used_range(wb, {"sheet": sheet})
            return {"file_path": file_path, "sheet": sheet, **result}
        except OfficeMCPError as e:
            return {"file_path": file_path, "sheet": sheet, "error": str(e)}

    @mcp.tool()
    def excel_clear_range(file_path: str, sheet: str, range: str, clear_type: str = "all") -> str:
        """清除范围内容.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            clear_type: 清除类型 (all/formulas/contents/comments)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _clear_range(wb, {"sheet": sheet, "range": range, "clear_type": clear_type})
            return f"已清除: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_copy_range(file_path: str, sheet: str, range: str) -> str:
        """复制范围到剪贴板.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _copy_range(wb, {"sheet": sheet, "range": range})
            return f"已复制: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_paste_range(file_path: str, sheet: str, target_cell: str = "A1", paste_type: str = "all") -> str:
        """粘贴范围.

        Args:
            file_path: 工作簿路径
            sheet: 目标工作表名称
            target_cell: 目标单元格
            paste_type: 粘贴类型 (all/formulas/values/formats)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _paste_range(wb, {
                "sheet": sheet,
                "target_cell": target_cell,
                "paste_type": paste_type,
            })
            return f"已粘贴: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_cut_range(file_path: str, sheet: str, range: str) -> str:
        """剪切范围.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _cut_range(wb, {"sheet": sheet, "range": range})
            return f"已剪切: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_delete_cells(file_path: str, sheet: str, range: str, shift: str = "left") -> str:
        """删除单元格.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            shift: 移动方向 (left/up)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _delete_cells(wb, {"sheet": sheet, "range": range, "shift": shift})
            return f"已删除: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_insert_cells(file_path: str, sheet: str, range: str, shift: str = "right") -> str:
        """插入单元格.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            shift: 移动方向 (right/down)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _insert_cells(wb, {"sheet": sheet, "range": range, "shift": shift})
            return f"已插入: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_row_height(file_path: str, sheet: str, row: int | str, height: float = 15.0) -> str:
        """设置行高.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            row: 行号或范围 (如 1 或 "1:3")
            height: 高度 (磅)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_row_height(wb, {"sheet": sheet, "row": row, "height": height})
            return f"已设置行高: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_column_width(file_path: str, sheet: str, column: str, width: float = 8.43) -> str:
        """设置列宽.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            column: 列标识 (如 "A" 或 "A:C")
            width: 宽度 (字符单位)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_column_width(wb, {"sheet": sheet, "column": column, "width": width})
            return f"已设置列宽: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_hide_rows(file_path: str, sheet: str, rows: str) -> str:
        """隐藏行.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            rows: 行号或范围 (如 "1" 或 "1:5")
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _hide_rows(wb, {"sheet": sheet, "rows": rows})
            return f"已隐藏: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Charts 类工具 (10 个) ============

    @mcp.tool()
    def excel_list_charts(file_path: str, sheet: str = "") -> dict:
        """列出工作表所有图表.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称 (可选, 不填则列出所有工作表的图表)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            charts = _list_charts(wb, {"sheet": sheet})
            return {"file_path": file_path, "count": len(charts), "charts": charts}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def excel_get_chart_info(file_path: str, sheet: str, chart_index: int = 1) -> dict:
        """获取图表信息.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引 (从 1 开始)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            info = _get_chart_info(wb, {"sheet": sheet, "chart_index": chart_index})
            return {"file_path": file_path, "sheet": sheet, "chart_index": chart_index, **info}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def excel_set_chart_title(file_path: str, sheet: str, chart_index: int, title: str) -> str:
        """设置图表标题.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            title: 标题文本
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_chart_title(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "title": title,
            })
            return f"已设置图表标题: {title}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_chart_legend(
        file_path: str,
        sheet: str,
        chart_index: int,
        show: bool = True,
        position: str = "bottom",
    ) -> str:
        """设置图表图例.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            show: 是否显示图例
            position: 图例位置 (bottom/top/left/right/corner)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_chart_legend(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "show": show,
                "position": position,
            })
            return f"已设置图例: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_chart_series(
        file_path: str,
        sheet: str,
        chart_index: int,
        series_name: str,
        values_range: str,
        categories_range: str = "",
    ) -> str:
        """添加图表系列.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            series_name: 系列名称
            values_range: 数值范围
            categories_range: 分类范围 (可选)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_chart_series(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "series_name": series_name,
                "values_range": values_range,
                "categories_range": categories_range,
            })
            return f"已添加系列: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_remove_chart_series(
        file_path: str, sheet: str, chart_index: int, series_index: int
    ) -> str:
        """移除图表系列.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            series_index: 系列索引 (从 1 开始)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _remove_chart_series(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "series_index": series_index,
            })
            return f"已移除系列: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_chart_axis(
        file_path: str,
        sheet: str,
        chart_index: int,
        axis: str = "x",
        title: str = "",
        min_scale: float | None = None,
        max_scale: float | None = None,
    ) -> str:
        """设置图表轴.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            axis: 轴类型 (x/y)
            title: 轴标题 (可选)
            min_scale: 最小值 (可选)
            max_scale: 最大值 (可选)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_chart_axis(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "axis": axis,
                "title": title,
                "min_scale": min_scale,
                "max_scale": max_scale,
            })
            return f"已设置轴: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_change_chart_type(
        file_path: str, sheet: str, chart_index: int, chart_type: str = "column"
    ) -> str:
        """更改图表类型.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            chart_type: 新图表类型 (column/bar/line/pie/scatter/area)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _change_chart_type(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "chart_type": chart_type,
            })
            return f"已更改图表类型: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_export_chart(file_path: str, sheet: str, chart_index: int, output_path: str) -> str:
        """导出图表为图片.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
            output_path: 输出图片路径
        """
        try:
            path = validate_path(file_path)
            out_path = validate_path(output_path)
            wb = office_manager.get_document(path)
            result = _export_chart(wb, {
                "sheet": sheet,
                "chart_index": chart_index,
                "output_path": str(out_path),
            })
            return f"已导出图表: {chart_index} -> {output_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_delete_chart(file_path: str, sheet: str, chart_index: int) -> str:
        """删除图表.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            chart_index: 图表索引
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _delete_chart(wb, {"sheet": sheet, "chart_index": chart_index})
            return f"已删除图表: {chart_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Format 类工具 (10 个) ============

    @mcp.tool()
    def excel_set_font(
        file_path: str,
        sheet: str,
        range: str,
        font_name: str = "",
        font_size: float | None = None,
        font_color: str = "",
    ) -> str:
        """设置字体.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            font_name: 字体名称 (可选)
            font_size: 字体大小 (可选)
            font_color: 字体颜色 (#RRGGBB, 可选)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_font(wb, {
                "sheet": sheet,
                "range": range,
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color,
            })
            return f"已设置字体: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_font_bold(file_path: str, sheet: str, range: str, bold: bool = True) -> str:
        """设置粗体.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            bold: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_font_bold(wb, {"sheet": sheet, "range": range, "bold": bold})
            return f"已设置粗体: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_font_italic(file_path: str, sheet: str, range: str, italic: bool = True) -> str:
        """设置斜体.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            italic: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_font_italic(wb, {"sheet": sheet, "range": range, "italic": italic})
            return f"已设置斜体: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_font_underline(
        file_path: str, sheet: str, range: str, underline: bool = True
    ) -> str:
        """设置下划线.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            underline: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_font_underline(wb, {"sheet": sheet, "range": range, "underline": underline})
            return f"已设置下划线: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_alignment(
        file_path: str,
        sheet: str,
        range: str,
        horizontal: str = "",
        vertical: str = "",
    ) -> str:
        """设置对齐.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            horizontal: 水平对齐 (left/center/right/general)
            vertical: 垂直对齐 (top/middle/bottom)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_alignment(wb, {
                "sheet": sheet,
                "range": range,
                "horizontal": horizontal,
                "vertical": vertical,
            })
            return f"已设置对齐: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_wrap_text(file_path: str, sheet: str, range: str, wrap: bool = True) -> str:
        """设置自动换行.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            wrap: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_wrap_text(wb, {"sheet": sheet, "range": range, "wrap": wrap})
            return f"已设置换行: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_indent(file_path: str, sheet: str, range: str, indent: int = 0) -> str:
        """设置缩进.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            indent: 缩进级别 (0-15)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_indent(wb, {"sheet": sheet, "range": range, "indent": indent})
            return f"已设置缩进: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_orientation(
        file_path: str, sheet: str, range: str, orientation: int = 0
    ) -> str:
        """设置文字方向.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            orientation: 角度 (0=水平, 90=垂直, 45=-45度)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_orientation(wb, {"sheet": sheet, "range": range, "orientation": orientation})
            return f"已设置文字方向: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_clear_format(file_path: str, sheet: str, range: str) -> str:
        """清除格式.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _clear_format(wb, {"sheet": sheet, "range": range})
            return f"已清除格式: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_copy_format(
        file_path: str, sheet: str, source_range: str, target_range: str
    ) -> str:
        """复制格式.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            source_range: 源格式范围
            target_range: 目标范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _copy_format(wb, {
                "sheet": sheet,
                "source_range": source_range,
                "target_range": target_range,
            })
            return f"已复制格式: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Page Setup 类工具 (10 个) ============

    @mcp.tool()
    def excel_set_page_orientation(
        file_path: str, sheet: str, orientation: str = "portrait"
    ) -> str:
        """设置页面方向.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            orientation: portrait(纵向)/landscape(横向)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_page_orientation(wb, {"sheet": sheet, "orientation": orientation})
            return f"已设置页面方向: {orientation}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_page_size(file_path: str, sheet: str, size: str = "A4") -> str:
        """设置页面大小.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            size: A4/A3/Letter/Legal 等
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_page_size(wb, {"sheet": sheet, "size": size})
            return f"已设置页面大小: {size}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_page_margins(
        file_path: str,
        sheet: str,
        top: float | None = None,
        bottom: float | None = None,
        left: float | None = None,
        right: float | None = None,
        header: float | None = None,
        footer: float | None = None,
    ) -> str:
        """设置页边距.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            top: 上边距 (英寸)
            bottom: 下边距 (英寸)
            left: 左边距 (英寸)
            right: 右边距 (英寸)
            header: 页眉边距 (英寸)
            footer: 页脚边距 (英寸)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            op = {"sheet": sheet}
            if top is not None:
                op["top"] = top
            if bottom is not None:
                op["bottom"] = bottom
            if left is not None:
                op["left"] = left
            if right is not None:
                op["right"] = right
            if header is not None:
                op["header"] = header
            if footer is not None:
                op["footer"] = footer
            result = _set_page_margins(wb, op)
            return f"已设置页边距: {sheet}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_header(file_path: str, sheet: str, text: str) -> str:
        """设置页眉.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            text: 页眉文本
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_header(wb, {"sheet": sheet, "text": text})
            return f"已设置页眉: {text}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_footer(file_path: str, sheet: str, text: str) -> str:
        """设置页脚.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            text: 页脚文本
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_footer(wb, {"sheet": sheet, "text": text})
            return f"已设置页脚: {text}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_print_title(
        file_path: str, sheet: str, rows: str = "", columns: str = ""
    ) -> str:
        """添加打印标题 (重复打印行/列).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            rows: 重复行 (如 "$1:$1")
            columns: 重复列 (如 "$A:$A")
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_print_title(wb, {"sheet": sheet, "rows": rows, "columns": columns})
            return f"已添加打印标题: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_print_area(file_path: str, sheet: str, range: str) -> str:
        """设置打印区域.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 打印区域 (如 "A1:D20")
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_print_area(wb, {"sheet": sheet, "range": range})
            return f"已设置打印区域: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_page_break(
        file_path: str, sheet: str, cell: str = "A20", break_type: str = "row"
    ) -> str:
        """设置分页符.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            cell: 分页符位置 (如 "A20")
            break_type: row(行)/column(列)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_page_break(wb, {"sheet": sheet, "cell": cell, "break_type": break_type})
            return f"已设置分页符: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_scale(file_path: str, sheet: str, scale: int = 100) -> str:
        """设置缩放比例.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            scale: 缩放比例 (10-400 百分比)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_scale(wb, {"sheet": sheet, "scale": scale})
            return f"已设置缩放: {scale}%"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_fit_to_page(
        file_path: str, sheet: str, fit_width: int = 1, fit_height: int = 0
    ) -> str:
        """设置适应页面.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            fit_width: 适应宽度 (1=单页宽, 0=自动)
            fit_height: 适应高度 (1=单页高, 0=自动)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_fit_to_page(wb, {
                "sheet": sheet,
                "fit_width": fit_width,
                "fit_height": fit_height,
            })
            return f"已设置适应页面: {fit_width}x{fit_height}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Formulas 公式 (7) ============

    @mcp.tool()
    def excel_set_array_formula(file_path: str, sheet: str, range: str, formula: str) -> str:
        """设置数组公式.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            formula: 公式字符串
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_array_formula(wb, {"sheet": sheet, "range": range, "formula": formula})
            return f"已设置数组公式: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_evaluate_formula(file_path: str, sheet: str, cell: str) -> str:
        """计算并返回公式结果.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            cell: 单元格地址
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _evaluate_formula(wb, {"sheet": sheet, "cell": cell})
            return f"{cell} = {result['value']}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_replace_formula(
        file_path: str, sheet: str, range: str, find: str, replace: str
    ) -> str:
        """替换范围内公式 (按字符串匹配).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            find: 查找字符串
            replace: 替换字符串
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _replace_formula(wb, {
                "sheet": sheet, "range": range,
                "find": find, "replace": replace,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_find_formula_cells(file_path: str, sheet: str, range: str = "A1:Z1000") -> str:
        """查找范围内所有含公式的单元格.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _find_formula_cells(wb, {"sheet": sheet, "range": range})
            import json
            return json.dumps(result, ensure_ascii=False, default=str)
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_convert_to_values(file_path: str, sheet: str, range: str) -> str:
        """将公式转换为静态值.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _convert_to_values(wb, {"sheet": sheet, "range": range})
            return f"已转换为值: {range}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_get_formula_info(file_path: str, sheet: str, cell: str) -> str:
        """获取公式信息 (类型/值/是否数组公式).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            cell: 单元格地址
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _get_formula_info(wb, {"sheet": sheet, "cell": cell})
            import json
            return json.dumps(result, ensure_ascii=False, default=str)
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_define_name(file_path: str, name: str, refers_to: str) -> str:
        """定义工作簿级名称.

        Args:
            file_path: 工作簿路径
            name: 名称
            refers_to: 引用表达式 (如 '=Sheet1!$A$1:$A$10')
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _define_name(wb, {"name": name, "refers_to": refers_to})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Tables 表格 (8) ============

    @mcp.tool()
    def excel_create_table(
        file_path: str, sheet: str, range: str,
        table_name: str = "", style_name: str = "TableStyleMedium2",
    ) -> str:
        """创建 Excel 表格 (ListObject).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 表格数据范围 (含表头)
            table_name: 表格名称
            style_name: 表格样式名
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            op = {"sheet": sheet, "range": range, "style_name": style_name}
            if table_name:
                op["table_name"] = table_name
            result = _create_table(wb, op)
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_list_tables(file_path: str, sheet: str = "") -> str:
        """列出工作簿中所有 Excel 表格.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称 (可选, 留空列出所有)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _list_tables(wb, {"sheet": sheet})
            import json
            return json.dumps(result, ensure_ascii=False, default=str)
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_resize_table(file_path: str, sheet: str, table_name: str, range: str) -> str:
        """调整 Excel 表格范围.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            table_name: 表格名称
            range: 新范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _resize_table(wb, {
                "sheet": sheet, "table_name": table_name, "range": range,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_table_style(file_path: str, sheet: str, table_name: str, style_name: str) -> str:
        """设置表格样式.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            table_name: 表格名称
            style_name: 样式名
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_table_style(wb, {
                "sheet": sheet, "table_name": table_name, "style_name": style_name,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_show_table_totals(file_path: str, sheet: str, table_name: str, show: bool = True) -> str:
        """显示/隐藏表格汇总行.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            table_name: 表格名称
            show: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _show_table_totals(wb, {
                "sheet": sheet, "table_name": table_name, "show": show,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_table_column(
        file_path: str, sheet: str, table_name: str, column_name: str, formula: str = ""
    ) -> str:
        """为表格添加计算列.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            table_name: 表格名称
            column_name: 新列名
            formula: 列公式
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_table_column(wb, {
                "sheet": sheet, "table_name": table_name,
                "column_name": column_name, "formula": formula,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_remove_table_column(file_path: str, sheet: str, table_name: str, column_name: str) -> str:
        """删除表格列.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            table_name: 表格名称
            column_name: 列名
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _remove_table_column(wb, {
                "sheet": sheet, "table_name": table_name, "column_name": column_name,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_delete_table(file_path: str, sheet: str, table_name: str) -> str:
        """删除 Excel 表格 (仅删除表格结构, 保留数据).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            table_name: 表格名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _delete_table(wb, {
                "sheet": sheet, "table_name": table_name,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Data 数据操作 (9) ============

    @mcp.tool()
    def excel_add_auto_filter(file_path: str, sheet: str, range: str = "") -> str:
        """添加自动筛选.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 数据范围 (留空使用 UsedRange)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_auto_filter(wb, {"sheet": sheet, "range": range})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_remove_auto_filter(file_path: str, sheet: str) -> str:
        """移除自动筛选.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _remove_auto_filter(wb, {"sheet": sheet})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_sort_range(
        file_path: str, sheet: str, range: str,
        key_column: str = "1", ascending: bool = True,
    ) -> str:
        """对范围内数据排序.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 数据范围
            key_column: 排序列号 (1-based) 或列地址
            ascending: True 升序 / False 降序
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            # 尝试转 int
            try:
                key_int = int(key_column)
            except (ValueError, TypeError):
                key_int = key_column
            result = _sort_range(wb, {
                "sheet": sheet, "range": range,
                "key_column": key_int, "ascending": ascending,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_advanced_filter(
        file_path: str, sheet: str, range: str, criteria_range: str,
        action: str = "filter", copy_to: str = "",
    ) -> str:
        """高级筛选 (就地筛选或复制到目标位置).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 数据范围
            criteria_range: 条件范围
            action: 'filter' 原地筛选 / 'copy' 复制
            copy_to: 复制目标 (action='copy' 时必填)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _advanced_filter(wb, {
                "sheet": sheet, "range": range,
                "criteria_range": criteria_range,
                "action": action, "copy_to": copy_to,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_remove_duplicates(file_path: str, sheet: str, range: str, columns: str = "") -> str:
        """删除重复行.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 数据范围
            columns: 判定列 (逗号分隔列号, 如 '1,2' 或留空为所有列)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _remove_duplicates(wb, {
                "sheet": sheet, "range": range, "columns": columns,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_group_rows(file_path: str, sheet: str, range: str) -> str:
        """分级显示 (组合行).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _group_rows(wb, {"sheet": sheet, "range": range})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_ungroup_rows(file_path: str, sheet: str, range: str) -> str:
        """取消行组合.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _ungroup_rows(wb, {"sheet": sheet, "range": range})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_group_columns(file_path: str, sheet: str, range: str) -> str:
        """分级显示 (组合列).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _group_columns(wb, {"sheet": sheet, "range": range})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_ungroup_columns(file_path: str, sheet: str, range: str) -> str:
        """取消列组合.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _ungroup_columns(wb, {"sheet": sheet, "range": range})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Protection 工作簿保护 (6) ============

    @mcp.tool()
    def excel_protect_workbook(file_path: str, password: str = "", structure: bool = True, windows: bool = False) -> str:
        """保护工作簿 (结构保护).

        Args:
            file_path: 工作簿路径
            password: 密码
            structure: 保护结构 (默认 True)
            windows: 保护窗口 (默认 False)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _protect_workbook(wb, {
                "password": password, "structure": structure, "windows": windows,
            })
            return f"已保护工作簿: structure={structure}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_unprotect_workbook(file_path: str, password: str = "") -> str:
        """撤销工作簿保护.

        Args:
            file_path: 工作簿路径
            password: 密码
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _unprotect_workbook(wb, {"password": password})
            return "已撤销工作簿保护"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_open_password(file_path: str, password: str) -> str:
        """设置打开密码 (在 Save As 加密后生效).

        Args:
            file_path: 工作簿路径
            password: 密码
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_open_password(wb, {"password": password})
            return "已设置打开密码 (调用 SaveAs 时生效)"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_write_reservation_password(file_path: str, password: str) -> str:
        """设置写保护密码 (推荐只读).

        Args:
            file_path: 工作簿路径
            password: 密码
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_write_reservation_password(wb, {"password": password})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_mark_as_final(file_path: str) -> str:
        """标记为最终状态 (Mark As Final).

        Args:
            file_path: 工作簿路径
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _mark_as_final(wb, {})
            return "已标记为最终状态"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_recommend_read_only(file_path: str, recommend: bool = True) -> str:
        """设置推荐只读 (ReadOnlyRecommended).

        Args:
            file_path: 工作簿路径
            recommend: True 启用 / False 关闭
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _recommend_read_only(wb, {"recommend": recommend})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Objects 对象操作 (5) ============

    @mcp.tool()
    def excel_add_image(
        file_path: str, sheet: str, image_path: str,
        cell: str = "A1", width: float | None = None, height: float | None = None,
    ) -> str:
        """插入图片.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            image_path: 图片文件路径
            cell: 锚定单元格
            width: 宽度 (磅, 可选)
            height: 高度 (磅, 可选)
        """
        try:
            path = validate_path(file_path)
            img_path = validate_path(image_path)  # 关键: 校验图片路径
            wb = office_manager.get_document(path)
            op = {"sheet": sheet, "image_path": str(img_path), "cell": cell}
            if width is not None:
                op["width"] = width
            if height is not None:
                op["height"] = height
            result = _add_image(wb, op)
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_list_shapes(file_path: str, sheet: str) -> str:
        """列出工作表所有形状 (图片/文本框/形状).

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _list_shapes(wb, {"sheet": sheet})
            import json
            return json.dumps(result, ensure_ascii=False, default=str)
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_delete_shape(file_path: str, sheet: str, index: int = 0, name: str = "") -> str:
        """删除形状.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            index: 形状索引 (1-based)
            name: 形状名称 (与 index 二选一)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            op: dict = {"sheet": sheet}
            if index > 0:
                op["index"] = index
            elif name:
                op["name"] = name
            else:
                return "错误: index 或 name 必填其一"
            result = _delete_shape(wb, op)
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_add_comment(file_path: str, sheet: str, cell: str, text: str, author: str = "AI") -> str:
        """添加批注.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            cell: 单元格
            text: 批注内容
            author: 作者
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_comment(wb, {
                "sheet": sheet, "cell": cell, "text": text, "author": author,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_delete_comment(file_path: str, sheet: str, cell: str) -> str:
        """删除批注.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            cell: 单元格
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _delete_comment(wb, {"sheet": sheet, "cell": cell})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ View 视图 (3) ============

    @mcp.tool()
    def excel_set_view_zoom(file_path: str, sheet: str, zoom: int = 100) -> str:
        """设置视图缩放.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            zoom: 缩放比例 (10-400)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_view_zoom(wb, {"sheet": sheet, "zoom": zoom})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_view_gridlines(file_path: str, sheet: str, show: bool = True) -> str:
        """设置是否显示网格线.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            show: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_view_gridlines(wb, {"sheet": sheet, "show": show})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_view_headings(file_path: str, sheet: str, show: bool = True) -> str:
        """设置是否显示行列标题.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            show: True/False
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_view_headings(wb, {"sheet": sheet, "show": show})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Calculation 计算 (3) ============

    @mcp.tool()
    def excel_recalculate(file_path: str, full: bool = True) -> str:
        """重新计算所有公式.

        Args:
            file_path: 工作簿路径
            full: True 全量重算 / False 仅脏数据
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _recalculate(wb, {"full": full})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_calculation_mode(file_path: str, mode: str = "auto") -> str:
        """设置计算模式.

        Args:
            file_path: 工作簿路径
            mode: auto / manual / semiauto
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_calculation_mode(wb, {"mode": mode})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def excel_set_iterative_calc(
        file_path: str, enable: bool = True,
        max_iterations: int = 100, max_change: float = 0.001,
    ) -> str:
        """启用/配置迭代计算.

        Args:
            file_path: 工作簿路径
            enable: True/False
            max_iterations: 最大迭代次数
            max_change: 最大变化量
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _set_iterative_calc(wb, {
                "enable": enable,
                "max_iterations": max_iterations,
                "max_change": max_change,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"
