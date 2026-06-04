"""Excel MCP 工具."""

from mcp.server.fastmcp import FastMCP

from office_mcp.config import settings
from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.operations.excel_ops import (
    apply_excel_operations,
    _add_data_validation,
    _add_conditional_format,
    _merge_cells,
    _set_borders,
    _add_named_range,
    _create_pivot_table,
    _import_data,
    _export_data,
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
    def excel_add_conditional_format(file_path: str, sheet: str, range: str, condition_type: str = "cell_value", operator: str = "greater", formula1: str = "", font_color: str = "", bg_color: str = "") -> str:
        """添加条件格式.

        Args:
            file_path: 工作簿路径
            sheet: 工作表名称
            range: 范围
            condition_type: 条件类型 (cell_value/formula)
            operator: 操作符 (greater/less/equal/between)
            formula1: 条件值
            font_color: 字体颜色 (#RRGGBB)
            bg_color: 背景颜色 (#RRGGBB)
        """
        try:
            path = validate_path(file_path)
            wb = office_manager.get_document(path)
            result = _add_conditional_format(wb, {"sheet": sheet, "range": range, "type": condition_type, "operator": operator, "formula1": formula1, "font_color": font_color, "bg_color": bg_color})
            return f"已添加条件格式: {range}"
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
            file_path: 工作簿路径
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
