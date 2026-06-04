"""Word MCP 工具."""

from pathlib import Path

from mcp.server.fastmcp import FastMCP

from office_mcp.config import settings
from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.operations.word_ops import apply_word_operations
from office_mcp.operations.word_ops import (
    _apply_style, _create_style, _list_styles,
    _add_bookmark, _goto_bookmark, _delete_bookmark, _insert_at_bookmark,
    _set_header, _set_footer, _add_page_number, _add_date_time, _insert_toc, _update_toc,
    _add_hyperlink, _mail_merge,
)


def register_word_tools(mcp: FastMCP) -> None:
    """注册 Word 相关工具."""

    @mcp.tool()
    def word_create_document(file_path: str, overwrite: bool = False) -> str:
        """创建新的 Word 文档.

        Args:
            file_path: 文档保存路径 (绝对路径)
            overwrite: 是否覆盖已存在的文件
        """
        try:
            path = validate_path(file_path)
            if path.exists() and not overwrite and not settings.default_overwrite:
                return f"错误: 文件已存在，请设置 overwrite=true 覆盖: {file_path}"
            office_manager.create_document(path, overwrite=overwrite)
            return f"已创建 Word 文档: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_open_document(file_path: str) -> str:
        """打开现有 Word 文档.

        Args:
            file_path: 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            office_manager.open_document(path)
            return f"已打开 Word 文档: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_apply_operations(file_path: str, operations: list[dict]) -> dict:
        """对 Word 文档执行批量操作.

        Args:
            file_path: 文档路径 (绝对路径)
            operations: 操作列表，每个操作是一个字典，包含 type 和其他参数

        支持的操作类型:
        - add_paragraph: 添加段落 (text, style, alignment)
        - insert_table: 插入表格 (rows, columns, data)
        - replace_text: 查找替换 (find, replace)
        - insert_image: 插入图片 (image_path, width, height)
        - add_page_break: 添加分页符
        - set_page_orientation: 设置页面方向 (orientation: portrait/landscape)
        - set_font: 设置字体 (font_name, font_size)
        - set_margins: 设置页边距 (top, bottom, left, right)
        - apply_style: 应用样式 (style_name, range)
        - create_style: 创建自定义样式 (name, font_name, font_size, bold, italic)
        - list_styles: 列出所有样式
        - save: 保存文档
        """
        try:
            path = validate_path(file_path)
            # 校验所有 image_path
            for op in operations:
                if op.get("type") == "insert_image":
                    img_path = op.get("image_path", "")
                    if img_path:
                        validate_path(img_path)

            doc = office_manager.get_document(path)
            results = apply_word_operations(doc, operations)
            return {"file_path": file_path, "results": results}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_export_pdf(file_path: str, output_path: str | None = None) -> str:
        """将 Word 文档导出为 PDF.

        Args:
            file_path: 文档路径 (绝对路径)
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
    def word_close_document(file_path: str, save: bool = True) -> str:
        """关闭 Word 文档.

        Args:
            file_path: 文档路径 (绝对路径)
            save: 是否保存更改
        """
        try:
            path = validate_path(file_path)
            office_manager.close_document(path, save=save)
            action = "保存并关闭" if save else "关闭(未保存)"
            return f"{action} Word 文档: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_apply_style(file_path: str, style_name: str, range_spec: str = "all") -> str:
        """应用样式到文档范围.

        Args:
            file_path: 文档路径 (绝对路径)
            style_name: 样式名称 (如 "Heading 1", "Normal")
            range_spec: 范围 ("all" 或其他)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _apply_style(doc, {"style_name": style_name, "range": range_spec})
            return f"已应用样式: {style_name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_create_style(file_path: str, name: str, font_name: str = "Arial", font_size: int = 12, bold: bool = False, italic: bool = False) -> str:
        """创建自定义样式.

        Args:
            file_path: 文档路径 (绝对路径)
            name: 样式名称
            font_name: 字体名称
            font_size: 字号
            bold: 是否加粗
            italic: 是否斜体
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _create_style(doc, {"name": name, "font_name": font_name, "font_size": font_size, "bold": bold, "italic": italic})
            return f"已创建样式: {name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_list_styles(file_path: str) -> dict:
        """列出文档中所有可用样式.

        Args:
            file_path: 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            styles = _list_styles(doc, {})
            return {"styles": styles, "count": len(styles)}
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def word_add_bookmark(file_path: str, name: str) -> str:
        """添加书签.

        Args:
            file_path: 文档路径 (绝对路径)
            name: 书签名称
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_bookmark(doc, {"name": name})
            return f"已添加书签: {name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_goto_bookmark(file_path: str, name: str) -> str:
        """跳转到书签.

        Args:
            file_path: 文档路径 (绝对路径)
            name: 书签名称
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _goto_bookmark(doc, {"name": name})
            return f"已跳转到书签: {name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_delete_bookmark(file_path: str, name: str) -> str:
        """删除书签.

        Args:
            file_path: 文档路径 (绝对路径)
            name: 书签名称
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_bookmark(doc, {"name": name})
            return f"已删除书签: {name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_insert_at_bookmark(file_path: str, name: str, text: str) -> str:
        """在书签位置插入文本.

        Args:
            file_path: 文档路径 (绝对路径)
            name: 书签名称
            text: 要插入的文本
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _insert_at_bookmark(doc, {"name": name, "text": text})
            return f"已在书签 {name} 位置插入文本"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_header(file_path: str, text: str, section: int = 1) -> str:
        """设置页眉.

        Args:
            file_path: 文档路径 (绝对路径)
            text: 页眉文本
            section: 节号 (默认第1节)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            _set_header(doc, {"text": text, "section": section})
            return f"已设置页眉: {text[:50]}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_footer(file_path: str, text: str, section: int = 1) -> str:
        """设置页脚.

        Args:
            file_path: 文档路径 (绝对路径)
            text: 页脚文本
            section: 节号 (默认第1节)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            _set_footer(doc, {"text": text, "section": section})
            return f"已设置页脚: {text[:50]}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_page_number(file_path: str, location: str = "footer", format: str = "decimal") -> str:
        """添加页码.

        Args:
            file_path: 文档路径 (绝对路径)
            location: 位置 (header 或 footer)
            format: 数字格式 (decimal, roman_upper, roman_lower, letter_upper, letter_lower)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            _add_page_number(doc, {"location": location, "format": format})
            return f"已在 {location} 添加页码"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_insert_toc(file_path: str, heading_levels: int = 3) -> str:
        """在文档末尾插入目录.

        Args:
            file_path: 文档路径 (绝对路径)
            heading_levels: 标题级别数 (1-9，默认3)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _insert_toc(doc, {"heading_levels": heading_levels})
            return f"已插入目录 ({heading_levels} 级标题)"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_update_toc(file_path: str) -> str:
        """更新目录（刷新页码）.

        Args:
            file_path: 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _update_toc(doc, {})
            return f"已更新目录"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_hyperlink(file_path: str, text: str, url: str, screen_tip: str = "") -> str:
        """在文档末尾添加超链接.

        Args:
            file_path: 文档路径 (绝对路径)
            text: 超链接显示文本
            url: 超链接地址
            screen_tip: 鼠标悬停提示文本 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_hyperlink(doc, {"text": text, "url": url, "screen_tip": screen_tip})
            return f"已添加超链接: {text} -> {url}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_mail_merge(
        file_path: str,
        data_source: str,
        connection: str = "",
        sql_statement: str = "",
    ) -> str:
        """对 Word 文档执行邮件合并.

        使用数据源文件 (Excel/CSV 等) 将数据填充到文档的邮件合并域中.

        Args:
            file_path: 模板文档路径 (绝对路径)
            data_source: 数据源文件路径 (如 Excel 文件)
            connection: ODBC 连接字符串 (可选)
            sql_statement: SQL 查询语句 (可选，用于筛选数据)
        """
        try:
            path = validate_path(file_path)
            validate_path(data_source)
            doc = office_manager.get_document(path)
            result = _mail_merge(doc, {
                "data_source": data_source,
                "connection": connection,
                "sql_statement": sql_statement,
            })
            return f"已执行邮件合并: {data_source}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_date_time(file_path: str, format: str = "%Y年%m月%d日") -> str:
        """在文档末尾添加日期时间.

        Args:
            file_path: 文档路径 (绝对路径)
            format: 日期格式 (Python strftime 格式)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            _add_date_time(doc, {"format": format})
            return f"已添加日期时间"
        except OfficeMCPError as e:
            return f"错误: {e}"
