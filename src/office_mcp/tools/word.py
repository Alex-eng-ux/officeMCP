"""Word MCP 工具."""

from pathlib import Path

from office_mcp.compat import FastMCP
from office_mcp.config import settings
from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.operations.word_ops import (
    _accept_all_changes,
    _add_bookmark,
    _add_bullet,
    _add_column,
    _add_comment,
    _add_date_time,
    _add_field,
    _add_hyperlink,
    _add_page_number,
    _add_row,
    _add_section_break,
    _add_smartart,
    # Styles 类
    _apply_bold,
    _apply_italic,
    _apply_style,
    _apply_underline,
    _check_typography,
    _compare_documents,
    _create_style,
    _crop_image,
    _delete_bookmark,
    _delete_column,
    _delete_comment,
    _delete_field,
    _delete_image,
    _delete_row,
    _delete_section,
    _get_cell,
    # Document 类
    _get_document_info,
    _get_document_properties,
    _get_field_result,
    _get_hyperlink,
    _get_image_info,
    _get_paragraph,
    _get_table_info,
    _goto_bookmark,
    _insert_at_bookmark,
    _insert_icon_to_word,
    _insert_toc,
    _list_comments,
    # Fields
    _list_fields,
    # Hyperlinks
    _list_hyperlinks,
    # Images
    _list_images,
    # Paragraphs 类
    _list_paragraphs,
    # Sections
    _list_sections,
    _list_styles,
    # Tables 类
    _list_tables,
    _mail_merge,
    _merge_cells,
    _merge_paragraphs,
    _reject_all_changes,
    _remove_bullet,
    _remove_hyperlink,
    _replace_image,
    _resize_image,
    _set_cell,
    _set_document_properties,
    # Protection
    _set_document_protection,
    _set_font_color,
    _set_font_size,
    _set_footer,
    _set_header,
    _set_highlight_color,
    _set_image_position,
    _set_image_wrap,
    _set_indent,
    _set_line_spacing,
    _set_paragraph_alignment,
    _set_paragraph_spacing,
    _set_password,
    _set_read_only,
    _set_section_columns,
    _set_section_margins,
    _set_section_orientation,
    _set_strikethrough,
    _set_subscript_superscript,
    _set_table_borders,
    _set_table_style,
    _split_cell,
    _split_paragraph,
    _track_changes,
    _unprotect_document,
    _update_field,
    _update_fields,
    _update_toc,
    apply_word_operations,
)
from office_mcp.utils.icons import search_icons


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
        - add_smartart: 添加 SmartArt (type: org_chart/cycle/pyramid/process/list, text_list: [])
        - add_field: 添加域 (field_type: DATE/TIME/PAGE/NUMPAGES/REF/AUTHOR/TITLE/FILENAME/HYPERLINK, code: "", preserve_formatting: bool)
        - update_fields: 更新所有域
        - mail_merge: 邮件合并 (data_source, output_path: 可选)
        - list_sections: 列出分节
        - add_section_break: 添加分节符 (break_type)
        - delete_section: 删除分节 (section)
        - set_section_orientation: 设置分节方向 (section, orientation)
        - set_section_margins: 设置分节页边距 (section, top, bottom, left, right)
        - set_section_columns: 设置分节列数 (section, columns, space_between)
        - list_fields: 列出域
        - update_field: 更新单个域 (field_index)
        - delete_field: 删除域 (field_index, keep_result)
        - get_field_result: 获取域结果 (field_index)
        - list_images: 列出图片
        - get_image_info: 获取图片信息 (image_index)
        - resize_image: 调整图片大小 (image_index, width, height, lock_aspect_ratio)
        - crop_image: 裁剪图片 (image_index, crop_left/right/top/bottom)
        - set_image_position: 设置图片位置 (image_index, left, top)
        - set_image_wrap: 设置图片文字环绕 (image_index, wrap_type)
        - replace_image: 替换图片 (image_index, new_image_path)
        - delete_image: 删除图片 (image_index)
        - list_hyperlinks: 列出超链接
        - get_hyperlink: 获取超链接信息 (hyperlink_index)
        - remove_hyperlink: 移除超链接 (hyperlink_index)
        - set_document_protection: 设置文档保护 (protection_type, password)
        - unprotect_document: 取消保护 (password)
        - set_read_only: 设置只读 (read_only)
        - set_password: 设置打开密码 (password)
        - accept_all_changes: 接受所有修订
        - reject_all_changes: 拒绝所有修订
        """
        try:
            path = validate_path(file_path)
            # 校验所有 image_path
            for op in operations:
                if op.get("type") == "insert_image":
                    img_path = op.get("image_path", "")
                    if img_path:
                        validate_path(img_path)

            doc = office_manager.ensure_document(path, activate=True)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            doc = office_manager.ensure_document(path, activate=False)
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
            data_path = validate_path(data_source)
            doc = office_manager.ensure_document(path, activate=True)
            result = _mail_merge(doc, {
                "data_source": str(data_path),
                "connection": connection,
                "sql_statement": sql_statement,
                "send_to_new_document": False,
            })
            return f"已执行邮件合并: {result}"
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

    @mcp.tool()
    def word_search_icons(query: str, limit: int = 10) -> dict:
        """搜索 Google Material Icons (2500+ 图标).

        Args:
            query: 搜索关键词 (如 "star", "check", "home", "arrow")
            limit: 返回结果数量上限 (默认 10)
        """
        try:
            results = search_icons(query, limit=limit)
            return {"icons": results, "count": len(results)}
        except Exception as e:
            return {"error": str(e), "icons": [], "count": 0}

    @mcp.tool()
    def word_insert_icon(
        file_path: str,
        icon_name: str = "",
        query: str = "",
        width: float = 24,
        height: float = 24,
        fill_color: str = "#000000",
        size: int = 24,
    ) -> str:
        """搜索并插入 Material Icon 到 Word 文档末尾.

        Args:
            file_path: 文档路径
            icon_name: 图标名称 (如 "star", "check", "home")
            query: 搜索关键词 (如果不知道具体名称，会自动搜索并取第一个)
            width: 宽度 (pt)
            height: 高度 (pt)
            fill_color: 填充颜色 (#RRGGBB)
            size: 图标尺寸 (px, 默认 24)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _insert_icon_to_word(doc, {
                "icon_name": icon_name,
                "query": query,
                "width": width,
                "height": height,
                "fill_color": fill_color,
                "size": size,
            })
            return f"已插入图标: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_smartart(
        file_path: str,
        smartart_type: str = "list",
        text_list: list[str] | None = None,
    ) -> str:
        """在 Word 文档末尾添加 SmartArt 图形.

        Args:
            file_path: 文档路径 (绝对路径)
            smartart_type: SmartArt 类型 (org_chart, cycle, pyramid, process, list)
            text_list: 文本列表，用于填充图形内容 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_smartart(doc, {
                "type": smartart_type,
                "text_list": text_list or [],
            })
            return f"已添加 SmartArt: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_field(
        file_path: str,
        field_type: str = "DATE",
        code: str = "",
        preserve_formatting: bool = True,
    ) -> str:
        """在 Word 文档末尾添加域.

        Args:
            file_path: 文档路径 (绝对路径)
            field_type: 域类型 (DATE, TIME, PAGE, NUMPAGES, REF)
            code: 域代码 (可选)
            preserve_formatting: 是否保留格式 (默认 True)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_field(doc, {
                "field_type": field_type,
                "code": code,
                "preserve_formatting": preserve_formatting,
            })
            return f"已添加域: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_update_fields(file_path: str, update_links: bool = False) -> str:
        """更新 Word 文档中的所有域.

        Args:
            file_path: 文档路径 (绝对路径)
            update_links: 是否更新链接 (默认 False)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _update_fields(doc, {
                "update_links": update_links,
            })
            return f"已更新域: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_mail_merge_enhanced(
        file_path: str,
        data_source: str,
        connection: str = "",
        sql_statement: str = "",
        output_path: str = "",
    ) -> str:
        """对 Word 文档执行邮件合并（增强版，支持输出到新文档）.

        使用数据源文件 (Excel/CSV 等) 将数据填充到文档的邮件合并域中.

        Args:
            file_path: 模板文档路径 (绝对路径)
            data_source: 数据源文件路径 (如 Excel 文件)
            connection: ODBC 连接字符串 (可选)
            sql_statement: SQL 查询语句 (可选，用于筛选数据)
            output_path: 输出到新文档的路径 (可选，不提供则合并到当前)
        """
        try:
            path = validate_path(file_path)
            data_path = validate_path(data_source)
            out_path = ""
            if output_path:
                out_path = str(validate_path(output_path))
            doc = office_manager.ensure_document(path, activate=True)
            result = _mail_merge(doc, {
                "data_source": str(data_path),
                "connection": connection,
                "sql_statement": sql_statement,
                "output_path": out_path,
                "send_to_new_document": True,
            })
            return f"已执行邮件合并: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_check_typography(file_path: str) -> dict:
        """检查 Word 文档排版问题.

        检查内容：
        - 孤立行和短行
        - 标点符号前后空格
        - 标题层级一致性

        Args:
            file_path: Word 文档路径 (绝对路径)

        Returns:
            包含问题列表的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.ensure_document(path, activate=False)
            issues = _check_typography(doc, {})
            return {
                "file_path": file_path,
                "issue_count": len(issues),
                "issues": issues
            }
        except OfficeMCPError as e:
            return {
                "file_path": file_path,
                "error": str(e),
                "issue_count": 0,
                "issues": []
            }

    # ================================================================
    # Sections 类工具 (6 个)
    # ================================================================

    @mcp.tool()
    def word_list_sections(file_path: str) -> dict:
        """列出文档中的所有分节.

        Args:
            file_path: Word 文档路径 (绝对路径)

        Returns:
            包含分节列表的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            sections = _list_sections(doc, {})
            return {"file_path": file_path, "sections": sections, "count": len(sections)}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "sections": [], "count": 0}

    @mcp.tool()
    def word_add_section_break(file_path: str, break_type: str = "next_page") -> str:
        """在文档末尾添加分节符.

        Args:
            file_path: Word 文档路径 (绝对路径)
            break_type: 分节符类型 (next_page, continuous, even_page, odd_page, column)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_section_break(doc, {"break_type": break_type})
            return f"已添加分节符: {break_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_delete_section(file_path: str, section: int) -> str:
        """删除指定分节.

        Args:
            file_path: Word 文档路径 (绝对路径)
            section: 分节索引 (从1开始)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_section(doc, {"section": section})
            return f"已删除分节: {section}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_section_orientation(
        file_path: str,
        orientation: str = "portrait",
        section: int = 1,
    ) -> str:
        """设置指定分节的页面方向.

        Args:
            file_path: Word 文档路径 (绝对路径)
            orientation: 页面方向 (portrait 或 landscape)
            section: 分节索引 (默认1)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_section_orientation(doc, {
                "section": section,
                "orientation": orientation,
            })
            return f"已设置分节 {section} 方向: {orientation}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_section_margins(
        file_path: str,
        section: int = 1,
        top: float | None = None,
        bottom: float | None = None,
        left: float | None = None,
        right: float | None = None,
    ) -> str:
        """设置指定分节的页边距 (单位: 磅).

        Args:
            file_path: Word 文档路径 (绝对路径)
            section: 分节索引 (默认1)
            top: 上边距 (磅)
            bottom: 下边距 (磅)
            left: 左边距 (磅)
            right: 右边距 (磅)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op = {"section": section}
            if top is not None:
                op["top"] = top
            if bottom is not None:
                op["bottom"] = bottom
            if left is not None:
                op["left"] = left
            if right is not None:
                op["right"] = right
            result = _set_section_margins(doc, op)
            return f"已设置分节 {section} 页边距"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_section_columns(
        file_path: str,
        columns: int = 1,
        section: int = 1,
        space_between: float = 36,
    ) -> str:
        """设置指定分节的列数.

        Args:
            file_path: Word 文档路径 (绝对路径)
            columns: 列数
            section: 分节索引 (默认1)
            space_between: 列间距 (磅, 默认 36)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_section_columns(doc, {
                "section": section,
                "columns": columns,
                "space_between": space_between,
            })
            return f"已设置分节 {section} 列数: {columns}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Fields 类工具 (5 个新, 已有 word_add_field/word_update_fields 扩展)
    # ================================================================

    @mcp.tool()
    def word_list_fields(file_path: str) -> dict:
        """列出文档中的所有域.

        Args:
            file_path: Word 文档路径 (绝对路径)

        Returns:
            包含域列表的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            fields = _list_fields(doc, {})
            return {"file_path": file_path, "fields": fields, "count": len(fields)}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "fields": [], "count": 0}

    @mcp.tool()
    def word_update_field(file_path: str, field_index: int) -> str:
        """更新单个域.

        Args:
            file_path: Word 文档路径 (绝对路径)
            field_index: 域索引 (从1开始)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _update_field(doc, {"field_index": field_index})
            return f"已更新域: {field_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_delete_field(file_path: str, field_index: int, keep_result: bool = False) -> str:
        """删除域.

        Args:
            file_path: Word 文档路径 (绝对路径)
            field_index: 域索引 (从1开始)
            keep_result: 是否保留域结果文本 (默认 False)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_field(doc, {
                "field_index": field_index,
                "keep_result": keep_result,
            })
            return f"已删除域: {field_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_get_field_result(file_path: str, field_index: int) -> dict:
        """获取域结果.

        Args:
            file_path: Word 文档路径 (绝对路径)
            field_index: 域索引 (从1开始)

        Returns:
            包含域代码和结果的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _get_field_result(doc, {"field_index": field_index})
            return {"file_path": file_path, **result}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    # ================================================================
    # Images 类工具 (8 个)
    # ================================================================

    @mcp.tool()
    def word_list_images(file_path: str) -> dict:
        """列出文档中的所有图片 (内联与浮动).

        Args:
            file_path: Word 文档路径 (绝对路径)

        Returns:
            包含图片列表的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            images = _list_images(doc, {})
            return {"file_path": file_path, "images": images, "count": len(images)}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "images": [], "count": 0}

    @mcp.tool()
    def word_get_image_info(file_path: str, image_index: int) -> dict:
        """获取指定图片的详细信息.

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引 (从1开始)

        Returns:
            包含图片详细信息的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _get_image_info(doc, {"image_index": image_index})
            return {"file_path": file_path, **result}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_resize_image(
        file_path: str,
        image_index: int,
        width: float | None = None,
        height: float | None = None,
        lock_aspect_ratio: bool = True,
    ) -> str:
        """调整图片大小.

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引 (从1开始)
            width: 宽度 (磅)
            height: 高度 (磅)
            lock_aspect_ratio: 锁定纵横比 (默认 True)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op = {
                "image_index": image_index,
                "lock_aspect_ratio": lock_aspect_ratio,
            }
            if width is not None:
                op["width"] = width
            if height is not None:
                op["height"] = height
            result = _resize_image(doc, op)
            return f"已调整图片 {image_index} 大小"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_crop_image(
        file_path: str,
        image_index: int,
        crop_left: float = 0,
        crop_right: float = 0,
        crop_top: float = 0,
        crop_bottom: float = 0,
    ) -> str:
        """裁剪图片 (仅支持浮动图片).

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引 (从1开始)
            crop_left: 左侧裁剪量 (磅)
            crop_right: 右侧裁剪量 (磅)
            crop_top: 顶部裁剪量 (磅)
            crop_bottom: 底部裁剪量 (磅)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _crop_image(doc, {
                "image_index": image_index,
                "crop_left": crop_left,
                "crop_right": crop_right,
                "crop_top": crop_top,
                "crop_bottom": crop_bottom,
            })
            return f"已裁剪图片: {image_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_image_position(
        file_path: str,
        image_index: int,
        left: float | None = None,
        top: float | None = None,
    ) -> str:
        """设置图片位置 (仅对浮动图片有效).

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引
            left: 左边距 (磅)
            top: 上边距 (磅)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op = {"image_index": image_index}
            if left is not None:
                op["left"] = left
            if top is not None:
                op["top"] = top
            result = _set_image_position(doc, op)
            return f"已设置图片 {image_index} 位置"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_image_wrap(file_path: str, image_index: int, wrap_type: str = "inline") -> str:
        """设置图片文字环绕类型.

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引
            wrap_type: 环绕类型 (none, square, tight, through, top_bottom, behind, in_front, inline)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_image_wrap(doc, {
                "image_index": image_index,
                "wrap_type": wrap_type,
            })
            return f"已设置图片 {image_index} 环绕: {wrap_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_replace_image(file_path: str, image_index: int, new_image_path: str) -> str:
        """替换图片.

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引
            new_image_path: 新图片绝对路径
        """
        try:
            path = validate_path(file_path)
            validate_path(new_image_path)
            doc = office_manager.get_document(path)
            result = _replace_image(doc, {
                "image_index": image_index,
                "new_image_path": new_image_path,
            })
            return f"已替换图片 {image_index} -> {new_image_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_delete_image(file_path: str, image_index: int) -> str:
        """删除图片.

        Args:
            file_path: Word 文档路径 (绝对路径)
            image_index: 图片索引
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_image(doc, {"image_index": image_index})
            return f"已删除图片: {image_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Hyperlinks 类工具 (3 个新, 已有 word_add_hyperlink)
    # ================================================================

    @mcp.tool()
    def word_list_hyperlinks(file_path: str) -> dict:
        """列出文档中的所有超链接.

        Args:
            file_path: Word 文档路径 (绝对路径)

        Returns:
            包含超链接列表的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            links = _list_hyperlinks(doc, {})
            return {"file_path": file_path, "hyperlinks": links, "count": len(links)}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "hyperlinks": [], "count": 0}

    @mcp.tool()
    def word_get_hyperlink(file_path: str, hyperlink_index: int) -> dict:
        """获取指定超链接的详细信息.

        Args:
            file_path: Word 文档路径 (绝对路径)
            hyperlink_index: 超链接索引 (从1开始)

        Returns:
            包含超链接详细信息的字典
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _get_hyperlink(doc, {"hyperlink_index": hyperlink_index})
            return {"file_path": file_path, **result}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_remove_hyperlink(file_path: str, hyperlink_index: int) -> str:
        """移除超链接 (保留文本).

        Args:
            file_path: Word 文档路径 (绝对路径)
            hyperlink_index: 超链接索引 (从1开始)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _remove_hyperlink(doc, {"hyperlink_index": hyperlink_index})
            return f"已移除超链接: {hyperlink_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Protection 类工具 (6 个)
    # ================================================================

    @mcp.tool()
    def word_set_document_protection(
        file_path: str,
        protection_type: str = "read_only",
        password: str = "",
    ) -> str:
        """设置文档保护.

        Args:
            file_path: Word 文档路径 (绝对路径)
            protection_type: 保护类型 (none, read_only, comments, tracked_changes, forms, all)
            password: 密码 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_document_protection(doc, {
                "protection_type": protection_type,
                "password": password,
            })
            return f"已设置文档保护: {protection_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_unprotect_document(file_path: str, password: str = "") -> str:
        """取消文档保护.

        Args:
            file_path: Word 文档路径 (绝对路径)
            password: 密码 (可选, 如果设置时提供了密码)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _unprotect_document(doc, {"password": password})
            return "已取消文档保护"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_read_only(file_path: str, read_only: bool = True) -> str:
        """设置文档为只读或取消只读.

        Args:
            file_path: Word 文档路径 (绝对路径)
            read_only: True 设置只读, False 取消只读 (默认 True)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_read_only(doc, {"read_only": read_only})
            return f"已设置只读: {read_only}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_password(file_path: str, password: str) -> str:
        """设置文档打开密码 (保存时生效).

        Args:
            file_path: Word 文档路径 (绝对路径)
            password: 打开密码
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_password(doc, {"password": password})
            return "已设置文档打开密码"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_accept_all_changes(file_path: str) -> str:
        """接受文档中的所有修订.

        Args:
            file_path: Word 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _accept_all_changes(doc, {})
            return f"已接受所有修订: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_reject_all_changes(file_path: str) -> str:
        """拒绝文档中的所有修订.

        Args:
            file_path: Word 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _reject_all_changes(doc, {})
            return f"已拒绝所有修订: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Document 类工具 (8 个新)
    # ================================================================

    @mcp.tool()
    def word_get_document_info(file_path: str) -> dict:
        """获取文档信息（页数、字数、作者等）.

        Args:
            file_path: Word 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            info = _get_document_info(doc, {})
            return {"file_path": file_path, **info}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_set_document_properties(file_path: str, properties: dict) -> str:
        """设置文档属性 (Title, Author, Subject, Keywords, Comments, Category, Manager, Company).

        Args:
            file_path: Word 文档路径 (绝对路径)
            properties: 字典，键值对形式. 例: {"title": "我的文档", "author": "张三"}
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_document_properties(doc, {"properties": properties})
            return f"已设置文档属性: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_get_document_properties(file_path: str) -> dict:
        """获取文档属性 (Title, Author, Subject, Keywords 等).

        Args:
            file_path: Word 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            properties = _get_document_properties(doc, {})
            return {"file_path": file_path, "properties": properties}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_add_comment(file_path: str, text: str, range_start: int = 0, range_end: int = 0) -> str:
        """添加批注.

        Args:
            file_path: Word 文档路径 (绝对路径)
            text: 批注内容
            range_start: 选区起始位置 (可选, 默认0表示在文档末尾添加)
            range_end: 选区结束位置 (可选, 默认0)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"text": text}
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _add_comment(doc, op)
            return f"已添加批注: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_list_comments(file_path: str) -> dict:
        """列出所有批注.

        Args:
            file_path: Word 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            comments = _list_comments(doc, {})
            return {"file_path": file_path, "count": len(comments), "comments": comments}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "count": 0, "comments": []}

    @mcp.tool()
    def word_delete_comment(file_path: str, index: int) -> str:
        """删除指定批注.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 批注索引 (1-based, 通过 word_list_comments 获取)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_comment(doc, {"index": index})
            return f"已删除批注: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_track_changes(file_path: str, enable: bool = True) -> str:
        """启用/禁用修订跟踪.

        Args:
            file_path: Word 文档路径 (绝对路径)
            enable: 是否启用修订跟踪 (默认 True)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _track_changes(doc, {"enable": enable})
            return f"已设置修订跟踪: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_compare_documents(file_path: str, compare_path: str, output_path: str = "") -> str:
        """比较两个文档的差异.

        Args:
            file_path: 当前 Word 文档路径 (绝对路径, 作为比较基准)
            compare_path: 要比较的文档路径 (绝对路径)
            output_path: 输出到新文档的路径 (可选, 不提供则结果写入当前文档)
        """
        try:
            path = validate_path(file_path)
            validate_path(compare_path)
            if output_path:
                validate_path(output_path)
            doc = office_manager.get_document(path)
            result = _compare_documents(doc, {
                "compare_path": compare_path,
                "output_path": output_path,
            })
            return f"已比较文档: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Paragraphs 类工具 (10 个新)
    # ================================================================

    @mcp.tool()
    def word_list_paragraphs(file_path: str, max_count: int = 0) -> dict:
        """列出文档中的所有段落.

        Args:
            file_path: Word 文档路径 (绝对路径)
            max_count: 最大返回数量 (可选, 0 表示返回全部)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {}
            if max_count > 0:
                op["max_count"] = max_count
            result = _list_paragraphs(doc, op)
            return {"file_path": file_path, **result}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "total": 0, "paragraphs": []}

    @mcp.tool()
    def word_get_paragraph(file_path: str, index: int) -> dict:
        """获取指定段落的内容、样式、对齐、缩进等信息.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            info = _get_paragraph(doc, {"index": index})
            return {"file_path": file_path, **info}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_set_paragraph_alignment(file_path: str, index: int, alignment: str) -> str:
        """设置段落对齐方式.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
            alignment: 对齐方式 (left, center, right, justify)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_paragraph_alignment(doc, {"index": index, "alignment": alignment})
            return f"已设置段落对齐: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_paragraph_spacing(
        file_path: str,
        index: int,
        space_before: float = -1,
        space_after: float = -1,
    ) -> str:
        """设置段落间距 (单位: 磅).

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
            space_before: 段前间距 (可选, 留空表示不修改)
            space_after: 段后间距 (可选, 留空表示不修改)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"index": index}
            if space_before >= 0:
                op["space_before"] = space_before
            if space_after >= 0:
                op["space_after"] = space_after
            result = _set_paragraph_spacing(doc, op)
            return f"已设置段落间距: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_line_spacing(file_path: str, index: int, rule: str, value: float = 0) -> str:
        """设置段落行距.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
            rule: 行距规则 (single, 1.5, double, multiple, at_least, exactly)
            value: 数值 (multiple 时为倍数, at_least/exactly 时为磅数)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_line_spacing(doc, {"index": index, "rule": rule, "value": value})
            return f"已设置行距: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_indent(
        file_path: str,
        index: int,
        left: float = -9999,
        right: float = -9999,
        first_line: float = -9999,
    ) -> str:
        """设置段落缩进 (单位: 磅).

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
            left: 左缩进 (可选)
            right: 右缩进 (可选)
            first_line: 首行缩进 (可选, 负数表示悬挂缩进)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"index": index}
            if left != -9999:
                op["left"] = left
            if right != -9999:
                op["right"] = right
            if first_line != -9999:
                op["first_line"] = first_line
            result = _set_indent(doc, op)
            return f"已设置缩进: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_bullet(file_path: str, index: int, character: str = "*", font_name: str = "Symbol") -> str:
        """为段落添加项目符号.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
            character: 项目符号字符 (可选, 默认 "*")
            font_name: 字体名称 (可选, 默认 "Symbol")
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_bullet(doc, {
                "index": index,
                "character": character,
                "font_name": font_name,
            })
            return f"已添加项目符号: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_remove_bullet(file_path: str, index: int) -> str:
        """移除段落的项目符号.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _remove_bullet(doc, {"index": index})
            return f"已移除项目符号: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_merge_paragraphs(file_path: str, start_index: int, end_index: int) -> str:
        """合并多个段落为单个段落.

        Args:
            file_path: Word 文档路径 (绝对路径)
            start_index: 起始段落索引 (1-based)
            end_index: 结束段落索引 (1-based, 包含)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _merge_paragraphs(doc, {
                "start_index": start_index,
                "end_index": end_index,
            })
            return f"已合并段落: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_split_paragraph(file_path: str, index: int, position: int = -1) -> str:
        """在指定位置拆分段落.

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 段落索引 (1-based)
            position: 段内字符位置 (可选, 留空表示在段尾拆分)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"index": index}
            if position >= 0:
                op["position"] = position
            result = _split_paragraph(doc, op)
            return f"已拆分段落: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Tables 类工具 (12 个新)
    # ================================================================

    @mcp.tool()
    def word_list_tables(file_path: str) -> dict:
        """列出文档中的所有表格.

        Args:
            file_path: Word 文档路径 (绝对路径)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.ensure_document(path, activate=False)
            tables = _list_tables(doc, {})
            return {"file_path": file_path, "count": len(tables), "tables": tables}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e), "count": 0, "tables": []}

    @mcp.tool()
    def word_get_table_info(file_path: str, index: int) -> dict:
        """获取表格信息 (行列数、样式).

        Args:
            file_path: Word 文档路径 (绝对路径)
            index: 表格索引 (1-based)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            info = _get_table_info(doc, {"index": index})
            return {"file_path": file_path, **info}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_set_cell(file_path: str, table_index: int, row: int, column: int, text: str) -> str:
        """设置单元格内容.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            row: 行号 (1-based)
            column: 列号 (1-based)
            text: 单元格文本
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_cell(doc, {
                "table_index": table_index,
                "row": row,
                "column": column,
                "text": text,
            })
            return f"已设置单元格: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_get_cell(file_path: str, table_index: int, row: int, column: int) -> dict:
        """获取单元格内容.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            row: 行号 (1-based)
            column: 列号 (1-based)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            text = _get_cell(doc, {
                "table_index": table_index,
                "row": row,
                "column": column,
            })
            return {"file_path": file_path, "table_index": table_index, "row": row, "column": column, "text": text}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def word_add_row(file_path: str, table_index: int, count: int = 1) -> str:
        """添加表格行.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            count: 添加行数 (可选, 默认 1)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_row(doc, {"table_index": table_index, "count": count})
            return f"已添加行: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_add_column(file_path: str, table_index: int, count: int = 1) -> str:
        """添加表格列.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            count: 添加列数 (可选, 默认 1)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _add_column(doc, {"table_index": table_index, "count": count})
            return f"已添加列: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_delete_row(file_path: str, table_index: int, row: int, count: int = 1) -> str:
        """删除表格行.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            row: 起始行号 (1-based)
            count: 删除行数 (可选, 默认 1)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_row(doc, {
                "table_index": table_index,
                "row": row,
                "count": count,
            })
            return f"已删除行: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_delete_column(file_path: str, table_index: int, column: int, count: int = 1) -> str:
        """删除表格列.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            column: 起始列号 (1-based)
            count: 删除列数 (可选, 默认 1)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _delete_column(doc, {
                "table_index": table_index,
                "column": column,
                "count": count,
            })
            return f"已删除列: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_merge_cells(
        file_path: str,
        table_index: int,
        start_row: int,
        start_column: int,
        end_row: int,
        end_column: int,
    ) -> str:
        """合并表格单元格 (矩形区域).

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            start_row: 起始行 (1-based)
            start_column: 起始列 (1-based)
            end_row: 结束行 (1-based, 包含)
            end_column: 结束列 (1-based, 包含)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _merge_cells(doc, {
                "table_index": table_index,
                "start_row": start_row,
                "start_column": start_column,
                "end_row": end_row,
                "end_column": end_column,
            })
            return f"已合并单元格: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_split_cell(file_path: str, table_index: int, row: int, column: int, rows: int = 1, columns: int = 1) -> str:
        """拆分表格单元格.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            row: 行号 (1-based)
            column: 列号 (1-based)
            rows: 拆分行数 (可选, 默认 1)
            columns: 拆分列数 (可选, 默认 1)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _split_cell(doc, {
                "table_index": table_index,
                "row": row,
                "column": column,
                "rows": rows,
                "columns": columns,
            })
            return f"已拆分单元格: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_table_borders(file_path: str, table_index: int, style: str = "single", line_width: float = 0.5) -> str:
        """设置表格边框.

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            style: 边框样式 (none, single, double, dotted, dashed, thick, 默认 single)
            line_width: 线条宽度磅数 (可选, 默认 0.5)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_table_borders(doc, {
                "table_index": table_index,
                "style": style,
                "line_width": line_width,
            })
            return f"已设置表格边框: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_table_style(file_path: str, table_index: int, style_name: str) -> str:
        """设置表格样式 (如 "Table Grid", "Light List").

        Args:
            file_path: Word 文档路径 (绝对路径)
            table_index: 表格索引 (1-based)
            style_name: 样式名称
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            result = _set_table_style(doc, {
                "table_index": table_index,
                "style_name": style_name,
            })
            return f"已设置表格样式: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ================================================================
    # Styles 类工具 (8 个新)
    # ================================================================

    @mcp.tool()
    def word_apply_bold(file_path: str, value: bool = True, index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置粗体.

        Args:
            file_path: Word 文档路径 (绝对路径)
            value: 是否加粗 (默认 True)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选, 与 range_end 同时设置)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"value": value}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _apply_bold(doc, op)
            return f"已设置粗体: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_apply_italic(file_path: str, value: bool = True, index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置斜体.

        Args:
            file_path: Word 文档路径 (绝对路径)
            value: 是否斜体 (默认 True)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"value": value}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _apply_italic(doc, op)
            return f"已设置斜体: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_apply_underline(file_path: str, value: bool = True, index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置下划线.

        Args:
            file_path: Word 文档路径 (绝对路径)
            value: 是否下划线 (默认 True)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"value": value}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _apply_underline(doc, op)
            return f"已设置下划线: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_font_color(file_path: str, color: str, index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置字体颜色.

        Args:
            file_path: Word 文档路径 (绝对路径)
            color: 颜色 (名称如 black/blue/red/green/yellow/white/gray/orange/purple, 或 #RRGGBB 十六进制)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"color": color}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _set_font_color(doc, op)
            return f"已设置字体颜色: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_font_size(file_path: str, size: float, index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置字体大小.

        Args:
            file_path: Word 文档路径 (绝对路径)
            size: 字号 (磅)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"size": size}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _set_font_size(doc, op)
            return f"已设置字体大小: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_highlight_color(file_path: str, color: str = "yellow", index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置高亮颜色.

        Args:
            file_path: Word 文档路径 (绝对路径)
            color: 颜色名称 (yellow, green, cyan, magenta, red, blue, dark_blue, teal, gray, black, white, none, 默认 yellow)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"color": color}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _set_highlight_color(doc, op)
            return f"已设置高亮颜色: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_strikethrough(file_path: str, value: bool = True, index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置删除线.

        Args:
            file_path: Word 文档路径 (绝对路径)
            value: 是否删除线 (默认 True)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"value": value}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _set_strikethrough(doc, op)
            return f"已设置删除线: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def word_set_subscript_superscript(file_path: str, mode: str = "normal", index: int = 0, range_start: int = 0, range_end: int = 0) -> str:
        """设置下标/上标.

        Args:
            file_path: Word 文档路径 (绝对路径)
            mode: 模式 (subscript=下标, superscript=上标, normal=正常, 默认 normal)
            index: 段落索引 (可选, 0 表示整个文档)
            range_start: 选区起始位置 (可选)
            range_end: 选区结束位置 (可选)
        """
        try:
            path = validate_path(file_path)
            doc = office_manager.get_document(path)
            op: dict = {"mode": mode}
            if index > 0:
                op["index"] = index
            if range_start and range_end:
                op["range_start"] = range_start
                op["range_end"] = range_end
            result = _set_subscript_superscript(doc, op)
            return f"已设置上下标: {result}"
        except OfficeMCPError as e:
            return f"错误: {e}"
