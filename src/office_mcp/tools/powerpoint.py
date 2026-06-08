"""PowerPoint MCP 工具."""

from office_mcp.compat import FastMCP
from office_mcp.config import settings
from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.operations.ppt_ops import (
    _add_animation,
    _add_audio,
    _add_chart,
    _add_comment,
    _add_connector,
    _add_freeform_shape,
    _add_line,
    _add_master_shape,
    _add_picture_extended,
    _add_picture_from_url,
    _add_ppt_hyperlink,
    _add_section,
    _add_smartart,
    _add_svg_icon,
    _add_textbox,
    _add_video,
    _align_shapes,
    _apply_layout,
    _batch_apply_format,
    _build_freeform_path,
    _clear_animations,
    _change_chart_type,
    _check_typography,
    _compare_presentations,
    _copy_animation,
    _copy_animation_from_shape,
    _copy_formatting,
    _copy_shape,
    _copy_to_clipboard,
    _create_from_template,
    _crop_picture,
    _delete_comment,
    _delete_node,
    _delete_shape,
    _distribute_shapes,
    _duplicate_shape,
    _duplicate_slide,
    _duplicate_slide_to_end,
    _export_html,
    _export_images,
    _export_shape,
    _flip_shape,
    _format_chart,
    _format_chart_axis,
    _format_connector,
    _format_shape,
    _get_active_window,
    _get_app_info,
    _get_chart_data,
    _get_node_positions,
    _get_ppt_group_items,
    _get_ppt_hyperlinks,
    _get_presentation_info,
    _get_properties,
    _get_screen_tip,
    _get_selection,
    _get_shape_count,
    _get_shape_info,
    _get_slide_info,
    _get_slide_layouts,
    _get_slide_notes,
    _get_slide_preview,
    _get_slide_size,
    _get_slideshow_status,
    _get_tags,
    _group_shapes,
    _insert_icon,
    _insert_node,
    _list_animations,
    _list_comments,
    _list_presentations,
    _list_shapes,
    _list_slides,
    _list_smartart_layouts,
    _list_templates,
    _lock_aspect_ratio,
    _merge_presentations,
    _merge_shapes,
    _modify_smartart,
    _move_slide,
    _navigate_to_slide,
    _paste_formatting,
    _ppt_add_table_column,
    _ppt_add_table_row,
    _ppt_batch_set_table_data,
    _ppt_clear_placeholder,
    _ppt_delete_table_column,
    _ppt_delete_table_row,
    _ppt_extract_text_as_markdown,
    _ppt_format_text_range,
    _ppt_find_replace_text,
    _ppt_get_text,
    _ppt_get_placeholder,
    _ppt_get_placeholder_type,
    _ppt_get_table_cells,
    _ppt_get_table_info,
    _ppt_get_textframe,
    _ppt_list_placeholders,
    _ppt_merge_table_cells,
    _ppt_resize_placeholder,
    _ppt_resize_table,
    _ppt_set_paragraph_format,
    _ppt_set_bullets,
    _ppt_set_fill,
    _ppt_set_font_color,
    _ppt_set_font_size,
    _ppt_set_line,
    _ppt_set_placeholder,
    _ppt_set_shadow,
    _ppt_set_text,
    _ppt_set_table_borders,
    _ppt_set_table_cells,
    _ppt_set_table_style,
    _ppt_split_table_cells,
    _redo,
    _remove_animation,
    _remove_ppt_hyperlink,
    _repair_presentation,
    _replace_font,
    _rotate_shape,
    _save_as,
    _search_icons,
    _select_shape,
    _set_chart_data,
    _set_chart_series,
    _set_animation_trigger,
    _set_default_font,
    _set_default_shape_style,
    _set_glow,
    _set_gradient_fill,
    _set_master_background,
    _set_media_settings,
    _set_node_editing_type,
    _set_node_positions,
    _set_picture_format,
    _set_properties,
    _set_reflection,
    _set_segment_type,
    _set_shape_visibility,
    _set_slide_notes_extended,
    _set_slide_number,
    _set_slide_size,
    _set_soft_edge,
    _set_tags,
    _set_theme_color,
    _set_theme_preset,
    _set_transition,
    _set_view,
    _set_window_state,
    _set_zorder,
    _slideshow_goto,
    _slideshow_next,
    _slideshow_previous,
    _start_slideshow,
    _stop_slideshow,
    _undo,
    _ungroup_ppt_shapes,
    _update_animation,
    _update_shape,
    apply_ppt_operations,
)


def register_ppt_tools(mcp: FastMCP) -> None:
    """注册 PowerPoint 相关工具."""

    @mcp.tool()
    def ppt_create_presentation(file_path: str, overwrite: bool = False) -> str:
        """创建新的 PowerPoint 演示文稿.

        Args:
            file_path: 演示文稿保存路径 (绝对路径)
            overwrite: 是否覆盖已存在的文件

        """
        try:
            path = validate_path(file_path)
            if path.exists() and not overwrite and not settings.default_overwrite:
                return f"错误: 文件已存在，请设置 overwrite=true 覆盖: {file_path}"
            office_manager.create_document(path, overwrite=overwrite)
            return f"已创建 PPT 演示文稿: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_open_presentation(file_path: str) -> str:
        """打开现有 PowerPoint 演示文稿.

        Args:
            file_path: 演示文稿路径 (绝对路径)

        """
        try:
            path = validate_path(file_path)
            office_manager.open_document(path)
            return f"已打开 PPT 演示文稿: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_apply_operations(file_path: str, operations: list[dict]) -> dict:
        """对 PowerPoint 演示文稿执行批量操作.

        Args:
            file_path: 演示文稿路径 (绝对路径)
            operations: 操作列表，每个操作是一个字典，包含 type 和其他参数

        支持的操作类型:
        - add_slide: 添加幻灯片 (layout: title/title_content/blank/section_header/two_content/comparison)
        - set_title: 设置标题 (slide_index, text)
        - add_text: 添加文本框 (slide_index, text, left, top, width, height)
        - insert_image: 插入图片 (slide_index, image_path, left, top, width, height)
        - insert_table: 插入表格 (slide_index, rows, columns, data, left, top, width, height)
        - set_slide_layout: 设置布局 (slide_index, layout)
        - delete_slide: 删除幻灯片 (slide_index)
        - set_background_color: 设置背景色 (slide_index, color)
        - add_shape: 添加形状 (slide_index, shape, left, top, width, height)
        - set_notes: 设置演讲者备注 (slide_index, text)
        - save: 保存演示文稿

        """
        try:
            path = validate_path(file_path)
            # 校验所有 image_path
            for op in operations:
                if op.get("type") == "insert_image":
                    img_path = op.get("image_path", "")
                    if img_path:
                        validate_path(img_path)

            presentation = office_manager.ensure_document(path, activate=True)
            results = apply_ppt_operations(presentation, operations)
            return {"file_path": file_path, "results": results}
        except OfficeMCPError as e:
            return {"file_path": file_path, "error": str(e)}

    @mcp.tool()
    def ppt_export_pdf(file_path: str, output_path: str | None = None) -> str:
        """将 PowerPoint 演示文稿导出为 PDF.

        Args:
            file_path: 演示文稿路径 (绝对路径)
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
    def ppt_close_presentation(file_path: str, save: bool = True) -> str:
        """关闭 PowerPoint 演示文稿.

        Args:
            file_path: 演示文稿路径 (绝对路径)
            save: 是否保存更改

        """
        try:
            path = validate_path(file_path)
            office_manager.close_document(path, save=save)
            action = "保存并关闭" if save else "关闭(未保存)"
            return f"{action} PPT 演示文稿: {file_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_animation(file_path: str, slide_index: int = 1, shape_index: int = 1, animation_type: str = "fade", trigger: str = "on_click", delay: float = 0, duration: float = None) -> str:
        """添加动画效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            animation_type: 动画类型 (appear/fade/wipe/push/dissolve/fly)
            trigger: 触发方式 (on_click/after_previous/with_previous)
            delay: 动画延迟时间 (秒)
            duration: 动画持续时间 (秒)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_animation(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "animation_type": animation_type,
                "trigger": trigger,
                "delay": delay,
                "duration": duration
            })
            return f"已添加动画效果: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_transition(file_path: str, slide_index: int = 1, transition_type: str = "fade", duration: float = 1.0) -> str:
        """设置幻灯片转换效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            transition_type: 转换类型 (none/fade/blind_down/push/wipe/split)
            duration: 持续时间 (秒)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_transition(ppt, {"slide_index": slide_index, "transition_type": transition_type, "duration": duration})
            return f"已设置转换效果: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_section(file_path: str, section_name: str, after_slide_index: int = 0) -> str:
        """添加分节.

        Args:
            file_path: 演示文稿路径
            section_name: 节名称
            after_slide_index: 在此幻灯片之后插入节

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_section(ppt, {"section_name": section_name, "after_slide_index": after_slide_index})
            return f"已添加分节: {section_name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_format_shape(file_path: str, slide_index: int = 1, shape_index: int = 1, fill_color: str = "", line_color: str = "", line_width: float = 1.0) -> str:
        """格式化形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            fill_color: 填充颜色 (#RRGGBB)
            line_color: 边框颜色 (#RRGGBB)
            line_width: 边框宽度 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _format_shape(ppt, {"slide_index": slide_index, "shape_index": shape_index, "fill_color": fill_color, "line_color": line_color, "line_width": line_width})
            return f"已格式化形状: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_slide_number(file_path: str, slide_index: int = 1, show: bool = True) -> str:
        """设置幻灯片编号显示.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            show: 是否显示

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_slide_number(ppt, {"slide_index": slide_index, "show": show})
            return f"已设置幻灯片编号: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_master_background(file_path: str, color: str = "#FFFFFF") -> str:
        """设置母版背景颜色.

        Args:
            file_path: 演示文稿路径
            color: 背景颜色 (#RRGGBB)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_master_background(ppt, {"color": color})
            return f"已设置母版背景颜色: {color}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_master_shape(
        file_path: str,
        shape: str = "rectangle",
        left: float = 0,
        top: float = 0,
        width: float = 100,
        height: float = 100,
        text: str = "",
        fill_color: str = "",
    ) -> str:
        """在母版上添加形状 (适用于所有基于此母版的幻灯片).

        Args:
            file_path: 演示文稿路径
            shape: 形状类型 (rectangle/oval/rounded_rectangle)
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt)
            height: 高度 (pt)
            text: 形状内文本
            fill_color: 填充颜色 (#RRGGBB)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_master_shape(ppt, {
                "shape": shape,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "text": text,
                "fill_color": fill_color,
            })
            return f"已在母版上添加形状: {shape}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_theme_color(
        file_path: str,
        color_name: str = "accent1",
        target: str = "fill",
        slide_index: int | None = None,
        shape_index: int = 1,
        tint: float = 0,
    ) -> str:
        """设置形状或文本的主题色而非硬编码 RGB.

        Args:
            file_path: 演示文稿路径
            color_name: 主题色名称 (dark1/light1/accent1-accent6/hyperlink)
            target: 目标 ("fill" / "font" / "line")
            slide_index: 幻灯片编号 (默认当前)
            shape_index: 形状编号
            tint: 色调偏移 (-1.0 到 1.0)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_theme_color(ppt, {
                "color_name": color_name,
                "target": target,
                "slide_index": slide_index,
                "shape_index": shape_index,
                "tint": tint,
            })
            return f"已设置主题色: {color_name} -> {target}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_theme_preset(
        file_path: str,
        preset: str = "",
        colors: dict | None = None,
    ) -> str:
        """应用预设主题色方案 (ocean/sunset/forest/corporate).

        Args:
            file_path: 演示文稿路径
            preset: 预设名称
            colors: 自定义颜色字典 (可选，覆盖 preset)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_theme_preset(ppt, {
                "preset": preset,
                "colors": colors or {},
            })
            return f"已应用主题预设: {preset or 'custom'}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_check_typography(file_path: str, slide_index: int | None = None) -> dict:
        """检查幻灯片排版质量问题 (短行、空文本框等).

        Args:
            file_path: 演示文稿路径
            slide_index: 指定幻灯片 (默认检查全部)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _check_typography(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_navigate_to_slide(file_path: str, slide_index: int) -> str:
        """实时导航到指定幻灯片，使修改立即可见.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片编号

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=True)
            result = _navigate_to_slide(ppt, {"slide_index": slide_index})
            return f"已导航到幻灯片: {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_create_from_template(
        template_name: str = "",
        template_path: str = "",
    ) -> str:
        """从用户模板文件夹创建演示文稿.

        Args:
            template_name: 模板文件名 (不含路径)
            template_path: 完整模板路径 (可选)

        """
        try:
            app = office_manager._get_app("ppt")
            result = _create_from_template(app, {
                "template_name": template_name,
                "template_path": template_path,
            })
            return f"已从模板创建: {template_name or template_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"
        except Exception as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_search_icons(
        file_path: str,
        query: str = "",
        limit: int = 20,
    ) -> dict:
        """搜索可用图标.

        Args:
            file_path: 演示文稿路径
            query: 搜索关键词
            limit: 返回数量上限

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _search_icons(ppt, {"query": query, "limit": limit})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_insert_icon(
        file_path: str,
        icon_name: str,
        slide_index: int = 1,
        left: float = 100,
        top: float = 100,
        width: float = 100,
        height: float = 100,
    ) -> str:
        """插入图标.

        Args:
            file_path: 演示文稿路径
            icon_name: 图标名称
            slide_index: 幻灯片编号
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt)
            height: 高度 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _insert_icon(ppt, {
                "slide_index": slide_index,
                "icon_name": icon_name,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            })
            return f"已插入图标: {icon_name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_smartart(
        file_path: str,
        smartart_type: str = "process",
        slide_index: int = 1,
        left: float = 100,
        top: float = 100,
        width: float = 400,
        height: float = 300,
        texts: list = None,
    ) -> str:
        """添加 SmartArt 图形.

        Args:
            file_path: 演示文稿路径
            smartart_type: SmartArt 类型 (org_chart/cycle/pyramid/process/list/hierarchy/relationship/matrix/picture)
            slide_index: 幻灯片索引
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt)
            height: 高度 (pt)
            texts: 文本列表，用于填充 SmartArt

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_smartart(ppt, {
                "slide_index": slide_index,
                "smartart_type": smartart_type,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "texts": texts or [],
            })
            return f"已添加 SmartArt 图形: {smartart_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_gradient_fill(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        gradient_type: str = "linear",
        colors: list = None,
        angle: float = 90,
    ) -> str:
        """设置形状的渐变填充.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            gradient_type: 渐变类型 (linear/radial/rectangular/path)
            colors: 颜色列表，每个元素为 {"color": "#RRGGBB", "position": 0-100}
            angle: 渐变角度 (仅 linear 类型)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_gradient_fill(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "gradient_type": gradient_type,
                "colors": colors or [{"color": "#000000", "position": 0}, {"color": "#FFFFFF", "position": 100}],
                "angle": angle,
            })
            return f"已设置渐变填充: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_freeform_shape(
        file_path: str,
        slide_index: int = 1,
        points: list = None,
        left: float = 0,
        top: float = 0,
        closed: bool = True,
    ) -> str:
        """添加自由路径形状 (预留接口，暂未完整实现).

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            points: 坐标点列表，每个元素为 [x, y]
            left: 左边距 (pt)
            top: 上边距 (pt)
            closed: 是否闭合路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_freeform_shape(ppt, {
                "slide_index": slide_index,
                "points": points or [],
                "left": left,
                "top": top,
                "closed": closed,
            })
            return f"已添加自由路径形状: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Export 类工具 ============

    @mcp.tool()
    def ppt_export_images(
        file_path: str,
        output_path: str,
        slide_indices: list | None = None,
        format: str = "png",
        width: int | None = None,
        height: int | None = None,
    ) -> dict:
        """导出幻灯片为图片.

        Args:
            file_path: 演示文稿路径
            output_path: 输出目录路径
            slide_indices: 幻灯片索引列表 (可选，默认全部)
            format: 图片格式 (png/jpg/gif/bmp)
            width: 图片宽度 (可选)
            height: 图片高度 (可选)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _export_images(ppt, {
                "output_path": output_path,
                "slide_indices": slide_indices or [],
                "format": format,
                "width": width,
                "height": height,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_slide_preview(
        file_path: str,
        slide_index: int = 1,
        width: int = 800,
        height: int = 600,
    ) -> dict:
        """获取幻灯片预览.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            width: 预览宽度
            height: 预览高度

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_slide_preview(ppt, {
                "slide_index": slide_index,
                "width": width,
                "height": height,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_copy_to_clipboard(
        file_path: str,
        slide_index: int = 1,
    ) -> str:
        """复制幻灯片到剪贴板.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _copy_to_clipboard(ppt, {"slide_index": slide_index})
            return f"已复制幻灯片到剪贴板: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_export_html(
        file_path: str,
        output_path: str,
        include_navigation: bool = True,
    ) -> str:
        """导出为 HTML.

        Args:
            file_path: 演示文稿路径
            output_path: HTML 输出路径
            include_navigation: 是否包含导航栏

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _export_html(ppt, {
                "output_path": output_path,
                "include_navigation": include_navigation,
            })
            return f"已导出 HTML: {output_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Slideshow 类工具 ============

    @mcp.tool()
    def ppt_start_slideshow(
        file_path: str,
        from_slide: int = 1,
        loop: bool = False,
    ) -> str:
        """开始放映.

        Args:
            file_path: 演示文稿路径
            from_slide: 开始幻灯片索引
            loop: 是否循环放映

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _start_slideshow(ppt, {
                "from_slide": from_slide,
                "loop": loop,
            })
            return f"已开始放映: from_slide={from_slide}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_stop_slideshow(file_path: str) -> str:
        """停止放映.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _stop_slideshow(ppt, {})
            return "已停止放映"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_slideshow_next(file_path: str) -> str:
        """放映下一页.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _slideshow_next(ppt, {})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_slideshow_previous(file_path: str) -> str:
        """放映上一页.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _slideshow_previous(ppt, {})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_slideshow_goto(
        file_path: str,
        slide_index: int = 1,
    ) -> str:
        """放映跳转到指定页.

        Args:
            file_path: 演示文稿路径
            slide_index: 目标幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _slideshow_goto(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_slideshow_status(file_path: str) -> dict:
        """获取放映状态.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_slideshow_status(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    # ============ Charts 类工具 ============

    @mcp.tool()
    def ppt_add_chart(
        file_path: str,
        slide_index: int = 1,
        chart_type: str = "column_clustered",
        left: float = 100,
        top: float = 100,
        width: float = 400,
        height: float = 300,
        data: list | None = None,
    ) -> str:
        """添加图表.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            chart_type: 图表类型 (column_clustered/column_stacked/bar_clustered/bar_stacked/line/line_markers/pie/pie_exploded/scatter/area/area_stacked/doughnut/radar/radar_filled/combo)
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt)
            height: 高度 (pt)
            data: 图表数据 (二维数组)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_chart(ppt, {
                "slide_index": slide_index,
                "chart_type": chart_type,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "data": data or [],
            })
            return f"已添加图表: slide {slide_index}, type={chart_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_chart_data(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        data: list | None = None,
    ) -> str:
        """设置图表数据.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (图表)
            data: 图表数据 (二维数组)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_chart_data(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "data": data or [],
            })
            return f"已设置图表数据: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_chart_data(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取图表数据.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (图表)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_chart_data(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_format_chart(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        title: str = "",
        show_legend: bool = True,
        legend_position: str = "bottom",
    ) -> str:
        """格式化图表.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (图表)
            title: 图表标题
            show_legend: 是否显示图例
            legend_position: 图例位置 (bottom/top/left/right)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _format_chart(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "title": title,
                "show_legend": show_legend,
                "legend_position": legend_position,
            })
            return f"已格式化图表: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_format_chart_axis(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        axis_type: str = "category",
        title: str = "",
        has_title: bool = True,
        has_gridlines: bool = True,
    ) -> str:
        """格式化图表轴.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (图表)
            axis_type: 轴类型 (category/value)
            title: 轴标题
            has_title: 是否显示轴标题
            has_gridlines: 是否显示网格线

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _format_chart_axis(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "axis_type": axis_type,
                "title": title,
                "has_title": has_title,
                "has_gridlines": has_gridlines,
            })
            return f"已格式化图表轴: slide {slide_index}, axis={axis_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_chart_series(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        series_index: int = 1,
        name: str = "",
        values: list | None = None,
        chart_type: str = "",
    ) -> str:
        """设置图表系列.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (图表)
            series_index: 系列索引
            name: 系列名称
            values: 系列值列表
            chart_type: 系列图表类型 (可选，用于组合图)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_chart_series(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "series_index": series_index,
                "name": name,
                "values": values or [],
                "chart_type": chart_type,
            })
            return f"已设置图表系列: slide {slide_index}, series {series_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_change_chart_type(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        chart_type: str = "column_clustered",
    ) -> str:
        """更改图表类型.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (图表)
            chart_type: 新图表类型

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _change_chart_type(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "chart_type": chart_type,
            })
            return f"已更改图表类型: slide {slide_index}, type={chart_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Animation 类工具 ============

    @mcp.tool()
    def ppt_list_animations(
        file_path: str,
        slide_index: int = 1,
    ) -> dict:
        """列出幻灯片动画.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_animations(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_update_animation(
        file_path: str,
        slide_index: int = 1,
        animation_index: int = 1,
        duration: float | None = None,
        delay: float | None = None,
        trigger: str | None = None,
    ) -> str:
        """更新动画属性.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            animation_index: 动画索引
            duration: 持续时间 (可选)
            delay: 延迟时间 (可选)
            trigger: 触发方式 (可选: on_click/after_previous/with_previous)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _update_animation(ppt, {
                "slide_index": slide_index,
                "animation_index": animation_index,
                "duration": duration,
                "delay": delay,
                "trigger": trigger,
            })
            return f"已更新动画: slide {slide_index}, animation {animation_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_remove_animation(
        file_path: str,
        slide_index: int = 1,
        animation_index: int = 1,
    ) -> str:
        """移除动画.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            animation_index: 动画索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _remove_animation(ppt, {
                "slide_index": slide_index,
                "animation_index": animation_index,
            })
            return f"已移除动画: slide {slide_index}, animation {animation_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_clear_animations(
        file_path: str,
        slide_index: int = 1,
    ) -> str:
        """清除所有动画.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _clear_animations(ppt, {"slide_index": slide_index})
            return f"已清除所有动画: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_animation_trigger(
        file_path: str,
        slide_index: int = 1,
        animation_index: int = 1,
        trigger: str = "on_click",
    ) -> str:
        """设置动画触发方式.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            animation_index: 动画索引
            trigger: 触发方式 (on_click/after_previous/with_previous)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_animation_trigger(ppt, {
                "slide_index": slide_index,
                "animation_index": animation_index,
                "trigger": trigger,
            })
            return f"已设置动画触发方式: slide {slide_index}, trigger={trigger}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_copy_animation(
        file_path: str,
        slide_index: int = 1,
        source_shape_index: int = 1,
        target_shape_index: int = 2,
    ) -> str:
        """复制动画到其他形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            source_shape_index: 源形状索引
            target_shape_index: 目标形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _copy_animation(ppt, {
                "slide_index": slide_index,
                "source_shape_index": source_shape_index,
                "target_shape_index": target_shape_index,
            })
            return f"已复制动画: from shape {source_shape_index} to shape {target_shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Comments 批注工具 ============

    @mcp.tool()
    def ppt_add_comment(
        file_path: str,
        slide_index: int = 1,
        text: str = "",
        author: str = "",
        left: float = 100,
        top: float = 100,
    ) -> str:
        """添加批注.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            text: 批注内容
            author: 作者名称
            left: 批注位置左边距 (pt)
            top: 批注位置上边距 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_comment(ppt, {
                "slide_index": slide_index,
                "text": text,
                "author": author,
                "left": left,
                "top": top,
            })
            return f"已添加批注: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_list_comments(
        file_path: str,
        slide_index: int | None = None,
    ) -> dict:
        """列出幻灯片上的批注.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引 (可选，默认列出所有)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_comments(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_delete_comment(
        file_path: str,
        slide_index: int = 1,
        comment_index: int = 1,
    ) -> str:
        """删除批注.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            comment_index: 批注索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _delete_comment(ppt, {
                "slide_index": slide_index,
                "comment_index": comment_index,
            })
            return f"已删除批注: slide {slide_index}, comment {comment_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Advanced 高级工具 ============

    @mcp.tool()
    def ppt_set_tags(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        tag_name: str = "",
        tag_value: str = "",
    ) -> str:
        """设置形状标签.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            tag_name: 标签名称
            tag_value: 标签值

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_tags(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "tag_name": tag_name,
                "tag_value": tag_value,
            })
            return f"已设置标签: {tag_name}={tag_value}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_tags(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取形状标签.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_tags(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_default_font(
        file_path: str,
        font_name: str = "",
        font_size: float | None = None,
        font_color: str = "",
    ) -> str:
        """设置演示文稿默认字体.

        Args:
            file_path: 演示文稿路径
            font_name: 字体名称
            font_size: 字体大小 (pt)
            font_color: 字体颜色 (#RRGGBB)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_default_font(ppt, {
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color,
            })
            return f"已设置默认字体: {font_name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_replace_font(
        file_path: str,
        old_font: str = "",
        new_font: str = "",
    ) -> str:
        """批量替换演示文稿中的字体.

        Args:
            file_path: 演示文稿路径
            old_font: 原字体名称
            new_font: 新字体名称

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _replace_font(ppt, {
                "old_font": old_font,
                "new_font": new_font,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_crop_picture(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        left: float = 0,
        top: float = 0,
        right: float = 0,
        bottom: float = 0,
    ) -> str:
        """裁剪图片.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            left: 左边裁剪比例 (0-1)
            top: 上边裁剪比例 (0-1)
            right: 右边裁剪比例 (0-1)
            bottom: 下边裁剪比例 (0-1)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _crop_picture(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "left": left,
                "top": top,
                "right": right,
                "bottom": bottom,
            })
            return f"已裁剪图片: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_picture_format(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        brightness: float | None = None,
        contrast: float | None = None,
        transparency: float | None = None,
    ) -> str:
        """设置图片格式 (亮度、对比度、透明度).

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            brightness: 亮度 (-1 到 1)
            contrast: 对比度 (-1 到 1)
            transparency: 透明度 (0 到 1)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_picture_format(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "brightness": brightness,
                "contrast": contrast,
                "transparency": transparency,
            })
            return f"已设置图片格式: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_export_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        output_path: str = "",
        format: str = "png",
    ) -> str:
        """导出形状为图片.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            output_path: 输出路径
            format: 图片格式 (png/jpg/emf/wmf)

        """
        try:
            path = validate_path(file_path)
            out_path = validate_path(output_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _export_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "output_path": str(out_path),
                "format": format,
            })
            return f"已导出形状: {output_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_shape_visibility(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        visible: bool = True,
    ) -> str:
        """设置形状可见性.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            visible: 是否可见

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_shape_visibility(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "visible": visible,
            })
            return f"已设置形状可见性: {visible}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_select_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> str:
        """选择形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _select_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return f"已选择形状: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_selection(file_path: str) -> dict:
        """获取当前选择.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_selection(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_view(
        file_path: str,
        view_type: str = "normal",
    ) -> str:
        """设置视图模式.

        Args:
            file_path: 演示文稿路径
            view_type: 视图类型 (normal/slide_sorter/notes_page/slide_master/handout_master)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_view(ppt, {"view_type": view_type})
            return f"已设置视图模式: {view_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_copy_animation_from_shape(
        file_path: str,
        slide_index: int = 1,
        source_shape_index: int = 1,
        target_shape_index: int = 2,
    ) -> str:
        """从形状复制动画到另一个形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            source_shape_index: 源形状索引
            target_shape_index: 目标形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _copy_animation_from_shape(ppt, {
                "slide_index": slide_index,
                "source_shape_index": source_shape_index,
                "target_shape_index": target_shape_index,
            })
            return f"已复制动画: {source_shape_index} -> {target_shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_picture_from_url(
        file_path: str,
        url: str = "",
        slide_index: int = 1,
        left: float = 100,
        top: float = 100,
        width: float | None = None,
        height: float | None = None,
    ) -> str:
        """从 URL 添加图片.

        Args:
            file_path: 演示文稿路径
            url: 图片 URL
            slide_index: 幻灯片索引
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt)
            height: 高度 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_picture_from_url(ppt, {
                "slide_index": slide_index,
                "url": url,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            })
            return f"已从 URL 添加图片: {url[:50]}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_svg_icon(
        file_path: str,
        svg_path: str = "",
        slide_index: int = 1,
        left: float = 100,
        top: float = 100,
        width: float = 72,
        height: float = 72,
    ) -> str:
        """添加 SVG 图标.

        Args:
            file_path: 演示文稿路径
            svg_path: SVG 文件路径
            slide_index: 幻灯片索引
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt)
            height: 高度 (pt)

        """
        try:
            path = validate_path(file_path)
            svg_validated = validate_path(svg_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_svg_icon(ppt, {
                "slide_index": slide_index,
                "svg_path": str(svg_validated),
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            })
            return f"已添加 SVG 图标: {svg_path}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_batch_apply_format(
        file_path: str,
        slide_index: int = 1,
        shape_indices: list = None,
        fill_color: str = "",
        line_color: str = "",
        line_width: float | None = None,
        font_name: str = "",
        font_size: float | None = None,
        font_color: str = "",
        bold: bool | None = None,
        italic: bool | None = None,
    ) -> str:
        """批量应用格式到多个形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_indices: 形状索引列表
            fill_color: 填充颜色 (#RRGGBB)
            line_color: 边框颜色 (#RRGGBB)
            line_width: 边框宽度 (pt)
            font_name: 字体名称
            font_size: 字体大小 (pt)
            font_color: 字体颜色 (#RRGGBB)
            bold: 是否加粗
            italic: 是否斜体

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _batch_apply_format(ppt, {
                "slide_index": slide_index,
                "shape_indices": shape_indices or [],
                "fill_color": fill_color,
                "line_color": line_color,
                "line_width": line_width,
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color,
                "bold": bold,
                "italic": italic,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_default_shape_style(
        file_path: str,
        fill_color: str = "",
        line_color: str = "",
        line_width: float = 1,
        line_visible: bool = True,
    ) -> str:
        """设置默认形状样式.

        Args:
            file_path: 演示文稿路径
            fill_color: 填充颜色 (#RRGGBB)
            line_color: 边框颜色 (#RRGGBB)
            line_width: 边框宽度 (pt)
            line_visible: 边框是否可见

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_default_shape_style(ppt, {
                "fill_color": fill_color,
                "line_color": line_color,
                "line_width": line_width,
                "line_visible": line_visible,
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_shape_count(
        file_path: str,
        slide_index: int | None = None,
    ) -> dict:
        """获取幻灯片上的形状数量.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引 (可选，默认统计所有)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_shape_count(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    # ============ Freeform 自由路径工具 ============

    @mcp.tool()
    def ppt_build_freeform_path(
        file_path: str,
        slide_index: int = 1,
        points: list = None,
        closed: bool = True,
        left: float = 100,
        top: float = 100,
        fill_color: str = "",
        line_color: str = "#000000",
        line_width: float = 1,
    ) -> str:
        """构建自由路径形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            points: 坐标点列表，每个元素为 [x, y]
            closed: 是否闭合路径
            left: 左边距 (pt)
            top: 上边距 (pt)
            fill_color: 填充颜色 (#RRGGBB)
            line_color: 边框颜色 (#RRGGBB)
            line_width: 边框宽度 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _build_freeform_path(ppt, {
                "slide_index": slide_index,
                "points": points or [],
                "closed": closed,
                "left": left,
                "top": top,
                "fill_color": fill_color,
                "line_color": line_color,
                "line_width": line_width,
            })
            return f"已构建自由路径: {len(points or [])} 个点"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_node_positions(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取自由形状的节点位置.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_node_positions(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_node_positions(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        node_positions: list = None,
    ) -> str:
        """设置自由形状的节点位置.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            node_positions: 节点位置列表，每个元素为 {"index": int, "x": float, "y": float}

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_node_positions(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "node_positions": node_positions or [],
            })
            return result
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_insert_node(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        after_index: int = 1,
        x: float = 0,
        y: float = 0,
        segment_type: str = "line",
        editing_type: str = "auto",
    ) -> str:
        """在自由形状中插入节点.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            after_index: 在此节点之后插入
            x: X 坐标
            y: Y 坐标
            segment_type: 线段类型 (line/curve)
            editing_type: 编辑类型 (auto/corner/smooth/symmetric)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _insert_node(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "after_index": after_index,
                "x": x,
                "y": y,
                "segment_type": segment_type,
                "editing_type": editing_type,
            })
            return f"已插入节点: ({x}, {y})"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_delete_node(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        node_index: int = 1,
    ) -> str:
        """删除自由形状中的节点.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            node_index: 节点索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _delete_node(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "node_index": node_index,
            })
            return f"已删除节点: {node_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_node_editing_type(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        node_index: int = 1,
        editing_type: str = "auto",
    ) -> str:
        """设置节点编辑类型.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            node_index: 节点索引
            editing_type: 编辑类型 (auto/corner/smooth/symmetric)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_node_editing_type(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "node_index": node_index,
                "editing_type": editing_type,
            })
            return f"已设置节点编辑类型: {editing_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_segment_type(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        node_index: int = 1,
        segment_type: str = "line",
    ) -> str:
        """设置线段类型.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            node_index: 节点索引
            segment_type: 线段类型 (line/curve)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_segment_type(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "node_index": node_index,
                "segment_type": segment_type,
            })
            return f"已设置线段类型: {segment_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Media 类工具 ============

    @mcp.tool()
    def ppt_add_video(
        file_path: str,
        video_path: str,
        slide_index: int = 1,
        left: float = 100,
        top: float = 100,
        width: float = None,
        height: float = None,
    ) -> str:
        """添加视频到幻灯片.

        Args:
            file_path: 演示文稿路径
            video_path: 视频文件路径
            slide_index: 幻灯片索引
            left: 左边距 (pt)
            top: 上边距 (pt)
            width: 宽度 (pt, 可选)
            height: 高度 (pt, 可选)

        """
        try:
            path = validate_path(file_path)
            video_p = validate_path(video_path)
            ppt = office_manager.ensure_document(path, activate=True)
            result = _add_video(ppt, {
                "slide_index": slide_index,
                "video_path": str(video_p),
                "left": left,
                "top": top,
                "width": width,
                "height": height,
            })
            return f"已添加视频: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_audio(
        file_path: str,
        audio_path: str,
        slide_index: int = 1,
        left: float = 100,
        top: float = 100,
        embed: bool = True,
    ) -> str:
        """添加音频到幻灯片.

        Args:
            file_path: 演示文稿路径
            audio_path: 音频文件路径
            slide_index: 幻灯片索引
            left: 左边距 (pt)
            top: 上边距 (pt)
            embed: 是否嵌入文档 (默认 True)

        """
        try:
            path = validate_path(file_path)
            audio_p = validate_path(audio_path)
            ppt = office_manager.ensure_document(path, activate=True)
            result = _add_audio(ppt, {
                "slide_index": slide_index,
                "audio_path": str(audio_p),
                "left": left,
                "top": top,
                "embed": embed,
            })
            return f"已添加音频: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_media_settings(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        play_on_click: bool = True,
        play_fullscreen: bool = False,
        play_loop: bool = False,
        play_automatically: bool = False,
        hide_when_not_playing: bool = False,
        volume: float = None,
    ) -> str:
        """设置媒体播放设置.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            play_on_click: 单击时播放
            play_fullscreen: 全屏播放
            play_loop: 循环播放
            play_automatically: 自动播放
            hide_when_not_playing: 不播放时隐藏
            volume: 音量 (0-1)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=True)
            result = _set_media_settings(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "play_on_click": play_on_click,
                "play_fullscreen": play_fullscreen,
                "play_loop": play_loop,
                "play_automatically": play_automatically,
                "hide_when_not_playing": hide_when_not_playing,
                "volume": volume,
            })
            return f"已设置媒体播放设置: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ SmartArt 类工具 ============

    @mcp.tool()
    def ppt_modify_smartart(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        node_index: int = None,
        text: str = None,
        add_node: bool = False,
        remove_node: int = None,
    ) -> str:
        """修改 SmartArt 图形.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            node_index: 节点索引 (1-based)
            text: 节点文本
            add_node: 添加节点
            remove_node: 删除节点索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _modify_smartart(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "node_index": node_index,
                "text": text,
                "add_node": add_node,
                "remove_node": remove_node,
            })
            return f"已修改 SmartArt: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_list_smartart_layouts(file_path: str) -> dict:
        """列出可用的 SmartArt 布局.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_smartart_layouts(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    # ============ Edit Operations 类工具 ============

    @mcp.tool()
    def ppt_undo(file_path: str, steps: int = 1) -> str:
        """撤销操作.

        Args:
            file_path: 演示文稿路径
            steps: 撤销步数 (默认 1)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _undo(ppt, {"steps": steps})
            return f"已撤销 {steps} 步操作"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_redo(file_path: str, steps: int = 1) -> str:
        """重做操作.

        Args:
            file_path: 演示文稿路径
            steps: 重做步数 (默认 1)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _redo(ppt, {"steps": steps})
            return f"已重做 {steps} 步操作"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_copy_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        target_slide_index: int = None,
        offset_left: float = 20,
        offset_top: float = 20,
    ) -> str:
        """复制形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 源幻灯片索引
            shape_index: 源形状索引
            target_slide_index: 目标幻灯片索引 (默认同一幻灯片)
            offset_left: 水平偏移 (pt)
            offset_top: 垂直偏移 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _copy_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "target_slide_index": target_slide_index or slide_index,
                "offset_left": offset_left,
                "offset_top": offset_top,
            })
            return f"已复制形状: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_copy_formatting(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> str:
        """复制形状格式.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 源形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _copy_formatting(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return f"已复制格式: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_paste_formatting(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> str:
        """粘贴格式到形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 目标形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _paste_formatting(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return f"已粘贴格式: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_duplicate_slide_to_end(
        file_path: str,
        slide_index: int = 1,
    ) -> str:
        """复制幻灯片到末尾.

        Args:
            file_path: 演示文稿路径
            slide_index: 要复制的幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _duplicate_slide_to_end(ppt, {
                "slide_index": slide_index,
            })
            return f"已复制幻灯片到末尾: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Layout 类工具 ============

    @mcp.tool()
    def ppt_align_shapes(
        file_path: str,
        slide_index: int = 1,
        shape_indices: list = None,
        align_type: str = "left",
        relative_to_slide: bool = False,
    ) -> str:
        """对齐多个形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_indices: 形状索引列表
            align_type: 对齐类型 (left/center/right/top/middle/bottom)
            relative_to_slide: 相对于幻灯片对齐 (默认 False，相对于形状组)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _align_shapes(ppt, {
                "slide_index": slide_index,
                "shape_indices": shape_indices or [],
                "align_type": align_type,
                "relative_to_slide": relative_to_slide,
            })
            return f"已对齐形状: slide {slide_index}, type={align_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_distribute_shapes(
        file_path: str,
        slide_index: int = 1,
        shape_indices: list = None,
        distribute_type: str = "horizontal",
    ) -> str:
        """分布多个形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_indices: 形状索引列表
            distribute_type: 分布类型 (horizontal/vertical)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _distribute_shapes(ppt, {
                "slide_index": slide_index,
                "shape_indices": shape_indices or [],
                "distribute_type": distribute_type,
            })
            return f"已分布形状: slide {slide_index}, type={distribute_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_slide_size(
        file_path: str,
        width: float = None,
        height: float = None,
        size_preset: str = "standard",
    ) -> str:
        """设置幻灯片尺寸.

        Args:
            file_path: 演示文稿路径
            width: 宽度 (pt)
            height: 高度 (pt)
            size_preset: 预设尺寸 (standard/widescreen/a4/custom)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_slide_size(ppt, {
                "width": width,
                "height": height,
                "size_preset": size_preset,
            })
            return f"已设置幻灯片尺寸: {size_preset}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_flip_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        flip_type: str = "horizontal",
    ) -> str:
        """翻转形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            flip_type: 翻转类型 (horizontal/vertical)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _flip_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "flip_type": flip_type,
            })
            return f"已翻转形状: slide {slide_index}, shape {shape_index}, type={flip_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_merge_shapes(
        file_path: str,
        slide_index: int = 1,
        shape_indices: list = None,
        merge_type: str = "union",
    ) -> str:
        """合并形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_indices: 形状索引列表
            merge_type: 合并类型 (union/combine/intersect/subtract)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _merge_shapes(ppt, {
                "slide_index": slide_index,
                "shape_indices": shape_indices or [],
                "merge_type": merge_type,
            })
            return f"已合并形状: slide {slide_index}, type={merge_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_rotate_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        angle: float = 0,
        relative: bool = True,
    ) -> str:
        """旋转形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            angle: 旋转角度 (度)
            relative: 是否相对旋转 (默认 True)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _rotate_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "angle": angle,
                "relative": relative,
            })
            return f"已旋转形状: slide {slide_index}, shape {shape_index}, angle={angle}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_lock_aspect_ratio(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        lock: bool = True,
    ) -> str:
        """锁定/解锁形状宽高比.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            lock: 是否锁定 (True 锁定, False 解锁)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _lock_aspect_ratio(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "lock": lock,
            })
            return f"已锁定宽高比: slide {slide_index}, shape {shape_index}, lock={lock}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Effects 类工具 ============

    @mcp.tool()
    def ppt_set_glow(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        color: str = "#FFFFFF",
        size: float = 10,
        transparency: float = 0.5,
    ) -> str:
        """设置形状发光效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            color: 发光颜色 (#RRGGBB)
            size: 发光大小 (pt)
            transparency: 透明度 (0-1)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_glow(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "color": color,
                "size": size,
                "transparency": transparency,
            })
            return f"已设置发光效果: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_reflection(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        transparency: float = 0.5,
        size: float = 0.5,
        blur: float = 5,
        distance: float = 10,
    ) -> str:
        """设置形状反射效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            transparency: 透明度 (0-1)
            size: 反射大小 (0-1)
            blur: 模糊程度 (pt)
            distance: 距离 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_reflection(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "transparency": transparency,
                "size": size,
                "blur": blur,
                "distance": distance,
            })
            return f"已设置反射效果: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_soft_edge(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        size: float = 10,
    ) -> str:
        """设置形状柔化边缘效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            size: 柔化大小 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_soft_edge(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "size": size,
            })
            return f"已设置柔化边缘效果: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ App 类工具 (5 个新工具) ============

    @mcp.tool()
    def ppt_get_app_info(file_path: str) -> dict:
        """获取 PowerPoint 应用信息（版本、路径）.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_app_info(ppt.Application, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_active_window(file_path: str) -> dict:
        """获取当前活动窗口信息.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_active_window(ppt.Application, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_window_state(file_path: str, state: str = "normal") -> str:
        """设置窗口状态（最大化/最小化/正常）.

        Args:
            file_path: 演示文稿路径
            state: 窗口状态 (maximized/minimized/normal)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_window_state(ppt.Application, {"state": state})
            return f"已设置窗口状态: {state}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_list_presentations(file_path: str) -> dict:
        """列出所有打开的演示文稿.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_presentations(ppt.Application, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_screen_tip(file_path: str) -> dict:
        """获取屏幕提示功能状态.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_screen_tip(ppt.Application, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    # ============ Presentation 类工具 (8 个新工具) ============

    @mcp.tool()
    def ppt_get_presentation_info(file_path: str) -> dict:
        """获取演示文稿信息（幻灯片数、大小）.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_presentation_info(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_list_templates(file_path: str) -> dict:
        """列出可用模板.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_templates(ppt.Application, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_properties(
        file_path: str,
        title: str = "",
        author: str = "",
        subject: str = "",
        keywords: str = "",
        comments: str = "",
    ) -> str:
        """设置演示文稿属性（标题、作者）.

        Args:
            file_path: 演示文稿路径
            title: 标题
            author: 作者
            subject: 主题
            keywords: 关键词
            comments: 备注

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_properties(ppt, {
                "title": title,
                "author": author,
                "subject": subject,
                "keywords": keywords,
                "comments": comments,
            })
            return "已设置演示文稿属性"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_properties(file_path: str) -> dict:
        """获取演示文稿属性.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_properties(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_save_as(file_path: str, output_path: str, format: str = "pptx") -> str:
        """另存为指定格式.

        Args:
            file_path: 演示文稿路径
            output_path: 输出路径
            format: 文件格式 (pptx/pdf/ppt/potx/ppsx/xps/gif/jpg/png/bmp/tif/wmv/mp4)

        """
        try:
            path = validate_path(file_path)
            out_path = validate_path(output_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _save_as(ppt, {"file_path": str(out_path), "format": format})
            return f"已另存为: {output_path}, 格式: {format}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_repair_presentation(file_path: str) -> dict:
        """修复损坏的演示文稿（检查并尝试修复常见问题）.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _repair_presentation(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_compare_presentations(
        file_path: str,
        presentation1_path: str,
        presentation2_path: str,
    ) -> dict:
        """比较两个演示文稿.

        Args:
            file_path: 当前演示文稿路径 (用于获取 Application)
            presentation1_path: 第一个演示文稿路径
            presentation2_path: 第二个演示文稿路径

        """
        try:
            path = validate_path(file_path)
            pres1_path = validate_path(presentation1_path)
            pres2_path = validate_path(presentation2_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _compare_presentations(ppt.Application, {
                "presentation1_path": str(pres1_path),
                "presentation2_path": str(pres2_path),
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_merge_presentations(
        file_path: str,
        target_path: str,
        source_paths: list[str],
        insert_position: int = -1,
    ) -> str:
        """合并多个演示文稿.

        Args:
            file_path: 当前演示文稿路径 (用于获取 Application)
            target_path: 目标演示文稿路径
            source_paths: 源演示文稿路径列表
            insert_position: 插入位置 (默认末尾)

        """
        try:
            path = validate_path(file_path)
            target = validate_path(target_path)
            sources = [validate_path(s) for s in source_paths]
            ppt = office_manager.ensure_document(path, activate=False)
            result = _merge_presentations(ppt.Application, {
                "target_path": str(target),
                "source_paths": [str(s) for s in sources],
                "insert_position": insert_position,
            })
            return f"已合并 {len(source_paths)} 个演示文稿"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Slides 类工具 (9 个新工具) ============

    @mcp.tool()
    def ppt_duplicate_slide(file_path: str, slide_index: int = 1, insert_after: int = None) -> str:
        """复制幻灯片.

        Args:
            file_path: 演示文稿路径
            slide_index: 要复制的幻灯片索引
            insert_after: 复制后插入位置 (默认原幻灯片之后)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _duplicate_slide(ppt, {
                "slide_index": slide_index,
                "insert_after": insert_after or slide_index,
            })
            return f"已复制幻灯片: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_move_slide(file_path: str, slide_index: int = 1, new_position: int = 1) -> str:
        """移动幻灯片位置.

        Args:
            file_path: 演示文稿路径
            slide_index: 要移动的幻灯片索引
            new_position: 新位置

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _move_slide(ppt, {
                "slide_index": slide_index,
                "new_position": new_position,
            })
            return f"已移动幻灯片: slide {slide_index} -> {new_position}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_list_slides(file_path: str) -> dict:
        """列出所有幻灯片信息.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_slides(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_slide_info(file_path: str, slide_index: int = 1) -> dict:
        """获取单个幻灯片信息.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_slide_info(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_slide_notes(file_path: str, slide_index: int = 1) -> dict:
        """获取幻灯片备注.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_slide_notes(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_slide_notes_extended(
        file_path: str,
        slide_index: int = 1,
        text: str = "",
        font_size: float = None,
        font_name: str = None,
    ) -> str:
        """设置幻灯片备注（扩展参数）.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            text: 备注文本
            font_size: 字体大小 (可选)
            font_name: 字体名称 (可选)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_slide_notes_extended(ppt, {
                "slide_index": slide_index,
                "text": text,
                "font_size": font_size,
                "font_name": font_name,
            })
            return f"已设置幻灯片备注: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_slide_layouts(file_path: str) -> dict:
        """获取可用布局列表.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_slide_layouts(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_apply_layout(
        file_path: str, slide_index: int = 1, layout: str = "title_content"
    ) -> str:
        """应用布局到幻灯片.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            layout: 布局名称或索引 (title/title_content/blank/section_header/two_content/comparison)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _apply_layout(ppt, {
                "slide_index": slide_index,
                "layout": layout,
            })
            return f"已应用布局: slide {slide_index}, layout {layout}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_slide_size(file_path: str) -> dict:
        """获取幻灯片尺寸.

        Args:
            file_path: 演示文稿路径

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_slide_size(ppt, {})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    # ============ Shapes 类工具 (10 个新工具) ============

    @mcp.tool()
    def ppt_list_shapes(file_path: str, slide_index: int = 1) -> dict:
        """列出幻灯片上的所有形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _list_shapes(ppt, {"slide_index": slide_index})
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_shape_info(file_path: str, slide_index: int = 1, shape_index: int = 1) -> dict:
        """获取形状详细信息.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_shape_info(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_update_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        left: float = None,
        top: float = None,
        width: float = None,
        height: float = None,
        rotation: float = None,
        name: str = None,
        visible: bool = None,
    ) -> str:
        """更新形状属性.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            left: 左边距 (可选)
            top: 上边距 (可选)
            width: 宽度 (可选)
            height: 高度 (可选)
            rotation: 旋转角度 (可选)
            name: 形状名称 (可选)
            visible: 是否可见 (可选)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _update_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "rotation": rotation,
                "name": name,
                "visible": visible,
            })
            return f"已更新形状属性: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_delete_shape(file_path: str, slide_index: int = 1, shape_index: int = 1) -> str:
        """删除形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _delete_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return f"已删除形状: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_zorder(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        zorder_action: str = "bring_to_front",
    ) -> str:
        """设置形状层级.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            zorder_action: 层级操作 (bring_to_front/send_to_back/bring_forward/send_backward)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _set_zorder(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "zorder_action": zorder_action,
            })
            return f"已设置形状层级: {zorder_action}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_line(
        file_path: str,
        slide_index: int = 1,
        begin_x: float = 100,
        begin_y: float = 100,
        end_x: float = 200,
        end_y: float = 100,
        line_color: str = "#000000",
        line_width: float = 1,
    ) -> str:
        """添加线条.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            begin_x: 起点 X 坐标
            begin_y: 起点 Y 坐标
            end_x: 终点 X 坐标
            end_y: 终点 Y 坐标
            line_color: 线条颜色 (#RRGGBB)
            line_width: 线条宽度 (pt)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_line(ppt, {
                "slide_index": slide_index,
                "begin_x": begin_x,
                "begin_y": begin_y,
                "end_x": end_x,
                "end_y": end_y,
                "line_color": line_color,
                "line_width": line_width,
            })
            return f"已添加线条: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_textbox_extended(
        file_path: str,
        slide_index: int = 1,
        text: str = "",
        left: float = 100,
        top: float = 100,
        width: float = 400,
        height: float = 100,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
        bold: bool = None,
        alignment: str = None,
    ) -> str:
        """添加文本框（扩展功能）.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            text: 文本内容
            left: 左边距
            top: 上边距
            width: 宽度
            height: 高度
            font_name: 字体名称 (可选)
            font_size: 字体大小 (可选)
            font_color: 字体颜色 (#RRGGBB) (可选)
            bold: 是否加粗 (可选)
            alignment: 对齐方式 (left/center/right) (可选)

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_textbox(ppt, {
                "slide_index": slide_index,
                "text": text,
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color,
                "bold": bold,
                "alignment": alignment,
            })
            return f"已添加文本框: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_picture_extended(
        file_path: str,
        slide_index: int = 1,
        image_path: str = "",
        left: float = 100,
        top: float = 100,
        width: float = None,
        height: float = None,
        link_to_file: bool = False,
        save_with_document: bool = True,
    ) -> str:
        """添加图片（扩展功能）.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            image_path: 图片路径
            left: 左边距
            top: 上边距
            width: 宽度 (可选，保持比例)
            height: 高度 (可选，保持比例)
            link_to_file: 是否链接到文件
            save_with_document: 是否随文档保存

        """
        try:
            path = validate_path(file_path)
            img_path = validate_path(image_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_picture_extended(ppt, {
                "slide_index": slide_index,
                "image_path": str(img_path),
                "left": left,
                "top": top,
                "width": width,
                "height": height,
                "link_to_file": link_to_file,
                "save_with_document": save_with_document,
            })
            return f"已添加图片: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_duplicate_shape(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        offset_left: float = 20,
        offset_top: float = 20,
    ) -> str:
        """复制形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            offset_left: 复制后 X 偏移
            offset_top: 复制后 Y 偏移

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _duplicate_shape(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "offset_left": offset_left,
                "offset_top": offset_top,
            })
            return f"已复制形状: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_group_shapes(
        file_path: str,
        slide_index: int = 1,
        shape_indices: list[int] = None,
    ) -> str:
        """组合形状.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_indices: 要组合的形状索引列表

        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _group_shapes(ppt, {
                "slide_index": slide_index,
                "shape_indices": shape_indices or [],
            })
            return f"已组合形状: {len(shape_indices or [])} 个形状"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Text 类工具 (10 个新工具) ============

    @mcp.tool()
    def ppt_set_text(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        text: str = "",
    ) -> str:
        """设置形状文本.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            text: 文本内容
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_text(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "text": text,
            })
            return f"已设置文本: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_text(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取形状文本.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_get_text(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_format_text_range(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        start: int = 1,
        length: int = 0,
        font_name: str = "",
        font_size: float = 0,
        font_color: str = "",
        bold: bool | None = None,
        italic: bool | None = None,
        underline: bool | None = None,
    ) -> str:
        """格式化文本范围.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            start: 开始位置 (1-based)
            length: 文本长度
            font_name: 字体名称
            font_size: 字体大小 (pt)
            font_color: 字体颜色 (#RRGGBB)
            bold: 是否粗体
            italic: 是否斜体
            underline: 是否下划线
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_format_text_range(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "start": start,
                "length": length,
                "font_name": font_name,
                "font_size": font_size,
                "font_color": font_color,
                "bold": bold,
                "italic": italic,
                "underline": underline,
            })
            return f"已格式化文本范围: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_paragraph_format(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        paragraph_index: int = 1,
        alignment: str = "left",
        line_spacing: float = 0,
        space_before: float = 0,
        space_after: float = 0,
    ) -> str:
        """设置段落格式.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            paragraph_index: 段落索引 (1-based)
            alignment: 对齐方式 (left/center/right/distribute)
            line_spacing: 行间距 (pt)
            space_before: 段前间距 (pt)
            space_after: 段后间距 (pt)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_paragraph_format(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "paragraph_index": paragraph_index,
                "alignment": alignment,
                "line_spacing": line_spacing,
                "space_before": space_before,
                "space_after": space_after,
            })
            return f"已设置段落格式: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_bullets(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        paragraph_index: int | None = None,
        bullet_type: str = "bullet",
        bullet_char: str = "",
        bullet_font: str = "",
        bullet_size: float = 0,
        indent_level: int = 0,
    ) -> str:
        """设置项目符号.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            paragraph_index: 段落索引 (可选，默认全部)
            bullet_type: 项目符号类型 (none/numbered/bullet/picture)
            bullet_char: 自定义符号字符
            bullet_font: 符号字体
            bullet_size: 符号大小 (pt)
            indent_level: 缩进级别 (0-9)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_bullets(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "paragraph_index": paragraph_index,
                "bullet_type": bullet_type,
                "bullet_char": bullet_char,
                "bullet_font": bullet_font,
                "bullet_size": bullet_size,
                "indent_level": indent_level,
            })
            return f"已设置项目符号: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_find_replace_text(
        file_path: str,
        slide_index: int | None = None,
        find_text: str = "",
        replace_text: str = "",
        match_case: bool = False,
        match_whole_word: bool = False,
    ) -> dict:
        """查找替换文本.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引 (可选，默认全部)
            find_text: 查找文本
            replace_text: 替换文本
            match_case: 是否区分大小写
            match_whole_word: 是否全字匹配
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_find_replace_text(ppt, {
                "slide_index": slide_index,
                "find_text": find_text,
                "replace_text": replace_text,
                "match_case": match_case,
                "match_whole_word": match_whole_word,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_textframe(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取文本框属性.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_get_textframe(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_extract_text_as_markdown(
        file_path: str,
        slide_index: int | None = None,
    ) -> dict:
        """提取所有文本为 Markdown.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引 (可选，默认全部)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_extract_text_as_markdown(ppt, {
                "slide_index": slide_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_font_size(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        font_size: float = 12,
        start: int | None = None,
        length: int | None = None,
    ) -> str:
        """设置字体大小.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            font_size: 字体大小 (pt)
            start: 开始位置 (可选)
            length: 文本长度 (可选)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_font_size(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "font_size": font_size,
                "start": start,
                "length": length,
            })
            return f"已设置字体大小: {font_size}pt"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_font_color(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        font_color: str = "#000000",
        start: int | None = None,
        length: int | None = None,
    ) -> str:
        """设置字体颜色.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            font_color: 字体颜色 (#RRGGBB)
            start: 开始位置 (可选)
            length: 文本长度 (可选)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_font_color(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "font_color": font_color,
                "start": start,
                "length": length,
            })
            return f"已设置字体颜色: {font_color}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Placeholders 类工具 (6 个新工具) ============

    @mcp.tool()
    def ppt_list_placeholders(
        file_path: str,
        slide_index: int = 1,
    ) -> dict:
        """列出幻灯片占位符.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_list_placeholders(ppt, {
                "slide_index": slide_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_get_placeholder(
        file_path: str,
        slide_index: int = 1,
        placeholder_index: int = 1,
    ) -> dict:
        """获取占位符内容.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            placeholder_index: 占位符索引 (1-based)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_get_placeholder(ppt, {
                "slide_index": slide_index,
                "placeholder_index": placeholder_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_placeholder(
        file_path: str,
        slide_index: int = 1,
        placeholder_index: int = 1,
        text: str = "",
    ) -> str:
        """设置占位符内容.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            placeholder_index: 占位符索引
            text: 文本内容
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_placeholder(ppt, {
                "slide_index": slide_index,
                "placeholder_index": placeholder_index,
                "text": text,
            })
            return f"已设置占位符: slide {slide_index}, placeholder {placeholder_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_clear_placeholder(
        file_path: str,
        slide_index: int = 1,
        placeholder_index: int = 1,
    ) -> str:
        """清空占位符.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            placeholder_index: 占位符索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_clear_placeholder(ppt, {
                "slide_index": slide_index,
                "placeholder_index": placeholder_index,
            })
            return f"已清空占位符: slide {slide_index}, placeholder {placeholder_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_placeholder_type(
        file_path: str,
        slide_index: int = 1,
        placeholder_index: int = 1,
    ) -> dict:
        """获取占位符类型.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            placeholder_index: 占位符索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_get_placeholder_type(ppt, {
                "slide_index": slide_index,
                "placeholder_index": placeholder_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_resize_placeholder(
        file_path: str,
        slide_index: int = 1,
        placeholder_index: int = 1,
        width: float | None = None,
        height: float | None = None,
        left: float | None = None,
        top: float | None = None,
    ) -> str:
        """调整占位符大小.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            placeholder_index: 占位符索引
            width: 宽度 (pt)
            height: 高度 (pt)
            left: 左边距 (可选)
            top: 上边距 (可选)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_resize_placeholder(ppt, {
                "slide_index": slide_index,
                "placeholder_index": placeholder_index,
                "width": width,
                "height": height,
                "left": left,
                "top": top,
            })
            return f"已调整占位符大小: slide {slide_index}, placeholder {placeholder_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Formatting 类工具 (3 个新工具) ============

    @mcp.tool()
    def ppt_set_fill(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        fill_type: str = "solid",
        color: str = "#FFFFFF",
        gradient_colors: list[str] | None = None,
        gradient_type: str = "linear",
        picture_path: str = "",
        transparency: float = 0,
    ) -> str:
        """设置填充（颜色/渐变/图片）.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            fill_type: 填充类型 (solid/gradient/picture/pattern/no_fill)
            color: 颜色 (#RRGGBB)
            gradient_colors: 渐变颜色列表
            gradient_type: 渐变类型 (linear/radial)
            picture_path: 图片路径
            transparency: 透明度 (0-100)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            if picture_path:
                validate_path(picture_path)
            result = _ppt_set_fill(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "fill_type": fill_type,
                "color": color,
                "gradient_colors": gradient_colors or [],
                "gradient_type": gradient_type,
                "picture_path": picture_path,
                "transparency": transparency,
            })
            return f"已设置填充: slide {slide_index}, shape {shape_index}, type={fill_type}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_line(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        line_color: str = "#000000",
        line_width: float = 1,
        dash_style: str = "solid",
        transparency: float = 0,
    ) -> str:
        """设置线条样式.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            line_color: 线条颜色 (#RRGGBB)
            line_width: 线条宽度 (pt)
            dash_style: 虚线样式 (solid/dash/dash_dot/dot/dash_dot_dot)
            transparency: 透明度 (0-100)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_line(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "line_color": line_color,
                "line_width": line_width,
                "dash_style": dash_style,
                "transparency": transparency,
            })
            return f"已设置线条样式: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_shadow(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        shadow_type: str = "offset",
        shadow_color: str = "#808080",
        shadow_blur: float = 4,
        shadow_offset_x: float = 3,
        shadow_offset_y: float = 3,
        shadow_transparency: float = 50,
    ) -> str:
        """设置阴影效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            shadow_type: 阴影类型 (none/offset/double/perspective)
            shadow_color: 阴影颜色 (#RRGGBB)
            shadow_blur: 模糊半径 (pt)
            shadow_offset_x: X 偏移 (pt)
            shadow_offset_y: Y 偏移 (pt)
            shadow_transparency: 透明度 (0-100)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_shadow(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "shadow_type": shadow_type,
                "shadow_color": shadow_color,
                "shadow_blur": shadow_blur,
                "shadow_offset_x": shadow_offset_x,
                "shadow_offset_y": shadow_offset_y,
                "shadow_transparency": shadow_transparency,
            })
            return f"已设置阴影效果: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Tables 类工具 (13 个新工具) ============

    @mcp.tool()
    def ppt_get_table_cells(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        start_row: int = 1,
        start_col: int = 1,
        end_row: int | None = None,
        end_col: int | None = None,
    ) -> dict:
        """获取表格单元格内容.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引 (表格)
            start_row: 开始行
            start_col: 开始列
            end_row: 结束行 (可选)
            end_col: 结束列 (可选)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_get_table_cells(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "start_row": start_row,
                "start_col": start_col,
                "end_row": end_row,
                "end_col": end_col,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_set_table_cells(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        row: int = 1,
        col: int = 1,
        text: str = "",
    ) -> str:
        """设置表格单元格内容.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            row: 行号
            col: 列号
            text: 文本内容
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_table_cells(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "row": row,
                "col": col,
                "text": text,
            })
            return f"已设置单元格: ({row},{col})"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_batch_set_table_data(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        data: list[list[str]] | None = None,
        start_row: int = 1,
        start_col: int = 1,
    ) -> str:
        """批量设置表格数据.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            data: 二维数组数据
            start_row: 开始行
            start_col: 开始列
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_batch_set_table_data(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "data": data or [],
                "start_row": start_row,
                "start_col": start_col,
            })
            return f"已批量设置表格数据: {len(data or [])} 行"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_merge_table_cells(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        start_row: int = 1,
        start_col: int = 1,
        end_row: int = 1,
        end_col: int = 1,
    ) -> str:
        """合并单元格.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            start_row: 开始行
            start_col: 开始列
            end_row: 结束行
            end_col: 结束列
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=True)
            result = _ppt_merge_table_cells(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "start_row": start_row,
                "start_col": start_col,
                "end_row": end_row,
                "end_col": end_col,
            })
            return f"已合并单元格: ({start_row},{start_col})-({end_row},{end_col})"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_split_table_cells(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        row: int = 1,
        col: int = 1,
    ) -> str:
        """拆分单元格.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            row: 行号
            col: 列号
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=True)
            result = _ppt_split_table_cells(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "row": row,
                "col": col,
            })
            return f"已拆分单元格: ({row},{col})"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_table_row(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        before_row: int | None = None,
    ) -> str:
        """添加表格行.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            before_row: 在此行之前插入 (可选，默认末尾)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_add_table_row(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "before_row": before_row,
            })
            return f"已添加表格行"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_delete_table_row(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        row: int = 1,
    ) -> str:
        """删除表格行.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            row: 行号
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_delete_table_row(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "row": row,
            })
            return f"已删除表格行: {row}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_add_table_column(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        before_col: int | None = None,
    ) -> str:
        """添加表格列.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            before_col: 在此列之前插入 (可选，默认末尾)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_add_table_column(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "before_col": before_col,
            })
            return f"已添加表格列"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_delete_table_column(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        col: int = 1,
    ) -> str:
        """删除表格列.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            col: 列号
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_delete_table_column(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "col": col,
            })
            return f"已删除表格列: {col}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_table_style(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        style_name: str = "",
    ) -> str:
        """设置表格样式.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            style_name: 样式名称或索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_table_style(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "style_name": style_name,
            })
            return f"已设置表格样式: {style_name}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_set_table_borders(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        border_color: str = "#000000",
        border_width: float = 1,
        border_style: str = "solid",
        apply_to: str = "all",
    ) -> str:
        """设置表格边框.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            border_color: 边框颜色 (#RRGGBB)
            border_width: 边框宽度 (pt)
            border_style: 边框样式 (solid/dash/dot)
            apply_to: 应用范围 (all/outer/inner/horizontal/vertical)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_set_table_borders(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "border_color": border_color,
                "border_width": border_width,
                "border_style": border_style,
                "apply_to": apply_to,
            })
            return f"已设置表格边框: apply_to={apply_to}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_table_info(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取表格信息.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_get_table_info(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_resize_table(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        width: float | None = None,
        height: float | None = None,
        left: float | None = None,
        top: float | None = None,
    ) -> str:
        """调整表格大小.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            width: 宽度 (pt)
            height: 高度 (pt)
            left: 左边距 (可选)
            top: 上边距 (可选)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ppt_resize_table(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "width": width,
                "height": height,
                "left": left,
                "top": top,
            })
            return f"已调整表格大小"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Connectors 连接线工具 ============

    @mcp.tool()
    def ppt_add_connector(
        file_path: str,
        slide_index: int = 1,
        connector_type: str = "elbow",
        start_shape: int = 1,
        start_connection: int = 1,
        end_shape: int = 2,
        end_connection: int = 1,
    ) -> str:
        """在形状之间添加连接线.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            connector_type: 连接线类型 (straight/elbow/curve)
            start_shape: 起始形状索引
            start_connection: 起始连接点索引
            end_shape: 目标形状索引
            end_connection: 目标连接点索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_connector(ppt, {
                "slide_index": slide_index,
                "connector_type": connector_type,
                "start_shape": start_shape,
                "start_connection": start_connection,
                "end_shape": end_shape,
                "end_connection": end_connection,
            })
            return f"已添加连接线: slide {slide_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_format_connector(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        color: str = None,
        weight: float = None,
        dash_style: str = None,
        arrow_begin: str = None,
        arrow_end: str = None,
    ) -> str:
        """格式化连接线外观.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            color: 线条颜色
            weight: 线条粗细
            dash_style: 虚线样式 (solid/dash/dash_dot)
            arrow_begin: 起始箭头 (none/arrow/stealth/diamond/oval)
            arrow_end: 结束箭头 (none/arrow/stealth/diamond/oval)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _format_connector(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "color": color,
                "weight": weight,
                "dash_style": dash_style,
                "arrow_begin": arrow_begin,
                "arrow_end": arrow_end,
            })
            return f"已格式化连接线"
        except OfficeMCPError as e:
            return f"错误: {e}"

    # ============ Groups 取消组合 / 获取组内项目 ============

    @mcp.tool()
    def ppt_ungroup_shapes(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> str:
        """取消组合.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 组合形状索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _ungroup_ppt_shapes(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return f"已取消组合: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_group_items(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> dict:
        """获取组合内所有项目.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 组合形状索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_ppt_group_items(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}

    # ============ Hyperlinks 超链接工具 ============

    @mcp.tool()
    def ppt_add_hyper_link(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
        url: str = "",
        sub_address: str = "",
    ) -> str:
        """添加超链接.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            url: 超链接 URL
            sub_address: 子地址 (如跳转到某张幻灯片)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _add_ppt_hyperlink(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
                "url": url,
                "sub_address": sub_address,
            })
            return f"已添加超链接: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"

    @mcp.tool()
    def ppt_get_hyperlinks(
        file_path: str,
        slide_index: int = 1,
    ) -> dict:
        """获取幻灯片上所有超链接.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _get_ppt_hyperlinks(ppt, {"slide_index": slide_index})
            return {"slide_index": slide_index, "hyperlinks": result}
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def ppt_remove_hyper_link(
        file_path: str,
        slide_index: int = 1,
        shape_index: int = 1,
    ) -> str:
        """移除超链接.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.ensure_document(path, activate=False)
            result = _remove_ppt_hyperlink(ppt, {
                "slide_index": slide_index,
                "shape_index": shape_index,
            })
            return f"已移除超链接: slide {slide_index}, shape {shape_index}"
        except OfficeMCPError as e:
            return f"错误: {e}"



