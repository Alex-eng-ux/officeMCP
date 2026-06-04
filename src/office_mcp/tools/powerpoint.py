"""PowerPoint MCP 工具."""

from mcp.server.fastmcp import FastMCP

from office_mcp.config import settings
from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.operations.ppt_ops import (
    _add_animation, _set_transition, _add_section,
    _format_shape, _set_slide_number,
    _set_master_background, _add_master_shape,
)
from office_mcp.operations.ppt_ops import apply_ppt_operations


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

            presentation = office_manager.get_document(path)
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
    def ppt_add_animation(file_path: str, slide_index: int = 1, shape_index: int = 1, animation_type: str = "fade") -> str:
        """添加动画效果.

        Args:
            file_path: 演示文稿路径
            slide_index: 幻灯片索引
            shape_index: 形状索引
            animation_type: 动画类型 (appear/fade/wipe/push/dissolve/fly)
        """
        try:
            path = validate_path(file_path)
            ppt = office_manager.get_document(path)
            result = _add_animation(ppt, {"slide_index": slide_index, "shape_index": shape_index, "animation_type": animation_type})
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
            ppt = office_manager.get_document(path)
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
            ppt = office_manager.get_document(path)
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
            ppt = office_manager.get_document(path)
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
            ppt = office_manager.get_document(path)
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
            ppt = office_manager.get_document(path)
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
            ppt = office_manager.get_document(path)
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
