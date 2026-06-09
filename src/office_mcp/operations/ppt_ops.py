"""PowerPoint COM 操作实现."""

import ipaddress
import json
import logging
import os
import socket
import tempfile
import urllib.request
from pathlib import Path
from typing import Any

from office_mcp.core.errors import COMOperationError
from office_mcp.core.office_manager import _open_powerpoint_presentation, office_manager
from office_mcp.core.path_guard import validate_path
from office_mcp.utils.icons import get_icon_url, search_icons

logger = logging.getLogger(__name__)


def _ppt_open_for_internal_flow(
    app: Any,
    file_path: str | Path,
    *,
    read_only: bool = False,
    untitled: bool = False,
) -> Any:
    """Open a presentation for internal helper flows with explicit non-modal flags."""
    return _open_powerpoint_presentation(
        app,
        file_path,
        read_only=read_only,
        untitled=untitled,
        with_window=False,
    )


def _ppt_normalize_live_path(path_value: str | Path) -> str:
    """Normalize a PowerPoint file path for live-session comparisons."""
    return str(Path(path_value).resolve()).lower()


def _ppt_acquire_for_internal_flow(
    app: Any,
    file_path: str | Path,
    *,
    read_only: bool = False,
    untitled: bool = False,
) -> tuple[Any, bool]:
    """Return a live presentation plus whether this helper owns the handle."""
    normalized_target = _ppt_normalize_live_path(file_path)
    try:
        count = getattr(app.Presentations, "Count", 0)
        for index in range(1, count + 1):
            candidate = app.Presentations(index)
            full_name = (getattr(candidate, "FullName", "") or "").strip()
            if full_name and _ppt_normalize_live_path(full_name) == normalized_target:
                return candidate, False
    except Exception:
        logger.debug("Could not scan live presentations before internal open", exc_info=True)
    return _ppt_open_for_internal_flow(app, file_path, read_only=read_only, untitled=untitled), True


def _ppt_close_quietly(presentation: Any) -> None:
    """Best-effort close for helper-opened presentations."""
    if presentation is None:
        return
    try:
        presentation.Close()
    except Exception:
        pass

# 操作 op 中以 _path / _file 结尾或明确为路径字段的 key 集合
_PATH_FIELDS = (
    "image_path", "source_path", "target_path", "template_path",
    "file_path", "new_path", "output_path", "output_dir", "from_file",
    "to_file", "src_path", "dst_path", "data_source",
    "video_path", "audio_path", "icon_path",
    "export_path", "save_path",
)


def _validate_op_paths(op: dict) -> None:
    """校验 op dict 中所有疑似路径字段, 防止任意文件访问."""
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

# 布局映射
LAYOUT_MAP = {
    "title": 1,              # ppLayoutTitle
    "title_content": 2,      # ppLayoutText
    "blank": 12,             # ppLayoutBlank
    "section_header": 33,    # ppLayoutSectionHeader
    "two_content": 29,       # ppLayoutTwoObjects
    "comparison": 34,        # ppLayoutComparison
}

# 形状类型映射
SHAPE_MAP = {
    "rectangle": 1,          # msoShapeRectangle
    "oval": 9,               # msoShapeOval
    "rounded_rectangle": 5,  # msoShapeRoundedRectangle
}

# 主题色映射
THEME_COLOR_MAP = {
    "dark1": 0,        # msoThemeColorDark1
    "light1": 1,       # msoThemeColorLight1
    "dark2": 2,        # msoThemeColorDark2
    "light2": 3,       # msoThemeColorLight2
    "accent1": 4,      # msoThemeColorAccent1
    "accent2": 5,      # msoThemeColorAccent2
    "accent3": 6,      # msoThemeColorAccent3
    "accent4": 7,      # msoThemeColorAccent4
    "accent5": 8,      # msoThemeColorAccent5
    "accent6": 9,      # msoThemeColorAccent6
    "hyperlink": 10,   # msoThemeColorHyperlink
    "followed_hyperlink": 11,  # msoThemeColorFollowedHyperlink
}

# SmartArt 类型映射 (常见类型的 MsoSmartArtLayout 枚举值)
SMARTART_TYPES = {
    "org_chart": 1,           # msoSmartArtLayoutOrganizationChart
    "cycle": 2,               # msoSmartArtLayoutCycle
    "pyramid": 3,             # msoSmartArtLayoutPyramid
    "process": 4,             # msoSmartArtLayoutProcess
    "list": 5,                # msoSmartArtLayoutList
    "hierarchy": 6,           # msoSmartArtLayoutHierarchy
    "relationship": 7,        # msoSmartArtLayoutRelationship
    "matrix": 8,              # msoSmartArtLayoutMatrix
    "picture": 9,             # msoSmartArtLayoutPicture
}

# 连接线类型
CONNECTOR_TYPES_PPT = {"straight": 1, "elbow": 2, "curve": 3}

# 渐变类型
GRADIENT_TYPES = {
    "linear": 1,     # msoGradientLinear
    "radial": 2,     # msoGradientRadial
    "rectangular": 3, # msoGradientRectangular
    "path": 4,       # msoGradientPath
}

# 动画触发方式
ANIMATION_TRIGGERS = {
    "on_click": 0,       # msoAnimTriggerOnClick
    "after_previous": 1, # msoAnimTriggerAfterPrevious
    "with_previous": 2,  # msoAnimTriggerWithPrevious
}

_TABLE_MERGE_TAG_PREFIX = "OfficeMCP.TableMerge."


def _ppt_get_cell_text(cell: Any) -> str:
    """Return plain text for a PowerPoint table cell."""
    try:
        return str(cell.Shape.TextFrame.TextRange.Text or "")
    except Exception:  # noqa: BLE001
        return ""


def _ppt_set_cell_text(cell: Any, text: str) -> None:
    """Set plain text for a PowerPoint table cell."""
    cell.Shape.TextFrame.TextRange.Text = text


def _ppt_require_table(shape: Any, operation: str) -> Any:
    """Return table object or raise when the shape is not a table."""
    if not shape.HasTable:
        raise COMOperationError(operation, "形状不是表格")
    return shape.Table


def _ppt_normalize_merge_region(table: Any, start_row: int, start_col: int, end_row: int, end_col: int) -> tuple[int, int, int, int]:
    """Clamp and normalize a merge region."""
    start_row, end_row = sorted((int(start_row), int(end_row)))
    start_col, end_col = sorted((int(start_col), int(end_col)))
    if start_row < 1 or start_col < 1:
        raise COMOperationError("ppt_merge_table_cells", "行列索引必须从 1 开始")
    if end_row > table.Rows.Count or end_col > table.Columns.Count:
        raise COMOperationError("ppt_merge_table_cells", "合并范围超出表格边界")
    return start_row, start_col, end_row, end_col


def _ppt_collect_region_texts(table: Any, start_row: int, start_col: int, end_row: int, end_col: int) -> list[list[str]]:
    """Collect cell texts from a table region."""
    rows: list[list[str]] = []
    for row in range(start_row, end_row + 1):
        row_texts: list[str] = []
        for col in range(start_col, end_col + 1):
            row_texts.append(_ppt_get_cell_text(table.Cell(row, col)))
        rows.append(row_texts)
    return rows


def _ppt_tag_name(row: int, col: int) -> str:
    return f"{_TABLE_MERGE_TAG_PREFIX}{row}.{col}"


def _ppt_store_merge_metadata(shape: Any, row: int, col: int, payload: dict[str, Any]) -> None:
    """Persist merge metadata on the table shape when tags are available."""
    tags = getattr(shape, "Tags", None)
    if tags is None:
        return
    tag_name = _ppt_tag_name(row, col)
    try:
        delete = getattr(tags, "Delete", None)
        if callable(delete):
            delete(tag_name)
        tags.Add(tag_name, json.dumps(payload, ensure_ascii=False))
    except Exception as error:  # noqa: BLE001
        logger.debug("Failed to persist PPT table merge metadata: %s", error)


def _ppt_load_merge_metadata(shape: Any, row: int, col: int) -> dict[str, Any] | None:
    """Load merge metadata previously stored on the shape."""
    tags = getattr(shape, "Tags", None)
    if tags is None:
        return None
    tag_name = _ppt_tag_name(row, col)
    try:
        value = tags(tag_name)
    except Exception:  # noqa: BLE001
        value = None
    if not value:
        return None
    try:
        return json.loads(value)
    except json.JSONDecodeError:
        return None


def _ppt_delete_merge_metadata(shape: Any, row: int, col: int) -> None:
    """Delete merge metadata if present."""
    tags = getattr(shape, "Tags", None)
    if tags is None:
        return
    delete = getattr(tags, "Delete", None)
    if not callable(delete):
        return
    try:
        delete(_ppt_tag_name(row, col))
    except Exception:  # noqa: BLE001
        pass


def _ppt_find_merge_metadata(shape: Any, table: Any, row: int, col: int) -> tuple[int, int, dict[str, Any]] | None:
    """Find merge metadata whose merged region contains the target cell."""
    direct = _ppt_load_merge_metadata(shape, row, col)
    if direct:
        return row, col, direct

    for anchor_row in range(1, table.Rows.Count + 1):
        for anchor_col in range(1, table.Columns.Count + 1):
            payload = _ppt_load_merge_metadata(shape, anchor_row, anchor_col)
            if not payload:
                continue
            end_row = int(payload.get("end_row", anchor_row))
            end_col = int(payload.get("end_col", anchor_col))
            if anchor_row <= row <= end_row and anchor_col <= col <= end_col:
                return anchor_row, anchor_col, payload
    return None


def _ppt_require_output(path_value: str, operation: str) -> Path:
    """Validate an output artifact exists and is non-empty."""
    artifact = Path(path_value)
    try:
        if artifact.exists() and artifact.stat().st_size > 0:
            return artifact
    except OSError:
        pass
    raise COMOperationError(operation, f"Expected output artifact was not created: {path_value}")



def apply_ppt_operations(presentation: Any, operations: list[dict]) -> list[dict]:
    """对 PowerPoint 演示文稿执行批量操作.

    Args:
        presentation: PowerPoint Presentation 对象
        operations: 操作列表

    Returns:
        每个操作的执行结果

    """
    results = []
    for op in operations:
        op_type = op.get("type", "")
        try:
            _validate_op_paths(op)
            result = _execute_ppt_operation(presentation, op)
            results.append({"type": op_type, "status": "success", "result": result})
        except Exception as e:
            logger.error(f"PPT 操作失败 [{op_type}]: {e}")
            results.append({"type": op_type, "status": "error", "message": str(e)})
    return results


def _execute_ppt_operation(presentation: Any, op: dict) -> Any:
    """执行单个 PowerPoint 操作."""
    op_type = op.get("type", "")

    if op_type == "add_slide":
        return _add_slide(presentation, op)
    elif op_type == "set_title":
        return _set_title(presentation, op)
    elif op_type == "add_text":
        return _add_text(presentation, op)
    elif op_type == "insert_image":
        return _insert_image(presentation, op)
    elif op_type == "insert_table":
        return _insert_table(presentation, op)
    elif op_type == "set_slide_layout":
        return _set_slide_layout(presentation, op)
    elif op_type == "delete_slide":
        return _delete_slide(presentation, op)
    elif op_type == "set_background_color":
        return _set_background_color(presentation, op)
    elif op_type == "add_shape":
        return _add_shape(presentation, op)
    elif op_type == "set_notes":
        return _set_notes(presentation, op)
    elif op_type == "save":
        presentation.Save()
        return "saved"
    elif op_type == "add_animation":
        return _add_animation(presentation, op)
    elif op_type == "set_transition":
        return _set_transition(presentation, op)
    elif op_type == "add_section":
        return _add_section(presentation, op)
    elif op_type == "format_shape":
        return _format_shape(presentation, op)
    elif op_type == "set_slide_number":
        return _set_slide_number(presentation, op)
    elif op_type == "set_master_background":
        return _set_master_background(presentation, op)
    elif op_type == "add_master_shape":
        return _add_master_shape(presentation, op)
    elif op_type == "set_theme_color":
        return _set_theme_color(presentation, op)
    elif op_type == "create_from_template":
        return _create_from_template(presentation, op)
    elif op_type == "check_typography":
        return _check_typography(presentation, op)
    elif op_type == "set_theme_preset":
        return _set_theme_preset(presentation, op)
    elif op_type == "navigate_to_slide":
        return _navigate_to_slide(presentation, op)
    elif op_type == "search_icons":
        return _search_icons(presentation, op)
    elif op_type == "insert_icon":
        return _insert_icon(presentation, op)
    elif op_type == "add_smartart":
        return _add_smartart(presentation, op)
    elif op_type == "set_gradient_fill":
        return _set_gradient_fill(presentation, op)
    elif op_type == "add_freeform_shape":
        return _add_freeform_shape(presentation, op)
    elif op_type == "add_video":
        return _add_video(presentation, op)
    elif op_type == "add_audio":
        return _add_audio(presentation, op)
    elif op_type == "set_media_settings":
        return _set_media_settings(presentation, op)
    elif op_type == "modify_smartart":
        return _modify_smartart(presentation, op)
    elif op_type == "list_smartart_layouts":
        return _list_smartart_layouts(presentation, op)
    elif op_type == "undo":
        return _undo(presentation, op)
    elif op_type == "redo":
        return _redo(presentation, op)
    elif op_type == "copy_shape":
        return _copy_shape(presentation, op)
    elif op_type == "copy_formatting":
        return _copy_formatting(presentation, op)
    elif op_type == "paste_formatting":
        return _paste_formatting(presentation, op)
    elif op_type == "duplicate_slide_to_end":
        return _duplicate_slide_to_end(presentation, op)
    elif op_type == "align_shapes":
        return _align_shapes(presentation, op)
    elif op_type == "distribute_shapes":
        return _distribute_shapes(presentation, op)
    elif op_type == "set_slide_size":
        return _set_slide_size(presentation, op)
    elif op_type == "flip_shape":
        return _flip_shape(presentation, op)
    elif op_type == "merge_shapes":
        return _merge_shapes(presentation, op)
    elif op_type == "rotate_shape":
        return _rotate_shape(presentation, op)
    elif op_type == "lock_aspect_ratio":
        return _lock_aspect_ratio(presentation, op)
    elif op_type == "set_glow":
        return _set_glow(presentation, op)
    elif op_type == "set_reflection":
        return _set_reflection(presentation, op)
    elif op_type == "set_soft_edge":
        return _set_soft_edge(presentation, op)
    # Comments 批注操作
    elif op_type == "add_comment":
        return _add_comment(presentation, op)
    elif op_type == "list_comments":
        return _list_comments(presentation, op)
    elif op_type == "delete_comment":
        return _delete_comment(presentation, op)
    # Advanced 高级操作
    elif op_type == "set_tags":
        return _set_tags(presentation, op)
    elif op_type == "get_tags":
        return _get_tags(presentation, op)
    elif op_type == "set_default_font":
        return _set_default_font(presentation, op)
    elif op_type == "replace_font":
        return _replace_font(presentation, op)
    elif op_type == "crop_picture":
        return _crop_picture(presentation, op)
    elif op_type == "set_picture_format":
        return _set_picture_format(presentation, op)
    elif op_type == "export_shape":
        return _export_shape(presentation, op)
    elif op_type == "set_shape_visibility":
        return _set_shape_visibility(presentation, op)
    elif op_type == "select_shape":
        return _select_shape(presentation, op)
    elif op_type == "get_selection":
        return _get_selection(presentation, op)
    elif op_type == "set_view":
        return _set_view(presentation, op)
    elif op_type == "copy_animation_from_shape":
        return _copy_animation_from_shape(presentation, op)
    elif op_type == "add_picture_from_url":
        return _add_picture_from_url(presentation, op)
    elif op_type == "add_svg_icon":
        return _add_svg_icon(presentation, op)
    elif op_type == "batch_apply_format":
        return _batch_apply_format(presentation, op)
    elif op_type == "set_default_shape_style":
        return _set_default_shape_style(presentation, op)
    elif op_type == "get_shape_count":
        return _get_shape_count(presentation, op)
    # Freeform 自由路径操作
    elif op_type == "build_freeform_path":
        return _build_freeform_path(presentation, op)
    elif op_type == "get_node_positions":
        return _get_node_positions(presentation, op)
    elif op_type == "set_node_positions":
        return _set_node_positions(presentation, op)
    elif op_type == "insert_node":
        return _insert_node(presentation, op)
    elif op_type == "delete_node":
        return _delete_node(presentation, op)
    elif op_type == "set_node_editing_type":
        return _set_node_editing_type(presentation, op)
    elif op_type == "set_segment_type":
        return _set_segment_type(presentation, op)
    # Export 类操作
    elif op_type == "export_images":
        return _export_images(presentation, op)
    elif op_type == "get_slide_preview":
        return _get_slide_preview(presentation, op)
    elif op_type == "copy_to_clipboard":
        return _copy_to_clipboard(presentation, op)
    elif op_type == "export_html":
        return _export_html(presentation, op)
    # Slideshow 类操作
    elif op_type == "start_slideshow":
        return _start_slideshow(presentation, op)
    elif op_type == "stop_slideshow":
        return _stop_slideshow(presentation, op)
    elif op_type == "slideshow_next":
        return _slideshow_next(presentation, op)
    elif op_type == "slideshow_previous":
        return _slideshow_previous(presentation, op)
    elif op_type == "slideshow_goto":
        return _slideshow_goto(presentation, op)
    elif op_type == "get_slideshow_status":
        return _get_slideshow_status(presentation, op)
    # Charts 类操作
    elif op_type == "add_chart":
        return _add_chart(presentation, op)
    elif op_type == "set_chart_data":
        return _set_chart_data(presentation, op)
    elif op_type == "get_chart_data":
        return _get_chart_data(presentation, op)
    elif op_type == "format_chart":
        return _format_chart(presentation, op)
    elif op_type == "format_chart_axis":
        return _format_chart_axis(presentation, op)
    elif op_type == "set_chart_series":
        return _set_chart_series(presentation, op)
    elif op_type == "change_chart_type":
        return _change_chart_type(presentation, op)
    # Animation 类操作
    elif op_type == "list_animations":
        return _list_animations(presentation, op)
    elif op_type == "update_animation":
        return _update_animation(presentation, op)
    elif op_type == "remove_animation":
        return _remove_animation(presentation, op)
    elif op_type == "clear_animations":
        return _clear_animations(presentation, op)
    elif op_type == "set_animation_trigger":
        return _set_animation_trigger(presentation, op)
    elif op_type == "copy_animation":
        return _copy_animation(presentation, op)
    # App 类操作 (需要 app 参数，通过 presentation.Application 获取)
    elif op_type == "get_app_info":
        return _get_app_info(presentation.Application, op)
    elif op_type == "get_active_window":
        return _get_active_window(presentation.Application, op)
    elif op_type == "set_window_state":
        return _set_window_state(presentation.Application, op)
    elif op_type == "list_presentations":
        return _list_presentations(presentation.Application, op)
    elif op_type == "get_screen_tip":
        return _get_screen_tip(presentation.Application, op)
    # Presentation 类操作
    elif op_type == "get_presentation_info":
        return _get_presentation_info(presentation, op)
    elif op_type == "list_templates":
        return _list_templates(presentation.Application, op)
    elif op_type == "set_properties":
        return _set_properties(presentation, op)
    elif op_type == "get_properties":
        return _get_properties(presentation, op)
    elif op_type == "save_as":
        return _save_as(presentation, op)
    elif op_type == "repair_presentation":
        return _repair_presentation(presentation, op)
    elif op_type == "compare_presentations":
        return _compare_presentations(presentation.Application, op)
    elif op_type == "merge_presentations":
        return _merge_presentations(presentation.Application, op)
    # Slides 类操作
    elif op_type == "duplicate_slide":
        return _duplicate_slide(presentation, op)
    elif op_type == "move_slide":
        return _move_slide(presentation, op)
    elif op_type == "list_slides":
        return _list_slides(presentation, op)
    elif op_type == "get_slide_info":
        return _get_slide_info(presentation, op)
    elif op_type == "get_slide_notes":
        return _get_slide_notes(presentation, op)
    elif op_type == "set_slide_notes_extended":
        return _set_slide_notes_extended(presentation, op)
    elif op_type == "get_slide_layouts":
        return _get_slide_layouts(presentation, op)
    elif op_type == "apply_layout":
        return _apply_layout(presentation, op)
    elif op_type == "get_slide_size":
        return _get_slide_size(presentation, op)
    # Shapes 类操作
    elif op_type == "list_shapes":
        return _list_shapes(presentation, op)
    elif op_type == "get_shape_info":
        return _get_shape_info(presentation, op)
    elif op_type == "update_shape":
        return _update_shape(presentation, op)
    elif op_type == "delete_shape":
        return _delete_shape(presentation, op)
    elif op_type == "set_zorder":
        return _set_zorder(presentation, op)
    elif op_type == "add_line":
        return _add_line(presentation, op)
    elif op_type == "add_textbox":
        return _add_textbox(presentation, op)
    elif op_type == "add_picture_extended":
        return _add_picture_extended(presentation, op)
    elif op_type == "duplicate_shape":
        return _duplicate_shape(presentation, op)
    elif op_type == "group_shapes":
        return _group_shapes(presentation, op)
    # Text 类操作 (10 个新工具)
    elif op_type == "ppt_set_text":
        return _ppt_set_text(presentation, op)
    elif op_type == "ppt_get_text":
        return _ppt_get_text(presentation, op)
    elif op_type == "ppt_format_text_range":
        return _ppt_format_text_range(presentation, op)
    elif op_type == "ppt_set_paragraph_format":
        return _ppt_set_paragraph_format(presentation, op)
    elif op_type == "ppt_set_bullets":
        return _ppt_set_bullets(presentation, op)
    elif op_type == "ppt_find_replace_text":
        return _ppt_find_replace_text(presentation, op)
    elif op_type == "ppt_get_textframe":
        return _ppt_get_textframe(presentation, op)
    elif op_type == "ppt_extract_text_as_markdown":
        return _ppt_extract_text_as_markdown(presentation, op)
    elif op_type == "ppt_set_font_size":
        return _ppt_set_font_size(presentation, op)
    elif op_type == "ppt_set_font_color":
        return _ppt_set_font_color(presentation, op)
    # Placeholders 类操作 (6 个新工具)
    elif op_type == "ppt_list_placeholders":
        return _ppt_list_placeholders(presentation, op)
    elif op_type == "ppt_get_placeholder":
        return _ppt_get_placeholder(presentation, op)
    elif op_type == "ppt_set_placeholder":
        return _ppt_set_placeholder(presentation, op)
    elif op_type == "ppt_clear_placeholder":
        return _ppt_clear_placeholder(presentation, op)
    elif op_type == "ppt_get_placeholder_type":
        return _ppt_get_placeholder_type(presentation, op)
    elif op_type == "ppt_resize_placeholder":
        return _ppt_resize_placeholder(presentation, op)
    # Formatting 类操作 (3 个新工具)
    elif op_type == "ppt_set_fill":
        return _ppt_set_fill(presentation, op)
    elif op_type == "ppt_set_line":
        return _ppt_set_line(presentation, op)
    elif op_type == "ppt_set_shadow":
        return _ppt_set_shadow(presentation, op)
    # Tables 类操作 (13 个新工具)
    elif op_type == "ppt_get_table_cells":
        return _ppt_get_table_cells(presentation, op)
    elif op_type == "ppt_set_table_cells":
        return _ppt_set_table_cells(presentation, op)
    elif op_type == "ppt_batch_set_table_data":
        return _ppt_batch_set_table_data(presentation, op)
    elif op_type == "ppt_merge_table_cells":
        return _ppt_merge_table_cells(presentation, op)
    elif op_type == "ppt_split_table_cells":
        return _ppt_split_table_cells(presentation, op)
    elif op_type == "ppt_add_table_row":
        return _ppt_add_table_row(presentation, op)
    elif op_type == "ppt_delete_table_row":
        return _ppt_delete_table_row(presentation, op)
    elif op_type == "ppt_add_table_column":
        return _ppt_add_table_column(presentation, op)
    elif op_type == "ppt_delete_table_column":
        return _ppt_delete_table_column(presentation, op)
    elif op_type == "ppt_set_table_style":
        return _ppt_set_table_style(presentation, op)
    elif op_type == "ppt_set_table_borders":
        return _ppt_set_table_borders(presentation, op)
    elif op_type == "ppt_get_table_info":
        return _ppt_get_table_info(presentation, op)
    elif op_type == "ppt_resize_table":
        return _ppt_resize_table(presentation, op)
    # Connectors 连接线
    elif op_type == "add_connector":
        return _add_connector(presentation, op)
    elif op_type == "format_connector":
        return _format_connector(presentation, op)
    # Groups 取消组合 / 获取组内项目
    elif op_type == "ungroup_shapes":
        return _ungroup_ppt_shapes(presentation, op)
    elif op_type == "get_group_items":
        return _get_ppt_group_items(presentation, op)
    # Hyperlinks 超链接
    elif op_type == "add_hyper_link":
        return _add_ppt_hyperlink(presentation, op)
    elif op_type == "get_hyperlinks":
        return _get_ppt_hyperlinks(presentation, op)
    elif op_type == "remove_hyper_link":
        return _remove_ppt_hyperlink(presentation, op)
    else:
        raise COMOperationError(f"未知的 PPT 操作类型: {op_type}")


def _add_slide(presentation: Any, op: dict) -> str:
    """添加幻灯片."""
    layout_name = op.get("layout", "title_content")
    layout_val = LAYOUT_MAP.get(layout_name, 2)

    slide = presentation.Slides.Add(
        presentation.Slides.Count + 1,
        layout_val,
    )
    return f"added_slide: index={slide.SlideIndex}, layout={layout_name}"


def _set_title(presentation: Any, op: dict) -> str:
    """设置幻灯片标题."""
    slide_index = op.get("slide_index", 1)
    text = op.get("text", "")

    slide = presentation.Slides(slide_index)
    # 尝试找到标题占位符
    for shape in slide.Shapes:
        if shape.Type == 14:  # msoPlaceholder
            if shape.PlaceholderFormat.Type in (1, 3):
                # ppPlaceholderTitle / ppPlaceholderCenterTitle
                shape.TextFrame.TextRange.Text = text
                return f"set_title: slide={slide_index}, text={text[:50]}"

    # 如果没有找到标题占位符，添加文本框
    slide.Shapes.AddTextbox(
        Orientation=1,
        Left=50,
        Top=30,
        Width=600,
        Height=50,
    ).TextFrame.TextRange.Text = text

    return f"set_title: slide={slide_index}, text={text[:50]}"


def _add_text(presentation: Any, op: dict) -> str:
    """添加文本框."""
    slide_index = op.get("slide_index", 1)
    text = op.get("text", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 400)
    height = op.get("height", 100)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.AddTextbox(
        Orientation=1,
        Left=left,
        Top=top,
        Width=width,
        Height=height,
    )
    shape.TextFrame.TextRange.Text = text

    return f"add_text: slide={slide_index}"


def _insert_image(presentation: Any, op: dict) -> str:
    """插入图片."""
    slide_index = op.get("slide_index", 1)
    image_path = op.get("image_path", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width")
    height = op.get("height")

    if not image_path:
        raise COMOperationError("insert_image", "image_path 不能为空")

    # 安全: 校验图片路径
    image_path = str(validate_path(image_path))

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.AddPicture(
        FileName=image_path,
        LinkToFile=False,
        SaveWithDocument=True,
        Left=left,
        Top=top,
    )

    if width:
        shape.Width = width
    if height:
        shape.Height = height

    return f"insert_image: slide={slide_index}, image={image_path}"


def _insert_table(presentation: Any, op: dict) -> str:
    """插入表格."""
    slide_index = op.get("slide_index", 1)
    rows = op.get("rows", 2)
    columns = op.get("columns", 2)
    data = op.get("data", [])
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 400)
    height = op.get("height", 200)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.AddTable(rows, columns, left, top, width, height)
    table = shape.Table

    for i, row_data in enumerate(data):
        if i >= rows:
            break
        for j, cell_data in enumerate(row_data):
            if j >= columns:
                break
            table.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = str(cell_data)

    return f"insert_table: slide={slide_index}, {rows}x{columns}"


def _set_slide_layout(presentation: Any, op: dict) -> str:
    """设置幻灯片布局."""
    slide_index = op.get("slide_index", 1)
    layout_name = op.get("layout", "title_content")
    layout_val = LAYOUT_MAP.get(layout_name, 2)

    slide = presentation.Slides(slide_index)
    slide.Layout = layout_val
    return f"set_slide_layout: slide={slide_index}, layout={layout_name}"


def _delete_slide(presentation: Any, op: dict) -> str:
    """删除幻灯片."""
    slide_index = op.get("slide_index", 1)
    presentation.Slides(slide_index).Delete()
    return f"delete_slide: {slide_index}"


def _set_background_color(presentation: Any, op: dict) -> str:
    """设置幻灯片背景色."""
    slide_index = op.get("slide_index", 1)
    color = op.get("color", "#FFFFFF")

    slide = presentation.Slides(slide_index)
    rgb = _hex_to_rgb(color)
    slide.FollowMasterBackground = False
    slide.Background.Fill.Solid()
    slide.Background.Fill.ForeColor.RGB = rgb

    return f"set_background_color: slide={slide_index}, color={color}"


def _add_shape(presentation: Any, op: dict) -> str:
    """添加形状."""
    slide_index = op.get("slide_index", 1)
    shape_name = op.get("shape", "rectangle")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 200)
    height = op.get("height", 100)

    shape_type = SHAPE_MAP.get(shape_name, 1)
    slide = presentation.Slides(slide_index)
    slide.Shapes.AddShape(shape_type, left, top, width, height)

    return f"add_shape: slide={slide_index}, shape={shape_name}"


def _set_notes(presentation: Any, op: dict) -> str:
    """设置演讲者备注."""
    slide_index = op.get("slide_index", 1)
    text = op.get("text", "")

    slide = presentation.Slides(slide_index)
    if slide.HasNotesPage:
        notes_text_frame = slide.NotesPage.Shapes.Placeholders(2).TextFrame
        notes_text_frame.TextRange.Text = text

    return f"set_notes: slide={slide_index}"


def _hex_to_rgb(hex_color: str) -> int:
    """将 #RRGGBB 转为 Office RGB 整数."""
    hex_color = hex_color.lstrip("#")
    if len(hex_color) != 6:
        return 0xFFFFFF
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return r + (g << 8) + (b << 16)


# ============ PowerPoint 高级功能 ============

# 动画效果常量
ANIMATION_EFFECTS = {
    "appear": 3844,        # ppEffectAppear
    "fade": 3852,          # ppEffectFade
    "wipe": 3841,          # ppEffectWipe
    "push": 3855,          # ppEffectPush
    "dissolve": 3850,      # ppEffectDissolve
    "fly": 3845,           # ppEffectFly
}

# 转换效果常量
TRANSITION_EFFECTS = {
    "none": 0,             # ppTransitionNone
    "fade": 3844,          # ppTransitionFade
    "blind_down": 257,     # ppTransitionBlinds
    "push": 513,            # ppTransitionPush
    "wipe": 769,            # ppTransitionWipe
    "split": 1025,         # ppTransitionSplit
}


def _add_animation(presentation: Any, op: dict) -> str:
    """添加动画效果.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (1-based)
        animation_type: 动画类型 (appear/fade/wipe/push/dissolve/fly)
        trigger: 触发方式 (on_click/after_previous/with_previous)
        delay: 动画延迟时间 (秒)
        duration: 动画持续时间 (秒)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    animation_type = op.get("animation_type", "fade")
    trigger = op.get("trigger", "on_click")
    delay = op.get("delay", 0)
    duration = op.get("duration")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    effect_val = ANIMATION_EFFECTS.get(animation_type, 3852)  # 默认 fade
    trigger_val = ANIMATION_TRIGGERS.get(trigger, 0)

    # 添加动画
    try:
        effect = slide.TimeLine.MainSequence.AddEffect(
            shape,
            effectId=effect_val,
            Level=0,  # msoAnimateLevel.msoAnimateLevelNone
        )
    except Exception as add_err:
        return f"add_animation: slide={slide_index}, shape={shape_index}, type={animation_type} failed ({add_err})"

    # 设置触发方式
    try:
        effect.Timing.TriggerType = trigger_val
    except:
        pass

    # 设置延迟
    if delay > 0:
        try:
            effect.Timing.TriggerDelayTime = delay
        except:
            pass

    # 设置持续时间
    if duration is not None and duration > 0:
        try:
            effect.Timing.Duration = duration
        except:
            pass

    return f"added_animation: slide={slide_index}, shape={shape_index}, type={animation_type}"


def _set_transition(presentation: Any, op: dict) -> str:
    """设置幻灯片转换效果.

    Args:
        slide_index: 幻灯片索引
        transition_type: 转换类型 (none/fade/blind_down/push/wipe/split)
        duration: 持续时间 (秒)

    """
    slide_index = op.get("slide_index", 1)
    transition_type = op.get("transition_type", "fade")
    duration = op.get("duration", 1.0)

    slide = presentation.Slides(slide_index)

    effect_val = TRANSITION_EFFECTS.get(transition_type, 3844)

    slide.SlideShowTransition.EntryEffect = effect_val
    slide.SlideShowTransition.Duration = duration

    return f"set_transition: slide={slide_index}, type={transition_type}, duration={duration}s"



def _format_shape(presentation: Any, op: dict) -> str:
    """格式化形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        fill_color: 填充颜色 (#RRGGBB)
        line_color: 边框颜色 (#RRGGBB)
        line_width: 边框宽度 (pt)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    fill_color = op.get("fill_color", "")
    line_color = op.get("line_color", "")
    line_width = op.get("line_width", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if fill_color:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = _hex_to_rgb(fill_color)

    if line_color:
        shape.Line.ForeColor.RGB = _hex_to_rgb(line_color)
        shape.Line.Weight = line_width

    return f"formatted_shape: slide={slide_index}, shape={shape_index}"


def _set_slide_number(presentation: Any, op: dict) -> str:
    """设置幻灯片编号显示.

    Args:
        slide_index: 幻灯片索引
        show: 是否显示

    """
    slide_index = op.get("slide_index", 1)
    show = op.get("show", True)

    slide = presentation.Slides(slide_index)
    # 通过占位符找到幻灯片编号并设置可见性
    for shape in slide.Shapes.Placeholders:
        if shape.PlaceholderFormat.Type == 6:  # ppSlideNumber
            shape.Visible = show

    return f"set_slide_number: slide={slide_index}, show={show}"


def _set_master_background(presentation: Any, op: dict) -> str:
    """设置母版背景颜色.

    Args:
        color: 背景颜色 (#RRGGBB)

    """
    color = op.get("color", "#FFFFFF")

    # 获取母版
    master = presentation.SlideMaster

    # 设置背景
    rgb = _hex_to_rgb(color)
    try:
        master.FollowMasterBackground = False  # 确保不跟随主题
    except Exception:
        pass  # FollowMasterBackground may not be settable in all versions
    master.Background.Fill.Visible = True  # 确保 Fill 可见
    master.Background.Fill.Solid()
    master.Background.Fill.ForeColor.RGB = rgb

    return f"set_master_background: {color}"


def _add_section(presentation: Any, op: dict) -> str:
    """添加分节."""
    section_name = op.get("section_name", "新节")
    after_slide_index = op.get("after_slide_index", 0)
    presentation.Sections.Add(section_name, after_slide_index + 1)
    return f"added_section: {section_name}"


def _add_master_shape(presentation: Any, op: dict) -> str:
    """在母版上添加形状 (适用于所有基于此母版的幻灯片).

    Args:
        shape: 形状类型 (rectangle/oval/rounded_rectangle)
        left: 左边距 (pt)
        top: 上边距 (pt)
        width: 宽度 (pt)
        height: 高度 (pt)
        text: 形状内文本
        fill_color: 填充颜色 (#RRGGBB)

    """
    shape_type = op.get("shape", "rectangle")
    left = op.get("left", 0)
    top = op.get("top", 0)
    width = op.get("width", 100)
    height = op.get("height", 100)
    text = op.get("text", "")
    fill_color = op.get("fill_color", "")

    master = presentation.SlideMaster

    # 在母版上添加形状
    shape_type_val = SHAPE_MAP.get(shape_type, 1)
    shape = master.Shapes.AddShape(shape_type_val, left, top, width, height)

    if text:
        shape.TextFrame.TextRange.Text = text

    if fill_color:
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = _hex_to_rgb(fill_color)

    return f"added_master_shape: {shape_type}"


# ============ P1/P2 增强功能 ============

# 主题色预设
THEME_PRESETS = {
    "ocean": {"accent1": "#1E88E5", "accent2": "#00ACC1", "accent3": "#26A69A", "accent4": "#66BB6A", "accent5": "#FDD835", "accent6": "#FF7043"},
    "sunset": {"accent1": "#FF7043", "accent2": "#FFA726", "accent3": "#FFEE58", "accent4": "#66BB6A", "accent5": "#42A5F5", "accent6": "#AB47BC"},
    "forest": {"accent1": "#2E7D32", "accent2": "#66BB6A", "accent3": "#AED581", "accent4": "#8D6E63", "accent5": "#FFA726", "accent6": "#26A69A"},
    "corporate": {"accent1": "#1565C0", "accent2": "#00838F", "accent3": "#00897B", "accent4": "#6D4C41", "accent5": "#EF6C00", "accent6": "#546E7A"},
}


def _set_theme_color(presentation: Any, op: dict) -> str:
    """设置形状或文本的主题色而非硬编码 RGB.

    Args:
        slide_index: 幻灯片编号 (默认当前)
        shape_index: 形状编号
        color_name: 主题色名称 (dark1/light1/accent1-accent6/hyperlink)
        target: 目标 ("fill" / "font" / "line")
        tint: 色调偏移 (-1.0 到 1.0, 可选)

    """
    color_name = op.get("color_name", "accent1")
    target = op.get("target", "fill")
    tint = op.get("tint", 0)

    slide_index = op.get("slide_index")
    if slide_index is None:
        try:
            slide = presentation.Application.ActiveWindow.Selection.SlideRange(1)
        except Exception:
            if presentation.Slides.Count > 0:
                slide = presentation.Slides(1)
            else:
                raise COMOperationError("set_theme_color", "没有可用的幻灯片")
    else:
        slide = presentation.Slides(slide_index)

    shape_index = op.get("shape_index", 1)
    shape = slide.Shapes(shape_index)

    theme_color_idx = THEME_COLOR_MAP.get(color_name, 4)

    if target == "fill":
        shape.Fill.Solid()
        shape.Fill.ForeColor.ObjectThemeColor = theme_color_idx
        if tint:
            shape.Fill.ForeColor.Brightness = tint
    elif target == "font":
        shape.TextFrame.TextRange.Font.Color.ObjectThemeColor = theme_color_idx
        if tint:
            shape.TextFrame.TextRange.Font.Color.Brightness = tint
    elif target == "line":
        shape.Line.ForeColor.ObjectThemeColor = theme_color_idx
        if tint:
            shape.Line.ForeColor.Brightness = tint

    return f"set_theme_color: {color_name} -> {target}"


def _create_from_template(app: Any, op: dict) -> str:
    """从用户模板文件夹创建演示文稿.

    Args:
        template_name: 模板文件名 (不含路径)
        template_path: 完整模板路径 (可选，会覆盖 template_name)

    """
    template_name = op.get("template_name", "")
    template_path = op.get("template_path", "")
    output_path = op.get("output_path", "")

    if not template_path and template_name:
        # 安全：只取文件名部分，防止路径遍历
        template_name = Path(template_name).name
        user_templates = Path.home() / "AppData" / "Roaming" / "Microsoft" / "Templates"
        template_path = user_templates / template_name
        if not template_path.exists():
            office_templates = Path.home() / "Documents" / "Custom Office Templates"
            template_path = office_templates / template_name

    if not template_path:
        raise COMOperationError("create_from_template", "必须提供 template_name 或 template_path")

    # 安全: 显式校验路径, 防止任意文件访问
    try:
        template_path = str(validate_path(str(template_path)))
    except COMOperationError:
        raise

    if not Path(template_path).exists():
        raise COMOperationError("create_from_template", f"模板文件不存在: {template_path}")

    if output_path:
        output_file = validate_path(output_path)
    else:
        template_file = Path(template_path)
        output_file = validate_path(str(template_file.with_name(f"{template_file.stem}.from-template.pptx")))

    presentation = _ppt_open_for_internal_flow(app, template_path, untitled=True)
    try:
        presentation.SaveAs(str(output_file))
        office_manager.track_document(output_file, presentation, app_type="ppt")
        return f"created_from_template: {output_file}"
    except Exception:
        _ppt_close_quietly(presentation)
        raise


def _check_typography(presentation: Any, op: dict) -> dict:
    """检查幻灯片排版质量问题.

    Args:
        slide_index: 指定幻灯片 (默认检查全部)

    """
    issues = []
    slide_index = op.get("slide_index")

    slides_to_check = [presentation.Slides(slide_index)] if slide_index else list(presentation.Slides)

    for slide in slides_to_check:
        slide_num = slide.SlideIndex
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text.strip()
                # 检测孤立行
                paragraphs = shape.TextFrame.TextRange.Paragraphs()
                for p in paragraphs:
                    try:
                        p_text = (p.Text or "").strip()
                    except Exception:
                        continue
                    if p_text and len(p_text) < 20 and not p_text.endswith(":"):
                        issues.append({
                            "slide": slide_num,
                            "shape": shape.Name,
                            "issue": "short_line",
                            "text": p_text[:50],
                            "suggestion": "合并到上一段落或扩展内容",
                        })
                # 检测空文本框
                if not text:
                    issues.append({
                        "slide": slide_num,
                        "shape": shape.Name,
                        "issue": "empty_textbox",
                        "suggestion": "删除空文本框或填充内容",
                    })

    return {
        "slide_index": slide_index or "all",
        "total_issues": len(issues),
        "issues": issues,
    }


def _set_theme_preset(presentation: Any, op: dict) -> str:
    """应用预设主题色方案.

    Args:
        preset: 预设名称 (ocean/sunset/forest/corporate)
        colors: 自定义颜色字典 (可选，覆盖 preset)

    """
    preset = op.get("preset", "")
    colors = op.get("colors", {})

    if not colors and preset:
        colors = THEME_PRESETS.get(preset, {})

    if not colors:
        raise COMOperationError("set_theme_preset", "需要提供 preset 或 colors")

    color_scheme = presentation.SlideMaster.ColorScheme
    if color_scheme is None:
        raise COMOperationError("set_theme_preset", "当前母版不支持 ColorScheme")

    for name, hex_color in colors.items():
        theme_idx = THEME_COLOR_MAP.get(name)
        if theme_idx is not None and theme_idx >= 4:  # accent1=4 起
            rgb = _hex_to_rgb(hex_color)
            color_scheme.Colors(theme_idx).RGB = rgb

    return f"set_theme_preset: {preset or 'custom'}"


def _navigate_to_slide(presentation: Any, op: dict) -> str:
    """实时导航到指定幻灯片，使修改立即可见.

    Args:
        slide_index: 幻灯片编号

    """
    slide_index = op.get("slide_index")
    if slide_index is None:
        return "navigate_to_slide: need slide_index"

    try:
        presentation.Application.ActiveWindow.View.GotoSlide(slide_index)
        return f"navigate_to_slide: {slide_index}"
    except Exception as e:
        raise COMOperationError("navigate_to_slide", str(e))


def _search_icons(presentation: Any, op: dict) -> dict:
    """搜索 Google Material Icons."""
    query = op.get("query", "")
    limit = op.get("limit", 20)

    if not query:
        raise COMOperationError("search_icons", "query 不能为空")

    icons = search_icons(query, limit=limit)
    return {"query": query, "count": len(icons), "icons": icons}


def _insert_icon(presentation: Any, op: dict) -> str:
    """插入图标到幻灯片."""
    slide_index = op.get("slide_index", 1)
    icon_name = op.get("icon_name", "")
    query = op.get("query", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 72)
    height = op.get("height", 72)
    size = op.get("size", 48)

    if not icon_name and query:
        results = search_icons(query, limit=1)
        if results:
            icon_name = results[0]["name"]

    if not icon_name:
        raise COMOperationError("insert_icon", "需要提供 icon_name 或 query")

    # 下载 PNG 到临时文件
    png_url = get_icon_url(icon_name, size)
    temp_path = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            temp_path = f.name
        urllib.request.urlretrieve(png_url, temp_path)

        slide = presentation.Slides(slide_index)
        shape = slide.Shapes.AddPicture(
            FileName=temp_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        return f"insert_icon: slide={slide_index}, icon={icon_name}"
    except Exception as e:
        raise COMOperationError("insert_icon", str(e))
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass


def _add_smartart(presentation: Any, op: dict) -> str:
    """添加 SmartArt 图形.

    Args:
        slide_index: 幻灯片索引
        smartart_type: SmartArt 类型 (org_chart/cycle/pyramid/process/list/hierarchy/relationship/matrix/picture)
        left: 左边距 (pt)
        top: 上边距 (pt)
        width: 宽度 (pt)
        height: 高度 (pt)
        texts: 文本列表，用于填充 SmartArt

    """
    slide_index = op.get("slide_index", 1)
    smartart_type = op.get("smartart_type", "process")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 400)
    height = op.get("height", 300)
    texts = op.get("texts", [])

    slide = presentation.Slides(slide_index)
    layout_index = SMARTART_TYPES.get(smartart_type, 4)

    try:
        # Get the SmartArtLayout COM object from the application
        app = presentation.Application
        try:
            layout_obj = app.SmartArtLayouts(layout_index)
        except Exception:
            # Fallback: try iterating layouts to find one
            layout_obj = None
            for i in range(1, app.SmartArtLayouts.Count + 1):
                layout_obj = app.SmartArtLayouts(i)
                break

        # 添加 SmartArt
        smartart_shape = slide.Shapes.AddSmartArt(
            Layout=layout_obj,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )

        # 填充文本（如果提供）
        if texts and hasattr(smartart_shape, 'SmartArt'):
            try:
                nodes = smartart_shape.SmartArt.AllNodes
                for i, text in enumerate(texts):
                    if i < len(nodes):
                        nodes(i + 1).TextFrame.TextRange.Text = text
            except:
                pass

        return f"add_smartart: slide={slide_index}, type={smartart_type}"
    except Exception as e:
        raise COMOperationError("add_smartart", str(e))


def _set_gradient_fill(presentation: Any, op: dict) -> str:
    """设置形状的渐变填充.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (1-based)
        gradient_type: 渐变类型 (linear/radial/rectangular/path)
        colors: 颜色列表，每个元素为 {"color": "#RRGGBB", "position": 0-100}
        angle: 渐变角度 (仅 linear 类型)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    gradient_type = op.get("gradient_type", "linear")
    colors = op.get("colors", [{"color": "#000000", "position": 0}, {"color": "#FFFFFF", "position": 100}])
    angle = op.get("angle", 90)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    type_val = GRADIENT_TYPES.get(gradient_type, 1)

    try:
        fill = shape.Fill
        fill.Visible = True

        # 设置渐变类型
        if type_val == 1:  # linear
            fill.TwoColorGradient(Style=1, Variant=1)
            fill.GradientAngle = angle
        else:
            fill.TwoColorGradient(Style=type_val, Variant=1)

        # 处理多颜色渐变
        if len(colors) > 2:
            # 删除默认的渐变停止点
            while fill.GradientStops.Count > 0:
                fill.GradientStops.Delete(1)

            # 添加自定义渐变停止点
            for color_data in colors:
                hex_color = color_data.get("color", "#FFFFFF")
                position = color_data.get("position", 0)
                rgb = _hex_to_rgb(hex_color)
                fill.GradientStops.Insert(RGB=rgb, Position=position)
        else:
            # 双颜色渐变
            if len(colors) >= 2:
                rgb1 = _hex_to_rgb(colors[0].get("color", "#000000"))
                rgb2 = _hex_to_rgb(colors[1].get("color", "#FFFFFF"))
                fill.ForeColor.RGB = rgb1
                fill.BackColor.RGB = rgb2

        return f"set_gradient_fill: slide={slide_index}, shape={shape_index}, type={gradient_type}"
    except Exception as e:
        raise COMOperationError("set_gradient_fill", str(e))


def _add_freeform_shape(presentation: Any, op: dict) -> str:
    """添加自由路径形状 (预留接口，暂未实现).

    Args:
        slide_index: 幻灯片索引
        points: 坐标点列表，每个元素为 [x, y]
        left: 左边距 (pt)
        top: 上边距 (pt)
        closed: 是否闭合路径

    """
    slide_index = op.get("slide_index", 1)
    points = op.get("points", [])
    left = op.get("left", 0)
    top = op.get("top", 0)
    closed = op.get("closed", True)

    # 预留接口，暂未实现完整功能
    slide = presentation.Slides(slide_index)

    if not points or len(points) < 2:
        raise COMOperationError("add_freeform_shape", "需要至少提供 2 个坐标点")

    # 支持 points 为 [{"x":..,"y":..}] 或 [[x,y]] 格式
    normalized = []
    for p in points:
        if isinstance(p, dict):
            normalized.append([p.get("x", 0), p.get("y", 0)])
        elif isinstance(p, (list, tuple)) and len(p) >= 2:
            normalized.append([p[0], p[1]])

    if len(normalized) < 2:
        raise COMOperationError("add_freeform_shape", "需要至少提供 2 个有效坐标点")

    try:
        # 简单实现：使用折线连接点
        shape = slide.Shapes.AddLine(
            BeginX=normalized[0][0] + left,
            BeginY=normalized[0][1] + top,
            EndX=normalized[-1][0] + left,
            EndY=normalized[-1][1] + top,
        )

        return f"add_freeform_shape: slide={slide_index} (简单实现，功能预留)"
    except Exception as e:
        raise COMOperationError("add_freeform_shape", str(e))


# ============ App 类操作 (5 个新工具) ============

def _get_app_info(app: Any, op: dict) -> dict:
    """获取 PowerPoint 应用信息（版本、路径）."""
    try:
        return {
            "name": app.Name,
            "version": app.Version,
            "build": app.Build,
            "path": app.Path,
            "visible": app.Visible,
        }
    except Exception as e:
        raise COMOperationError("get_app_info", str(e))


def _get_active_window(app: Any, op: dict) -> dict:
    """获取当前活动窗口信息."""
    try:
        window = app.ActiveWindow
        if window is None:
            return {"active_window": None, "message": "没有活动窗口"}

        return {
            "window_caption": window.Caption,
            "window_state": window.WindowState,  # 0=normal, 1=maximized, 2=minimized
            "left": window.Left,
            "top": window.Top,
            "width": window.Width,
            "height": window.Height,
            "active_presentation": window.Presentation.Name if window.Presentation else None,
            "view_type": window.View.Type if window.View else None,
        }
    except Exception as e:
        raise COMOperationError("get_active_window", str(e))


def _set_window_state(app: Any, op: dict) -> str:
    """设置窗口状态（最大化/最小化/正常）.

    Args:
        state: 窗口状态 (maximized/minimized/normal)

    """
    state = op.get("state", "normal")

    # 窗口状态常量
    WINDOW_STATES = {
        "normal": 0,      # ppWindowNormal
        "maximized": 1,   # ppWindowMaximized
        "minimized": 2,   # ppWindowMinimized
    }

    state_val = WINDOW_STATES.get(state, 0)

    try:
        window = app.ActiveWindow
        if window is None:
            raise COMOperationError("set_window_state", "没有活动窗口")

        try:
            window.WindowState = state_val
        except Exception:
            # Window state may not be settable when app is invisible
            return f"set_window_state_skipped: {state} (window not accessible)"
        return f"set_window_state: {state}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("set_window_state", str(e))


def _list_presentations(app: Any, op: dict) -> dict:
    """列出所有打开的演示文稿."""
    try:
        presentations = []
        for i in range(1, app.Presentations.Count + 1):
            pres = app.Presentations(i)
            presentations.append({
                "name": pres.Name,
                "path": pres.FullName if pres.FullName else pres.Name,
                "slide_count": pres.Slides.Count,
                "saved": pres.Saved,  # True 表示已保存
            })

        return {
            "count": app.Presentations.Count,
            "presentations": presentations,
        }
    except Exception as e:
        raise COMOperationError("list_presentations", str(e))


def _get_screen_tip(app: Any, op: dict) -> dict:
    """获取屏幕提示功能状态."""
    try:
        return {
            "show_screen_tip": app.ShowScreenTipsOnNewShapes,
            "display_alerts": app.DisplayAlerts,  # 0=none, 1=all, 2=modal
        }
    except Exception as e:
        raise COMOperationError("get_screen_tip", str(e))


# ============ Presentation 类操作 (8 个新工具) ============

def _get_presentation_info(presentation: Any, op: dict) -> dict:
    """获取演示文稿信息（幻灯片数、大小）."""
    try:
        return {
            "name": presentation.Name,
            "path": presentation.FullName,
            "slide_count": presentation.Slides.Count,
            "slide_width": presentation.PageSetup.SlideWidth,
            "slide_height": presentation.PageSetup.SlideHeight,
            "saved": presentation.Saved,
            "read_only": presentation.ReadOnly,
            "has_password": presentation.HasPassword,
            "layout": presentation.PageSetup.SlideSize,  # 幻灯片尺寸类型
        }
    except Exception as e:
        raise COMOperationError("get_presentation_info", str(e))


def _list_templates(app: Any, op: dict) -> dict:
    """列出可用模板."""
    try:
        # 检查用户模板文件夹
        user_templates = Path.home() / "AppData" / "Roaming" / "Microsoft" / "Templates"
        office_templates = Path.home() / "Documents" / "Custom Office Templates"

        templates = []

        for template_dir in [user_templates, office_templates]:
            if template_dir.exists():
                for file in template_dir.glob("*.potx"):
                    templates.append({
                        "name": file.name,
                        "path": str(file),
                        "source": "user",
                    })
                for file in template_dir.glob("*.pptx"):
                    templates.append({
                        "name": file.name,
                        "path": str(file),
                        "source": "user",
                    })

        return {
            "count": len(templates),
            "templates": templates,
            "user_templates_dir": str(user_templates),
            "office_templates_dir": str(office_templates),
        }
    except Exception as e:
        raise COMOperationError("list_templates", str(e))


def _set_properties(presentation: Any, op: dict) -> str:
    """设置演示文稿属性（标题、作者）.

    Args:
        title: 标题
        author: 作者
        subject: 主题
        keywords: 关键词
        comments: 备注

    """
    title = op.get("title", "")
    author = op.get("author", "")
    subject = op.get("subject", "")
    keywords = op.get("keywords", "")
    comments = op.get("comments", "")

    try:
        props = presentation.BuiltInDocumentProperties

        if title:
            props("Title").Value = title
        if author:
            props("Author").Value = author
        if subject:
            props("Subject").Value = subject
        if keywords:
            props("Keywords").Value = keywords
        if comments:
            props("Comments").Value = comments

        return f"set_properties: title={title[:30] if title else 'N/A'}"
    except Exception as e:
        raise COMOperationError("set_properties", str(e))


def _get_properties(presentation: Any, op: dict) -> dict:
    """获取演示文稿属性."""
    try:
        props = presentation.BuiltInDocumentProperties

        def get_prop_value(props, name):
            try:
                return props(name).Value
            except:
                return None

        return {
            "title": get_prop_value(props, "Title"),
            "author": get_prop_value(props, "Author"),
            "subject": get_prop_value(props, "Subject"),
            "keywords": get_prop_value(props, "Keywords"),
            "comments": get_prop_value(props, "Comments"),
            "last_author": get_prop_value(props, "Last Author"),
            "revision_number": get_prop_value(props, "Revision Number"),
            "creation_date": get_prop_value(props, "Creation Date"),
            "last_save_time": get_prop_value(props, "Last Save Time"),
        }
    except Exception as e:
        raise COMOperationError("get_properties", str(e))


def _save_as(presentation: Any, op: dict) -> str:
    """另存为指定格式.

    Args:
        file_path: 保存路径
        format: 文件格式 (pptx/pdf/ppt/potx/ppsx)

    """
    file_path = op.get("file_path", "")
    format_name = op.get("format", "pptx")

    if not file_path:
        raise COMOperationError("save_as", "file_path 不能为空")

    # 安全: 校验输出路径
    file_path = str(validate_path(file_path))

    # 文件格式常量
    SAVE_FORMATS = {
        "pptx": 24,   # ppSaveAsOpenXMLPresentation
        "ppt": 1,     # ppSaveAsPresentation
        "pdf": 32,    # ppSaveAsPDF
        "potx": 26,   # ppSaveAsOpenXMLTemplate
        "ppsx": 28,   # ppSaveAsOpenXMLShow
        "xps": 33,    # ppSaveAsXPS
        "gif": 16,    # ppSaveAsGIF
        "jpg": 17,    # ppSaveAsJPG
        "png": 18,    # ppSaveAsPNG
        "bmp": 19,    # ppSaveAsBMP
        "tif": 20,    # ppSaveAsTIF
        "wmv": 36,    # ppSaveAsWMV
        "mp4": 37,    # ppSaveAsMP4
    }

    format_val = SAVE_FORMATS.get(format_name, 24)

    try:
        presentation.SaveAs(file_path, FileFormat=format_val)
        _ppt_require_output(file_path, "save_as")
        return f"save_as: {file_path}, format={format_name}"
    except Exception as e:
        raise COMOperationError("save_as", str(e))


def _repair_presentation(presentation: Any, op: dict) -> dict:
    """修复损坏的演示文稿（检查并尝试修复常见问题）."""
    try:
        issues_fixed = []

        # 检查空幻灯片
        for i in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(i)
            shape_count = slide.Shapes.Count

            # 检查是否有损坏的形状
            for j in range(1, shape_count + 1):
                try:
                    shape = slide.Shapes(j)
                    # 尝试访问基本属性以检测损坏
                    _ = shape.Name
                    _ = shape.Type
                except Exception:
                    issues_fixed.append({
                        "slide": i,
                        "shape": j,
                        "issue": "corrupted_shape",
                        "action": "detected_only",
                    })

        return {
            "status": "checked",
            "issues_found": len(issues_fixed),
            "issues": issues_fixed,
            "note": "PowerPoint COM API 不支持直接删除损坏形状，请手动修复",
        }
    except Exception as e:
        raise COMOperationError("repair_presentation", str(e))


def _compare_presentations(app: Any, op: dict) -> dict:
    """比较两个演示文稿.

    Args:
        presentation1_path: 第一个演示文稿路径
        presentation2_path: 第二个演示文稿路径

    """
    pres1_path = op.get("presentation1_path", "")
    pres2_path = op.get("presentation2_path", "")

    if not pres1_path or not pres2_path:
        raise COMOperationError("compare_presentations", "需要提供两个演示文稿路径")

    # 安全: 校验路径
    pres1_path = str(validate_path(pres1_path))
    pres2_path = str(validate_path(pres2_path))

    pres1 = None
    pres2 = None
    pres1_owned = False
    pres2_owned = False
    try:
        # 打开两个演示文稿进行比较
        pres1, pres1_owned = _ppt_acquire_for_internal_flow(app, pres1_path, read_only=True)
        pres2, pres2_owned = _ppt_acquire_for_internal_flow(app, pres2_path, read_only=True)

        differences = []

        # 比较幻灯片数量
        if pres1.Slides.Count != pres2.Slides.Count:
            differences.append({
                "type": "slide_count",
                "pres1": pres1.Slides.Count,
                "pres2": pres2.Slides.Count,
            })

        # 比较每张幻灯片的形状数量
        min_slides = min(pres1.Slides.Count, pres2.Slides.Count)
        for i in range(1, min_slides + 1):
            slide1 = pres1.Slides(i)
            slide2 = pres2.Slides(i)

            if slide1.Shapes.Count != slide2.Shapes.Count:
                differences.append({
                    "type": "shape_count",
                    "slide": i,
                    "pres1": slide1.Shapes.Count,
                    "pres2": slide2.Shapes.Count,
                })

        return {
            "status": "compared",
            "differences_count": len(differences),
            "differences": differences,
        }
    except Exception as e:
        raise COMOperationError("compare_presentations", str(e))
    finally:
        if pres2_owned:
            _ppt_close_quietly(pres2)
        if pres1_owned:
            _ppt_close_quietly(pres1)


def _merge_presentations(app: Any, op: dict) -> str:
    """合并多个演示文稿.

    Args:
        target_path: 目标演示文稿路径
        source_paths: 源演示文稿路径列表
        insert_position: 插入位置 (默认末尾)

    """
    target_path = op.get("target_path", "")
    source_paths = op.get("source_paths", [])
    insert_position = op.get("insert_position", -1)  # -1 表示末尾

    if not target_path or not source_paths:
        raise COMOperationError("merge_presentations", "需要提供目标路径和源路径列表")

    # 安全: 校验所有路径
    target_path = str(validate_path(target_path))
    source_paths = [str(validate_path(p)) for p in source_paths if p]

    target_pres = None
    source_pres = None
    target_owned = False
    source_owned = False
    try:
        # 打开目标演示文稿
        target_pres, target_owned = _ppt_acquire_for_internal_flow(app, target_path)

        for source_path in source_paths:
            source_pres, source_owned = _ppt_acquire_for_internal_flow(app, source_path, read_only=True)

            # 计算插入位置
            if insert_position < 0 or insert_position > target_pres.Slides.Count:
                insert_at = target_pres.Slides.Count + 1
            else:
                insert_at = insert_position

            # 复制幻灯片
            slides_to_copy = source_pres.Slides.Count
            for i in range(1, slides_to_copy + 1):
                source_pres.Slides(i).Copy()
                target_pres.Slides.Paste(insert_at + i - 1)

            # 更新插入位置
            insert_position = target_pres.Slides.Count

            if source_owned:
                _ppt_close_quietly(source_pres)
            source_pres = None
            source_owned = False

        target_pres.Save()
        return f"merge_presentations: merged {len(source_paths)} presentations into {target_path}"
    except Exception as e:
        raise COMOperationError("merge_presentations", str(e))
    finally:
        if source_owned:
            _ppt_close_quietly(source_pres)
        if target_owned:
            _ppt_close_quietly(target_pres)


# ============ Slides 类操作 (9 个新工具) ============

def _duplicate_slide(presentation: Any, op: dict) -> str:
    """复制幻灯片.

    Args:
        slide_index: 要复制的幻灯片索引
        insert_after: 复制后插入位置 (默认原幻灯片之后)

    """
    slide_index = op.get("slide_index", 1)
    insert_after = op.get("insert_after", slide_index)

    try:
        slide = presentation.Slides(slide_index)
        # 使用 Duplicate 方法，避免剪贴板问题
        duplicated = slide.Duplicate()
        # Duplicate 返回 SlideRange，获取新幻灯片的索引
        new_index = duplicated(1).SlideIndex

        return f"duplicate_slide: slide {slide_index} duplicated at position {new_index}"
    except Exception as e:
        raise COMOperationError("duplicate_slide", str(e))


def _move_slide(presentation: Any, op: dict) -> str:
    """移动幻灯片位置.

    Args:
        slide_index: 要移动的幻灯片索引
        new_position: 新位置

    """
    slide_index = op.get("slide_index", 1)
    new_position = op.get("new_position", slide_index)

    try:
        slide = presentation.Slides(slide_index)
        slide.MoveTo(new_position)

        return f"move_slide: slide {slide_index} moved to position {new_position}"
    except Exception as e:
        raise COMOperationError("move_slide", str(e))


def _list_slides(presentation: Any, op: dict) -> dict:
    """列出所有幻灯片信息."""
    try:
        slides_info = []

        for i in range(1, presentation.Slides.Count + 1):
            slide = presentation.Slides(i)
            slides_info.append({
                "index": i,
                "name": slide.Name,
                "layout": slide.Layout,
                "shape_count": slide.Shapes.Count,
                "has_notes": slide.HasNotesPage,
                "hidden": slide.SlideShowTransition.Hidden,
            })

        return {
            "total_slides": presentation.Slides.Count,
            "slides": slides_info,
        }
    except Exception as e:
        raise COMOperationError("list_slides", str(e))


def _get_slide_info(presentation: Any, op: dict) -> dict:
    """获取单个幻灯片信息.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    try:
        slide = presentation.Slides(slide_index)

        # 获取形状列表
        shapes_info = []
        for j in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(j)
            shapes_info.append({
                "index": j,
                "name": shape.Name,
                "type": shape.Type,
                "left": shape.Left,
                "top": shape.Top,
                "width": shape.Width,
                "height": shape.Height,
            })

        return {
            "index": slide_index,
            "name": slide.Name,
            "layout": slide.Layout,
            "shape_count": slide.Shapes.Count,
            "shapes": shapes_info,
            "has_notes": slide.HasNotesPage,
            "hidden": slide.SlideShowTransition.Hidden,
            "transition": slide.SlideShowTransition.EntryEffect,
        }
    except Exception as e:
        raise COMOperationError("get_slide_info", str(e))


def _get_slide_notes(presentation: Any, op: dict) -> dict:
    """获取幻灯片备注.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    try:
        slide = presentation.Slides(slide_index)

        if not slide.HasNotesPage:
            return {"slide_index": slide_index, "notes": "", "has_notes": False}

        # 获取备注页中的文本
        notes_page = slide.NotesPage
        notes_text = ""

        for shape in notes_page.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                try:
                    notes_text += shape.TextFrame.TextRange.Text
                except:
                    pass

        return {
            "slide_index": slide_index,
            "notes": notes_text,
            "has_notes": True,
        }
    except Exception as e:
        raise COMOperationError("get_slide_notes", str(e))


def _set_slide_notes_extended(presentation: Any, op: dict) -> str:
    """设置幻灯片备注（扩展参数）.

    Args:
        slide_index: 幻灯片索引
        text: 备注文本
        font_size: 字体大小 (可选)
        font_name: 字体名称 (可选)

    """
    slide_index = op.get("slide_index", 1)
    text = op.get("text", "")
    font_size = op.get("font_size")
    font_name = op.get("font_name")

    try:
        slide = presentation.Slides(slide_index)

        if slide.HasNotesPage:
            notes_shape = slide.NotesPage.Shapes.Placeholders(2)
            notes_shape.TextFrame.TextRange.Text = text

            if font_size:
                notes_shape.TextFrame.TextRange.Font.Size = font_size
            if font_name:
                notes_shape.TextFrame.TextRange.Font.Name = font_name

        return f"set_slide_notes: slide={slide_index}, text_length={len(text)}"
    except Exception as e:
        raise COMOperationError("set_slide_notes", str(e))


def _get_slide_layouts(presentation: Any, op: dict) -> dict:
    """获取可用布局列表."""
    try:
        layouts = []

        # 获取母版的布局
        master = presentation.SlideMaster
        for i in range(1, master.CustomLayouts.Count + 1):
            layout = master.CustomLayouts(i)
            layouts.append({
                "index": i,
                "name": layout.Name,
                "type": layout.Type if hasattr(layout, "Type") else None,
            })

        # 同时返回常用布局映射
        return {
            "custom_layouts_count": master.CustomLayouts.Count,
            "layouts": layouts,
            "common_layouts": {
                "title": 1,
                "title_content": 2,
                "blank": 12,
                "section_header": 33,
                "two_content": 29,
                "comparison": 34,
            },
        }
    except Exception as e:
        raise COMOperationError("get_slide_layouts", str(e))


def _apply_layout(presentation: Any, op: dict) -> str:
    """应用布局到幻灯片.

    Args:
        slide_index: 幻灯片索引
        layout: 布局名称或索引

    """
    slide_index = op.get("slide_index", 1)
    layout = op.get("layout", "title_content")

    try:
        slide = presentation.Slides(slide_index)

        # 如果是名称，转换为索引
        if isinstance(layout, str):
            layout_val = LAYOUT_MAP.get(layout, 2)
        else:
            layout_val = layout

        slide.Layout = layout_val

        return f"apply_layout: slide={slide_index}, layout={layout}"
    except Exception as e:
        raise COMOperationError("apply_layout", str(e))


def _get_slide_size(presentation: Any, op: dict) -> dict:
    """获取幻灯片尺寸."""
    try:
        return {
            "width": presentation.PageSetup.SlideWidth,
            "height": presentation.PageSetup.SlideHeight,
            "width_inches": presentation.PageSetup.SlideWidth / 72,  # 1 inch = 72 points
            "height_inches": presentation.PageSetup.SlideHeight / 72,
            "slide_size_type": presentation.PageSetup.SlideSize,
            "size_types": {
                0: "custom",
                1: "on_screen_4x3",
                2: "letter_paper",
                3: "a4_paper",
                4: "35mm",
                5: "overhead",
                6: "banner",
                7: "on_screen_16x9",
                8: "on_screen_16x10",
            },
        }
    except Exception as e:
        raise COMOperationError("get_slide_size", str(e))


# ============ Shapes 类操作 (10 个新工具) ============

def _list_shapes(presentation: Any, op: dict) -> dict:
    """列出幻灯片上的所有形状.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    try:
        slide = presentation.Slides(slide_index)
        shapes_info = []

        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            shape_info = {
                "index": i,
                "name": shape.Name,
                "type": shape.Type,
                "type_name": _get_shape_type_name(shape.Type),
                "left": shape.Left,
                "top": shape.Top,
                "width": shape.Width,
                "height": shape.Height,
                "rotation": shape.Rotation,
                "visible": shape.Visible,
                "has_text_frame": shape.HasTextFrame,
            }

            # 如果有文本，获取文本内容
            if shape.HasTextFrame and shape.TextFrame.HasText:
                try:
                    shape_info["text"] = shape.TextFrame.TextRange.Text[:100]
                except:
                    shape_info["text"] = ""

            shapes_info.append(shape_info)

        return {
            "slide_index": slide_index,
            "total_shapes": slide.Shapes.Count,
            "shapes": shapes_info,
        }
    except Exception as e:
        raise COMOperationError("list_shapes", str(e))


def _get_shape_type_name(shape_type: int) -> str:
    """获取形状类型名称."""
    SHAPE_TYPE_NAMES = {
        1: "auto_shape",
        7: "embedded_object",
        9: "line",
        10: "picture",
        12: "placeholder",
        13: "text_box",
        14: "mso_placeholder",
        17: "chart",
        19: "table",
        24: "media",
        26: "smart_art",
        28: "web_video",
    }
    return SHAPE_TYPE_NAMES.get(shape_type, f"type_{shape_type}")


def _get_shape_info(presentation: Any, op: dict) -> dict:
    """获取形状详细信息.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)

        info = {
            "index": shape_index,
            "name": shape.Name,
            "type": shape.Type,
            "type_name": _get_shape_type_name(shape.Type),
            "left": shape.Left,
            "top": shape.Top,
            "width": shape.Width,
            "height": shape.Height,
            "rotation": shape.Rotation,
            "visible": shape.Visible,
            "z_order_position": shape.ZOrderPosition,
            "has_text_frame": shape.HasTextFrame,
        }

        # 填充信息
        try:
            fill = shape.Fill
            info["fill"] = {
                "visible": fill.Visible,
                "type": fill.Type,
            }
        except:
            info["fill"] = {"visible": False}

        # 线条信息
        try:
            line = shape.Line
            info["line"] = {
                "visible": line.Visible,
                "weight": line.Weight,
            }
        except:
            info["line"] = {"visible": False}

        # 文本信息
        if shape.HasTextFrame:
            try:
                text_frame = shape.TextFrame
                info["text_frame"] = {
                    "has_text": text_frame.HasText,
                    "text": text_frame.TextRange.Text if text_frame.HasText else "",
                    "font_name": text_frame.TextRange.Font.Name,
                    "font_size": text_frame.TextRange.Font.Size,
                }
            except:
                info["text_frame"] = {"has_text": False}

        return info
    except Exception as e:
        raise COMOperationError("get_shape_info", str(e))


def _update_shape(presentation: Any, op: dict) -> str:
    """更新形状属性.

    Args:
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
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    left = op.get("left")
    top = op.get("top")
    width = op.get("width")
    height = op.get("height")
    rotation = op.get("rotation")
    name = op.get("name")
    visible = op.get("visible")

    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)

        if left is not None:
            shape.Left = left
        if top is not None:
            shape.Top = top
        if width is not None:
            shape.Width = width
        if height is not None:
            shape.Height = height
        if rotation is not None:
            shape.Rotation = rotation
        if name is not None:
            shape.Name = name
        if visible is not None:
            shape.Visible = visible

        return f"update_shape: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("update_shape", str(e))


def _delete_shape(presentation: Any, op: dict) -> str:
    """删除形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    try:
        slide = presentation.Slides(slide_index)
        slide.Shapes(shape_index).Delete()

        return f"delete_shape: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("delete_shape", str(e))


def _set_zorder(presentation: Any, op: dict) -> str:
    """设置形状层级.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        zorder_action: 层级操作 (bring_to_front/send_to_back/bring_forward/send_backward)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    zorder_action = op.get("zorder_action", "bring_to_front")

    # ZOrder 操作常量
    ZORDER_ACTIONS = {
        "bring_to_front": 0,      # msoBringToFront
        "send_to_back": 1,        # msoSendToBack
        "bring_forward": 2,       # msoBringForward
        "send_backward": 3,       # msoSendBackward
    }

    action_val = ZORDER_ACTIONS.get(zorder_action, 0)

    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)
        shape.ZOrder(action_val)

        return f"set_zorder: slide={slide_index}, shape={shape_index}, action={zorder_action}"
    except Exception as e:
        raise COMOperationError("set_zorder", str(e))


def _add_line(presentation: Any, op: dict) -> str:
    """添加线条.

    Args:
        slide_index: 幻灯片索引
        begin_x: 起点 X 坐标
        begin_y: 起点 Y 坐标
        end_x: 终点 X 坐标
        end_y: 终点 Y 坐标
        line_color: 线条颜色 (#RRGGBB)
        line_width: 线条宽度 (pt)

    """
    slide_index = op.get("slide_index", 1)
    begin_x = op.get("begin_x", 100)
    begin_y = op.get("begin_y", 100)
    end_x = op.get("end_x", 200)
    end_y = op.get("end_y", 100)
    line_color = op.get("line_color", "#000000")
    line_width = op.get("line_width", 1)

    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes.AddLine(begin_x, begin_y, end_x, end_y)

        if line_color:
            shape.Line.ForeColor.RGB = _hex_to_rgb(line_color)
        if line_width:
            shape.Line.Weight = line_width

        return f"add_line: slide={slide_index}"
    except Exception as e:
        raise COMOperationError("add_line", str(e))


def _add_textbox(presentation: Any, op: dict) -> str:
    """添加文本框.

    Args:
        slide_index: 幻灯片索引
        text: 文本内容
        left: 左边距
        top: 上边距
        width: 宽度
        height: 高度
        font_name: 字体名称 (可选)
        font_size: 字体大小 (可选)
        font_color: 字体颜色 (可选)
        bold: 是否加粗 (可选)
        alignment: 对齐方式 (left/center/right)

    """
    slide_index = op.get("slide_index", 1)
    text = op.get("text", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 400)
    height = op.get("height", 100)
    font_name = op.get("font_name")
    font_size = op.get("font_size")
    font_color = op.get("font_color")
    bold = op.get("bold")
    alignment = op.get("alignment")

    # 对齐方式常量
    ALIGNMENT_MAP = {
        "left": 1,     # ppAlignLeft
        "center": 2,   # ppAlignCenter
        "right": 3,    # ppAlignRight
    }

    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes.AddTextbox(1, left, top, width, height)  # Orientation=1 (horizontal)

        shape.TextFrame.TextRange.Text = text

        if font_name:
            shape.TextFrame.TextRange.Font.Name = font_name
        if font_size:
            shape.TextFrame.TextRange.Font.Size = font_size
        if font_color:
            shape.TextFrame.TextRange.Font.Color.RGB = _hex_to_rgb(font_color)
        if bold is not None:
            shape.TextFrame.TextRange.Font.Bold = bold
        if alignment:
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = ALIGNMENT_MAP.get(alignment, 1)

        return f"add_textbox: slide={slide_index}, text_length={len(text)}"
    except Exception as e:
        raise COMOperationError("add_textbox", str(e))


def _add_picture_extended(presentation: Any, op: dict) -> str:
    """添加图片（扩展功能）.

    Args:
        slide_index: 幻灯片索引
        image_path: 图片路径
        left: 左边距
        top: 上边距
        width: 宽度 (可选，保持比例)
        height: 高度 (可选，保持比例)
        link_to_file: 是否链接到文件
        save_with_document: 是否随文档保存

    """
    slide_index = op.get("slide_index", 1)
    image_path = op.get("image_path", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width")
    height = op.get("height")
    link_to_file = op.get("link_to_file", False)
    save_with_document = op.get("save_with_document", True)

    if not image_path:
        raise COMOperationError("add_picture", "image_path 不能为空")

    # 安全: 校验图片路径
    image_path = str(validate_path(image_path))

    try:
        slide = presentation.Slides(slide_index)

        # 添加图片
        shape = slide.Shapes.AddPicture(
            FileName=image_path,
            LinkToFile=link_to_file,
            SaveWithDocument=save_with_document,
            Left=left,
            Top=top,
        )

        # 设置尺寸
        if width:
            shape.Width = width
        if height:
            shape.Height = height

        return f"add_picture: slide={slide_index}, image={image_path}, linked={link_to_file}"
    except Exception as e:
        raise COMOperationError("add_picture", str(e))


def _duplicate_shape(presentation: Any, op: dict) -> str:
    """复制形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        offset_left: 复制后 X 偏移 (可选)
        offset_top: 复制后 Y 偏移 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    offset_left = op.get("offset_left", 20)
    offset_top = op.get("offset_top", 20)

    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)

        # 复制形状
        shape.Copy()
        new_shape = slide.Shapes.Paste()

        # 调整位置
        new_shape.Left = shape.Left + offset_left
        new_shape.Top = shape.Top + offset_top

        return f"duplicate_shape: slide={slide_index}, shape={shape_index} duplicated"
    except Exception as e:
        raise COMOperationError("duplicate_shape", str(e))


def _group_shapes(presentation: Any, op: dict) -> str:
    """组合形状.

    Args:
        slide_index: 幻灯片索引
        shape_indices: 要组合的形状索引列表

    """
    slide_index = op.get("slide_index", 1)
    shape_indices = op.get("shape_indices", [])

    if not shape_indices or len(shape_indices) < 2:
        raise COMOperationError("group_shapes", "需要至少提供 2 个形状索引")

    try:
        slide = presentation.Slides(slide_index)

        # 获取要组合的形状
        shapes_to_group = []
        for idx in shape_indices:
            shapes_to_group.append(slide.Shapes(idx))

        # 组合形状
        # 注意：COM API 需要通过 ShapeRange 来组合
        shape_range = slide.Shapes.Range(shape_indices)
        try:
            grouped_shape = shape_range.Group()
        except Exception as group_err:
            err_text = str(group_err)
            if "禁止分组" in err_text or "group" in err_text.lower() or "不能" in err_text:
                return f"group_shapes_skipped: 所选形状不支持分组（可能包含占位符等不可分组形状），slide={slide_index}"
            raise

        return f"group_shapes: slide={slide_index}, grouped {len(shape_indices)} shapes"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("group_shapes", str(e))


# ============ Text 类操作 (10 个新工具) ============

def _ppt_set_text(presentation: Any, op: dict) -> str:
    """设置形状文本.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        text: 文本内容

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    text = op.get("text", "")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        raise COMOperationError("ppt_set_text", "形状没有文本框")

    shape.TextFrame.TextRange.Text = text
    return f"set_text: slide={slide_index}, shape={shape_index}"


def _ppt_get_text(presentation: Any, op: dict) -> dict:
    """获取形状文本.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        return {"slide_index": slide_index, "shape_index": shape_index, "text": "", "has_text": False}

    text = shape.TextFrame.TextRange.Text
    return {"slide_index": slide_index, "shape_index": shape_index, "text": text, "has_text": True}


def _ppt_format_text_range(presentation: Any, op: dict) -> str:
    """格式化文本范围.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        start: 开始位置 (1-based)
        length: 文本长度
        font_name: 字体名称
        font_size: 字体大小
        font_color: 字体颜色 (#RRGGBB)
        bold: 是否粗体
        italic: 是否斜体
        underline: 是否下划线

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    start = op.get("start", 1)
    length = op.get("length", 0)
    font_name = op.get("font_name", "")
    font_size = op.get("font_size", 0)
    font_color = op.get("font_color", "")
    bold = op.get("bold")
    italic = op.get("italic")
    underline = op.get("underline")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        raise COMOperationError("ppt_format_text_range", "形状没有文本框")

    text_range = shape.TextFrame.TextRange.Characters(start, length)

    if font_name:
        text_range.Font.Name = font_name
    if font_size > 0:
        text_range.Font.Size = font_size
    if font_color:
        text_range.Font.Color.RGB = _hex_to_rgb(font_color)
    if bold is not None:
        text_range.Font.Bold = bold
    if italic is not None:
        text_range.Font.Italic = italic
    if underline is not None:
        text_range.Font.Underline = underline

    return f"format_text_range: slide={slide_index}, shape={shape_index}, start={start}, length={length}"


def _ppt_set_paragraph_format(presentation: Any, op: dict) -> str:
    """设置段落格式.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        paragraph_index: 段落索引 (1-based)
        alignment: 对齐方式 (left/center/right/distribute)
        line_spacing: 行间距 (pt)
        space_before: 段前间距 (pt)
        space_after: 段后间距 (pt)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    paragraph_index = op.get("paragraph_index", 1)
    alignment = op.get("alignment", "left")
    line_spacing = op.get("line_spacing", 0)
    space_before = op.get("space_before", 0)
    space_after = op.get("space_after", 0)

    # 对齐方式映射
    ALIGNMENT_MAP = {
        "left": 1,        # ppAlignLeft
        "center": 2,      # ppAlignCenter
        "right": 3,       # ppAlignRight
        "distribute": 4,  # ppAlignDistribute
    }

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        raise COMOperationError("ppt_set_paragraph_format", "形状没有文本框")

    text_range = shape.TextFrame.TextRange
    # Use TextRange directly for single-paragraph shapes,
    # or iterate Paragraphs collection for multi-paragraph
    try:
        paragraphs = text_range.Paragraphs()
        if paragraph_index > paragraphs.Count:
            raise COMOperationError("ppt_set_paragraph_format", f"段落索引超出范围: {paragraph_index}")
        para = paragraphs(paragraph_index)
    except Exception:
        # Fallback: operate on the entire TextRange if Paragraphs() is not supported
        para = text_range

    if alignment:
        para.ParagraphFormat.Alignment = ALIGNMENT_MAP.get(alignment, 1)
    if line_spacing > 0:
        para.ParagraphFormat.LineSpacing = line_spacing
    if space_before > 0:
        para.ParagraphFormat.SpaceBefore = space_before
    if space_after > 0:
        para.ParagraphFormat.SpaceAfter = space_after

    return f"set_paragraph_format: slide={slide_index}, shape={shape_index}, paragraph={paragraph_index}"


def _ppt_set_bullets(presentation: Any, op: dict) -> str:
    """设置项目符号.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        paragraph_index: 段落索引 (可选，默认全部)
        bullet_type: 项目符号类型 (none/numbered/bullet/picture)
        bullet_char: 自定义符号字符 (可选)
        bullet_font: 符号字体 (可选)
        bullet_size: 符号大小 (可选)
        indent_level: 缩进级别 (0-9)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    paragraph_index = op.get("paragraph_index")
    bullet_type = op.get("bullet_type", "bullet")
    bullet_char = op.get("bullet_char", "")
    bullet_font = op.get("bullet_font", "")
    bullet_size = op.get("bullet_size", 0)
    indent_level = op.get("indent_level", 0)

    # 项目符号类型映射
    BULLET_TYPE_MAP = {
        "none": 0,       # ppBulletNone
        "numbered": 2,   # ppBulletNumbered
        "bullet": 1,     # ppBulletUnnumbered
        "picture": 3,    # ppBulletPicture
    }

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        raise COMOperationError("ppt_set_bullets", "形状没有文本框")

    text_range = shape.TextFrame.TextRange

    if paragraph_index:
        paragraphs = text_range.Paragraphs()
        if paragraph_index > paragraphs.Count:
            raise COMOperationError("ppt_set_bullets", f"段落索引超出范围: {paragraph_index}")
        para = paragraphs(paragraph_index)
    else:
        para = text_range

    para.ParagraphFormat.Bullet.Type = BULLET_TYPE_MAP.get(bullet_type, 1)

    if bullet_char and bullet_type == "bullet":
        para.ParagraphFormat.Bullet.Character = ord(bullet_char[0])
    if bullet_font:
        para.ParagraphFormat.Bullet.Font.Name = bullet_font
    if bullet_size > 0:
        para.ParagraphFormat.Bullet.Size = bullet_size
    if indent_level > 0:
        para.ParagraphFormat.IndentLevel = indent_level

    return f"set_bullets: slide={slide_index}, shape={shape_index}, type={bullet_type}"


def _ppt_find_replace_text(presentation: Any, op: dict) -> dict:
    """查找替换文本.

    Args:
        slide_index: 幻灯片索引 (可选，默认全部)
        find_text: 查找文本
        replace_text: 替换文本
        match_case: 是否区分大小写
        match_whole_word: 是否全字匹配

    """
    slide_index = op.get("slide_index")
    find_text = op.get("find_text", "")
    replace_text = op.get("replace_text", "")
    match_case = op.get("match_case", False)
    match_whole_word = op.get("match_whole_word", False)

    if not find_text:
        raise COMOperationError("ppt_find_replace_text", "find_text 不能为空")

    replaced_count = 0
    found_locations = []

    slides_to_check = [presentation.Slides(slide_index)] if slide_index else list(presentation.Slides)

    for slide in slides_to_check:
        slide_num = slide.SlideIndex
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text_range = shape.TextFrame.TextRange
                original_text = text_range.Text

                # 简单查找替换
                if find_text in original_text:
                    count = original_text.count(find_text)
                    replaced_count += count
                    found_locations.append({
                        "slide": slide_num,
                        "shape": shape.Name,
                        "count": count
                    })

                    # 执行替换
                    new_text = original_text.replace(find_text, replace_text)
                    text_range.Text = new_text

    return {
        "replaced_count": replaced_count,
        "found_locations": found_locations,
        "find_text": find_text,
        "replace_text": replace_text
    }


def _ppt_get_textframe(presentation: Any, op: dict) -> dict:
    """获取文本框属性.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    result = {
        "slide_index": slide_index,
        "shape_index": shape_index,
        "has_textframe": shape.HasTextFrame,
    }

    if shape.HasTextFrame:
        tf = shape.TextFrame
        result["has_text"] = tf.HasText
        result["margin_left"] = tf.MarginLeft
        result["margin_right"] = tf.MarginRight
        result["margin_top"] = tf.MarginTop
        result["margin_bottom"] = tf.MarginBottom
        result["word_wrap"] = tf.WordWrap
        result["autosize"] = tf.AutoSize

        if tf.HasText:
            result["text"] = tf.TextRange.Text
            result["text_length"] = len(tf.TextRange.Text)
            result["paragraph_count"] = tf.TextRange.Paragraphs().Count

    return result


def _ppt_extract_text_as_markdown(presentation: Any, op: dict) -> dict:
    """提取所有文本为 Markdown.

    Args:
        slide_index: 幻灯片索引 (可选，默认全部)

    """
    slide_index = op.get("slide_index")

    markdown_lines = []
    slides_to_check = [presentation.Slides(slide_index)] if slide_index else list(presentation.Slides)

    for slide in slides_to_check:
        slide_num = slide.SlideIndex
        markdown_lines.append(f"\n## Slide {slide_num}\n")

        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text = shape.TextFrame.TextRange.Text.strip()
                if text:
                    # 检查是否为标题占位符
                    is_title = False
                    try:
                        if shape.Type == 14:  # msoPlaceholder
                            if shape.PlaceholderFormat.Type in (1, 3):  # ppPlaceholderTitle
                                is_title = True
                    except:
                        pass

                    if is_title:
                        markdown_lines.append(f"### {text}\n")
                    else:
                        # 处理段落
                        paragraphs = shape.TextFrame.TextRange.Paragraphs()
                        for para in paragraphs:
                            try:
                                para_text = (para.Text or "").strip()
                            except:
                                continue
                            if para_text:
                                # 检查是否有项目符号
                                try:
                                    bullet_type = para.ParagraphFormat.Bullet.Type
                                    if bullet_type == 1:  # bullet
                                        markdown_lines.append(f"- {para_text}")
                                    elif bullet_type == 2:  # numbered
                                        markdown_lines.append(f"1. {para_text}")
                                    else:
                                        markdown_lines.append(para_text)
                                except:
                                    markdown_lines.append(para_text)

    return {
        "slide_index": slide_index or "all",
        "markdown": "\n".join(markdown_lines)
    }


def _ppt_set_font_size(presentation: Any, op: dict) -> str:
    """设置字体大小.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        font_size: 字体大小 (pt)
        start: 开始位置 (可选)
        length: 文本长度 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    font_size = op.get("font_size", 12)
    start = op.get("start")
    length = op.get("length")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        raise COMOperationError("ppt_set_font_size", "形状没有文本框")

    if start and length:
        text_range = shape.TextFrame.TextRange.Characters(start, length)
    else:
        text_range = shape.TextFrame.TextRange

    text_range.Font.Size = font_size
    return f"set_font_size: slide={slide_index}, shape={shape_index}, size={font_size}pt"


def _ppt_set_font_color(presentation: Any, op: dict) -> str:
    """设置字体颜色.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        font_color: 字体颜色 (#RRGGBB)
        start: 开始位置 (可选)
        length: 文本长度 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    font_color = op.get("font_color", "#000000")
    start = op.get("start")
    length = op.get("length")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTextFrame:
        raise COMOperationError("ppt_set_font_color", "形状没有文本框")

    if start and length:
        text_range = shape.TextFrame.TextRange.Characters(start, length)
    else:
        text_range = shape.TextFrame.TextRange

    text_range.Font.Color.RGB = _hex_to_rgb(font_color)
    return f"set_font_color: slide={slide_index}, shape={shape_index}, color={font_color}"


# ============ Placeholders 类操作 (6 个新工具) ============

# 占位符类型映射
PLACEHOLDER_TYPES = {
    1: "title",           # ppPlaceholderTitle
    2: "body",            # ppPlaceholderBody
    3: "center_title",    # ppPlaceholderCenterTitle
    4: "center_body",     # ppPlaceholderCenterBody
    5: "subtitle",        # ppPlaceholderSubtitle
    6: "slide_number",    # ppSlideNumber
    7: "header",          # ppHeader
    8: "footer",          # ppFooter
    9: "date",            # ppDate
    10: "object",         # ppPlaceholderObject
    11: "chart",          # ppPlaceholderChart
    12: "table",          # ppPlaceholderTable
    13: "clip_art",       # ppPlaceholderClipArt
    14: "org_chart",      # ppPlaceholderOrgChart
    15: "media",          # ppPlaceholderMedia
    16: "title_only",     # ppPlaceholderTitleOnly
    17: "vertical_text",  # ppPlaceholderVerticalText
    18: "vertical_object", # ppPlaceholderVerticalObject
    19: "picture",        # ppPlaceholderPicture
}


def _ppt_list_placeholders(presentation: Any, op: dict) -> dict:
    """列出幻灯片占位符.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    slide = presentation.Slides(slide_index)
    placeholders = []

    for shape in slide.Shapes.Placeholders:
        try:
            ph_type = shape.PlaceholderFormat.Type
            ph_name = PLACEHOLDER_TYPES.get(ph_type, f"unknown_{ph_type}")
            placeholders.append({
                "index": shape.PlaceholderFormat.Index,
                "type": ph_name,
                "type_id": ph_type,
                "name": shape.Name,
                "has_text": shape.HasTextFrame and shape.TextFrame.HasText,
            })
        except:
            pass

    return {
        "slide_index": slide_index,
        "placeholder_count": len(placeholders),
        "placeholders": placeholders
    }


def _ppt_get_placeholder(presentation: Any, op: dict) -> dict:
    """获取占位符内容.

    Args:
        slide_index: 幻灯片索引
        placeholder_index: 占位符索引 (1-based)

    """
    slide_index = op.get("slide_index", 1)
    placeholder_index = op.get("placeholder_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.Placeholders(placeholder_index)

    result = {
        "slide_index": slide_index,
        "placeholder_index": placeholder_index,
        "type": PLACEHOLDER_TYPES.get(shape.PlaceholderFormat.Type, "unknown"),
        "name": shape.Name,
    }

    if shape.HasTextFrame:
        result["has_text"] = shape.TextFrame.HasText
        if shape.TextFrame.HasText:
            result["text"] = shape.TextFrame.TextRange.Text

    return result


def _ppt_set_placeholder(presentation: Any, op: dict) -> str:
    """设置占位符内容.

    Args:
        slide_index: 幻灯片索引
        placeholder_index: 占位符索引
        text: 文本内容

    """
    slide_index = op.get("slide_index", 1)
    placeholder_index = op.get("placeholder_index", 1)
    text = op.get("text", "")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.Placeholders(placeholder_index)

    if shape.HasTextFrame:
        shape.TextFrame.TextRange.Text = text
        return f"set_placeholder: slide={slide_index}, placeholder={placeholder_index}"
    else:
        raise COMOperationError("ppt_set_placeholder", "占位符不支持文本")


def _ppt_clear_placeholder(presentation: Any, op: dict) -> str:
    """清空占位符.

    Args:
        slide_index: 幻灯片索引
        placeholder_index: 占位符索引

    """
    slide_index = op.get("slide_index", 1)
    placeholder_index = op.get("placeholder_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.Placeholders(placeholder_index)

    if shape.HasTextFrame:
        shape.TextFrame.TextRange.Text = ""
        return f"clear_placeholder: slide={slide_index}, placeholder={placeholder_index}"
    else:
        return f"clear_placeholder: slide={slide_index}, placeholder={placeholder_index} (无文本)"


def _ppt_get_placeholder_type(presentation: Any, op: dict) -> dict:
    """获取占位符类型.

    Args:
        slide_index: 幻灯片索引
        placeholder_index: 占位符索引

    """
    slide_index = op.get("slide_index", 1)
    placeholder_index = op.get("placeholder_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.Placeholders(placeholder_index)

    type_id = shape.PlaceholderFormat.Type
    type_name = PLACEHOLDER_TYPES.get(type_id, f"unknown_{type_id}")

    return {
        "slide_index": slide_index,
        "placeholder_index": placeholder_index,
        "type_id": type_id,
        "type_name": type_name
    }


def _ppt_resize_placeholder(presentation: Any, op: dict) -> str:
    """调整占位符大小.

    Args:
        slide_index: 幻灯片索引
        placeholder_index: 占位符索引
        width: 宽度 (pt)
        height: 高度 (pt)
        left: 左边距 (可选)
        top: 上边距 (可选)

    """
    slide_index = op.get("slide_index", 1)
    placeholder_index = op.get("placeholder_index", 1)
    width = op.get("width")
    height = op.get("height")
    left = op.get("left")
    top = op.get("top")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes.Placeholders(placeholder_index)

    if width:
        shape.Width = width
    if height:
        shape.Height = height
    if left:
        shape.Left = left
    if top:
        shape.Top = top

    return f"resize_placeholder: slide={slide_index}, placeholder={placeholder_index}"


# ============ Formatting 类操作 (3 个新工具) ============

def _ppt_set_fill(presentation: Any, op: dict) -> str:
    """设置填充（颜色/渐变/图片）.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        fill_type: 填充类型 (solid/gradient/picture/pattern/no_fill)
        color: 颜色 (#RRGGBB, solid 类型)
        gradient_colors: 渐变颜色列表 (gradient 类型)
        gradient_type: 渐变类型 (linear/radial)
        picture_path: 图片路径 (picture 类型)
        transparency: 透明度 (0-100)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    fill_type = op.get("fill_type", "solid")
    color = op.get("color", "#FFFFFF")
    gradient_colors = op.get("gradient_colors", [])
    gradient_type = op.get("gradient_type", "linear")
    picture_path = op.get("picture_path", "")
    transparency = op.get("transparency", 0)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    fill = shape.Fill

    if fill_type == "no_fill":
        fill.Visible = False
    elif fill_type == "solid":
        fill.Visible = True
        fill.Solid()
        fill.ForeColor.RGB = _hex_to_rgb(color)
        if transparency > 0:
            fill.Transparency = transparency / 100.0
    elif fill_type == "gradient":
        fill.Visible = True
        type_val = GRADIENT_TYPES.get(gradient_type, 1)
        if len(gradient_colors) >= 2:
            fill.TwoColorGradient(Style=type_val, Variant=1)
            fill.ForeColor.RGB = _hex_to_rgb(gradient_colors[0])
            fill.BackColor.RGB = _hex_to_rgb(gradient_colors[1])
        else:
            fill.OneColorGradient(Style=type_val, Variant=1)
            fill.ForeColor.RGB = _hex_to_rgb(color)
    elif fill_type == "picture":
        if picture_path:
            # 安全: 校验图片路径
            picture_path = str(validate_path(picture_path))
            fill.Visible = True
            fill.UserPicture(picture_path)
    elif fill_type == "pattern":
        fill.Visible = True
        fill.Patterned(1)  # msoPatternMixed
        fill.ForeColor.RGB = _hex_to_rgb(color)

    return f"set_fill: slide={slide_index}, shape={shape_index}, type={fill_type}"


def _ppt_set_line(presentation: Any, op: dict) -> str:
    """设置线条样式.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        line_color: 线条颜色 (#RRGGBB)
        line_width: 线条宽度 (pt)
        line_style: 线条样式 (solid/dash/dash_dot/dot)
        dash_style: 虚线样式 (solid/dash/dash_dot/dot/dash_dot_dot)
        transparency: 透明度 (0-100)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    line_color = op.get("line_color", "#000000")
    line_width = op.get("line_width", 1)
    line_style = op.get("line_style", "solid")
    dash_style = op.get("dash_style", "solid")
    transparency = op.get("transparency", 0)

    # 线条样式映射
    LINE_STYLE_MAP = {
        "solid": 1,        # msoLineSolid
        "dash": 2,         # msoLineDash
        "dash_dot": 3,     # msoLineDashDot
        "dot": 4,          # msoLineDot
        "dash_dot_dot": 5, # msoLineDashDotDot
    }

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    line = shape.Line
    line.Visible = True
    line.ForeColor.RGB = _hex_to_rgb(line_color)
    line.Weight = line_width

    if dash_style:
        line.DashStyle = LINE_STYLE_MAP.get(dash_style, 1)

    if transparency > 0:
        line.Transparency = transparency / 100.0

    return f"set_line: slide={slide_index}, shape={shape_index}"


def _ppt_set_shadow(presentation: Any, op: dict) -> str:
    """设置阴影效果.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        shadow_type: 鐢影类型 (none/offset/double/perspective)
        shadow_color: 鐢影颜色 (#RRGGBB)
        shadow_blur: 模糊半径 (pt)
        shadow_offset_x: X 偏移 (pt)
        shadow_offset_y: Y 偏移 (pt)
        shadow_transparency: 透明度 (0-100)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    shadow_type = op.get("shadow_type", "offset")
    shadow_color = op.get("shadow_color", "#808080")
    shadow_blur = op.get("shadow_blur", 4)
    shadow_offset_x = op.get("shadow_offset_x", 3)
    shadow_offset_y = op.get("shadow_offset_y", 3)
    shadow_transparency = op.get("shadow_transparency", 50)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    shadow = shape.Shadow

    if shadow_type == "none":
        shadow.Visible = False
    else:
        shadow.Visible = True
        shadow.ForeColor.RGB = _hex_to_rgb(shadow_color)
        shadow.Blur = shadow_blur
        shadow.OffsetX = shadow_offset_x
        shadow.OffsetY = shadow_offset_y
        shadow.Transparency = shadow_transparency / 100.0

    return f"set_shadow: slide={slide_index}, shape={shape_index}, type={shadow_type}"


# ============ Tables 类操作 (13 个新工具) ============

def _ppt_get_table_cells(presentation: Any, op: dict) -> dict:
    """获取表格单元格内容.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (表格)
        start_row: 开始行 (可选)
        start_col: 开始列 (可选)
        end_row: 结束行 (可选)
        end_col: 结束列 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    start_row = op.get("start_row", 1)
    start_col = op.get("start_col", 1)
    end_row = op.get("end_row")
    end_col = op.get("end_col")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_get_table_cells", "形状不是表格")

    table = shape.Table
    rows = table.Rows.Count
    cols = table.Columns.Count

    if end_row is None:
        end_row = rows
    if end_col is None:
        end_col = cols

    cells = []
    for r in range(start_row, min(end_row + 1, rows + 1)):
        row_data = []
        for c in range(start_col, min(end_col + 1, cols + 1)):
            try:
                cell_text = table.Cell(r, c).Shape.TextFrame.TextRange.Text
            except:
                cell_text = ""
            row_data.append(cell_text)
        cells.append(row_data)

    return {
        "slide_index": slide_index,
        "shape_index": shape_index,
        "rows": rows,
        "columns": cols,
        "cells": cells
    }


def _ppt_set_table_cells(presentation: Any, op: dict) -> str:
    """设置表格单元格内容.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        row: 行号
        col: 列号
        text: 文本内容

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    row = op.get("row", 1)
    col = op.get("col", 1)
    text = op.get("text", "")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_set_table_cells", "形状不是表格")

    table = shape.Table
    table.Cell(row, col).Shape.TextFrame.TextRange.Text = text

    return f"set_table_cells: slide={slide_index}, shape={shape_index}, cell({row},{col})"


def _ppt_batch_set_table_data(presentation: Any, op: dict) -> str:
    """批量设置表格数据.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        data: 二维数组数据
        start_row: 开始行 (可选)
        start_col: 开始列 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    data = op.get("data", [])
    start_row = op.get("start_row", 1)
    start_col = op.get("start_col", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_batch_set_table_data", "形状不是表格")

    table = shape.Table

    for i, row_data in enumerate(data):
        for j, cell_data in enumerate(row_data):
            r = start_row + i
            c = start_col + j
            if r <= table.Rows.Count and c <= table.Columns.Count:
                table.Cell(r, c).Shape.TextFrame.TextRange.Text = str(cell_data)

    return f"batch_set_table_data: slide={slide_index}, shape={shape_index}, {len(data)} rows"


def _ppt_merge_table_cells(presentation: Any, op: dict) -> str:
    """合并单元格.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        start_row: 开始行
        start_col: 开始列
        end_row: 结束行
        end_col: 结束列

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    start_row = op.get("start_row", 1)
    start_col = op.get("start_col", 1)
    end_row = op.get("end_row", 1)
    end_col = op.get("end_col", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    table = _ppt_require_table(shape, "ppt_merge_table_cells")
    start_row, start_col, end_row, end_col = _ppt_normalize_merge_region(
        table, start_row, start_col, end_row, end_col
    )
    texts = _ppt_collect_region_texts(table, start_row, start_col, end_row, end_col)
    merged_text = "\n".join(text.strip() for row_texts in texts for text in row_texts if text.strip())

    _ppt_store_merge_metadata(
        shape,
        start_row,
        start_col,
        {
            "end_row": end_row,
            "end_col": end_col,
            "texts": texts,
        },
    )

    for r in range(start_row, end_row + 1):
        for c in range(start_col, end_col + 1):
            _ppt_set_cell_text(table.Cell(r, c), "")

    try:
        cell = table.Cell(start_row, start_col)
        cell.Merge(table.Cell(end_row, end_col))
    except Exception as merge_error:
        # Merge failed - restore original texts
        logger.warning("PPT merge failed, restoring original texts: %s", merge_error)
        for row_offset, r in enumerate(range(start_row, end_row + 1)):
            for col_offset, c in enumerate(range(start_col, end_col + 1)):
                if row_offset < len(texts) and col_offset < len(texts[row_offset]):
                    try:
                        _ppt_set_cell_text(table.Cell(r, c), texts[row_offset][col_offset])
                    except Exception:
                        logger.debug("Could not restore text to cell (%d,%d) after merge failure", r, c)
        _ppt_delete_merge_metadata(shape, start_row, start_col)
        raise COMOperationError("ppt_merge_table_cells", f"Merge failed: {merge_error}") from merge_error

    _ppt_set_cell_text(table.Cell(start_row, start_col), merged_text)

    return f"merge_table_cells: slide={slide_index}, shape={shape_index}, ({start_row},{start_col})-({end_row},{end_col})"


def _ppt_split_table_cells(presentation: Any, op: dict) -> str:
    """拆分单元格.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        row: 行号
        col: 列号

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    row = op.get("row", 1)
    col = op.get("col", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    table = _ppt_require_table(shape, "ppt_split_table_cells")
    if row < 1 or col < 1 or row > table.Rows.Count or col > table.Columns.Count:
        raise COMOperationError("ppt_split_table_cells", "拆分单元格超出表格边界")

    # Phase 1: Snapshot pre-split state
    metadata_match = _ppt_find_merge_metadata(shape, table, row, col)
    metadata_anchor_row = row
    metadata_anchor_col = col
    metadata = None
    if metadata_match:
        metadata_anchor_row, metadata_anchor_col, metadata = metadata_match
    merged_text = _ppt_get_cell_text(table.Cell(row, col))
    pre_rows = table.Rows.Count
    pre_cols = table.Columns.Count

    # Phase 2: Execute split
    cell = table.Cell(row, col)
    # Check if the cell is actually part of a merged region before splitting.
    # In PowerPoint COM, Split() only works on merged cells.
    try:
        is_merged = False
        try:
            # A merged cell has ColSpan > 1 or RowSpan > 1
            is_merged = cell.ColSpan > 1 or cell.RowSpan > 1
        except Exception:
            # If we cannot determine merge status, attempt split anyway
            is_merged = True

        if not is_merged:
            return f"split_table_cells: slide={slide_index}, shape={shape_index}, ({row},{col}) skipped (cell not merged)"

        cell.Split()
    except COMOperationError:
        raise
    except Exception as split_error:
        # Split failed — cell may not actually be merged despite appearing so
        return f"split_table_cells: slide={slide_index}, shape={shape_index}, ({row},{col}) skipped (cell not merged: {split_error})"

    # Phase 3: Restore texts after split
    post_rows = table.Rows.Count
    post_cols = table.Columns.Count

    if metadata:
        # We have stored metadata about the original pre-merge state
        orig_end_row = int(metadata.get("end_row", metadata_anchor_row))
        orig_end_col = int(metadata.get("end_col", metadata_anchor_col))
        texts = metadata.get("texts") or []

        # After split, the grid may have expanded. Calculate the actual region.
        # The split cell should now span from (row, col) to (orig_end_row, orig_end_col)
        # but as individual cells again.
        actual_end_row = min(orig_end_row, post_rows)
        actual_end_col = min(orig_end_col, post_cols)

        for row_offset, target_row in enumerate(range(metadata_anchor_row, actual_end_row + 1)):
            for col_offset, target_col in enumerate(range(metadata_anchor_col, actual_end_col + 1)):
                original_text = ""
                if row_offset < len(texts) and col_offset < len(texts[row_offset]):
                    original_text = str(texts[row_offset][col_offset] or "")
                try:
                    _ppt_set_cell_text(table.Cell(target_row, target_col), original_text)
                except Exception:
                    # Cell may not exist if grid structure is unexpected
                    logger.debug("Could not restore text to cell (%d,%d) after split", target_row, target_col)

        _ppt_delete_merge_metadata(shape, metadata_anchor_row, metadata_anchor_col)
    else:
        # No metadata - distribute merged text to the first cell, clear others
        # After split, the cell at (row, col) should be the top-left of the split region
        _ppt_set_cell_text(table.Cell(row, col), merged_text)
        # Try to clear any newly created adjacent cells
        # We don't know the exact split region, so just clear cells that might be empty
        for r in range(row, min(row + 2, post_rows + 1)):
            for c in range(col, min(col + 2, post_cols + 1)):
                if r == row and c == col:
                    continue
                try:
                    cell_text = _ppt_get_cell_text(table.Cell(r, c))
                    if not cell_text.strip():
                        _ppt_set_cell_text(table.Cell(r, c), "")
                except Exception:
                    pass

    return f"split_table_cells: slide={slide_index}, shape={shape_index}, cell({row},{col})"


def _ppt_add_table_row(presentation: Any, op: dict) -> str:
    """添加表格行.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        before_row: 在此行之前插入 (可选，默认末尾)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    before_row = op.get("before_row")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_add_table_row", "形状不是表格")

    table = shape.Table

    if before_row:
        table.Rows.Add(before_row)
    else:
        table.Rows.Add()

    return f"add_table_row: slide={slide_index}, shape={shape_index}"


def _ppt_delete_table_row(presentation: Any, op: dict) -> str:
    """删除表格行.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        row: 行号

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    row = op.get("row", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_delete_table_row", "形状不是表格")

    table = shape.Table
    table.Rows(row).Delete()

    return f"delete_table_row: slide={slide_index}, shape={shape_index}, row={row}"


def _ppt_add_table_column(presentation: Any, op: dict) -> str:
    """添加表格列.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        before_col: 在此列之前插入 (可选，默认末尾)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    before_col = op.get("before_col")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_add_table_column", "形状不是表格")

    table = shape.Table

    if before_col:
        table.Columns.Add(before_col)
    else:
        table.Columns.Add()

    return f"add_table_column: slide={slide_index}, shape={shape_index}"


def _ppt_delete_table_column(presentation: Any, op: dict) -> str:
    """删除表格列.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        col: 列号

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    col = op.get("col", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_delete_table_column", "形状不是表格")

    table = shape.Table
    table.Columns(col).Delete()

    return f"delete_table_column: slide={slide_index}, shape={shape_index}, col={col}"


def _ppt_set_table_style(presentation: Any, op: dict) -> str:
    """设置表格样式.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        style_name: 样式名称或索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    style_name = op.get("style_name", "")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_set_table_style", "形状不是表格")

    # 表格样式应用
    try:
        table = shape.Table
        # PowerPoint 表格样式通过 TableStyle 属性设置
        # 这里简化处理，实际可能需要遍历可用样式
        if style_name:
            shape.Table.Apply()
    except:
        pass

    return f"set_table_style: slide={slide_index}, shape={shape_index}, style={style_name}"


def _ppt_set_table_borders(presentation: Any, op: dict) -> str:
    """设置表格边框.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        border_color: 边框颜色 (#RRGGBB)
        border_width: 边框宽度 (pt)
        border_style: 边框样式 (solid/dash/dot)
        apply_to: 应用范围 (all/outer/inner/horizontal/vertical)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    border_color = op.get("border_color", "#000000")
    border_width = op.get("border_width", 1)
    border_style = op.get("border_style", "solid")
    apply_to = op.get("apply_to", "all")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_set_table_borders", "形状不是表格")

    table = shape.Table
    rgb = _hex_to_rgb(border_color)

    # 边框样式映射
    BORDER_STYLE_MAP = {
        "solid": 1,
        "dash": 2,
        "dot": 4,
    }

    style_val = BORDER_STYLE_MAP.get(border_style, 1)

    # 设置边框
    skipped_cells = 0
    for r in range(1, table.Rows.Count + 1):
        for c in range(1, table.Columns.Count + 1):
            try:
                cell = table.Cell(r, c)
                borders = cell.Shape.Line

                should_apply = False
                if apply_to == "all":
                    should_apply = True
                elif apply_to == "outer":
                    if r == 1 or r == table.Rows.Count or c == 1 or c == table.Columns.Count:
                        should_apply = True
                elif apply_to == "inner":
                    if r > 1 and r < table.Rows.Count and c > 1 and c < table.Columns.Count:
                        should_apply = True

                if should_apply:
                    try:
                        borders.Visible = True
                    except Exception:
                        pass
                    try:
                        borders.ForeColor.RGB = rgb
                    except Exception:
                        pass
                    try:
                        borders.Weight = border_width
                    except Exception:
                        pass
            except Exception:
                skipped_cells += 1
                continue

    result = f"set_table_borders: slide={slide_index}, shape={shape_index}, apply_to={apply_to}"
    if skipped_cells:
        result += f", skipped_cells={skipped_cells}"
    return result


def _ppt_get_table_info(presentation: Any, op: dict) -> dict:
    """获取表格信息.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "is_table": False
        }

    table = shape.Table

    return {
        "slide_index": slide_index,
        "shape_index": shape_index,
        "is_table": True,
        "rows": table.Rows.Count,
        "columns": table.Columns.Count,
        "width": shape.Width,
        "height": shape.Height,
        "name": shape.Name
    }


def _ppt_resize_table(presentation: Any, op: dict) -> str:
    """调整表格大小.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        width: 宽度 (pt)
        height: 高度 (pt)
        left: 左边距 (可选)
        top: 上边距 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    width = op.get("width")
    height = op.get("height")
    left = op.get("left")
    top = op.get("top")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not shape.HasTable:
        raise COMOperationError("ppt_resize_table", "形状不是表格")

    if width:
        shape.Width = width
    if height:
        shape.Height = height
    if left:
        shape.Left = left
    if top:
        shape.Top = top

    return f"resize_table: slide={slide_index}, shape={shape_index}"


# ============ Export 类功能 ============

# 导出格式常量
EXPORT_FORMATS = {
    "png": 18,   # ppShapeFormatPNG
    "jpg": 19,   # ppShapeFormatJPG
    "gif": 17,   # ppShapeFormatGIF
    "bmp": 20,   # ppShapeFormatBMP
}


def _export_images(presentation: Any, op: dict) -> dict:
    """导出幻灯片为图片.

    Args:
        output_path: 输出目录路径
        slide_indices: 幻灯片索引列表 (可选，默认全部)
        format: 图片格式 (png/jpg/gif/bmp)
        width: 图片宽度 (可选)
        height: 图片高度 (可选)

    """
    output_path = op.get("output_path", "")
    slide_indices = op.get("slide_indices", [])
    format_name = op.get("format", "png")
    width = op.get("width")
    height = op.get("height")

    if not output_path:
        raise COMOperationError("export_images", "output_path 不能为空")

    # 安全: 校验输出路径
    output_path = str(validate_path(output_path))

    output_dir = Path(output_path)
    if not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=True)

    format_val = EXPORT_FORMATS.get(format_name, 18)

    exported_files = []
    slides_to_export = slide_indices if slide_indices else range(1, presentation.Slides.Count + 1)

    for idx in slides_to_export:
        try:
            slide = presentation.Slides(idx)
            file_name = f"slide_{idx}.{format_name}"
            file_path = output_dir / file_name

            if width and height:
                slide.Export(str(file_path), format_val, width, height)
            else:
                slide.Export(str(file_path), format_val)

            _ppt_require_output(str(file_path), "export_images")
            exported_files.append(str(file_path))
        except Exception as e:
            logger.warning(f"导出幻灯片 {idx} 失败: {e}")

    return {
        "output_path": output_path,
        "format": format_name,
        "exported_count": len(exported_files),
        "files": exported_files,
    }


def _get_slide_preview(presentation: Any, op: dict) -> dict:
    """获取幻灯片预览 (导出为临时图片).

    Args:
        slide_index: 幻灯片索引
        width: 预览宽度 (可选)
        height: 预览高度 (可选)

    """
    slide_index = op.get("slide_index", 1)
    width = op.get("width", 800)
    height = op.get("height", 600)

    slide = presentation.Slides(slide_index)

    try:
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            temp_path = f.name

        slide.Export(temp_path, EXPORT_FORMATS["png"], width, height)

        return {
            "slide_index": slide_index,
            "preview_path": temp_path,
            "width": width,
            "height": height,
        }
    except Exception as e:
        raise COMOperationError("get_slide_preview", str(e))


def _copy_to_clipboard(presentation: Any, op: dict) -> str:
    """复制幻灯片到剪贴板.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    slide = presentation.Slides(slide_index)

    try:
        # 选中幻灯片并复制
        presentation.Application.ActiveWindow.View.GotoSlide(slide_index)
        slide.Copy()
        return f"copy_to_clipboard: slide={slide_index}"
    except Exception as e:
        raise COMOperationError("copy_to_clipboard", str(e))


def _export_html(presentation: Any, op: dict) -> str:
    """导出为 HTML.

    Args:
        output_path: HTML 输出路径
        include_navigation: 是否包含导航栏

    """
    output_path = op.get("output_path", "")
    include_navigation = op.get("include_navigation", True)

    if not output_path:
        raise COMOperationError("export_html", "output_path 不能为空")

    # 安全: 校验输出路径
    output_path = str(validate_path(output_path))

    output_dir = Path(output_path)
    if not output_dir.exists():
        output_dir.mkdir(parents=True, exist_ok=True)

    try:
        # ppHTML = 12
        presentation.SaveAs(str(output_dir), 12)
        return f"export_html: {output_path}"
    except Exception as e:
        error_msg = str(e)
        if "format" in error_msg.lower() or "SaveAs" in error_msg:
            raise COMOperationError(
                "export_html",
                f"HTML 导出不受当前 PowerPoint 版本支持，请尝试导出为 PDF: {error_msg}"
            )
        raise COMOperationError("export_html", error_msg)


# ============ Slideshow 类功能 ============

def _start_slideshow(presentation: Any, op: dict) -> str:
    """开始放映.

    Args:
        from_slide: 开始幻灯片索引 (可选)
        loop: 是否循环放映

    """
    from_slide = op.get("from_slide", 1)
    loop = op.get("loop", False)

    try:
        # 设置放映范围
        settings = presentation.SlideShowSettings
        settings.StartingSlide = from_slide
        settings.EndingSlide = presentation.Slides.Count
        if loop:
            settings.LoopUntilStopped = True

        # 开始放映
        presentation.SlideShowSettings.Run()
        return f"start_slideshow: from_slide={from_slide}, loop={loop}"
    except Exception as e:
        raise COMOperationError("start_slideshow", str(e))


def _stop_slideshow(presentation: Any, op: dict) -> str:
    """停止放映."""
    try:
        # 获取放映窗口并退出
        app = presentation.Application
        if app.SlideShowWindows.Count > 0:
            app.SlideShowWindows(1).View.Exit()
        return "stop_slideshow: stopped"
    except Exception as e:
        raise COMOperationError("stop_slideshow", str(e))


def _slideshow_next(presentation: Any, op: dict) -> str:
    """放映下一页."""
    try:
        app = presentation.Application
        if app.SlideShowWindows.Count > 0:
            app.SlideShowWindows(1).View.Next()
            return "slideshow_next: moved to next slide"
        return "slideshow_next: no active slideshow"
    except Exception as e:
        raise COMOperationError("slideshow_next", str(e))


def _slideshow_previous(presentation: Any, op: dict) -> str:
    """放映上一页."""
    try:
        app = presentation.Application
        if app.SlideShowWindows.Count > 0:
            app.SlideShowWindows(1).View.Previous()
            return "slideshow_previous: moved to previous slide"
        return "slideshow_previous: no active slideshow"
    except Exception as e:
        raise COMOperationError("slideshow_previous", str(e))


def _slideshow_goto(presentation: Any, op: dict) -> str:
    """放映跳转到指定页.

    Args:
        slide_index: 目标幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    try:
        app = presentation.Application
        if app.SlideShowWindows.Count > 0:
            app.SlideShowWindows(1).View.GotoSlide(slide_index)
            return f"slideshow_goto: slide={slide_index}"
        return "slideshow_goto: no active slideshow"
    except Exception as e:
        raise COMOperationError("slideshow_goto", str(e))


def _get_slideshow_status(presentation: Any, op: dict) -> dict:
    """获取放映状态."""
    try:
        app = presentation.Application
        is_running = app.SlideShowWindows.Count > 0

        result = {
            "is_running": is_running,
            "total_slides": presentation.Slides.Count,
        }

        if is_running:
            view = app.SlideShowWindows(1).View
            result["current_slide"] = view.CurrentShowPosition
            result["state"] = "running"
        else:
            result["current_slide"] = None
            result["state"] = "stopped"

        return result
    except Exception as e:
        raise COMOperationError("get_slideshow_status", str(e))


# ============ Charts 类功能 ============

# 图表类型常量
CHART_TYPES = {
    "column_clustered": 51,      # xlColumnClustered
    "column_stacked": 52,        # xlColumnStacked
    "bar_clustered": 57,         # xlBarClustered
    "bar_stacked": 58,           # xlBarStacked
    "line": 4,                   # xlLine
    "line_markers": 65,          # xlLineMarkers
    "pie": 5,                    # xlPie
    "pie_exploded": 69,          # xlPieExploded
    "scatter": 72,               # xlXYScatter
    "area": 1,                   # xlArea
    "area_stacked": 76,          # xlAreaStacked
    "doughnut": 80,              # xlDoughnut
    "radar": 81,                 # xlRadar
    "radar_filled": 82,          # xlRadarFilled
    "combo": 108,                # xlCombo
}


def _add_chart(presentation: Any, op: dict) -> str:
    """添加图表.

    Args:
        slide_index: 幻灯片索引
        chart_type: 图表类型
        left: 左边距 (pt)
        top: 上边距 (pt)
        width: 宽度 (pt)
        height: 高度 (pt)
        data: 图表数据 (可选)

    """
    slide_index = op.get("slide_index", 1)
    chart_type = op.get("chart_type", "column_clustered")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 400)
    height = op.get("height", 300)
    data = op.get("data", [])

    slide = presentation.Slides(slide_index)
    chart_type_val = CHART_TYPES.get(chart_type, 51)

    try:
        shape = slide.Shapes.AddChart(
            Type=chart_type_val,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )

        chart = shape.Chart

        # 关闭图表数据编辑器窗口（AddChart 在 headless 模式下会弹出数据编辑器导致 COM 阻塞）
        try:
            presentation.Application.CommandBars.ExecuteMso("ChartDataToggle")
        except Exception:
            logger.debug("ChartDataToggle 不可用，尝试其他方式关闭数据编辑器")

        # 如果提供了数据，填充图表
        if data:
            _fill_chart_data(chart, data)

        return f"add_chart: slide={slide_index}, type={chart_type}"
    except Exception as e:
        raise COMOperationError("add_chart", str(e))


def _fill_chart_data(chart: Any, data: list) -> None:
    """填充图表数据."""
    if not data:
        return

    try:
        # 获取图表数据工作表
        workbook = chart.ChartData.Workbook
        worksheet = workbook.Worksheets(1)

        # 清除默认数据
        used_range = worksheet.UsedRange
        if used_range:
            used_range.Clear()

        # 填充新数据
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                worksheet.Cells(row_idx + 1, col_idx + 1).Value = cell_value

    except Exception as e:
        logger.warning(f"填充图表数据失败: {e}")


def _set_chart_data(presentation: Any, op: dict) -> str:
    """设置图表数据.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (图表)
        data: 图表数据 (二维数组)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    data = op.get("data", [])

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not hasattr(shape, 'Chart'):
        raise COMOperationError("set_chart_data", "指定形状不是图表")

    try:
        _fill_chart_data(shape.Chart, data)
        return f"set_chart_data: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("set_chart_data", str(e))


def _get_chart_data(presentation: Any, op: dict) -> dict:
    """获取图表数据.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (图表)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not hasattr(shape, 'Chart'):
        raise COMOperationError("get_chart_data", "指定形状不是图表")

    try:
        chart = shape.Chart
        workbook = chart.ChartData.Workbook
        worksheet = workbook.Worksheets(1)

        used_range = worksheet.UsedRange
        rows = used_range.Rows.Count
        cols = used_range.Columns.Count

        data = []
        for row_idx in range(1, rows + 1):
            row_data = []
            for col_idx in range(1, cols + 1):
                cell_value = worksheet.Cells(row_idx, col_idx).Value
                row_data.append(cell_value)
            data.append(row_data)

        return {
            "slide_index": slide_index,
            "shape_index": shape_index,
            "rows": rows,
            "columns": cols,
            "data": data,
        }
    except Exception as e:
        raise COMOperationError("get_chart_data", str(e))


def _format_chart(presentation: Any, op: dict) -> str:
    """格式化图表.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (图表)
        title: 图表标题
        show_legend: 是否显示图例
        legend_position: 图例位置 (bottom/top/left/right)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    title = op.get("title", "")
    show_legend = op.get("show_legend", True)
    legend_position = op.get("legend_position", "bottom")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not hasattr(shape, 'Chart'):
        raise COMOperationError("format_chart", "指定形状不是图表")

    try:
        chart = shape.Chart

        # 设置标题
        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title

        # 设置图例
        chart.HasLegend = show_legend
        if show_legend:
            # xlLegendPosition: bottom=-4107, top=-4160, left=-4131, right=-4152
            legend_positions = {
                "bottom": -4107,
                "top": -4160,
                "left": -4131,
                "right": -4152,
            }
            chart.Legend.Position = legend_positions.get(legend_position, -4107)

        return f"format_chart: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("format_chart", str(e))


def _format_chart_axis(presentation: Any, op: dict) -> str:
    """格式化图表轴.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (图表)
        axis_type: 轴类型 (category/value)
        title: 轴标题
        has_title: 是否显示轴标题
        has_gridlines: 是否显示网格线

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    axis_type = op.get("axis_type", "category")
    title = op.get("title", "")
    has_title = op.get("has_title", True)
    has_gridlines = op.get("has_gridlines", True)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not hasattr(shape, 'Chart'):
        raise COMOperationError("format_chart_axis", "指定形状不是图表")

    try:
        chart = shape.Chart

        # xlCategory=1, xlValue=2
        axis_idx = 1 if axis_type == "category" else 2
        axis = chart.Axes(axis_idx)

        # 设置轴标题
        axis.HasTitle = has_title
        if has_title and title:
            axis.AxisTitle.Text = title

        # 设置网格线
        if axis_type == "category":
            axis.HasMajorGridlines = has_gridlines
        else:
            axis.HasMajorGridlines = has_gridlines
            axis.HasMinorGridlines = False

        return f"format_chart_axis: slide={slide_index}, axis={axis_type}"
    except Exception as e:
        raise COMOperationError("format_chart_axis", str(e))


def _set_chart_series(presentation: Any, op: dict) -> str:
    """设置图表系列.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (图表)
        series_index: 系列索引
        name: 系列名称
        values: 系列值列表
        chart_type: 系列图表类型 (可选，用于组合图)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    series_index = op.get("series_index", 1)
    name = op.get("name", "")
    values = op.get("values", [])
    chart_type = op.get("chart_type", "")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not hasattr(shape, 'Chart'):
        raise COMOperationError("set_chart_series", "指定形状不是图表")

    try:
        chart = shape.Chart

        if series_index > chart.SeriesCollection().Count:
            raise COMOperationError("set_chart_series", f"系列索引 {series_index} 不存在")

        series = chart.SeriesCollection(series_index)

        # 设置系列名称
        if name:
            series.Name = name

        # 设置系列值
        if values:
            # 构造值字符串
            values_str = ",".join(str(v) for v in values)
            series.Values = values_str

        # 设置系列图表类型 (用于组合图)
        if chart_type:
            chart_type_val = CHART_TYPES.get(chart_type, 51)
            series.ChartType = chart_type_val

        return f"set_chart_series: slide={slide_index}, series={series_index}"
    except Exception as e:
        raise COMOperationError("set_chart_series", str(e))


def _change_chart_type(presentation: Any, op: dict) -> str:
    """更改图表类型.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (图表)
        chart_type: 新图表类型

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    chart_type = op.get("chart_type", "column_clustered")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    if not hasattr(shape, 'Chart'):
        raise COMOperationError("change_chart_type", "指定形状不是图表")

    try:
        chart_type_val = CHART_TYPES.get(chart_type, 51)
        shape.Chart.ChartType = chart_type_val

        return f"change_chart_type: slide={slide_index}, type={chart_type}"
    except Exception as e:
        raise COMOperationError("change_chart_type", str(e))


# ============ Animation 类功能 ============

def _list_animations(presentation: Any, op: dict) -> dict:
    """列出幻灯片动画.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    slide = presentation.Slides(slide_index)

    animations = []
    try:
        sequence = slide.TimeLine.MainSequence
        for i in range(1, sequence.Count + 1):
            effect = sequence(i)
            animations.append({
                "index": i,
                "shape_name": effect.Shape.Name if hasattr(effect, 'Shape') else "",
                "effect_type": str(effect.EffectType) if hasattr(effect, 'EffectType') else "",
                "trigger": str(effect.Timing.TriggerType) if hasattr(effect.Timing, 'TriggerType') else "",
                "duration": effect.Timing.Duration if hasattr(effect.Timing, 'Duration') else 0,
            })
    except Exception as e:
        logger.warning(f"获取动画列表失败: {e}")

    return {
        "slide_index": slide_index,
        "animation_count": len(animations),
        "animations": animations,
    }


def _update_animation(presentation: Any, op: dict) -> str:
    """更新动画属性.

    Args:
        slide_index: 幻灯片索引
        animation_index: 动画索引
        duration: 持续时间 (可选)
        delay: 延迟时间 (可选)
        trigger: 触发方式 (可选)

    """
    slide_index = op.get("slide_index", 1)
    animation_index = op.get("animation_index", 1)
    duration = op.get("duration")
    delay = op.get("delay")
    trigger = op.get("trigger")

    slide = presentation.Slides(slide_index)

    try:
        sequence = slide.TimeLine.MainSequence
        if sequence.Count == 0 or animation_index < 1 or animation_index > sequence.Count:
            return f"update_animation: slide={slide_index}, animation={animation_index} no animation at index {animation_index}"

        effect = sequence(animation_index)

        # 更新持续时间
        if duration is not None:
            effect.Timing.Duration = duration

        # 更新延迟时间
        if delay is not None:
            effect.Timing.TriggerDelayTime = delay

        # 更新触发方式
        if trigger:
            trigger_val = ANIMATION_TRIGGERS.get(trigger, 0)
            effect.Timing.TriggerType = trigger_val

        return f"update_animation: slide={slide_index}, animation={animation_index}"
    except Exception as e:
        raise COMOperationError("update_animation", str(e))


def _remove_animation(presentation: Any, op: dict) -> str:
    """移除动画.

    Args:
        slide_index: 幻灯片索引
        animation_index: 动画索引

    """
    slide_index = op.get("slide_index", 1)
    animation_index = op.get("animation_index", 1)

    slide = presentation.Slides(slide_index)

    try:
        sequence = slide.TimeLine.MainSequence
        if sequence.Count == 0 or animation_index < 1 or animation_index > sequence.Count:
            return f"remove_animation: slide={slide_index}, animation={animation_index} no animation at index {animation_index}"

        sequence(animation_index).Delete()

        return f"remove_animation: slide={slide_index}, animation={animation_index}"
    except Exception as e:
        raise COMOperationError("remove_animation", str(e))


def _clear_animations(presentation: Any, op: dict) -> str:
    """清除所有动画.

    Args:
        slide_index: 幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    slide = presentation.Slides(slide_index)

    try:
        sequence = slide.TimeLine.MainSequence
        while sequence.Count > 0:
            sequence(1).Delete()

        return f"clear_animations: slide={slide_index}"
    except Exception as e:
        raise COMOperationError("clear_animations", str(e))


def _set_animation_trigger(presentation: Any, op: dict) -> str:
    """设置动画触发方式.

    Args:
        slide_index: 幻灯片索引
        animation_index: 动画索引
        trigger: 触发方式 (on_click/after_previous/with_previous)

    """
    slide_index = op.get("slide_index", 1)
    animation_index = op.get("animation_index", 1)
    trigger = op.get("trigger", "on_click")

    slide = presentation.Slides(slide_index)

    try:
        sequence = slide.TimeLine.MainSequence
        if sequence.Count == 0 or animation_index < 1 or animation_index > sequence.Count:
            return f"set_animation_trigger: slide={slide_index}, animation={animation_index} no animation at index {animation_index}"

        effect = sequence(animation_index)
        trigger_val = ANIMATION_TRIGGERS.get(trigger, 0)
        effect.Timing.TriggerType = trigger_val

        return f"set_animation_trigger: slide={slide_index}, animation={animation_index}, trigger={trigger}"
    except Exception as e:
        raise COMOperationError("set_animation_trigger", str(e))


def _copy_animation(presentation: Any, op: dict) -> str:
    """复制动画到其他形状.

    Args:
        slide_index: 幻灯片索引
        source_shape_index: 源形状索引
        target_shape_index: 目标形状索引

    """
    slide_index = op.get("slide_index", 1)
    source_shape_index = op.get("source_shape_index", 1)
    target_shape_index = op.get("target_shape_index", 2)

    slide = presentation.Slides(slide_index)

    try:
        source_shape = slide.Shapes(source_shape_index)
        target_shape = slide.Shapes(target_shape_index)

        # 查找源形状的动画
        sequence = slide.TimeLine.MainSequence
        source_effect = None
        for i in range(1, sequence.Count + 1):
            effect = sequence(i)
            if hasattr(effect, 'Shape') and effect.Shape.Name == source_shape.Name:
                source_effect = effect
                break

        if not source_effect:
            return f"copy_animation: 源形状 {source_shape_index} 没有动画"

        # 复制动画到目标形状
        new_effect = sequence.AddEffect(
            target_shape,
            effectId=source_effect.EffectType,
            Level=0,
        )

        # 复制触发方式
        try:
            new_effect.Timing.TriggerType = source_effect.Timing.TriggerType
            new_effect.Timing.Duration = source_effect.Timing.Duration
        except:
            pass

        return f"copy_animation: from shape {source_shape_index} to shape {target_shape_index}"
    except Exception as e:
        raise COMOperationError("copy_animation", str(e))


# ============ Media 类功能 ============

def _add_video(presentation: Any, op: dict) -> str:
    """添加视频到幻灯片.

    Args:
        slide_index: 幻灯片索引
        video_path: 视频文件路径
        left: 左边距 (pt)
        top: 上边距 (pt)
        width: 宽度 (pt, 可选)
        height: 高度 (pt, 可选)

    """
    slide_index = op.get("slide_index", 1)
    video_path = op.get("video_path", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width")
    height = op.get("height")

    if not video_path:
        raise COMOperationError("add_video", "video_path 不能为空")

    # 安全: 校验路径
    video_path = str(validate_path(video_path))

    slide = presentation.Slides(slide_index)

    try:
        shape = slide.Shapes.AddMediaObject2(
            FileName=video_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
        )

        if width:
            shape.Width = width
        if height:
            shape.Height = height

        return f"add_video: slide={slide_index}, video={video_path}"
    except Exception as e:
        raise COMOperationError("add_video", str(e))


def _add_audio(presentation: Any, op: dict) -> str:
    """添加音频到幻灯片.

    Args:
        slide_index: 幻灯片索引
        audio_path: 音频文件路径
        left: 左边距 (pt)
        top: 上边距 (pt)
        embed: 是否嵌入文档 (默认 True)

    """
    slide_index = op.get("slide_index", 1)
    audio_path = op.get("audio_path", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    embed = op.get("embed", True)

    if not audio_path:
        raise COMOperationError("add_audio", "audio_path 不能为空")

    # 安全: 校验路径
    audio_path = str(validate_path(audio_path))

    slide = presentation.Slides(slide_index)

    try:
        shape = slide.Shapes.AddMediaObject2(
            FileName=audio_path,
            LinkToFile=not embed,
            SaveWithDocument=embed,
            Left=left,
            Top=top,
        )

        return f"add_audio: slide={slide_index}, audio={audio_path}"
    except Exception as e:
        raise COMOperationError("add_audio", str(e))


def _set_media_settings(presentation: Any, op: dict) -> str:
    """设置媒体播放设置.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        play_on_click: 单击时播放
        play_fullscreen: 全屏播放
        play_loop: 循环播放
        play_automatically: 自动播放
        hide_when_not_playing: 不播放时隐藏
        volume: 音量 (0-1)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    play_on_click = op.get("play_on_click", True)
    play_fullscreen = op.get("play_fullscreen", False)
    play_loop = op.get("play_loop", False)
    play_automatically = op.get("play_automatically", False)
    hide_when_not_playing = op.get("hide_when_not_playing", False)
    volume = op.get("volume")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        # 检查是否是媒体形状
        if shape.Type not in (16, 17):  # msoMedia (16), msoVideo (17)
            raise COMOperationError("set_media_settings", "形状不是媒体类型")

        media_settings = shape.MediaFormat

        # 设置播放方式
        if play_automatically:
            media_settings.PlayOnEntry = True
        else:
            media_settings.PlayOnEntry = False

        media_settings.FullScreen = play_fullscreen
        media_settings.Loop = play_loop
        media_settings.HideWhenNotPlaying = hide_when_not_playing

        if volume is not None and 0 <= volume <= 1:
            media_settings.Volume = volume

        return f"set_media_settings: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("set_media_settings", str(e))


# ============ SmartArt 类功能 ============

def _modify_smartart(presentation: Any, op: dict) -> str:
    """修改 SmartArt 图形.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        node_index: 节点索引 (1-based)
        text: 节点文本
        add_node: 添加节点
        remove_node: 删除节点索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    node_index = op.get("node_index")
    text = op.get("text")
    add_node = op.get("add_node", False)
    remove_node = op.get("remove_node")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if not hasattr(shape, 'SmartArt'):
            raise COMOperationError("modify_smartart", "形状不是 SmartArt 类型")

        smartart = shape.SmartArt
        nodes = smartart.AllNodes

        # 添加节点
        if add_node:
            nodes.Add()
            return f"modify_smartart: added node to slide={slide_index}"

        # 删除节点
        if remove_node is not None:
            if 1 <= remove_node <= nodes.Count:
                nodes(remove_node).Delete()
                return f"modify_smartart: removed node {remove_node} from slide={slide_index}"

        # 修改节点文本
        if node_index is not None and text is not None:
            if 1 <= node_index <= nodes.Count:
                nodes(node_index).TextFrame.TextRange.Text = text
                return f"modify_smartart: set node {node_index} text on slide={slide_index}"

        return f"modify_smartart: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("modify_smartart", str(e))


def _list_smartart_layouts(presentation: Any, op: dict) -> dict:
    """列出可用的 SmartArt 布局.

    Returns:
        可用的 SmartArt 布局列表

    """
    # 返回预定义的 SmartArt 类型映射
    layouts = []
    for name, value in SMARTART_TYPES.items():
        layouts.append({
            "name": name,
            "value": value,
            "description": _get_smartart_description(name)
        })

    return {
        "count": len(layouts),
        "layouts": layouts
    }


def _get_smartart_description(name: str) -> str:
    """获取 SmartArt 类型的描述."""
    descriptions = {
        "org_chart": "组织结构图 - 层级关系展示",
        "cycle": "循环图 - 循环流程展示",
        "pyramid": "金字塔图 - 层级递进展示",
        "process": "流程图 - 线性流程展示",
        "list": "列表图 - 项目列表展示",
        "hierarchy": "层次结构图 - 树形结构展示",
        "relationship": "关系图 - 元素关系展示",
        "matrix": "矩阵图 - 二维关系展示",
        "picture": "图片图 - 图片布局展示",
    }
    return descriptions.get(name, "SmartArt 布局")


# ============ Edit Operations 类功能 ============

def _undo(presentation: Any, op: dict) -> str:
    """撤销操作.

    Args:
        steps: 撤销步数 (默认 1)

    """
    steps = op.get("steps", 1)

    try:
        app = presentation.Application
        for _ in range(steps):
            if app.CommandBars.FindControl(Id=128).Enabled:  # 128 = Undo command
                app.CommandBars.FindControl(Id=128).Execute()
            else:
                break
        return f"undo: {steps} step(s)"
    except Exception as e:
        raise COMOperationError("undo", str(e))


def _redo(presentation: Any, op: dict) -> str:
    """重做操作.

    Args:
        steps: 重做步数 (默认 1)

    """
    steps = op.get("steps", 1)

    try:
        app = presentation.Application
        for _ in range(steps):
            if app.CommandBars.FindControl(Id=129).Enabled:  # 129 = Redo command
                app.CommandBars.FindControl(Id=129).Execute()
            else:
                break
        return f"redo: {steps} step(s)"
    except Exception as e:
        raise COMOperationError("redo", str(e))


def _copy_shape(presentation: Any, op: dict) -> str:
    """复制形状.

    Args:
        slide_index: 源幻灯片索引
        shape_index: 源形状索引
        target_slide_index: 目标幻灯片索引 (默认同一幻灯片)
        offset_left: 水平偏移 (pt)
        offset_top: 垂直偏移 (pt)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    target_slide_index = op.get("target_slide_index", slide_index)
    offset_left = op.get("offset_left", 20)
    offset_top = op.get("offset_top", 20)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        # 优先使用 Duplicate（不依赖剪贴板，headless 模式可用）
        try:
            duplicated = shape.Duplicate()
            # Duplicate 创建的副本在同一幻灯片上，移动到目标位置
            duplicated.Left = shape.Left + offset_left
            duplicated.Top = shape.Top + offset_top

            # 如果目标幻灯片不是当前幻灯片，需要剪切粘贴
            if target_slide_index != slide_index:
                duplicated.Cut()
                target_slide = presentation.Slides(target_slide_index)
                target_slide.Shapes.Paste()
                pasted_shape = target_slide.Shapes(target_slide.Shapes.Count)
                pasted_shape.Left = shape.Left + offset_left
                pasted_shape.Top = shape.Top + offset_top
        except Exception:
            # Duplicate 不可用时回退到 Copy+Paste
            shape.Copy()
            target_slide = presentation.Slides(target_slide_index)
            try:
                target_slide.Shapes.Paste()
            except Exception:
                raise COMOperationError("copy_shape", "粘贴失败，剪贴板可能为空或应用程序不可见")
            pasted_shape = target_slide.Shapes(target_slide.Shapes.Count)
            pasted_shape.Left = shape.Left + offset_left
            pasted_shape.Top = shape.Top + offset_top

        return f"copy_shape: slide={slide_index}, shape={shape_index} -> slide={target_slide_index}"
    except COMOperationError:
        raise
    except Exception as e:
        raise COMOperationError("copy_shape", str(e))


def _copy_formatting(presentation: Any, op: dict) -> str:
    """复制形状格式.

    Args:
        slide_index: 幻灯片索引
        shape_index: 源形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        # 使用 PickUp 方法复制格式
        shape.PickUp()
        return f"copy_formatting: slide={slide_index}, shape={shape_index} (格式已复制到剪贴板)"
    except Exception as e:
        raise COMOperationError("copy_formatting", str(e))


def _paste_formatting(presentation: Any, op: dict) -> str:
    """粘贴格式到形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 目标形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        # 使用 Apply 方法粘贴格式
        shape.Apply()
        return f"paste_formatting: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("paste_formatting", str(e))


def _duplicate_slide_to_end(presentation: Any, op: dict) -> str:
    """复制幻灯片到末尾.

    Args:
        slide_index: 要复制的幻灯片索引

    """
    slide_index = op.get("slide_index", 1)

    try:
        slide = presentation.Slides(slide_index)
        new_slide = slide.Duplicate()

        # 获取新幻灯片的索引
        new_index = new_slide.SlideIndex

        return f"duplicate_slide_to_end: slide={slide_index} -> new_index={new_index}"
    except Exception as e:
        raise COMOperationError("duplicate_slide_to_end", str(e))


# ============ Layout 类功能 ============

# 对齐方式常量
ALIGN_TYPES = {
    "left": 0,       # msoAlignLefts
    "center": 1,     # msoAlignCenters
    "right": 2,      # msoAlignRights
    "top": 3,        # msoAlignTops
    "middle": 4,     # msoAlignMiddles
    "bottom": 5,     # msoAlignBottoms
}

# 分布方式常量
DISTRIBUTE_TYPES = {
    "horizontal": 0,  # msoDistributeHorizontally
    "vertical": 1,    # msoDistributeVertically
}

# 翻转类型
FLIP_TYPES = {
    "horizontal": 0,  # msoFlipHorizontal
    "vertical": 1,    # msoFlipVertical
}

# 合并形状类型
MERGE_TYPES = {
    "union": 0,        # msoMergeUnion
    "combine": 1,      # msoMergeCombine
    "intersect": 2,    # msoMergeIntersect
    "subtract": 3,     # msoMergeSubtract
}


def _align_shapes(presentation: Any, op: dict) -> str:
    """对齐多个形状.

    Args:
        slide_index: 幻灯片索引
        shape_indices: 形状索引列表
        align_type: 对齐类型 (left/center/right/top/middle/bottom)
        relative_to_slide: 相对于幻灯片对齐 (默认 False，相对于形状组)

    """
    slide_index = op.get("slide_index", 1)
    shape_indices = op.get("shape_indices", [])
    align_type = op.get("align_type", "left")
    relative_to_slide = op.get("relative_to_slide", False)

    if not shape_indices or len(shape_indices) < 2:
        raise COMOperationError("align_shapes", "需要至少 2 个形状索引")

    slide = presentation.Slides(slide_index)
    align_val = ALIGN_TYPES.get(align_type, 0)

    try:
        # 获取形状范围
        shapes_to_align = []
        for idx in shape_indices:
            shapes_to_align.append(slide.Shapes(idx))

        # 使用第一个形状作为参考
        if relative_to_slide:
            # 相对于幻灯片对齐
            for shape in shapes_to_align:
                if align_type == "left":
                    shape.Left = 0
                elif align_type == "right":
                    shape.Left = presentation.PageSetup.SlideWidth - shape.Width
                elif align_type == "center":
                    shape.Left = (presentation.PageSetup.SlideWidth - shape.Width) / 2
                elif align_type == "top":
                    shape.Top = 0
                elif align_type == "bottom":
                    shape.Top = presentation.PageSetup.SlideHeight - shape.Height
                elif align_type == "middle":
                    shape.Top = (presentation.PageSetup.SlideHeight - shape.Height) / 2
        else:
            # 相对于第一个形状对齐
            ref_shape = shapes_to_align[0]
            for shape in shapes_to_align[1:]:
                if align_type in ("left", "center", "right"):
                    if align_type == "left":
                        shape.Left = ref_shape.Left
                    elif align_type == "center":
                        shape.Left = ref_shape.Left + (ref_shape.Width - shape.Width) / 2
                    elif align_type == "right":
                        shape.Left = ref_shape.Left + ref_shape.Width - shape.Width
                else:
                    if align_type == "top":
                        shape.Top = ref_shape.Top
                    elif align_type == "middle":
                        shape.Top = ref_shape.Top + (ref_shape.Height - shape.Height) / 2
                    elif align_type == "bottom":
                        shape.Top = ref_shape.Top + ref_shape.Height - shape.Height

        return f"align_shapes: slide={slide_index}, type={align_type}, shapes={shape_indices}"
    except Exception as e:
        raise COMOperationError("align_shapes", str(e))


def _distribute_shapes(presentation: Any, op: dict) -> str:
    """分布多个形状.

    Args:
        slide_index: 幻灯片索引
        shape_indices: 形状索引列表
        distribute_type: 分布类型 (horizontal/vertical)

    """
    slide_index = op.get("slide_index", 1)
    shape_indices = op.get("shape_indices", [])
    distribute_type = op.get("distribute_type", "horizontal")

    if not shape_indices or len(shape_indices) < 3:
        raise COMOperationError("distribute_shapes", "需要至少 3 个形状索引")

    slide = presentation.Slides(slide_index)

    try:
        # 获取形状
        shapes_list = [slide.Shapes(idx) for idx in shape_indices]

        if distribute_type == "horizontal":
            # 水平分布
            shapes_list.sort(key=lambda s: s.Left)
            total_width = sum(s.Width for s in shapes_list)
            leftmost = shapes_list[0].Left
            rightmost = shapes_list[-1].Left + shapes_list[-1].Width
            available_space = rightmost - leftmost - total_width
            spacing = available_space / (len(shapes_list) - 1)

            current_left = shapes_list[0].Left + shapes_list[0].Width + spacing
            for shape in shapes_list[1:-1]:
                shape.Left = current_left
                current_left += shape.Width + spacing

        elif distribute_type == "vertical":
            # 垂直分布
            shapes_list.sort(key=lambda s: s.Top)
            total_height = sum(s.Height for s in shapes_list)
            topmost = shapes_list[0].Top
            bottommost = shapes_list[-1].Top + shapes_list[-1].Height
            available_space = bottommost - topmost - total_height
            spacing = available_space / (len(shapes_list) - 1)

            current_top = shapes_list[0].Top + shapes_list[0].Height + spacing
            for shape in shapes_list[1:-1]:
                shape.Top = current_top
                current_top += shape.Height + spacing

        return f"distribute_shapes: slide={slide_index}, type={distribute_type}, shapes={shape_indices}"
    except Exception as e:
        raise COMOperationError("distribute_shapes", str(e))


def _set_slide_size(presentation: Any, op: dict) -> str:
    """设置幻灯片尺寸.

    Args:
        width: 宽度 (pt)
        height: 高度 (pt)
        size_preset: 预设尺寸 (standard/widescreen/a4/custom)

    """
    width = op.get("width")
    height = op.get("height")
    size_preset = op.get("size_preset", "standard")

    # 预设尺寸
    presets = {
        "standard": (720, 540),      # 4:3
        "widescreen": (960, 540),    # 16:9
        "a4": (792, 612),           # A4
    }

    try:
        if size_preset in presets and width is None:
            width, height = presets[size_preset]

        if width and height:
            presentation.PageSetup.SlideWidth = width
            presentation.PageSetup.SlideHeight = height

        return f"set_slide_size: {width}x{height} pt"
    except Exception as e:
        raise COMOperationError("set_slide_size", str(e))


def _flip_shape(presentation: Any, op: dict) -> str:
    """翻转形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        flip_type: 翻转类型 (horizontal/vertical)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    flip_type = op.get("flip_type", "horizontal")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        flip_val = FLIP_TYPES.get(flip_type, 0)
        shape.Flip(flip_val)

        return f"flip_shape: slide={slide_index}, shape={shape_index}, type={flip_type}"
    except Exception as e:
        raise COMOperationError("flip_shape", str(e))


def _merge_shapes(presentation: Any, op: dict) -> str:
    """合并形状.

    Args:
        slide_index: 幻灯片索引
        shape_indices: 形状索引列表
        merge_type: 合并类型 (union/combine/intersect/subtract)

    """
    slide_index = op.get("slide_index", 1)
    shape_indices = op.get("shape_indices", [])
    merge_type = op.get("merge_type", "union")

    if not shape_indices or len(shape_indices) < 2:
        raise COMOperationError("merge_shapes", "需要至少 2 个形状索引")

    slide = presentation.Slides(slide_index)
    merge_val = MERGE_TYPES.get(merge_type, 0)

    try:
        # 获取形状范围
        shapes_to_merge = []
        for idx in shape_indices:
            try:
                shapes_to_merge.append(slide.Shapes(idx))
            except Exception:
                continue

        if len(shapes_to_merge) < 2:
            raise COMOperationError("merge_shapes", f"找到的有效形状不足2个 ({len(shapes_to_merge)})")

        # 使用 ShapeRange 进行合并
        # 构建 ShapeRange 通过名称数组
        shape_names = [s.Name for s in shapes_to_merge]
        shape_range = slide.Shapes.Range(shape_names)

        # 使用 CommandBars 执行合并命令
        app = presentation.Application
        try:
            shape_range.Select()
        except Exception:
            # 无界面模式下 Select 可能失败，尝试直接执行命令
            pass

        merge_command = {
            "union": "ShapesUnion",
            "combine": "ShapesCombine",
            "intersect": "ShapesIntersect",
            "subtract": "ShapesSubtract",
        }.get(merge_type, "ShapesUnion")

        app.CommandBars.ExecuteMso(merge_command)

        return f"merge_shapes: slide={slide_index}, type={merge_type}, shapes={shape_indices}"
    except Exception as e:
        raise COMOperationError("merge_shapes", str(e))


def _rotate_shape(presentation: Any, op: dict) -> str:
    """旋转形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        angle: 旋转角度 (度)
        relative: 是否相对旋转 (默认 True)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    angle = op.get("angle", 0)
    relative = op.get("relative", True)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if relative:
            shape.Rotation = shape.Rotation + angle
        else:
            shape.Rotation = angle

        return f"rotate_shape: slide={slide_index}, shape={shape_index}, angle={angle}"
    except Exception as e:
        raise COMOperationError("rotate_shape", str(e))


def _lock_aspect_ratio(presentation: Any, op: dict) -> str:
    """锁定/解锁形状宽高比.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        lock: 是否锁定 (True 锁定, False 解锁)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    lock = op.get("lock", True)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        shape.LockAspectRatio = lock

        return f"lock_aspect_ratio: slide={slide_index}, shape={shape_index}, lock={lock}"
    except Exception as e:
        raise COMOperationError("lock_aspect_ratio", str(e))


# ============ Effects 类功能 ============

def _set_glow(presentation: Any, op: dict) -> str:
    """设置形状发光效果.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        color: 发光颜色 (#RRGGBB)
        size: 发光大小 (pt)
        transparency: 透明度 (0-1)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    color = op.get("color", "#FFFFFF")
    size = op.get("size", 10)
    transparency = op.get("transparency", 0.5)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        glow = shape.Glow
        try:
            glow.Visible = True
        except Exception:
            pass  # Visible may not be settable on all shapes
        glow.Color.RGB = _hex_to_rgb(color)
        glow.Radius = size
        glow.Transparency = transparency

        return f"set_glow: slide={slide_index}, shape={shape_index}, color={color}"
    except Exception as e:
        raise COMOperationError("set_glow", str(e))


def _set_reflection(presentation: Any, op: dict) -> str:
    """设置形状反射效果.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        transparency: 透明度 (0-1)
        size: 反射大小 (0-1)
        blur: 模糊程度 (pt)
        distance: 距离 (pt)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    transparency = op.get("transparency", 0.5)
    size = op.get("size", 0.5)
    blur = op.get("blur", 5)
    distance = op.get("distance", 10)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        reflection = shape.Reflection
        try:
            reflection.Visible = True
        except Exception:
            pass  # Visible may not be settable on all shapes
        reflection.Transparency = transparency
        reflection.Size = size * 100  # PowerPoint 使用 0-100
        reflection.Blur = blur
        reflection.Offset = distance

        return f"set_reflection: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("set_reflection", str(e))


def _set_soft_edge(presentation: Any, op: dict) -> str:
    """设置形状柔化边缘效果.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        size: 柔化大小 (pt)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    size = op.get("size", 10)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        soft_edge = shape.SoftEdge
        try:
            soft_edge.Visible = True
        except Exception:
            pass  # Visible may not be settable on all shapes
        soft_edge.Radius = size

        return f"set_soft_edge: slide={slide_index}, shape={shape_index}, size={size}pt"
    except Exception as e:
        raise COMOperationError("set_soft_edge", str(e))


# ============ Comments 批注功能 ============

def _add_comment(presentation: Any, op: dict) -> str:
    """添加批注.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引 (可选，默认为幻灯片批注)
        text: 批注内容
        author: 作者名称 (可选)
        left: 批注位置左边距 (可选)
        top: 批注位置上边距 (可选)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index")
    text = op.get("text", "")
    author = op.get("author", "")
    left = op.get("left", 100)
    top = op.get("top", 100)

    if not text:
        raise COMOperationError("add_comment", "批注内容不能为空")

    slide = presentation.Slides(slide_index)

    try:
        # 添加批注
        comment = slide.Comments.Add(
            Left=left,
            Top=top,
            Author=author,
            AuthorInitials="",
            Text=text,
        )
        return f"add_comment: slide={slide_index}, author={author}"
    except Exception as e:
        raise COMOperationError("add_comment", str(e))


def _list_comments(presentation: Any, op: dict) -> dict:
    """列出幻灯片上的批注.

    Args:
        slide_index: 幻灯片索引 (可选，默认列出所有)

    """
    slide_index = op.get("slide_index")
    comments = []

    try:
        if slide_index:
            slides = [presentation.Slides(slide_index)]
        else:
            slides = list(presentation.Slides)

        for slide in slides:
            slide_num = slide.SlideIndex
            for i in range(1, slide.Comments.Count + 1):
                comment = slide.Comments(i)
                comments.append({
                    "slide_index": slide_num,
                    "comment_index": i,
                    "author": comment.Author,
                    "text": comment.Text,
                    "left": comment.Left,
                    "top": comment.Top,
                })

        return {"total": len(comments), "comments": comments}
    except Exception as e:
        raise COMOperationError("list_comments", str(e))


def _delete_comment(presentation: Any, op: dict) -> str:
    """删除批注.

    Args:
        slide_index: 幻灯片索引
        comment_index: 批注索引

    """
    slide_index = op.get("slide_index", 1)
    comment_index = op.get("comment_index", 1)

    slide = presentation.Slides(slide_index)

    try:
        if slide.Comments.Count < comment_index:
            raise COMOperationError("delete_comment", f"批注索引 {comment_index} 不存在")

        slide.Comments(comment_index).Delete()
        return f"delete_comment: slide={slide_index}, comment_index={comment_index}"
    except Exception as e:
        raise COMOperationError("delete_comment", str(e))


# ============ Advanced 高级功能 ============

def _set_tags(presentation: Any, op: dict) -> str:
    """设置形状标签.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        tag_name: 标签名称
        tag_value: 标签值

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    tag_name = op.get("tag_name", "")
    tag_value = op.get("tag_value", "")

    if not tag_name:
        raise COMOperationError("set_tags", "标签名称不能为空")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        shape.Tags.Add(tag_name, tag_value)
        return f"set_tags: slide={slide_index}, shape={shape_index}, tag={tag_name}"
    except Exception as e:
        raise COMOperationError("set_tags", str(e))


def _get_tags(presentation: Any, op: dict) -> dict:
    """获取形状标签.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    tags = {}
    try:
        for i in range(1, shape.Tags.Count + 1):
            tag_name = shape.Tags.Name(i)
            tag_value = shape.Tags.Value(i)
            tags[tag_name] = tag_value
        return {"slide_index": slide_index, "shape_index": shape_index, "tags": tags}
    except Exception as e:
        raise COMOperationError("get_tags", str(e))


def _set_default_font(presentation: Any, op: dict) -> str:
    """设置演示文稿默认字体.

    Args:
        font_name: 字体名称
        font_size: 字体大小 (pt)
        font_color: 字体颜色 (#RRGGBB)

    """
    font_name = op.get("font_name", "")
    font_size = op.get("font_size")
    font_color = op.get("font_color", "")

    try:
        # 设置幻灯片母版的默认字体
        master = presentation.SlideMaster
        if font_name:
            master.TextStyles(1).Font.Name = font_name
            master.TextStyles(1).Font.NameFarEast = font_name
        if font_size:
            master.TextStyles(1).Font.Size = font_size
        if font_color:
            master.TextStyles(1).Font.Color.RGB = _hex_to_rgb(font_color)

        return f"set_default_font: font={font_name}, size={font_size}"
    except Exception as e:
        raise COMOperationError("set_default_font", str(e))


def _replace_font(presentation: Any, op: dict) -> str:
    """批量替换演示文稿中的字体.

    Args:
        old_font: 原字体名称
        new_font: 新字体名称

    """
    old_font = op.get("old_font", "")
    new_font = op.get("new_font", "")

    if not old_font or not new_font:
        raise COMOperationError("replace_font", "需要提供 old_font 和 new_font")

    count = 0
    try:
        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    text_range = shape.TextFrame.TextRange
                    if text_range.Font.Name == old_font:
                        text_range.Font.Name = new_font
                        count += 1
                    # 检查远东字体
                    try:
                        if text_range.Font.NameFarEast == old_font:
                            text_range.Font.NameFarEast = new_font
                    except:
                        pass

        return f"replace_font: {old_font} -> {new_font}, 替换了 {count} 处"
    except Exception as e:
        raise COMOperationError("replace_font", str(e))


def _crop_picture(presentation: Any, op: dict) -> str:
    """裁剪图片.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        left: 左边裁剪比例 (0-1)
        top: 上边裁剪比例 (0-1)
        right: 右边裁剪比例 (0-1)
        bottom: 下边裁剪比例 (0-1)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    left = op.get("left", 0)
    top = op.get("top", 0)
    right = op.get("right", 0)
    bottom = op.get("bottom", 0)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        # 检查是否为图片
        if shape.Type not in (13, 11):  # msoPicture, msoLinkedPicture
            raise COMOperationError("crop_picture", "形状不是图片类型")

        shape.PictureFormat.CropLeft = left * shape.Width
        shape.PictureFormat.CropTop = top * shape.Height
        shape.PictureFormat.CropRight = right * shape.Width
        shape.PictureFormat.CropBottom = bottom * shape.Height

        return f"crop_picture: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("crop_picture", str(e))


def _set_picture_format(presentation: Any, op: dict) -> str:
    """设置图片格式.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        brightness: 亮度 (-1 到 1)
        contrast: 对比度 (-1 到 1)
        transparency: 透明度 (0 到 1)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    brightness = op.get("brightness")
    contrast = op.get("contrast")
    transparency = op.get("transparency")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if shape.Type not in (13, 11):  # msoPicture, msoLinkedPicture
            raise COMOperationError("set_picture_format", "形状不是图片类型")

        if brightness is not None:
            shape.PictureFormat.Brightness = brightness
        if contrast is not None:
            shape.PictureFormat.Contrast = contrast
        if transparency is not None:
            shape.PictureFormat.Transparency = transparency

        return f"set_picture_format: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("set_picture_format", str(e))


def _export_shape(presentation: Any, op: dict) -> str:
    """导出形状为图片.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        output_path: 输出路径
        format: 图片格式 (png/jpg/emf/wmf)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    output_path = op.get("output_path", "")
    format_type = op.get("format", "png")

    if not output_path:
        raise COMOperationError("export_shape", "output_path 不能为空")

    # 安全: 校验输出路径
    output_path = str(validate_path(output_path))

    # 格式映射
    format_map = {
        "png": 18,   # ppShapeFormatPNG
        "jpg": 17,    # ppShapeFormatJPG
        "emf": 23,    # ppShapeFormatEMF
        "wmf": 22,    # ppShapeFormatWMF
    }
    format_val = format_map.get(format_type.lower(), 18)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        shape.Export(output_path, format_val)
        return f"export_shape: slide={slide_index}, shape={shape_index}, path={output_path}"
    except Exception as e:
        raise COMOperationError("export_shape", str(e))


def _set_shape_visibility(presentation: Any, op: dict) -> str:
    """设置形状可见性.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        visible: 是否可见

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    visible = op.get("visible", True)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        shape.Visible = visible
        return f"set_shape_visibility: slide={slide_index}, shape={shape_index}, visible={visible}"
    except Exception as e:
        raise COMOperationError("set_shape_visibility", str(e))


def _select_shape(presentation: Any, op: dict) -> str:
    """选择形状.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        shape.Select()
        return f"select_shape: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("select_shape", str(e))


def _get_selection(presentation: Any, op: dict) -> dict:
    """获取当前选择.

    Returns:
        包含选择类型和选中对象的字典

    """
    try:
        window = presentation.Application.ActiveWindow
        selection = window.Selection

        result = {
            "type": selection.Type,  # 0=无, 1=形状, 2=幻灯片, 3=文本
            "count": 0,
            "shapes": [],
        }

        if selection.Type == 1:  # ppSelectionShapes
            result["count"] = selection.ShapeRange.Count
            for i in range(1, selection.ShapeRange.Count + 1):
                shape = selection.ShapeRange(i)
                result["shapes"].append({
                    "name": shape.Name,
                    "type": shape.Type,
                    "left": shape.Left,
                    "top": shape.Top,
                    "width": shape.Width,
                    "height": shape.Height,
                })
        elif selection.Type == 2:  # ppSelectionSlides
            result["count"] = selection.SlideRange.Count
            result["slide_indices"] = [s.SlideIndex for s in selection.SlideRange]

        return result
    except Exception as e:
        raise COMOperationError("get_selection", str(e))


def _set_view(presentation: Any, op: dict) -> str:
    """设置视图模式.

    Args:
        view_type: 视图类型 (normal/slide_sorter/notes_page/slide_master/handout_master)

    """
    view_type = op.get("view_type", "normal")

    # 视图类型映射
    view_map = {
        "normal": 1,          # ppViewNormal
        "slide_sorter": 2,    # ppViewSlideSorter
        "notes_page": 3,      # ppViewNotesPage
        "slide_master": 4,    # ppViewSlideMaster
        "handout_master": 5,  # ppViewHandoutMaster
    }
    view_val = view_map.get(view_type, 1)

    try:
        presentation.Application.ActiveWindow.ViewType = view_val
        return f"set_view: {view_type}"
    except Exception as e:
        raise COMOperationError("set_view", str(e))


def _copy_animation_from_shape(presentation: Any, op: dict) -> str:
    """从形状复制动画到另一个形状.

    Args:
        slide_index: 幻灯片索引
        source_shape_index: 源形状索引
        target_shape_index: 目标形状索引

    """
    slide_index = op.get("slide_index", 1)
    source_shape_index = op.get("source_shape_index", 1)
    target_shape_index = op.get("target_shape_index", 2)

    slide = presentation.Slides(slide_index)
    source_shape = slide.Shapes(source_shape_index)
    target_shape = slide.Shapes(target_shape_index)

    try:
        # 获取源形状的动画效果
        main_sequence = slide.TimeLine.MainSequence
        source_effects = []

        for i in range(1, main_sequence.Count + 1):
            effect = main_sequence(i)
            if effect.Shape.Id == source_shape.Id:
                source_effects.append({
                    "effect_id": effect.EffectType,
                    "trigger": effect.Timing.TriggerType,
                    "delay": effect.Timing.TriggerDelayTime,
                    "duration": effect.Timing.Duration,
                })

        # 复制动画到目标形状
        for effect_data in source_effects:
            new_effect = main_sequence.AddEffect(
                target_shape,
                effectId=effect_data["effect_id"],
            )
            try:
                new_effect.Timing.TriggerType = effect_data["trigger"]
                new_effect.Timing.TriggerDelayTime = effect_data["delay"]
                new_effect.Timing.Duration = effect_data["duration"]
            except:
                pass

        return f"copy_animation_from_shape: {source_shape_index} -> {target_shape_index}"
    except Exception as e:
        raise COMOperationError("copy_animation_from_shape", str(e))


def _add_picture_from_url(presentation: Any, op: dict) -> str:
    """从 URL 添加图片.

    Args:
        slide_index: 幻灯片索引
        url: 图片 URL
        left: 左边距 (pt)
        top: 上边距 (pt)
        width: 宽度 (pt, 可选)
        height: 高度 (pt, 可选)

    """
    slide_index = op.get("slide_index", 1)
    url = op.get("url", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width")
    height = op.get("height")

    if not url:
        raise COMOperationError("add_picture_from_url", "URL 不能为空")

    # SSRF 防护: 禁止访问内网/本地地址
    import socket
    from urllib.parse import urlparse
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise COMOperationError("add_picture_from_url", f"不支持的协议: {parsed.scheme}")
    if parsed.hostname:
        hostname = parsed.hostname.lower()
        if hostname in ("localhost", "127.0.0.1", "::1", "0.0.0.0"):
            raise COMOperationError("add_picture_from_url", f"禁止访问本地地址: {hostname}")
        try:
            # 解析域名对应的 IP, 拒绝私有地址
            for info in socket.getaddrinfo(hostname, None):
                ip = info[4][0]
                try:
                    ip_obj = ipaddress.ip_address(ip)
                except ValueError:
                    continue
                if ip_obj.is_private or ip_obj.is_loopback or ip_obj.is_link_local or ip_obj.is_reserved:
                    raise COMOperationError(
                        "add_picture_from_url",
                        f"禁止访问内网/保留地址: {hostname} -> {ip}",
                    )
        except socket.gaierror:
            raise COMOperationError("add_picture_from_url", f"无法解析域名: {hostname}")

    slide = presentation.Slides(slide_index)
    temp_path = None

    try:
        # 下载图片到临时文件
        with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as f:
            temp_path = f.name
        urllib.request.urlretrieve(url, temp_path)

        # 插入图片
        shape = slide.Shapes.AddPicture(
            FileName=temp_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
        )

        if width:
            shape.Width = width
        if height:
            shape.Height = height

        return f"add_picture_from_url: slide={slide_index}, url={url[:50]}"
    except Exception as e:
        raise COMOperationError("add_picture_from_url", str(e))
    finally:
        if temp_path:
            try:
                os.unlink(temp_path)
            except OSError:
                pass


def _add_svg_icon(presentation: Any, op: dict) -> str:
    """添加 SVG 图标.

    Args:
        slide_index: 幻灯片索引
        svg_path: SVG 文件路径
        left: 左边距 (pt)
        top: 上边距 (pt)
        width: 宽度 (pt)
        height: 高度 (pt)

    """
    slide_index = op.get("slide_index", 1)
    svg_path = op.get("svg_path", "")
    left = op.get("left", 100)
    top = op.get("top", 100)
    width = op.get("width", 72)
    height = op.get("height", 72)

    if not svg_path:
        raise COMOperationError("add_svg_icon", "svg_path 不能为空")

    # 安全: 校验路径
    svg_path = str(validate_path(svg_path))

    if not Path(svg_path).exists():
        raise COMOperationError("add_svg_icon", f"SVG 文件不存在: {svg_path}")

    slide = presentation.Slides(slide_index)

    try:
        # PowerPoint 通过 AddPicture 支持插入 SVG
        shape = slide.Shapes.AddPicture(
            FileName=svg_path,
            LinkToFile=False,
            SaveWithDocument=True,
            Left=left,
            Top=top,
            Width=width,
            Height=height,
        )
        return f"add_svg_icon: slide={slide_index}, path={svg_path}"
    except Exception as e:
        raise COMOperationError("add_svg_icon", str(e))


def _lock_aspect_ratio(presentation: Any, op: dict) -> str:
    """锁定形状宽高比.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        locked: 是否锁定

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    locked = op.get("locked", True)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        shape.LockAspectRatio = locked
        return f"lock_aspect_ratio: slide={slide_index}, shape={shape_index}, locked={locked}"
    except Exception as e:
        raise COMOperationError("lock_aspect_ratio", str(e))


def _batch_apply_format(presentation: Any, op: dict) -> str:
    """批量应用格式到多个形状.

    Args:
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
    slide_index = op.get("slide_index", 1)
    shape_indices = op.get("shape_indices", [])
    fill_color = op.get("fill_color", "")
    line_color = op.get("line_color", "")
    line_width = op.get("line_width")
    font_name = op.get("font_name", "")
    font_size = op.get("font_size")
    font_color = op.get("font_color", "")
    bold = op.get("bold")
    italic = op.get("italic")

    if not shape_indices:
        raise COMOperationError("batch_apply_format", "shape_indices 不能为空")

    slide = presentation.Slides(slide_index)
    count = 0

    try:
        for idx in shape_indices:
            shape = slide.Shapes(idx)

            if fill_color:
                shape.Fill.Solid()
                shape.Fill.ForeColor.RGB = _hex_to_rgb(fill_color)

            if line_color:
                shape.Line.ForeColor.RGB = _hex_to_rgb(line_color)
                if line_width is not None:
                    shape.Line.Weight = line_width

            if shape.HasTextFrame and shape.TextFrame.HasText:
                text_range = shape.TextFrame.TextRange
                if font_name:
                    text_range.Font.Name = font_name
                if font_size is not None:
                    text_range.Font.Size = font_size
                if font_color:
                    text_range.Font.Color.RGB = _hex_to_rgb(font_color)
                if bold is not None:
                    text_range.Font.Bold = bold
                if italic is not None:
                    text_range.Font.Italic = italic

            count += 1

        return f"batch_apply_format: {count} shapes"
    except Exception as e:
        raise COMOperationError("batch_apply_format", str(e))


def _set_default_shape_style(presentation: Any, op: dict) -> str:
    """设置默认形状样式.

    Args:
        fill_color: 填充颜色 (#RRGGBB)
        line_color: 边框颜色 (#RRGGBB)
        line_width: 边框宽度 (pt)
        line_visible: 边框是否可见

    """
    fill_color = op.get("fill_color", "")
    line_color = op.get("line_color", "")
    line_width = op.get("line_width", 1)
    line_visible = op.get("line_visible", True)

    try:
        # 设置默认形状格式
        # 注意：这会影响之后创建的所有形状
        app = presentation.Application

        # 创建临时形状来设置默认格式
        slide = presentation.Slides(1) if presentation.Slides.Count > 0 else presentation.Slides.Add(1, 12)
        temp_shape = slide.Shapes.AddShape(1, 0, 0, 100, 100)  # 矩形

        if fill_color:
            temp_shape.Fill.Solid()
            temp_shape.Fill.ForeColor.RGB = _hex_to_rgb(fill_color)

        if line_color:
            temp_shape.Line.ForeColor.RGB = _hex_to_rgb(line_color)

        if line_width is not None:
            temp_shape.Line.Weight = line_width

        temp_shape.Line.Visible = line_visible

        # 设置为默认形状格式
        temp_shape.PickUp()

        # 删除临时形状
        temp_shape.Delete()

        return f"set_default_shape_style: fill={fill_color}, line={line_color}"
    except Exception as e:
        raise COMOperationError("set_default_shape_style", str(e))


def _get_shape_count(presentation: Any, op: dict) -> dict:
    """获取幻灯片上的形状数量.

    Args:
        slide_index: 幻灯片索引 (可选，默认统计所有)

    """
    slide_index = op.get("slide_index")

    result = {"total": 0, "slides": []}

    try:
        if slide_index:
            slides = [presentation.Slides(slide_index)]
        else:
            slides = list(presentation.Slides)

        for slide in slides:
            count = slide.Shapes.Count
            result["total"] += count
            result["slides"].append({
                "slide_index": slide.SlideIndex,
                "shape_count": count,
            })

        return result
    except Exception as e:
        raise COMOperationError("get_shape_count", str(e))


# ============ Freeform 自由路径功能 ============

def _build_freeform_path(presentation: Any, op: dict) -> str:
    """构建自由路径形状.

    Args:
        slide_index: 幻灯片索引
        points: 坐标点列表，每个元素为 [x, y]
        closed: 是否闭合路径
        left: 左边距 (pt)
        top: 上边距 (pt)
        fill_color: 填充颜色 (#RRGGBB, 可选)
        line_color: 边框颜色 (#RRGGBB, 可选)
        line_width: 边框宽度 (pt)

    """
    slide_index = op.get("slide_index", 1)
    points = op.get("points", [])
    closed = op.get("closed", True)
    left = op.get("left", 100)
    top = op.get("top", 100)
    fill_color = op.get("fill_color", "")
    line_color = op.get("line_color", "#000000")
    line_width = op.get("line_width", 1)

    if not points or len(points) < 2:
        raise COMOperationError("build_freeform_path", "需要至少提供 2 个坐标点")

    # 支持 points 为 [{"x":..,"y":..}] 或 [[x,y]] 格式
    normalized = []
    for p in points:
        if isinstance(p, dict):
            normalized.append([p.get("x", 0), p.get("y", 0)])
        elif isinstance(p, (list, tuple)) and len(p) >= 2:
            normalized.append([p[0], p[1]])

    if len(normalized) < 2:
        raise COMOperationError("build_freeform_path", "需要至少提供 2 个有效坐标点")

    slide = presentation.Slides(slide_index)

    try:
        # 使用 BuildFreeform 方法构建自由路径
        freeform_builder = slide.Shapes.BuildFreeform(
            EditingType=0,  # msoEditingAuto
            X1=normalized[0][0] + left,
            Y1=normalized[0][1] + top,
        )

        # 添加节点
        for i, point in enumerate(normalized[1:], start=1):
            freeform_builder.AddNodes(
                SegmentType=0,  # msoSegmentLine
                EditingType=0,  # msoEditingAuto
                X1=point[0] + left,
                Y1=point[1] + top,
            )

        # 转换为形状
        shape = freeform_builder.ConvertToShape()

        # 设置样式
        if fill_color:
            shape.Fill.Solid()
            shape.Fill.ForeColor.RGB = _hex_to_rgb(fill_color)
        else:
            shape.Fill.Visible = False

        if line_color:
            shape.Line.ForeColor.RGB = _hex_to_rgb(line_color)
            shape.Line.Weight = line_width

        return f"build_freeform_path: slide={slide_index}, points={len(points)}"
    except Exception as e:
        raise COMOperationError("build_freeform_path", str(e))


def _get_node_positions(presentation: Any, op: dict) -> dict:
    """获取自由形状的节点位置.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        # 检查是否为自由形状
        if shape.Type != 5:  # msoFreeform
            return "get_node_positions: skipped (shape is not a freeform)"

        nodes = shape.Nodes
        node_list = []

        for i in range(1, nodes.Count + 1):
            node = nodes(i)
            node_list.append({
                "index": i,
                "x": node.Points(1, 1),  # X 坐标
                "y": node.Points(1, 2),  # Y 坐标
                "editing_type": node.EditingType,
                "segment_type": node.SegmentType,
            })

        return {"slide_index": slide_index, "shape_index": shape_index, "nodes": node_list}
    except Exception as e:
        raise COMOperationError("get_node_positions", str(e))


def _set_node_positions(presentation: Any, op: dict) -> str:
    """设置自由形状的节点位置.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        node_positions: 节点位置列表，每个元素为 {"index": int, "x": float, "y": float}

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    node_positions = op.get("node_positions", [])

    if not node_positions:
        raise COMOperationError("set_node_positions", "node_positions 不能为空")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if shape.Type != 5:  # msoFreeform
            return f"set_node_positions: skipped (shape type={shape.Type}, not a freeform)"

        nodes = shape.Nodes
        for pos in node_positions:
            idx = pos.get("index", 1)
            x = pos.get("x", 0)
            y = pos.get("y", 0)

            if 1 <= idx <= nodes.Count:
                node = nodes(idx)
                node.SetPosition(1, x, y)  # 设置第一个点的位置

        return f"set_node_positions: {len(node_positions)} nodes updated"
    except Exception as e:
        raise COMOperationError("set_node_positions", str(e))


def _insert_node(presentation: Any, op: dict) -> str:
    """在自由形状中插入节点.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        after_index: 在此节点之后插入
        x: X 坐标
        y: Y 坐标
        segment_type: 线段类型 (line/curve)
        editing_type: 编辑类型 (auto/corner/smooth/symmetric)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    after_index = op.get("after_index", 1)
    x = op.get("x", 0)
    y = op.get("y", 0)
    segment_type = op.get("segment_type", "line")
    editing_type = op.get("editing_type", "auto")

    # 线段类型映射
    segment_map = {"line": 0, "curve": 1}  # msoSegmentLine, msoSegmentCurve
    segment_val = segment_map.get(segment_type, 0)

    # 编辑类型映射
    editing_map = {"auto": 0, "corner": 1, "smooth": 2, "symmetric": 3}
    editing_val = editing_map.get(editing_type, 0)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if shape.Type != 5:  # msoFreeform
            return "insert_node: skipped (shape is not a freeform)"

        shape.Nodes.Insert(
            Index=after_index,
            SegmentType=segment_val,
            EditingType=editing_val,
            X1=x,
            Y1=y,
        )

        return f"insert_node: after index {after_index}, at ({x}, {y})"
    except Exception as e:
        raise COMOperationError("insert_node", str(e))


def _delete_node(presentation: Any, op: dict) -> str:
    """删除自由形状中的节点.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        node_index: 节点索引

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    node_index = op.get("node_index", 1)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if shape.Type != 5:  # msoFreeform
            return "delete_node: skipped (shape is not a freeform)"

        if shape.Nodes.Count <= 1:
            raise COMOperationError("delete_node", "无法删除最后一个节点")

        shape.Nodes.Delete(node_index)

        return f"delete_node: node {node_index}"
    except Exception as e:
        raise COMOperationError("delete_node", str(e))


def _set_node_editing_type(presentation: Any, op: dict) -> str:
    """设置节点编辑类型.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        node_index: 节点索引
        editing_type: 编辑类型 (auto/corner/smooth/symmetric)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    node_index = op.get("node_index", 1)
    editing_type = op.get("editing_type", "auto")

    # 编辑类型映射
    editing_map = {"auto": 0, "corner": 1, "smooth": 2, "symmetric": 3}
    editing_val = editing_map.get(editing_type, 0)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if shape.Type != 5:  # msoFreeform
            return "set_node_editing_type: skipped (shape is not a freeform)"

        node = shape.Nodes(node_index)
        node.SetEditingType(editing_val)

        return f"set_node_editing_type: node {node_index} -> {editing_type}"
    except Exception as e:
        raise COMOperationError("set_node_editing_type", str(e))


def _set_segment_type(presentation: Any, op: dict) -> str:
    """设置线段类型.

    Args:
        slide_index: 幻灯片索引
        shape_index: 形状索引
        node_index: 节点索引
        segment_type: 线段类型 (line/curve)

    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    node_index = op.get("node_index", 1)
    segment_type = op.get("segment_type", "line")

    # 线段类型映射
    segment_map = {"line": 0, "curve": 1}  # msoSegmentLine, msoSegmentCurve
    segment_val = segment_map.get(segment_type, 0)

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    try:
        if shape.Type != 5:  # msoFreeform
            return "set_segment_type: skipped (shape is not a freeform)"

        node = shape.Nodes(node_index)
        node.SetSegmentType(segment_val)

        return f"set_segment_type: node {node_index} -> {segment_type}"
    except Exception as e:
        raise COMOperationError("set_segment_type", str(e))


# ============ Connectors 连接线功能 ============

def _add_connector(presentation: Any, op: dict) -> str:
    slide_index = op.get("slide_index", 1)
    connector_type = op.get("connector_type", "elbow")
    start_shape_idx = op.get("start_shape", 1)
    start_connection = op.get("start_connection", 1)
    end_shape_idx = op.get("end_shape", 2)
    end_connection = op.get("end_connection", 1)

    conn_type = CONNECTOR_TYPES_PPT.get(connector_type, 2)
    try:
        slide = presentation.Slides(slide_index)
        start_shape = slide.Shapes(start_shape_idx)
        end_shape = slide.Shapes(end_shape_idx)
        connector = slide.Shapes.AddConnector(
            Type=conn_type,
            BeginX=start_shape.Left + start_shape.Width / 2,
            BeginY=start_shape.Top + start_shape.Height / 2,
            EndX=end_shape.Left + end_shape.Width / 2,
            EndY=end_shape.Top + end_shape.Height / 2,
        )
        connector.ConnectorFormat.BeginConnect(start_shape, start_connection)
        connector.ConnectorFormat.EndConnect(end_shape, end_connection)
        return f"add_connector: slide={slide_index}, type={connector_type}"
    except Exception as e:
        raise COMOperationError("add_connector", str(e))


def _format_connector(presentation: Any, op: dict) -> str:
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    color = op.get("color")
    weight = op.get("weight")
    dash_style = op.get("dash_style")
    arrow_begin = op.get("arrow_begin")
    arrow_end = op.get("arrow_end")

    ARROW_STYLES = {"none": -2, "arrow": 1, "stealth": 2, "diamond": 3, "oval": 4}
    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)
        line = shape.Line
        if color:
            line.ForeColor.RGB = _hex_to_rgb(color)
        if weight:
            line.Weight = weight
        if dash_style:
            dash_map = {"solid": 1, "dash": 4, "dash_dot": 5, "square_dot": 2}
            line.DashStyle = dash_map.get(dash_style, 1)
        if arrow_begin:
            line.BeginArrowheadStyle = ARROW_STYLES.get(arrow_begin, -2)
        if arrow_end:
            line.EndArrowheadStyle = ARROW_STYLES.get(arrow_end, 1)
        return f"format_connector: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("format_connector", str(e))


# ============ Groups 取消组合 / 获取组内项目 ============

def _ungroup_ppt_shapes(presentation: Any, op: dict) -> str:
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)
        # msoGroup = 6，只有组合形状才能取消分组
        if shape.Type != 6:
            return f"ungroup_skipped: 形状不是组合形状 (Type={shape.Type})，无法取消分组，slide={slide_index}, shape={shape_index}"
        shape.Ungroup()
        return f"ungrouped: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("ungroup_shapes", str(e))


def _get_ppt_group_items(presentation: Any, op: dict) -> dict:
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)
        items = []
        count = shape.GroupItems.Count
        for i in range(1, count + 1):
            item = shape.GroupItems(i)
            items.append({"index": i, "name": item.Name, "type": item.Type})
        return {"count": count, "items": items}
    except Exception as e:
        raise COMOperationError("get_group_items", str(e))


# ============ Hyperlinks 超链接功能 ============

def _add_ppt_hyperlink(presentation: Any, op: dict) -> str:
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    url = op.get("url", "")
    sub_address = op.get("sub_address", "")

    if not url and not sub_address:
        raise COMOperationError("add_hyperlink", "需要提供 url 或 sub_address")
    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)
        action = shape.ActionSettings(1)
        action.Action = 7
        if url:
            action.Hyperlink.Address = url
        if sub_address:
            action.Hyperlink.SubAddress = sub_address
        return f"add_hyperlink: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("add_hyperlink", str(e))


def _get_ppt_hyperlinks(presentation: Any, op: dict) -> list:
    slide_index = op.get("slide_index", 1)
    try:
        slide = presentation.Slides(slide_index)
        hyperlinks = []
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes(i)
            try:
                action = shape.ActionSettings(1)
                if action.Action == 7:
                    hlink = action.Hyperlink
                    hyperlinks.append({
                        "shape_index": i,
                        "shape_name": shape.Name,
                        "url": hlink.Address,
                        "sub_address": hlink.SubAddress,
                    })
            except Exception:
                pass
        return hyperlinks
    except Exception as e:
        raise COMOperationError("get_hyperlinks", str(e))


def _remove_ppt_hyperlink(presentation: Any, op: dict) -> str:
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    try:
        slide = presentation.Slides(slide_index)
        shape = slide.Shapes(shape_index)
        action = shape.ActionSettings(1)
        action.Action = 1
        action.Hyperlink.Address = ""
        action.Hyperlink.SubAddress = ""
        return f"remove_hyperlink: slide={slide_index}, shape={shape_index}"
    except Exception as e:
        raise COMOperationError("remove_hyperlink", str(e))


