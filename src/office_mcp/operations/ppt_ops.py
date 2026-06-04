"""PowerPoint COM 操作实现."""

import logging
from pathlib import Path
from typing import Any

from office_mcp.core.errors import COMOperationError

logger = logging.getLogger(__name__)

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
            if shape.PlaceholderFormat.Type in (1, 3):  # ppPlaceholderTitle / ppPlaceholderCenterTitle
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
    """
    slide_index = op.get("slide_index", 1)
    shape_index = op.get("shape_index", 1)
    animation_type = op.get("animation_type", "fade")

    slide = presentation.Slides(slide_index)
    shape = slide.Shapes(shape_index)

    effect_val = ANIMATION_EFFECTS.get(animation_type, 3852)  # 默认 fade

    # 添加动画
    effect = slide.TimeLine.MainSequence.AddEffect(
        shape,
        EffectId=effect_val,
        Level=0,  # msoAnimateLevel.msoAnimateLevelNone
    )

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
    master.FollowMasterBackground = False  # 确保不跟随主题
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
