"""Google Material Icons 搜索和下载."""
import json
import logging
import os
import re
import tempfile
import urllib.request
from pathlib import Path

logger = logging.getLogger(__name__)

ICON_NAME_RE = re.compile(r'^[a-z0-9_]+$')

MATERIAL_ICONS_API = "https://fonts.google.com/metadata/icons"


def search_icons(query: str, limit: int = 10) -> list[dict]:
    """搜索 Google Material Icons.

    Args:
        query: 搜索关键词
        limit: 返回结果数量上限

    Returns:
        匹配的图标列表，每项包含 name 和 version
    """
    try:
        req = urllib.request.Request(MATERIAL_ICONS_API)
        with urllib.request.urlopen(req, timeout=10) as resp:
            text = resp.read().decode("utf-8")
            # 响应格式: )]}'\n{"kind":"webfonts#metadata","items":[...]}
            json_text = text.split("\n", 1)[1]
            data = json.loads(json_text)
            icons = data.get("items", [])
            # 过滤匹配的图标
            matched = [
                {"name": ic["name"], "version": ic.get("version", 0)}
                for ic in icons
                if query.lower() in ic["name"].lower()
            ]
            return matched[:limit]
    except Exception as e:
        logger.warning(f"搜索 Material Icons 失败: {e}")
        return []


def get_icon_svg(name: str, fill_color: str = "#000000", size: int = 24) -> str:
    """获取指定图标的 SVG 内容.

    Args:
        name: 图标名称 (如 "star", "check", "home")
        fill_color: 填充颜色 (#RRGGBB)
        size: 尺寸 (px)

    Returns:
        SVG 字符串，失败返回空字符串
    """
    if not ICON_NAME_RE.match(name):
        raise ValueError(f"非法图标名称: {name}，仅允许小写字母、数字和下划线")
    svg_url = (
        f"https://fonts.gstatic.com/s/i/short-term/release/"
        f"materialsymbolsoutlined/{name}/default/{size}px.svg"
    )
    try:
        req = urllib.request.Request(svg_url)
        with urllib.request.urlopen(req, timeout=10) as resp:
            svg = resp.read().decode("utf-8")
            # 替换 fill color
            if fill_color:
                svg = svg.replace('fill="currentColor"', f'fill="{fill_color}"')
                svg = svg.replace("fill:currentColor", f"fill:{fill_color}")
            return svg
    except Exception as e:
        logger.warning(f"获取图标 SVG 失败: {name}: {e}")
        return ""


def get_icon_url(name: str, size: int = 24) -> str:
    """获取图标 PNG 下载 URL.

    Args:
        name: 图标名称
        size: 图标尺寸 (px)

    Returns:
        PNG 下载 URL
    """
    if not ICON_NAME_RE.match(name):
        raise ValueError(f"非法图标名称: {name}，仅允许小写字母、数字和下划线")
    return (
        f"https://fonts.gstatic.com/s/i/short-term/release/"
        f"materialsymbolsoutlined/{name}/default/{size}px.png"
    )
