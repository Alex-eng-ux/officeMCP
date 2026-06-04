"""路径安全校验."""

import os
from pathlib import Path

from office_mcp.config import get_allowed_dirs
from office_mcp.core.errors import PathNotAllowedError

# 禁止访问的系统目录
BLOCKED_PREFIXES = [
    r"C:\Windows",
    r"C:\Program Files",
    r"C:\Program Files (x86)",
    r"C:\ProgramData",
    r"C:\$Recycle.Bin",
]


def validate_path(file_path: str) -> Path:
    """校验文件路径是否允许操作.

    Args:
        file_path: 文件路径

    Returns:
        解析后的绝对路径

    Raises:
        PathNotAllowedError: 路径不允许操作
    """
    if not file_path:
        raise PathNotAllowedError("空路径")

    # 展开环境变量
    expanded = os.path.expandvars(file_path)
    path = Path(expanded).resolve()

    # 必须是绝对路径
    if not path.is_absolute():
        raise PathNotAllowedError(str(path))

    path_str = str(path)

    # 检查系统目录
    for blocked in BLOCKED_PREFIXES:
        if path_str.lower().startswith(blocked.lower()):
            raise PathNotAllowedError(str(path))

    # 检查允许目录
    allowed_dirs = get_allowed_dirs()
    if not allowed_dirs:
        raise PathNotAllowedError(str(path))

    for allowed in allowed_dirs:
        try:
            path.relative_to(allowed)
            return path
        except ValueError:
            continue

    raise PathNotAllowedError(str(path))


def validate_paths(*paths: str) -> list[Path]:
    """批量校验路径.

    Args:
        paths: 多个文件路径

    Returns:
        解析后的绝对路径列表
    """
    return [validate_path(p) for p in paths if p]
