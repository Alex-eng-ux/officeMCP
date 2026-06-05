"""Path validation helpers for Office MCP."""

from __future__ import annotations

import os
from pathlib import Path

from office_mcp.config import get_allowed_dirs
from office_mcp.core.errors import PathNotAllowedError

BLOCKED_PREFIXES = [
    r"C:\Windows",
    r"C:\Program Files",
    r"C:\Program Files (x86)",
    r"C:\ProgramData",
    r"C:\$Recycle.Bin",
]


def _allowed_dir_strings() -> list[str]:
    """Return allowed directories as normalized strings for diagnostics."""
    return [str(path) for path in get_allowed_dirs()]


def validate_path(file_path: str) -> Path:
    """Validate that a path is safe and inside the allowed scope."""
    if not file_path:
        raise PathNotAllowedError("<empty>", _allowed_dir_strings())

    expanded = os.path.expanduser(os.path.expandvars(file_path))
    raw_path = Path(expanded)
    if not raw_path.is_absolute():
        raise PathNotAllowedError(file_path, _allowed_dir_strings())

    path = raw_path.resolve()
    path_str = str(path)

    for blocked in BLOCKED_PREFIXES:
        if path_str.lower().startswith(blocked.lower()):
            raise PathNotAllowedError(str(path), _allowed_dir_strings())

    allowed_dirs = get_allowed_dirs()
    if not allowed_dirs:
        raise PathNotAllowedError(str(path), [])

    for allowed in allowed_dirs:
        try:
            path.relative_to(allowed)
            return path
        except ValueError:
            continue

    raise PathNotAllowedError(str(path), [str(item) for item in allowed_dirs])


def validate_paths(*paths: str) -> list[Path]:
    """Validate multiple paths."""
    return [validate_path(path) for path in paths if path]
