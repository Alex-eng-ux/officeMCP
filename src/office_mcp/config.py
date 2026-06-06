"""Configuration management for Office MCP."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

_WORKSPACE_HINT_VARS = (
    "OFFICE_MCP_WORKSPACE_DIRS",
    "PROJECT_ROOT",
    "WORKSPACE_ROOT",
    "INIT_CWD",
    "PWD",
    "CLAUDE_PROJECT_DIR",
    "CURSOR_WORKSPACE_ROOT",
    "VSCODE_CWD",
)


def _get_env_bool(name: str, default: bool) -> bool:
    """Read a boolean environment variable."""
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


@dataclass
class Settings:
    """Office MCP server configuration."""

    allowed_dirs: str
    auto_discover_dirs: bool
    default_overwrite: bool
    backup_before_edit: bool
    visible: bool
    word_visible: bool
    excel_visible: bool
    ppt_visible: bool


def _load_visible_setting(app_name: str, default: bool) -> bool:
    """Read per-app visibility with a shared fallback."""
    specific_name = f"OFFICE_MCP_{app_name.upper()}_VISIBLE"
    if os.environ.get(specific_name) is not None:
        return _get_env_bool(specific_name, default)
    return _get_env_bool("OFFICE_MCP_VISIBLE", default)


def _load_settings() -> Settings:
    """Load settings from the environment."""
    shared_visible = _get_env_bool("OFFICE_MCP_VISIBLE", True)
    explicit_allowed_dirs = os.environ.get("OFFICE_MCP_ALLOWED_DIRS")
    return Settings(
        allowed_dirs=explicit_allowed_dirs or "",
        auto_discover_dirs=_get_env_bool(
            "OFFICE_MCP_AUTO_DISCOVER_DIRS",
            explicit_allowed_dirs is None,
        ),
        default_overwrite=_get_env_bool("OFFICE_MCP_DEFAULT_OVERWRITE", False),
        backup_before_edit=_get_env_bool("OFFICE_MCP_BACKUP_BEFORE_EDIT", True),
        visible=shared_visible,
        word_visible=_load_visible_setting("word", True),
        excel_visible=_load_visible_setting("excel", True),
        ppt_visible=_load_visible_setting("ppt", True),
    )


settings = _load_settings()


def _split_dir_list(raw_value: str) -> list[str]:
    """Split a semicolon-separated path list into trimmed entries."""
    return [part.strip() for part in raw_value.split(";") if part.strip()]


def _iter_discovered_dirs() -> list[Path]:
    """Collect common workspace and user-facing directories."""
    candidates: list[str] = []

    user_profile = os.environ.get("USERPROFILE", "").strip()
    if user_profile:
        candidates.extend(
            [
                user_profile,
                os.path.join(user_profile, "Documents"),
                os.path.join(user_profile, "Desktop"),
            ]
        )

    for env_name in _WORKSPACE_HINT_VARS:
        raw_value = os.environ.get(env_name, "")
        if raw_value:
            candidates.extend(_split_dir_list(raw_value))

    try:
        candidates.append(str(Path.cwd()))
    except Exception:
        pass

    resolved: list[Path] = []
    seen: set[str] = set()
    for candidate in candidates:
        expanded = os.path.expanduser(os.path.expandvars(candidate))
        if not expanded:
            continue
        try:
            path = Path(expanded).resolve()
        except Exception:
            continue
        key = os.path.normcase(str(path))
        if key in seen:
            continue
        seen.add(key)
        resolved.append(path)
    return resolved


def get_allowed_dirs() -> list[Path]:
    """Return the effective list of allowed directories."""
    configured: list[Path] = []
    for raw_dir in _split_dir_list(settings.allowed_dirs):
        expanded = os.path.expanduser(os.path.expandvars(raw_dir))
        try:
            configured.append(Path(expanded).resolve())
        except Exception:
            continue

    sources = configured
    if not configured:
        sources = _iter_discovered_dirs()
    elif settings.auto_discover_dirs:
        sources = [*configured, *_iter_discovered_dirs()]

    combined: list[Path] = []
    seen: set[str] = set()
    for path in sources:
        key = os.path.normcase(str(path))
        if key in seen:
            continue
        seen.add(key)
        combined.append(path)
    return combined
