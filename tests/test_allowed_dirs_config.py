from __future__ import annotations

from importlib import reload
from pathlib import Path


def _reload_config_module():
    import office_mcp.config as config_module

    return reload(config_module)


def test_visibility_defaults_to_visible_for_all_apps(monkeypatch) -> None:
    monkeypatch.delenv("OFFICE_MCP_VISIBLE", raising=False)
    monkeypatch.delenv("OFFICE_MCP_WORD_VISIBLE", raising=False)
    monkeypatch.delenv("OFFICE_MCP_EXCEL_VISIBLE", raising=False)
    monkeypatch.delenv("OFFICE_MCP_PPT_VISIBLE", raising=False)

    config_module = _reload_config_module()

    assert config_module.settings.visible is True
    assert config_module.settings.word_visible is True
    assert config_module.settings.excel_visible is True
    assert config_module.settings.ppt_visible is True


def test_visibility_allows_app_specific_overrides(monkeypatch) -> None:
    monkeypatch.setenv("OFFICE_MCP_VISIBLE", "true")
    monkeypatch.setenv("OFFICE_MCP_WORD_VISIBLE", "false")
    monkeypatch.setenv("OFFICE_MCP_EXCEL_VISIBLE", "true")
    monkeypatch.setenv("OFFICE_MCP_PPT_VISIBLE", "false")

    config_module = _reload_config_module()

    assert config_module.settings.visible is True
    assert config_module.settings.word_visible is False
    assert config_module.settings.excel_visible is True
    assert config_module.settings.ppt_visible is False


def test_default_allowed_dirs_fall_back_to_discovered_workspace_paths(
    tmp_path: Path, monkeypatch
) -> None:
    workspace_root = (tmp_path / "workspace").resolve()
    monkeypatch.delenv("OFFICE_MCP_ALLOWED_DIRS", raising=False)
    monkeypatch.delenv("OFFICE_MCP_AUTO_DISCOVER_DIRS", raising=False)
    monkeypatch.setenv("PROJECT_ROOT", str(workspace_root))

    config_module = _reload_config_module()
    allowed_dirs = config_module.get_allowed_dirs()

    assert config_module.settings.allowed_dirs == ""
    assert config_module.settings.auto_discover_dirs is True
    assert workspace_root in allowed_dirs


def test_explicit_allowed_dirs_disable_auto_discovery_by_default(
    tmp_path: Path, monkeypatch
) -> None:
    explicit_root = (tmp_path / "explicit-root").resolve()
    detected_root = (tmp_path / "detected-root").resolve()
    monkeypatch.setenv("OFFICE_MCP_ALLOWED_DIRS", str(explicit_root))
    monkeypatch.delenv("OFFICE_MCP_AUTO_DISCOVER_DIRS", raising=False)
    monkeypatch.setenv("PROJECT_ROOT", str(detected_root))

    config_module = _reload_config_module()
    allowed_dirs = config_module.get_allowed_dirs()

    assert config_module.settings.allowed_dirs == str(explicit_root)
    assert config_module.settings.auto_discover_dirs is False
    assert explicit_root in allowed_dirs
    assert detected_root not in allowed_dirs


def test_get_allowed_dirs_supports_multiple_project_roots(tmp_path: Path, monkeypatch) -> None:
    config_module = _reload_config_module()
    workspace_root = (tmp_path / "workspace").resolve()
    project_root = (tmp_path / "project").resolve()
    monkeypatch.setattr(
        config_module.settings,
        "allowed_dirs",
        f"{workspace_root}; {project_root}",
    )

    allowed_dirs = config_module.get_allowed_dirs()

    assert workspace_root in allowed_dirs
    assert project_root in allowed_dirs


def test_get_allowed_dirs_prefers_explicit_roots_without_auto_discovery(
    tmp_path: Path, monkeypatch
) -> None:
    config_module = _reload_config_module()
    explicit_root = (tmp_path / "explicit-root").resolve()
    monkeypatch.setattr(config_module.settings, "allowed_dirs", str(explicit_root))
    monkeypatch.setattr(config_module.settings, "auto_discover_dirs", False)
    monkeypatch.setenv("PROJECT_ROOT", str((tmp_path / "detected-root").resolve()))

    allowed_dirs = config_module.get_allowed_dirs()

    assert explicit_root in allowed_dirs
    assert len(allowed_dirs) == 1


def test_get_allowed_dirs_can_merge_detected_roots_when_enabled(
    tmp_path: Path, monkeypatch
) -> None:
    config_module = _reload_config_module()
    explicit_root = (tmp_path / "explicit-root").resolve()
    detected_root = (tmp_path / "detected-root").resolve()
    monkeypatch.setattr(config_module.settings, "allowed_dirs", str(explicit_root))
    monkeypatch.setattr(config_module.settings, "auto_discover_dirs", True)
    monkeypatch.setenv("PROJECT_ROOT", str(detected_root))

    allowed_dirs = config_module.get_allowed_dirs()

    assert explicit_root in allowed_dirs
    assert detected_root in allowed_dirs


def test_get_allowed_dirs_expands_environment_variables_for_project_roots(
    tmp_path: Path, monkeypatch
) -> None:
    config_module = _reload_config_module()
    workspace_root = (tmp_path / "workspace").resolve()
    project_root = workspace_root / "client-a"
    monkeypatch.setenv("OFFICE_MCP_TEST_ROOT", str(workspace_root))
    monkeypatch.setattr(
        config_module.settings,
        "allowed_dirs",
        r"%OFFICE_MCP_TEST_ROOT%;%OFFICE_MCP_TEST_ROOT%\client-a",
    )

    allowed_dirs = config_module.get_allowed_dirs()

    assert workspace_root in allowed_dirs
    assert project_root.resolve() in allowed_dirs
