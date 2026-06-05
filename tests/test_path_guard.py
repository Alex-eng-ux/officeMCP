from __future__ import annotations

from pathlib import Path

import pytest

from office_mcp.core import path_guard
from office_mcp.core.errors import PathNotAllowedError


def test_validate_path_allows_file_inside_allowed_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    allowed_dir = tmp_path.resolve()
    target = allowed_dir / "report.docx"
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [allowed_dir])

    assert path_guard.validate_path(str(target)) == target.resolve()


def test_validate_path_rejects_relative_path(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    allowed_dir = tmp_path.resolve()
    monkeypatch.chdir(allowed_dir)
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [allowed_dir])

    with pytest.raises(PathNotAllowedError) as exc_info:
        path_guard.validate_path("slides.pptx")

    assert "slides.pptx" in str(exc_info.value)
    assert exc_info.value.suggestion


def test_validate_path_rejects_empty_value(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [Path.cwd().resolve()])

    with pytest.raises(PathNotAllowedError) as exc_info:
        path_guard.validate_path("")

    assert exc_info.value.suggestion


def test_validate_path_allows_nested_file_when_project_root_is_allowed(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    project_root = tmp_path.resolve() / "client-project"
    nested_target = project_root / "docs" / "drafts" / "spec.docx"
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [project_root])

    assert path_guard.validate_path(str(nested_target)) == nested_target.resolve()


def test_validate_path_rejects_path_outside_project_roots_with_diagnostics(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    allowed_dir = (tmp_path / "project-a").resolve()
    disallowed_target = (tmp_path / "project-b" / "notes.docx").resolve()
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [allowed_dir])

    with pytest.raises(PathNotAllowedError) as exc_info:
        path_guard.validate_path(str(disallowed_target))

    assert str(disallowed_target) in str(exc_info.value)
    assert "OFFICE_MCP_ALLOWED_DIRS" in exc_info.value.suggestion


@pytest.mark.skipif(Path("C:/").drive.upper() != "C:", reason="Windows path guard check")
def test_validate_path_blocks_windows_system_directory(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [Path("C:/").resolve()])

    with pytest.raises(PathNotAllowedError):
        path_guard.validate_path(r"C:\Windows\System32\drivers\etc\hosts")


def test_validate_path_diagnostics_list_allowed_roots(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    allowed_dirs = [(tmp_path / "workspace").resolve(), (tmp_path / "assets").resolve()]
    outside_target = (tmp_path / "external" / "deck.pptx").resolve()
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: allowed_dirs)

    with pytest.raises(PathNotAllowedError) as exc_info:
        path_guard.validate_path(str(outside_target))

    for allowed_dir in allowed_dirs:
        assert str(allowed_dir) in str(exc_info.value)
