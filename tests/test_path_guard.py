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

    with pytest.raises(PathNotAllowedError):
        path_guard.validate_path("slides.pptx")


def test_validate_path_rejects_empty_value(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [Path.cwd().resolve()])

    with pytest.raises(PathNotAllowedError):
        path_guard.validate_path("")


@pytest.mark.skipif(Path("C:/").drive.upper() != "C:", reason="Windows path guard check")
def test_validate_path_blocks_windows_system_directory(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [Path("C:/").resolve()])

    with pytest.raises(PathNotAllowedError):
        path_guard.validate_path(r"C:\Windows\System32\drivers\etc\hosts")
