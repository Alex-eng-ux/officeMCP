from __future__ import annotations

import sys
from pathlib import Path

import office_mcp

if sys.version_info >= (3, 11):
    import tomllib
else:
    import tomli as tomllib


ROOT = Path(__file__).resolve().parents[1]


def load_project_metadata() -> dict:
    return tomllib.loads((ROOT / "pyproject.toml").read_text(encoding="utf-8"))


def test_package_version_matches_pyproject() -> None:
    project = load_project_metadata()["project"]
    assert office_mcp.__version__ == project["version"]


def test_project_urls_use_real_repository() -> None:
    urls = load_project_metadata()["project"]["urls"]
    for url in urls.values():
        assert "yourusername" not in url
        assert "Alex-eng-ux/officeMCP" in url


def test_console_script_targets_server_main() -> None:
    scripts = load_project_metadata()["project"]["scripts"]
    assert scripts["office-mcp"] == "office_mcp.server:main"
