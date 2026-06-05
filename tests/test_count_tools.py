from __future__ import annotations

import pytest

from count_tools import (
    _resolve_tool_registry,
    format_summary,
    summarize_tool_names,
    validate_summary,
)


def test_summarize_tool_names_groups_known_prefixes() -> None:
    summary = summarize_tool_names(
        [
            "excel_open",
            "excel_open",
            "word_open",
            "ppt_open",
            "office_cleanup",
        ]
    )

    assert summary == {
        "Excel": 1,
        "Word": 1,
        "PowerPoint": 1,
        "Other": 1,
        "Total": 4,
    }


def test_format_summary_emits_human_readable_lines() -> None:
    lines = format_summary(
        {
            "Excel": 2,
            "Word": 3,
            "PowerPoint": 4,
            "Other": 1,
            "Total": 10,
        }
    )

    assert lines == [
        "Total: 10",
        "Excel: 2",
        "Word: 3",
        "PowerPoint: 4",
        "Other: 1",
    ]


def test_validate_summary_enforces_thresholds() -> None:
    summary = {
        "Excel": 1,
        "Word": 1,
        "PowerPoint": 0,
        "Other": 0,
        "Total": 2,
    }

    with pytest.raises(ValueError, match="at least 3 tools"):
        validate_summary(summary, min_total=3)

    with pytest.raises(ValueError, match="ppt_"):
        validate_summary(summary, required_prefixes=["ppt_"])


def test_resolve_tool_registry_supports_direct_and_manager_layouts() -> None:
    direct_tools = {"word_open_document": object()}
    managed_tools = {"ppt_open_presentation": object()}

    class DirectMCP:
        tools = direct_tools

    class Manager:
        _tools = managed_tools

    class ManagedMCP:
        _tool_manager = Manager()

    assert _resolve_tool_registry(DirectMCP()) is direct_tools
    assert _resolve_tool_registry(ManagedMCP()) is managed_tools
