from __future__ import annotations

import json
from pathlib import Path

import pytest

from office_mcp.tools import excel as excel_tools


class FakeMCP:
    def __init__(self) -> None:
        self.tools: dict[str, object] = {}

    def tool(self):  # noqa: ANN202
        def decorator(func):
            self.tools[func.__name__] = func
            return func

        return decorator


@pytest.fixture
def excel_tool_map() -> dict[str, object]:
    fake_mcp = FakeMCP()
    excel_tools.register_excel_tools(fake_mcp)
    return fake_mcp.tools


def _stub_validate_path(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(excel_tools, "validate_path", lambda value: Path(value))


def _fail_get_document(*args, **kwargs):  # noqa: ANN002, ANN003
    raise AssertionError("tool should use office_manager.ensure_document instead of get_document")


def test_excel_apply_operations_uses_recovery_path_for_live_workbook(
    excel_tool_map: dict[str, object], monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook = object()
    workbook_path = tmp_path / "sheet.xlsx"
    calls: list[tuple[Path, bool]] = []

    _stub_validate_path(monkeypatch)
    monkeypatch.setattr(excel_tools.office_manager, "get_document", _fail_get_document)
    monkeypatch.setattr(
        excel_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or workbook,
    )
    monkeypatch.setattr(
        excel_tools,
        "apply_excel_operations",
        lambda wb, operations: {"workbook_id": id(wb), "operations": operations},
    )

    result = excel_tool_map["excel_apply_operations"](str(workbook_path), [{"type": "save"}])

    assert calls == [(workbook_path, True)]
    assert result == {
        "file_path": str(workbook_path),
        "results": {"workbook_id": id(workbook), "operations": [{"type": "save"}]},
    }


def test_excel_check_typography_uses_recovery_path(
    excel_tool_map: dict[str, object], monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook = object()
    workbook_path = tmp_path / "sheet.xlsx"
    calls: list[tuple[Path, bool]] = []

    _stub_validate_path(monkeypatch)
    monkeypatch.setattr(excel_tools.office_manager, "get_document", _fail_get_document)
    monkeypatch.setattr(
        excel_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or workbook,
    )
    monkeypatch.setattr(
        excel_tools,
        "_check_typography",
        lambda wb, options: [{"sheet": options["sheet"], "workbook_id": id(wb)}],
    )

    result = excel_tool_map["excel_check_typography"](str(workbook_path), sheet="Summary")

    assert calls == [(workbook_path, False)]
    assert result == {
        "file_path": str(workbook_path),
        "sheet": "Summary",
        "issue_count": 1,
        "issues": [{"sheet": "Summary", "workbook_id": id(workbook)}],
    }


@pytest.mark.parametrize(
    ("tool_name", "operation_name", "tool_args", "operation_result", "expected_result"),
    [
        (
            "excel_list_worksheets",
            "_list_worksheets",
            (),
            ["Sheet1", "Sheet2"],
            {"file_path": None, "count": 2, "worksheets": ["Sheet1", "Sheet2"]},
        ),
        (
            "excel_get_worksheet_info",
            "_get_worksheet_info",
            ("Summary",),
            {"name": "Summary", "used_range": "A1:C9"},
            {"file_path": None, "name": "Summary", "used_range": "A1:C9"},
        ),
        (
            "excel_list_used_range",
            "_list_used_range",
            ("Summary",),
            {"address": "A1:C9", "rows": 9, "columns": 3},
            {"file_path": None, "sheet": "Summary", "address": "A1:C9", "rows": 9, "columns": 3},
        ),
    ],
)
def test_excel_read_list_tools_use_recovery_path(
    excel_tool_map: dict[str, object],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    tool_name: str,
    operation_name: str,
    tool_args: tuple[object, ...],
    operation_result: object,
    expected_result: dict[str, object],
) -> None:
    workbook = object()
    workbook_path = tmp_path / "sheet.xlsx"
    calls: list[tuple[Path, bool]] = []

    _stub_validate_path(monkeypatch)
    monkeypatch.setattr(excel_tools.office_manager, "get_document", _fail_get_document)
    monkeypatch.setattr(
        excel_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or workbook,
    )
    monkeypatch.setattr(excel_tools, operation_name, lambda wb, options: operation_result)

    result = excel_tool_map[tool_name](str(workbook_path), *tool_args)

    assert calls == [(workbook_path, False)]
    assert result == {**expected_result, "file_path": str(workbook_path)}


def test_excel_list_tables_uses_recovery_path(
    excel_tool_map: dict[str, object], monkeypatch: pytest.MonkeyPatch, tmp_path: Path
) -> None:
    workbook = object()
    workbook_path = tmp_path / "sheet.xlsx"
    calls: list[tuple[Path, bool]] = []

    _stub_validate_path(monkeypatch)
    monkeypatch.setattr(excel_tools.office_manager, "get_document", _fail_get_document)
    monkeypatch.setattr(
        excel_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or workbook,
    )
    monkeypatch.setattr(
        excel_tools,
        "_list_tables",
        lambda wb, options: [{"name": "SalesTable", "sheet": options["sheet"], "workbook_id": id(wb)}],
    )

    result = excel_tool_map["excel_list_tables"](str(workbook_path), sheet="Summary")

    assert calls == [(workbook_path, False)]
    assert json.loads(result) == [{"name": "SalesTable", "sheet": "Summary", "workbook_id": id(workbook)}]


@pytest.mark.parametrize(
    ("tool_name", "tool_kwargs", "expected_activate", "stub_name", "expected_fragment"),
    [
        (
            "excel_add_data_validation",
            {
                "sheet": "Summary",
                "range": "A2:A20",
                "validation_type": "whole",
                "formula1": "1",
                "operator": "greater_equal",
                "formula2": "",
            },
            False,
            "_add_data_validation",
            "A2:A20",
        ),
        (
            "excel_add_slicer",
            {
                "target_sheet": "Dashboard",
                "pivot_sheet": "Pivot",
                "field_name": "Region",
                "pivot_name": "PivotTable_1",
            },
            False,
            "_add_slicer",
            "Region",
        ),
        (
            "excel_create_pivot_table",
            {
                "source_sheet": "Raw",
                "source_range": "A1:D20",
                "target_sheet": "Pivot",
                "target_cell": "A3",
                "row_fields": ["Region"],
                "column_fields": ["Month"],
                "data_fields": {"Sales": "sum"},
            },
            True,
            "_create_pivot_table",
            "workbook_id",
        ),
        (
            "excel_set_view_zoom",
            {"sheet": "Summary", "zoom": 125},
            True,
            "_set_view_zoom",
            "workbook_id",
        ),
        (
            "excel_set_view_gridlines",
            {"sheet": "Summary", "show": False},
            True,
            "_set_view_gridlines",
            "workbook_id",
        ),
        (
            "excel_set_view_headings",
            {"sheet": "Summary", "show": False},
            True,
            "_set_view_headings",
            "workbook_id",
        ),
    ],
)
def test_excel_advanced_tools_route_through_ensure_document(
    excel_tool_map: dict[str, object],
    monkeypatch: pytest.MonkeyPatch,
    tmp_path: Path,
    tool_name: str,
    tool_kwargs: dict[str, object],
    expected_activate: bool,
    stub_name: str,
    expected_fragment: str,
) -> None:
    workbook = object()
    workbook_path = tmp_path / "sheet.xlsx"
    calls: list[tuple[Path, bool]] = []

    _stub_validate_path(monkeypatch)
    monkeypatch.setattr(excel_tools.office_manager, "get_document", _fail_get_document)
    monkeypatch.setattr(
        excel_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or workbook,
    )
    monkeypatch.setattr(excel_tools, stub_name, lambda wb, options: {"workbook_id": id(wb), "options": options})

    result = excel_tool_map[tool_name](str(workbook_path), **tool_kwargs)

    assert calls == [(workbook_path, expected_activate)]
    assert expected_fragment in str(result)
