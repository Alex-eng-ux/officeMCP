from __future__ import annotations

from office_mcp.core.errors import COMOperationError
from office_mcp.operations.excel_ops import (
    _add_data_validation,
    _add_slicer,
    _create_pivot_table,
    _goal_seek,
    _protect_worksheet,
)


class _FakeValidation:
    def __init__(self) -> None:
        self.deleted = False
        self.add_calls: list[dict[str, object]] = []
        self.IgnoreBlank = None
        self.InCellDropdown = None
        self.ShowInput = None
        self.InputTitle = ""
        self.InputMessage = ""
        self.ShowError = None
        self.ErrorTitle = ""
        self.ErrorMessage = ""

    def Delete(self) -> None:  # noqa: N802
        self.deleted = True

    def Add(self, **kwargs) -> None:  # noqa: N802
        self.add_calls.append(kwargs)


class _FakeRange:
    def __init__(self, address: str, worksheet=None) -> None:
        self.Address = address
        self.Worksheet = worksheet
        self.Validation = _FakeValidation()

    def GoalSeek(self, Goal, ChangingCell):  # noqa: N802, ANN001
        self.goal_seek_call = {"Goal": Goal, "ChangingCell": ChangingCell}
        return True


class _FakeSheet:
    def __init__(self, name: str = "Sheet1") -> None:
        self.Name = name
        self.ranges: dict[str, _FakeRange] = {}
        self.protect_calls: list[dict[str, object]] = []
        self.UsedRange = type("UsedRange", (), {"Address": "$A$1:$F$10"})()

    def Range(self, address: str):  # noqa: N802
        if address not in self.ranges:
            self.ranges[address] = _FakeRange(address, self)
        return self.ranges[address]

    def Protect(self, **kwargs) -> None:  # noqa: N802
        self.protect_calls.append(kwargs)

    def PivotTables(self):  # noqa: N802
        return self._pivot_tables


class _FakeWorksheets:
    def __init__(self, sheets: dict[str, _FakeSheet]) -> None:
        self._sheets = sheets
        self._added_count = 0

    def __call__(self, name: str):
        if name not in self._sheets:
            raise KeyError(name)
        return self._sheets[name]

    def Add(self):  # noqa: N802
        self._added_count += 1
        sheet = _FakeSheet(f"Added{self._added_count}")
        self._sheets[sheet.Name] = sheet
        return sheet


def test_add_data_validation_sets_extended_options() -> None:
    sheet = _FakeSheet()
    workbook = type("Workbook", (), {"Worksheets": _FakeWorksheets({"Sheet1": sheet})})()

    result = _add_data_validation(
        workbook,
        {
            "sheet": "Sheet1",
            "range": "A1:A10",
            "type": "list",
            "formula1": "Yes,No",
            "ignore_blank": False,
            "in_cell_dropdown": False,
            "show_input": True,
            "input_title": "Allowed",
            "input_message": "Choose one",
            "show_error": True,
            "error_title": "Invalid",
            "error_message": "Must match the list",
            "error_style": "warning",
        },
    )

    validation = sheet.Range("A1:A10").Validation
    assert result == "added_data_validation: A1:A10 (list)"
    assert validation.deleted is True
    assert validation.add_calls == [{"Type": 3, "AlertStyle": 2, "Formula1": "Yes,No"}]
    assert validation.IgnoreBlank is False
    assert validation.InCellDropdown is False
    assert validation.InputTitle == "Allowed"
    assert validation.ErrorMessage == "Must match the list"


def test_add_data_validation_requires_formula2_for_between() -> None:
    sheet = _FakeSheet()
    workbook = type("Workbook", (), {"Worksheets": _FakeWorksheets({"Sheet1": sheet})})()

    try:
        _add_data_validation(
            workbook,
            {
                "sheet": "Sheet1",
                "range": "B1:B10",
                "type": "whole",
                "formula1": "1",
                "operator": "between",
            },
        )
    except COMOperationError as exc:
        assert "formula2 is required" in str(exc)
    else:
        raise AssertionError("between validation should require formula2")


def test_create_pivot_table_uses_used_range_when_source_range_empty() -> None:
    source_sheet = _FakeSheet("Raw")
    target_sheet = _FakeSheet("Pivot")
    sheets = {"Raw": source_sheet, "Pivot": target_sheet}

    class FakePivotField:
        def __init__(self, name: str) -> None:
            self.name = name
            self.Orientation = None
            self.Position = None

    class FakePivotTable:
        def __init__(self) -> None:
            self.fields: dict[str, FakePivotField] = {}
            self.data_fields: list[tuple[str, str, int]] = []
            self.TableStyle2 = None

        def PivotFields(self, name: str):  # noqa: N802
            self.fields.setdefault(name, FakePivotField(name))
            return self.fields[name]

        def AddDataField(self, field, caption: str, function: int) -> None:  # noqa: N802, ANN001
            self.data_fields.append((field.name, caption, function))

    class FakePivotCache:
        def __init__(self) -> None:
            self.create_table_call = None
            self.pivot_table = FakePivotTable()

        def CreatePivotTable(self, **kwargs):  # noqa: N802
            self.create_table_call = kwargs
            return self.pivot_table

    class FakePivotCaches:
        def __init__(self) -> None:
            self.create_call = None
            self.cache = FakePivotCache()

        def Create(self, **kwargs):  # noqa: N802
            self.create_call = kwargs
            return self.cache

    pivot_caches = FakePivotCaches()
    workbook = type(
        "Workbook",
        (),
        {
            "Worksheets": _FakeWorksheets(sheets),
            "PivotCaches": pivot_caches,
        },
    )()

    result = _create_pivot_table(
        workbook,
        {
            "source_sheet": "Raw",
            "source_range": "",
            "target_sheet": "Pivot",
            "target_cell": "A3",
            "row_fields": ["Region"],
            "column_fields": ["Month"],
            "filter_fields": ["Year"],
            "data_fields": {"Sales": "sum"},
            "pivot_name": "PivotSales",
            "style_name": "PivotStyleLight16",
        },
    )

    assert result == "created_pivot_table: Pivot!A3 (PivotSales)"
    assert pivot_caches.create_call == {"SourceType": 1, "SourceData": "'Raw'!$A$1:$F$10"}
    pivot_table = pivot_caches.cache.pivot_table
    assert pivot_table.fields["Region"].Orientation == 1
    assert pivot_table.fields["Month"].Orientation == 2
    assert pivot_table.fields["Year"].Orientation == 3
    assert pivot_table.data_fields == [("Sales", "sum_Sales", -4157)]
    assert pivot_table.TableStyle2 == "PivotStyleLight16"


def test_add_slicer_falls_back_to_add_when_add2_unavailable() -> None:
    target_sheet = _FakeSheet("Dashboard")
    pivot_sheet = _FakeSheet("Pivot")
    pivot_table = type("PivotTable", (), {"Name": "PivotTable_1"})()
    pivot_sheet._pivot_tables = type(
        "PivotTables",
        (),
        {"Count": 1, "__call__": lambda self, index: pivot_table},
    )()

    class FakeSlicer:
        pass

    class FakeSlicers:
        def __init__(self) -> None:
            self.calls: list[dict[str, object]] = []

        def Add(self, **kwargs):  # noqa: N802
            self.calls.append(kwargs)
            return FakeSlicer()

    class FakeSlicerCache:
        def __init__(self) -> None:
            self.Slicers = FakeSlicers()

    class FakeSlicerCaches:
        Count = 0

        def Add2(self, pivot_table, field_name):  # noqa: N802, ANN001
            raise RuntimeError("Add2 unsupported")

        def Add(self, pivot_table, field_name):  # noqa: N802, ANN001
            self.add_call = (pivot_table, field_name)
            return FakeSlicerCache()

    workbook = type(
        "Workbook",
        (),
        {
            "Worksheets": _FakeWorksheets({"Dashboard": target_sheet, "Pivot": pivot_sheet}),
            "SlicerCaches": FakeSlicerCaches(),
        },
    )()

    result = _add_slicer(
        workbook,
        {
            "target_sheet": "Dashboard",
            "pivot_sheet": "Pivot",
            "pivot_name": "PivotTable_1",
            "field_name": "Region",
            "left": 20,
            "top": 30,
            "width": 240,
            "height": 180,
        },
    )

    assert result == "added_slicer: Region"
    assert workbook.SlicerCaches.add_call == (pivot_table, "Region")


def test_protect_worksheet_passes_granular_flags() -> None:
    sheet = _FakeSheet()
    workbook = type("Workbook", (), {"Worksheets": _FakeWorksheets({"Sheet1": sheet})})()

    result = _protect_worksheet(
        workbook,
        {
            "sheet": "Sheet1",
            "password": "secret",
            "user_interface_only": True,
            "allow_filtering": True,
            "allow_sorting": True,
            "allow_using_pivot_tables": True,
        },
    )

    assert result == "protected_worksheet: Sheet1"
    assert sheet.protect_calls == [{
        "Password": "secret",
        "DrawingObjects": True,
        "Contents": True,
        "Scenarios": True,
        "UserInterfaceOnly": True,
        "AllowFormattingCells": False,
        "AllowFormattingColumns": False,
        "AllowFormattingRows": False,
        "AllowInsertingColumns": False,
        "AllowInsertingRows": False,
        "AllowInsertingHyperlinks": False,
        "AllowDeletingColumns": False,
        "AllowDeletingRows": False,
        "AllowSorting": True,
        "AllowFiltering": True,
        "AllowUsingPivotTables": True,
    }]


def test_goal_seek_calls_native_excel_api() -> None:
    sheet = _FakeSheet()
    workbook = type("Workbook", (), {"Worksheets": _FakeWorksheets({"Model": sheet})})()

    result = _goal_seek(
        workbook,
        {
            "sheet": "Model",
            "set_cell": "D10",
            "goal": 100.0,
            "changing_cell": "B2",
        },
    )

    assert result == "goal_seek: set_cell=D10, changing_cell=B2, converged=True"
    assert sheet.Range("D10").goal_seek_call == {"Goal": 100.0, "ChangingCell": sheet.Range("B2")}
