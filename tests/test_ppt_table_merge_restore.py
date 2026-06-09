from __future__ import annotations

import json

from office_mcp.operations.ppt_ops import _ppt_merge_table_cells, _ppt_split_table_cells


class FakeTextRange:
    def __init__(self, text: str = "") -> None:
        self.Text = text


class FakeTextFrame:
    def __init__(self, text: str = "") -> None:
        self.TextRange = FakeTextRange(text)


class FakeShapeCell:
    def __init__(self, text: str = "") -> None:
        self.TextFrame = FakeTextFrame(text)


class FakeCell:
    def __init__(self, table: "FakeTable", row: int, col: int, text: str = "") -> None:
        self.table = table
        self.row = row
        self.col = col
        self.Shape = FakeShapeCell(text)
        self.RowSpan = 1
        self.ColSpan = 1

    def Merge(self, other: "FakeCell") -> None:  # noqa: N802
        self.table.last_merge = (self.row, self.col, other.row, other.col)
        start_row, end_row = sorted((self.row, other.row))
        start_col, end_col = sorted((self.col, other.col))
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = self.table.Cell(row, col)
                cell.RowSpan = end_row - start_row + 1
                cell.ColSpan = end_col - start_col + 1

    def Split(self) -> None:  # noqa: N802
        self.table.last_split = (self.row, self.col)
        for cell in self.table._cells.values():
            cell.RowSpan = 1
            cell.ColSpan = 1


class FakeCount:
    def __init__(self, count: int) -> None:
        self.Count = count


class FakeTable:
    def __init__(self, cells: list[list[str]]) -> None:
        self._cells = {
            (row_index + 1, col_index + 1): FakeCell(self, row_index + 1, col_index + 1, text)
            for row_index, row in enumerate(cells)
            for col_index, text in enumerate(row)
        }
        self.Rows = FakeCount(len(cells))
        self.Columns = FakeCount(len(cells[0]))
        self.last_merge: tuple[int, int, int, int] | None = None
        self.last_split: tuple[int, int] | None = None

    def Cell(self, row: int, col: int) -> FakeCell:  # noqa: N802
        return self._cells[(row, col)]


class FakeTags:
    def __init__(self) -> None:
        self.values: dict[str, str] = {}

    def Add(self, name: str, value: str) -> None:  # noqa: N802
        self.values[name] = value

    def Delete(self, name: str) -> None:  # noqa: N802
        self.values.pop(name, None)

    def __call__(self, name: str) -> str:
        return self.values.get(name, "")


class FakeShape:
    def __init__(self, table: FakeTable) -> None:
        self.HasTable = True
        self.Table = table
        self.Tags = FakeTags()


class FakeShapes:
    def __init__(self, shape: FakeShape) -> None:
        self._shape = shape

    def __call__(self, index: int) -> FakeShape:
        assert index == 1
        return self._shape


class FakeSlide:
    def __init__(self, shape: FakeShape) -> None:
        self.Shapes = FakeShapes(shape)


class FakeSlides:
    def __init__(self, slide: FakeSlide) -> None:
        self._slide = slide

    def __call__(self, index: int) -> FakeSlide:
        assert index == 1
        return self._slide


class FakePresentation:
    def __init__(self, cells: list[list[str]]) -> None:
        table = FakeTable(cells)
        shape = FakeShape(table)
        self.shape = shape
        self.table = table
        self.Slides = FakeSlides(FakeSlide(shape))


def test_ppt_merge_table_cells_stores_region_metadata_and_combines_text() -> None:
    presentation = FakePresentation(
        [
            ["R1C1", "R1C2"],
            ["R2C1", "R2C2"],
        ]
    )

    result = _ppt_merge_table_cells(
        presentation,
        {"slide_index": 1, "shape_index": 1, "start_row": 1, "start_col": 1, "end_row": 2, "end_col": 2},
    )

    assert "merge_table_cells" in result
    assert presentation.table.last_merge == (1, 1, 2, 2)
    assert presentation.table.Cell(1, 1).Shape.TextFrame.TextRange.Text == "R1C1\nR1C2\nR2C1\nR2C2"

    metadata = json.loads(presentation.shape.Tags("OfficeMCP.TableMerge.1.1"))
    assert metadata["end_row"] == 2
    assert metadata["end_col"] == 2
    assert metadata["texts"] == [["R1C1", "R1C2"], ["R2C1", "R2C2"]]


def test_ppt_split_table_cells_restores_original_texts_from_metadata() -> None:
    presentation = FakePresentation(
        [
            ["A", "B"],
            ["C", "D"],
        ]
    )
    _ppt_merge_table_cells(
        presentation,
        {"slide_index": 1, "shape_index": 1, "start_row": 1, "start_col": 1, "end_row": 2, "end_col": 2},
    )

    result = _ppt_split_table_cells(
        presentation,
        {"slide_index": 1, "shape_index": 1, "row": 1, "col": 1},
    )

    assert "split_table_cells" in result
    assert presentation.table.last_split == (1, 1)
    assert presentation.table.Cell(1, 1).Shape.TextFrame.TextRange.Text == "A"
    assert presentation.table.Cell(1, 2).Shape.TextFrame.TextRange.Text == "B"
    assert presentation.table.Cell(2, 1).Shape.TextFrame.TextRange.Text == "C"
    assert presentation.table.Cell(2, 2).Shape.TextFrame.TextRange.Text == "D"
    assert presentation.shape.Tags("OfficeMCP.TableMerge.1.1") == ""


def test_ppt_split_table_cells_restores_original_texts_from_non_anchor_cell() -> None:
    presentation = FakePresentation(
        [
            ["A", "B"],
            ["C", "D"],
        ]
    )
    _ppt_merge_table_cells(
        presentation,
        {"slide_index": 1, "shape_index": 1, "start_row": 1, "start_col": 1, "end_row": 2, "end_col": 2},
    )

    result = _ppt_split_table_cells(
        presentation,
        {"slide_index": 1, "shape_index": 1, "row": 2, "col": 2},
    )

    assert "split_table_cells" in result
    assert presentation.table.last_split == (2, 2)
    assert presentation.table.Cell(1, 1).Shape.TextFrame.TextRange.Text == "A"
    assert presentation.table.Cell(1, 2).Shape.TextFrame.TextRange.Text == "B"
    assert presentation.table.Cell(2, 1).Shape.TextFrame.TextRange.Text == "C"
    assert presentation.table.Cell(2, 2).Shape.TextFrame.TextRange.Text == "D"
    assert presentation.shape.Tags("OfficeMCP.TableMerge.1.1") == ""
