from __future__ import annotations

import zipfile
from pathlib import Path

from office_mcp.operations.word_ops import (
    _add_paragraph,
    _build_mail_merge_open_kwargs,
    _list_excel_sheet_names,
)


def _write_minimal_workbook(path: Path, sheet_name: str = "Customers") -> None:
    workbook_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="{sheet_name}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/workbook.xml", workbook_xml)


def test_list_excel_sheet_names_reads_visible_worksheets(tmp_path: Path) -> None:
    workbook_path = tmp_path / "mail-merge.xlsx"
    _write_minimal_workbook(workbook_path, sheet_name="Clients")

    assert _list_excel_sheet_names(workbook_path) == ["Clients"]


def test_build_mail_merge_open_kwargs_adds_explicit_excel_defaults(tmp_path: Path) -> None:
    workbook_path = tmp_path / "mail-merge.xlsx"
    _write_minimal_workbook(workbook_path, sheet_name="Clients")

    kwargs = _build_mail_merge_open_kwargs(str(workbook_path), "", "")

    assert kwargs["Name"] == str(workbook_path)
    assert kwargs["Connection"].startswith("Provider=Microsoft.ACE.OLEDB.12.0;")
    assert kwargs["SQLStatement"] == "SELECT * FROM [Clients$]"
    assert kwargs["SubType"] == 16


def test_build_mail_merge_open_kwargs_preserves_user_overrides(tmp_path: Path) -> None:
    workbook_path = tmp_path / "mail-merge.xlsx"
    _write_minimal_workbook(workbook_path, sheet_name="Clients")

    kwargs = _build_mail_merge_open_kwargs(
        str(workbook_path),
        "DSN=CustomSource;",
        "SELECT * FROM [VIP$]",
    )

    assert kwargs["Connection"] == "DSN=CustomSource;"
    assert kwargs["SQLStatement"] == "SELECT * FROM [VIP$]"


def test_add_paragraph_appends_without_overwriting_previous_content() -> None:
    class FakeParagraph:
        def __init__(self, text: str = "") -> None:
            self.Range = type("Range", (), {"Text": text})()
            self.Style = None
            self.Alignment = None

    class FakeParagraphs:
        def __init__(self) -> None:
            self.items = [FakeParagraph("")]

        @property
        def Count(self) -> int:  # noqa: N802
            return len(self.items)

        def __call__(self, index: int) -> FakeParagraph:
            return self.items[index - 1]

    class FakeRange:
        def __init__(self, doc: "FakeDoc") -> None:
            self._doc = doc

        def Collapse(self, Direction=0) -> None:  # noqa: N802, ANN001
            return None

        def InsertAfter(self, text: str) -> None:  # noqa: N802
            self._doc.paragraphs.items.append(FakeParagraph(text))

    class FakeDoc:
        def __init__(self) -> None:
            self.paragraphs = FakeParagraphs()
            self.Content = FakeRange(self)

        @property
        def Paragraphs(self) -> FakeParagraphs:  # noqa: N802
            return self.paragraphs

    doc = FakeDoc()

    _add_paragraph(doc, {"text": "First paragraph", "style": "Normal", "alignment": "left"})
    _add_paragraph(doc, {"text": "Second paragraph", "style": "Normal", "alignment": "left"})

    assert [p.Range.Text for p in doc.Paragraphs.items[1:]] == [
        "First paragraph\r",
        "Second paragraph\r",
    ]
    assert doc.Paragraphs.items[1].Style == "Normal"
