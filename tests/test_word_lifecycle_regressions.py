from __future__ import annotations

from pathlib import Path

from office_mcp.operations import word_ops
from office_mcp.operations.word_ops import _delete_section, _mail_merge


class _FakeBoundaryRange:
    def __init__(self, end: int) -> None:
        self.End = end
        self.delete_called = False

    def Delete(self) -> None:  # noqa: N802
        self.delete_called = True


class _FakeSections:
    def __init__(self, ends: list[int]) -> None:
        self._ranges = [_FakeBoundaryRange(end) for end in ends]

    @property
    def Count(self) -> int:  # noqa: N802
        return len(self._ranges)

    def __call__(self, index: int):
        return type("Section", (), {"Range": self._ranges[index - 1]})()


def test_delete_section_removes_only_middle_section_break() -> None:
    calls: list[tuple[int, int]] = []

    class FakeDoc:
        def __init__(self) -> None:
            self.Sections = _FakeSections([10, 20, 30])

        def Range(self, start: int, end: int):  # noqa: N802
            calls.append((start, end))
            return type("DeletionRange", (), {"Delete": lambda self: None})()

    doc = FakeDoc()

    result = _delete_section(doc, {"section": 2})

    assert result == "deleted_section: 2"
    assert calls == [(19, 20)]
    assert doc.Sections._ranges[1].delete_called is False


def test_delete_section_removes_previous_break_for_last_section() -> None:
    calls: list[tuple[int, int]] = []

    class FakeDoc:
        def __init__(self) -> None:
            self.Sections = _FakeSections([10, 20, 30])

        def Range(self, start: int, end: int):  # noqa: N802
            calls.append((start, end))
            return type("DeletionRange", (), {"Delete": lambda self: None})()

    doc = FakeDoc()

    result = _delete_section(doc, {"section": 3})

    assert result == "deleted_section: 3"
    assert calls == [(19, 20)]


def test_mail_merge_tracks_saved_output_document(tmp_path: Path, monkeypatch) -> None:
    output_path = tmp_path / "merged.docx"
    tracked: list[tuple[Path, object, str | None]] = []

    class FakeMergedDocument:
        def __init__(self) -> None:
            self.FullName = ""
            self.Name = "Merged.docx"
            self.saved_paths: list[str] = []
            self.closed = False

        def SaveAs(self, path: str) -> None:  # noqa: N802
            self.saved_paths.append(path)
            self.FullName = path

        def Close(self, SaveChanges=False) -> None:  # noqa: N802, ANN001
            self.closed = True

    class FakeDocuments:
        def __init__(self, docs: list[object]) -> None:
            self.docs = docs

        @property
        def Count(self) -> int:  # noqa: N802
            return len(self.docs)

        def Item(self, index: int):  # noqa: N802
            return self.docs[index - 1]

    class FakeMailMerge:
        def __init__(self, app, merged_doc) -> None:
            self._app = app
            self._merged_doc = merged_doc
            self.Destination = None
            self.open_kwargs = None

        def OpenDataSource(self, **kwargs) -> None:  # noqa: N802
            self.open_kwargs = kwargs

        def Execute(self, Pause=False) -> None:  # noqa: N802, ANN001
            self._app.Documents.docs.append(self._merged_doc)
            self._app.ActiveDocument = self._merged_doc

    class FakeOptions:
        ConfirmConversions = True

    class FakeApp:
        def __init__(self, template_doc, merged_doc) -> None:
            self.DisplayAlerts = 1
            self.ScreenUpdating = True
            self.Options = FakeOptions()
            self.ActiveDocument = template_doc
            self.Documents = FakeDocuments([template_doc])
            self._merged_doc = merged_doc

    class FakeTemplateDocument:
        def __init__(self) -> None:
            self.FullName = str(tmp_path / "template.docx")
            self.Name = "template.docx"
            self.Application = None
            self.MailMerge = None

    template_doc = FakeTemplateDocument()
    merged_doc = FakeMergedDocument()
    app = FakeApp(template_doc, merged_doc)
    template_doc.Application = app
    template_doc.MailMerge = FakeMailMerge(app, merged_doc)

    monkeypatch.setattr(
        word_ops,
        "_retry_word_mail_merge_call",
        lambda stage, callable_obj, *args, **kwargs: callable_obj(*args, **kwargs),
    )
    monkeypatch.setattr(
        word_ops.office_manager,
        "track_document",
        lambda file_path, doc, app_type=None: tracked.append((file_path, doc, app_type)) or doc,
    )

    result = _mail_merge(
        template_doc,
        {
            "data_source": str(tmp_path / "data.xlsx"),
            "output_path": str(output_path),
        },
    )

    assert result == f"mail_merge_executed_to_file: {output_path}"
    assert merged_doc.saved_paths == [str(output_path)]
    assert merged_doc.closed is False
    assert tracked == [(output_path, merged_doc, "word")]
