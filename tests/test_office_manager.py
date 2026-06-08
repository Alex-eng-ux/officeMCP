from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest

from office_mcp.compat import FallbackFastMCP
from office_mcp.config import settings
from office_mcp.core.errors import COMOperationError, OfficeNotInstalledError
from office_mcp.core.office_manager import OfficeManager
from office_mcp.operations.word_ops import (
    _get_document_info,
    _is_document_protected,
)
from office_mcp.tools.office import register_office_tools


def test_normalize_app_type_accepts_powerpoint_alias() -> None:
    manager = OfficeManager()

    assert manager._normalize_app_type("powerpoint") == "ppt"


def test_activate_app_rejects_mismatched_file_type(tmp_path: Path) -> None:
    manager = OfficeManager()
    ppt_path = tmp_path / "deck.pptx"

    with pytest.raises(COMOperationError, match=r"app_type=word.*deck\.pptx"):
        manager.activate_app("word", ppt_path)


def test_get_document_rebinds_when_cached_handle_is_stale(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "sheet.xlsx"
    stale_doc = object()
    rebound_doc = object()

    manager._documents[str(doc_path.resolve())] = stale_doc
    manager._doc_types[str(doc_path.resolve())] = "excel"

    monkeypatch.setattr(manager, "_is_document_alive", lambda doc: doc is rebound_doc)
    monkeypatch.setattr(manager, "_find_open_document", lambda file_path, app_type: rebound_doc)

    doc = manager.get_document(doc_path)

    assert doc is rebound_doc
    assert manager._documents[str(doc_path.resolve())] is rebound_doc


def test_get_document_reopens_when_cached_handle_is_stale_and_live_app_has_no_document(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "notes.docx"
    stale_doc = object()
    reopened_doc = object()

    manager._documents[str(doc_path.resolve())] = stale_doc
    manager._doc_types[str(doc_path.resolve())] = "word"

    monkeypatch.setattr(manager, "_is_document_alive", lambda doc: False)
    monkeypatch.setattr(manager, "_find_open_document", lambda file_path, app_type: None)
    monkeypatch.setattr(manager, "open_document", lambda file_path: reopened_doc)

    doc = manager.get_document(doc_path)

    assert doc is reopened_doc


def test_ensure_document_marks_target_active(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "deck.pptx"
    live_doc = object()

    monkeypatch.setattr(manager, "get_document", lambda file_path, require_active=False: live_doc)

    doc = manager.ensure_document(doc_path, activate=True)

    assert doc is live_doc
    assert manager._active_file == doc_path


def test_ensure_document_activates_live_document_when_requested(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "sheet.xlsx"

    class FakeSheet:
        def __init__(self) -> None:
            self.activated = False

        def Activate(self) -> None:  # noqa: N802
            self.activated = True

    class FakeWindow:
        def __init__(self) -> None:
            self.activated = False

        def Activate(self) -> None:  # noqa: N802
            self.activated = True

    class FakeWindows:
        Count = 1

        def __init__(self, window: FakeWindow) -> None:
            self.window = window

        def __call__(self, index: int) -> FakeWindow:
            assert index == 1
            return self.window

    class FakeWorkbook:
        def __init__(self) -> None:
            self.activated = False
            self.window = FakeWindow()
            self.ActiveSheet = FakeSheet()
            self.Windows = FakeWindows(self.window)

        def Activate(self) -> None:  # noqa: N802
            self.activated = True

    fake_doc = FakeWorkbook()
    monkeypatch.setattr(manager, "get_document", lambda file_path, require_active=False: fake_doc)

    doc = manager.ensure_document(doc_path, activate=True)

    assert doc is fake_doc
    assert fake_doc.activated is True
    assert fake_doc.window.activated is True
    assert fake_doc.ActiveSheet.activated is True


def test_close_document_keeps_tracking_when_close_fails(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "draft.docx"
    path_key = str(doc_path.resolve())

    class FakeDoc:
        FullName = path_key

        def Save(self) -> None:  # noqa: N802
            return None

        def Close(self, SaveChanges=False) -> None:  # noqa: N802, ANN001
            raise RuntimeError("still open")

    fake_doc = FakeDoc()
    manager._documents[path_key] = fake_doc
    manager._doc_types[path_key] = "word"

    monkeypatch.setattr(manager, "_retry_on_modal", lambda func, *args, **kwargs: func(*args))

    manager.close_document(doc_path, save=True)

    assert manager._documents[path_key] is fake_doc
    assert manager._doc_types[path_key] == "word"


def test_track_document_adopts_existing_handle(tmp_path: Path) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "merged.docx"
    doc = object()

    tracked = manager.track_document(doc_path, doc, app_type="word")

    assert tracked is doc
    assert manager._documents[str(doc_path.resolve())] is doc
    assert manager._doc_types[str(doc_path.resolve())] == "word"


def test_open_document_uses_parameterized_powerpoint_open(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    manager = OfficeManager()
    deck_path = tmp_path / "deck.pptx"
    deck_path.touch()
    open_calls: list[dict[str, object]] = []
    fake_pres = object()

    class FakePresentations:
        def Open(self, path: str, **kwargs):  # noqa: N802
            open_calls.append({"path": path, **kwargs})
            return fake_pres

    fake_app = type("App", (), {"Presentations": FakePresentations()})()

    monkeypatch.setattr(manager, "_find_open_document", lambda file_path, app_type: None)
    monkeypatch.setattr(manager, "_get_app", lambda app_type: fake_app)
    monkeypatch.setattr(manager, "_retry_on_modal", lambda func, *args, **kwargs: func(*args))

    doc = manager.open_document(deck_path)

    assert doc is fake_pres
    assert open_calls == [{
        "path": str(deck_path),
        "ReadOnly": False,
        "Untitled": False,
        "WithWindow": settings.ppt_visible,
    }]


def test_dispatch_app_maps_class_not_registered_to_office_not_installed(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    manager = OfficeManager()

    class FakeError(Exception):
        hresult = -2147221164

    fake_client = types.SimpleNamespace(
        GetActiveObject=lambda progid: None,
        DispatchEx=lambda progid: (_ for _ in ()).throw(FakeError("Class not registered")),
        Dispatch=lambda progid: (_ for _ in ()).throw(FakeError("Class not registered")),
    )
    fake_win32com = types.SimpleNamespace(client=fake_client)

    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(CoInitialize=lambda: None))
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    with pytest.raises(OfficeNotInstalledError, match="Class not registered"):
        manager._dispatch_app("excel")


def test_dispatch_app_maps_busy_failures_to_com_operation_error(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    manager = OfficeManager()

    class FakeBusyError(Exception):
        hresult = -2147418111

    fake_client = types.SimpleNamespace(
        GetActiveObject=lambda progid: None,
        DispatchEx=lambda progid: (_ for _ in ()).throw(FakeBusyError("Call was rejected by callee")),
        Dispatch=lambda progid: (_ for _ in ()).throw(FakeBusyError("Call was rejected by callee")),
    )
    fake_win32com = types.SimpleNamespace(client=fake_client)

    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(CoInitialize=lambda: None))
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)

    with pytest.raises(COMOperationError, match="busy, blocked by a dialog, or disconnected"):
        manager._dispatch_app("excel")


def test_create_document_saves_new_word_document_to_target_path(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "proposal.docx"

    class FakeDocument:
        def __init__(self) -> None:
            self.saved_paths: list[str] = []
            self.FullName = ""

        def SaveAs(self, path: str) -> None:  # noqa: N802
            self.saved_paths.append(path)
            Path(path).touch()
            self.FullName = path

    class FakeDocuments:
        def __init__(self, doc: FakeDocument) -> None:
            self._doc = doc

        def Add(self) -> FakeDocument:  # noqa: N802
            return self._doc

    class FakeWordApp:
        def __init__(self, doc: FakeDocument) -> None:
            self.Documents = FakeDocuments(doc)

    fake_doc = FakeDocument()
    fake_app = FakeWordApp(fake_doc)

    monkeypatch.setattr(manager, "_get_app", lambda app_type: fake_app)
    monkeypatch.setattr(manager, "_retry_on_modal", lambda func, *args, **kwargs: func(*args))

    doc = manager.create_document(doc_path, overwrite=True)

    assert doc is fake_doc
    assert fake_doc.saved_paths == [str(doc_path.resolve())]
    assert manager._documents[str(doc_path.resolve())] is fake_doc
    assert manager._doc_types[str(doc_path.resolve())] == "word"


def test_create_document_creates_parent_directories_before_initial_save(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    manager = OfficeManager()
    doc_path = tmp_path / "nested" / "folder" / "tracker.xlsx"

    class FakeDocument:
        def __init__(self) -> None:
            self.saved_paths: list[str] = []
            self.FullName = ""

        def SaveAs(self, path: str) -> None:  # noqa: N802
            self.saved_paths.append(path)
            Path(path).touch()
            self.FullName = path

    class FakeWorkbooks:
        def __init__(self, doc: FakeDocument) -> None:
            self._doc = doc

        def Add(self) -> FakeDocument:  # noqa: N802
            return self._doc

    class FakeExcelApp:
        def __init__(self, doc: FakeDocument) -> None:
            self.Workbooks = FakeWorkbooks(doc)

    fake_doc = FakeDocument()
    fake_app = FakeExcelApp(fake_doc)

    monkeypatch.setattr(manager, "_get_app", lambda app_type: fake_app)
    monkeypatch.setattr(manager, "_retry_on_modal", lambda func, *args, **kwargs: func(*args))

    doc = manager.create_document(doc_path, overwrite=True)

    assert doc is fake_doc
    assert doc_path.parent.exists()
    assert doc_path.exists()
    assert fake_doc.saved_paths == [str(doc_path.resolve())]


def test_register_office_tools_exposes_create_document() -> None:
    mcp = FallbackFastMCP("office-test")

    register_office_tools(mcp)

    assert "office_create_document" in mcp.tools


def test_is_document_protected_returns_false_for_no_protection() -> None:
    doc = types.SimpleNamespace(ProtectionType=0)

    assert _is_document_protected(doc) is False


def test_is_document_protected_returns_false_for_negative_one_no_protection() -> None:
    doc = types.SimpleNamespace(ProtectionType=-1)

    assert _is_document_protected(doc) is False


def test_is_document_protected_returns_true_for_non_zero_protection() -> None:
    doc = types.SimpleNamespace(ProtectionType=3)

    assert _is_document_protected(doc) is True


def test_is_document_protected_returns_false_when_probe_raises() -> None:
    class FakeDoc:
        @property
        def ProtectionType(self):  # noqa: N802
            raise RuntimeError("COM probe failed")

    assert _is_document_protected(FakeDoc()) is False


def test_get_document_info_tolerates_compute_statistics_failures() -> None:
    class FakeProperty:
        def __init__(self, value) -> None:
            self.Value = value

    class FakeProperties:
        def __call__(self, name: str) -> FakeProperty:
            mapping = {
                "Author": "Tester",
                "Title": "Doc",
                "Subject": "Subj",
                "Creation Date": "2026-06-07",
                "Last Save Time": "2026-06-07",
                "Revision Number": 1,
            }
            return FakeProperty(mapping[name])

    class FakeDoc:
        ProtectionType = -1
        TrackRevisions = False
        BuiltInDocumentProperties = FakeProperties()

        def ComputeStatistics(self, stat_id: int) -> int:  # noqa: N802
            if stat_id == 2:
                raise RuntimeError("boom")
            return 7

    info = _get_document_info(FakeDoc(), {})

    assert info["page_count"] == 0
    assert info["word_count"] == 7
    assert info["character_count"] == 7
    assert info["paragraph_count"] == 7
    assert info["line_count"] == 7
    assert info["author"] == "Tester"
    assert info["is_protected"] is False
