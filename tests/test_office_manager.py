from __future__ import annotations

from pathlib import Path
import sys
import types

import pytest

from office_mcp.core.errors import COMOperationError, OfficeNotInstalledError
from office_mcp.core.office_manager import OfficeManager


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
