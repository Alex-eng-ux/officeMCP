"""Office COM application and document lifecycle management."""

from __future__ import annotations

import contextlib
import datetime
import logging
import os
import shutil
import time
from pathlib import Path
from typing import Any

import psutil

from office_mcp.config import settings
from office_mcp.core.errors import (
    COMOperationError,
    DocumentAlreadyOpenError,
    DocumentNotOpenError,
    OfficeNotInstalledError,
)

logger = logging.getLogger(__name__)


WORD_PROGID = "Word.Application"
EXCEL_PROGID = "Excel.Application"
PPT_PROGID = "PowerPoint.Application"

PDF_FORMAT_WORD = 17
PDF_FORMAT_EXCEL = 57
PDF_FORMAT_PPT = 32

RETRYABLE_HRESULTS = {
    -2147418111,  # RPC_E_CALL_REJECTED
    -2147418110,  # RPC_E_SERVERCALL_RETRYLATER
    -2147417848,  # RPC_E_DISCONNECTED
    -2147417846,  # RPC_E_SYS_CALL_FAILED / busy variants seen in Office automation
    -2147023170,  # RPC_S_CALL_FAILED
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE
    -2292178814,  # OLE_E_PROMPTSAVECANCELLED
}


def _is_descendant_of(pid: int, ancestor_pid: int, max_depth: int = 6) -> bool:
    """Return whether pid descends from ancestor_pid."""
    try:
        current = psutil.Process(pid)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return False

    seen: set[int] = set()
    for _ in range(max_depth):
        if current.pid in seen:
            return False
        seen.add(current.pid)
        if current.pid == ancestor_pid:
            return True
        try:
            current = current.parent()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            return False
        if current is None:
            return False
    return False


def _get_app_visibility(app_type: str) -> bool:
    """Return visibility for a specific Office application."""
    visibility_map = {
        "word": settings.word_visible,
        "excel": settings.excel_visible,
        "ppt": settings.ppt_visible,
    }
    return visibility_map[app_type]


def _normalize_path_key(file_path: Path) -> str:
    """Normalize a file path for dictionary keys."""
    return str(file_path.resolve())


def _path_exists_and_nonempty(file_path: Path) -> bool:
    """Return whether a file exists and is non-empty."""
    try:
        return file_path.exists() and file_path.stat().st_size > 0
    except OSError:
        return False


def _open_powerpoint_presentation(
    app: Any,
    file_path: Path | str,
    *,
    read_only: bool = False,
    untitled: bool = False,
    with_window: bool | None = None,
) -> Any:
    """Open a PowerPoint presentation with explicit modal-safe flags."""
    if with_window is None:
        with_window = _get_app_visibility("ppt")
    return app.Presentations.Open(
        str(file_path),
        ReadOnly=bool(read_only),
        Untitled=bool(untitled),
        WithWindow=bool(with_window),
    )


class OfficeManager:
    """Manage Office COM app instances and tracked documents."""

    def __init__(self) -> None:
        self._apps: dict[str, Any] = {}
        self._documents: dict[str, Any] = {}
        self._doc_types: dict[str, str] = {}
        self._active_file: Path | None = None

    def _is_retryable_com_exception(self, error: Exception) -> bool:
        """Return whether an exception looks like a transient Office COM failure."""
        hresult = getattr(error, "hresult", None)
        if hresult in RETRYABLE_HRESULTS:
            return True

        text = str(error).lower()
        retry_markers = (
            "call was rejected by callee",
            "server unavailable",
            "server threw an exception",
            "remote procedure call failed",
            "rpc server is unavailable",
            "object invoked has disconnected",
            "application is busy",
        )
        return any(marker in text for marker in retry_markers)

    def _retry_on_modal(
        self,
        func,
        *args,
        max_retries: int = 15,
        retry_interval: float = 1.0,
    ):
        """Retry Office COM calls that fail while Office is busy or modal."""
        last_error: Exception | None = None
        for attempt in range(max_retries):
            try:
                return func(*args)
            except Exception as error:  # noqa: BLE001
                last_error = error
                if not self._is_retryable_com_exception(error) or attempt >= max_retries - 1:
                    raise
                time.sleep(retry_interval)
        raise COMOperationError("retry_on_modal", str(last_error))

    def _normalize_app_type(self, app_type: str) -> str:
        """Normalize user-facing app type values."""
        normalized = (app_type or "").strip().lower()
        if normalized in {"powerpoint", "power point"}:
            normalized = "ppt"

        if normalized not in {"word", "excel", "ppt"}:
            raise COMOperationError(f"Unsupported app type: {app_type}")

        return normalized

    def _get_app_progid(self, app_type: str) -> str:
        """Return the COM progid for an app type."""
        progid_map = {
            "word": WORD_PROGID,
            "excel": EXCEL_PROGID,
            "ppt": PPT_PROGID,
        }
        return progid_map[self._normalize_app_type(app_type)]

    def _get_collection_name(self, app_type: str) -> str:
        """Return the Office collection name for open documents."""
        return {
            "word": "Documents",
            "excel": "Workbooks",
            "ppt": "Presentations",
        }[self._normalize_app_type(app_type)]

    def _remember_document(self, file_path: Path, app_type: str, doc: Any) -> Any:
        """Track a document handle."""
        path_key = _normalize_path_key(file_path)
        self._documents[path_key] = doc
        self._doc_types[path_key] = app_type
        return doc

    def track_document(self, file_path: Path, doc: Any, app_type: str | None = None) -> Any:
        """Public helper to adopt an already-open Office document into manager state."""
        normalized_type = app_type or self._get_app_type_by_path(file_path)
        return self._remember_document(file_path, normalized_type, doc)

    def _forget_document(self, file_path: Path) -> None:
        """Forget a tracked document handle."""
        path_key = _normalize_path_key(file_path)
        self._documents.pop(path_key, None)
        self._doc_types.pop(path_key, None)
        if self._active_file and _normalize_path_key(self._active_file) == path_key:
            self._active_file = None

    def _probe_app(self, app_type: str, app: Any) -> bool:
        """Best-effort liveness probe for an app object."""
        try:
            _ = app.Visible
            _ = getattr(app, self._get_collection_name(app_type))
            return True
        except Exception as error:  # noqa: BLE001
            if self._is_retryable_com_exception(error):
                return False
            return False

    def _dispatch_app(self, app_type: str) -> Any:
        """Create a fresh Office app handle."""
        try:
            import pythoncom
        except ImportError:
            pythoncom = None

        try:
            import win32com.client
        except ImportError as error:
            raise OfficeNotInstalledError(
                detail="pywin32 is not installed",
                suggestion="Install pywin32 and Microsoft Office, then retry.",
            ) from error

        if pythoncom is not None:
            with contextlib.suppress(Exception):
                pythoncom.CoInitialize()

        progid = self._get_app_progid(app_type)

        def _finalize_app(app: Any, origin: str) -> Any:
            try:
                app.Visible = _get_app_visibility(app_type)
            except Exception as vis_err:  # noqa: BLE001
                # PowerPoint may reject Visible=False when presentations are open
                logger.debug("Could not set %s Visible=%s: %s", progid, _get_app_visibility(app_type), vis_err)
            # Suppress modal dialogs for unattended automation where supported.
            if app_type in ("excel", "word", "ppt"):
                try:
                    app.DisplayAlerts = 0  # wdAlertsNone / xlAlertsNone
                except Exception:  # noqa: BLE001
                    logger.debug("Could not set DisplayAlerts=0 for %s", progid)
            self._apps[app_type] = app
            logger.info("Attached Office app via %s: %s", origin, progid)
            return app

        def _dispatch_new() -> Any:
            try:
                app = win32com.client.DispatchEx(progid)
                return _finalize_app(app, "DispatchEx")
            except Exception as error:  # noqa: BLE001
                logger.debug("DispatchEx failed for %s: %s", progid, error)
                app = win32com.client.Dispatch(progid)
                return _finalize_app(app, "Dispatch")

        def _attach_existing() -> Any | None:
            get_active = getattr(win32com.client, "GetActiveObject", None)
            if get_active is None:
                return None
            try:
                app = get_active(progid)
            except Exception as error:  # noqa: BLE001
                logger.debug("GetActiveObject failed for %s: %s", progid, error)
                return None
            if app is None:
                return None
            return _finalize_app(app, "GetActiveObject")

        try:
            existing = self._retry_on_modal(_attach_existing)
            if existing is not None:
                return existing
            return self._retry_on_modal(_dispatch_new)
        except Exception as error:  # noqa: BLE001
            hresult = getattr(error, "hresult", None)
            detail = str(error)
            lower_detail = detail.lower()
            if hresult in {-2147221005, -2147221164} or any(
                marker in lower_detail
                for marker in (
                    "invalid class string",
                    "class not registered",
                    "activex component can't create object",
                    "library not registered",
                )
            ):
                raise OfficeNotInstalledError(detail=detail) from error
            if self._is_retryable_com_exception(error):
                raise COMOperationError(
                    f"{app_type}_dispatch",
                    f"Office application is busy, blocked by a dialog, or disconnected: {detail}",
                ) from error
            raise COMOperationError(
                f"{app_type}_dispatch",
                f"Unable to start or attach to {progid}: {detail}",
            ) from error

    def _get_app(self, app_type: str) -> Any:
        """Return a healthy app handle, recreating it if needed."""
        app_type = self._normalize_app_type(app_type)
        app = self._apps.get(app_type)
        if app is not None and self._probe_app(app_type, app):
            return app
        if app is not None:
            logger.warning("Discarding stale Office app handle: %s", app_type)
            self._apps.pop(app_type, None)
        return self._dispatch_app(app_type)

    def _get_app_type_by_path(self, file_path: Path) -> str:
        """Infer the Office app type from a file extension."""
        ext = file_path.suffix.lower()
        if ext in {".docx", ".doc", ".docm"}:
            return "word"
        if ext in {".xlsx", ".xls", ".xlsm"}:
            return "excel"
        if ext in {".pptx", ".ppt", ".pptm"}:
            return "ppt"
        raise COMOperationError("file_type", f"Unsupported file type: {ext}")

    def _backup_if_needed(self, file_path: Path) -> None:
        """Create a timestamped backup before editing when enabled."""
        if not settings.backup_before_edit or not file_path.exists():
            return
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_name = f"{file_path.name}.office-mcp-backup-{timestamp}{file_path.suffix}"
        backup_path = file_path.parent / backup_name
        try:
            shutil.copy2(file_path, backup_path)
            logger.info("Created Office MCP backup: %s", backup_path)
        except Exception as error:  # noqa: BLE001
            logger.warning("Backup failed for %s: %s", file_path, error)

    def _iter_live_documents(self, app_type: str, app: Any) -> list[Any]:
        """Return the live documents/workbooks/presentations from an app."""
        collection = getattr(app, self._get_collection_name(app_type))
        return list(collection)

    def _find_open_document(self, file_path: Path, app_type: str) -> Any | None:
        """Find an already-open document in the live Office app."""
        target = _normalize_path_key(file_path)
        app = self._get_app(app_type)
        try:
            for candidate in self._iter_live_documents(app_type, app):
                try:
                    full_name = getattr(candidate, "FullName", "")
                except Exception:  # noqa: BLE001
                    continue
                if full_name and _normalize_path_key(Path(full_name)) == target:
                    return candidate
        except Exception as error:  # noqa: BLE001
            if self._is_retryable_com_exception(error):
                logger.warning("Failed while scanning live %s documents: %s", app_type, error)
                return None
            raise
        return None

    def _is_document_alive(self, doc: Any) -> bool:
        """Probe whether a cached document handle still responds."""
        try:
            _ = getattr(doc, "Name")
            return True
        except Exception as error:  # noqa: BLE001
            if self._is_retryable_com_exception(error):
                return False
            return False

    def _rebind_document(self, file_path: Path, app_type: str) -> Any:
        """Recover a stale document handle by finding or reopening the document."""
        rebound = self._find_open_document(file_path, app_type)
        if rebound is not None:
            logger.info("Rebound live %s document handle: %s", app_type, file_path)
            return self._remember_document(file_path, app_type, rebound)

        logger.info("Reopening %s document after stale Office session: %s", app_type, file_path)
        self._forget_document(file_path)
        return self.open_document(file_path)

    def open_document(self, file_path: Path) -> Any:
        """Open a document/workbook/presentation."""
        path_key = _normalize_path_key(file_path)
        cached = self._documents.get(path_key)
        if cached is not None:
            if self._is_document_alive(cached):
                raise DocumentAlreadyOpenError(path_key)
            self._forget_document(file_path)

        app_type = self._get_app_type_by_path(file_path)
        existing = self._find_open_document(file_path, app_type)
        if existing is not None:
            logger.info("Adopted already-open %s document: %s", app_type, file_path)
            return self._remember_document(file_path, app_type, existing)

        app = self._get_app(app_type)
        self._backup_if_needed(file_path)

        def _open() -> Any:
            if app_type == "word":
                return app.Documents.Open(str(file_path))
            if app_type == "excel":
                return app.Workbooks.Open(str(file_path))
            if app_type == "ppt":
                return _open_powerpoint_presentation(app, file_path)
            raise COMOperationError("open_document", f"Unsupported app type: {app_type}")

        try:
            doc = self._retry_on_modal(_open)
            logger.info("Opened %s document: %s", app_type, path_key)
            return self._remember_document(file_path, app_type, doc)
        except Exception as error:  # noqa: BLE001
            raise COMOperationError("open_document", str(error)) from error

    def create_document(self, file_path: Path, overwrite: bool = False) -> Any:
        """Create a new blank Office document."""
        path_key = _normalize_path_key(file_path)
        cached = self._documents.get(path_key)
        if cached is not None:
            if self._is_document_alive(cached):
                raise DocumentAlreadyOpenError(path_key)
            self._forget_document(file_path)

        if file_path.exists() and not overwrite:
            raise COMOperationError("create_document", f"File already exists: {path_key}")

        file_path.parent.mkdir(parents=True, exist_ok=True)

        app_type = self._get_app_type_by_path(file_path)
        app = self._get_app(app_type)

        def _create() -> Any:
            if app_type == "word":
                return app.Documents.Add()
            if app_type == "excel":
                return app.Workbooks.Add()
            if app_type == "ppt":
                pres = app.Presentations.Add()
                # New presentations have 0 slides; add one so Slides(1) works
                if pres.Slides.Count == 0:
                    pres.Slides.Add(1, 2)  # 2 = ppLayoutText (title + content)
                return pres
            raise COMOperationError("create_document", f"Unsupported app type: {app_type}")

        def _initial_save(doc: Any) -> None:
            if app_type == "excel":
                doc.SaveAs(path_key)
                return
            if app_type in {"word", "ppt"}:
                doc.SaveAs(path_key)
                return
            raise COMOperationError("create_document", f"Unsupported app type: {app_type}")

        def _verify_initial_save(doc: Any) -> None:
            if not file_path.exists():
                raise COMOperationError(
                    "create_document",
                    f"Initial save did not materialize on disk: {path_key}",
                )

            full_name = getattr(doc, "FullName", "") or ""
            if full_name and _normalize_path_key(Path(full_name)) != path_key:
                raise COMOperationError(
                    "create_document",
                    f"Document saved to unexpected path: {full_name}",
                )

        try:
            doc = self._retry_on_modal(_create)
            self._retry_on_modal(_initial_save, doc)
            self._retry_on_modal(_verify_initial_save, doc)
            logger.info("Created %s document handle for: %s", app_type, path_key)
            return self._remember_document(file_path, app_type, doc)
        except Exception as error:  # noqa: BLE001
            raise COMOperationError("create_document", str(error)) from error

    def get_document(self, file_path: Path, require_active: bool = False) -> Any:
        """Return a live document handle, recovering if the cached handle is stale."""
        if require_active and self._active_file and _normalize_path_key(self._active_file) != _normalize_path_key(file_path):
            raise DocumentNotOpenError(
                f"Document is not active: {file_path}. "
                f"Current active file: {self._active_file}. "
                "Call office_activate first if this tool requires the active document lock."
            )

        path_key = _normalize_path_key(file_path)
        app_type = self._get_app_type_by_path(file_path)
        doc = self._documents.get(path_key)
        if doc is None:
            rebound = self._find_open_document(file_path, app_type)
            if rebound is not None:
                return self._remember_document(file_path, app_type, rebound)
            return self.open_document(file_path)

        if self._is_document_alive(doc):
            return doc

        logger.warning("Cached %s handle is stale, attempting recovery: %s", app_type, file_path)
        return self._rebind_document(file_path, app_type)

    def ensure_document(self, file_path: Path, activate: bool = False) -> Any:
        """Return a live document and optionally make it active."""
        doc = self.get_document(file_path, require_active=False)
        if activate:
            self._activate_document(file_path, doc)
            self._active_file = file_path
        return doc

    def _activate_document(self, file_path: Path, doc: Any) -> None:
        """Best-effort activation for window-sensitive Office operations."""
        app_type = self._get_app_type_by_path(file_path)

        with contextlib.suppress(Exception):
            doc.Activate()

        if app_type == "excel":
            with contextlib.suppress(Exception):
                windows = getattr(doc, "Windows", None)
                if windows is not None and windows.Count >= 1:
                    windows(1).Activate()
            with contextlib.suppress(Exception):
                sheet = getattr(doc, "ActiveSheet", None)
                if sheet is not None:
                    sheet.Activate()
        elif app_type == "ppt":
            with contextlib.suppress(Exception):
                windows = getattr(doc, "Windows", None)
                if windows is not None and windows.Count >= 1:
                    windows(1).Activate()
        elif app_type == "word":
            with contextlib.suppress(Exception):
                active_window = getattr(doc, "ActiveWindow", None)
                if active_window is not None:
                    active_window.Activate()

    def activate_app(self, app_type: str, file_path: Path | None = None) -> dict[str, Any]:
        """Activate an Office app and optionally a specific file."""
        normalized_app_type = self._normalize_app_type(app_type)
        if file_path is not None:
            inferred = self._get_app_type_by_path(file_path)
            if inferred != normalized_app_type:
                raise COMOperationError(
                    "activate_app",
                    f"app_type={normalized_app_type} does not match file type {inferred}: {file_path}",
                )
            self.ensure_document(file_path, activate=True)
        else:
            self._get_app(normalized_app_type)
            self._active_file = None

        return {
            "app_type": normalized_app_type,
            "active_file": str(self._active_file) if self._active_file else None,
            "open_documents": list(self._documents.keys()),
        }

    def close_document(self, file_path: Path, save: bool = True) -> None:
        """Close a tracked document with phased save→probe→close approach."""
        path_key = _normalize_path_key(file_path)
        doc = self._documents.get(path_key)
        if doc is None:
            raise DocumentNotOpenError(path_key)

        app_type = self._doc_types.get(path_key)
        save_succeeded = False

        # Phase 1: Save (with Save vs SaveAs distinction)
        if save:
            try:
                def _save() -> None:
                    if app_type == "excel":
                        # For Excel: check doc.Path to decide Save vs SaveAs
                        doc_path = doc.Path
                        if doc_path:
                            # Workbook has been saved before — fast in-place save
                            logger.info("Using Save() for already-saved Excel workbook: %s", path_key)
                            doc.Save()
                        else:
                            # New/unsaved workbook — full SaveAs required
                            logger.info("Using SaveAs() for new Excel workbook: %s", path_key)
                            doc.SaveAs(path_key)
                    elif app_type in {"word", "ppt"}:
                        # For Word/PPT: check FullName to decide Save vs SaveAs
                        try:
                            full_name = doc.FullName
                            if full_name and _normalize_path_key(Path(full_name)) == path_key:
                                logger.info("Using Save() for already-saved %s document: %s", app_type, path_key)
                                doc.Save()
                            else:
                                logger.info("Using SaveAs() for new %s document: %s", app_type, path_key)
                                doc.SaveAs(path_key)
                        except Exception:  # noqa: BLE001
                            logger.info("Falling back to SaveAs() for %s document: %s", app_type, path_key)
                            doc.SaveAs(path_key)

                self._retry_on_modal(_save)
                save_succeeded = True
                logger.info("Saved Office file before close: %s", path_key)
            except Exception as save_error:  # noqa: BLE001
                logger.warning("Save failed for %s: %s", path_key, save_error)

        # Phase 2: Probe (verify save succeeded)
        if save and save_succeeded:
            try:
                if app_type == "excel":
                    saved_flag = doc.Saved
                    if saved_flag:
                        logger.info("Excel save probe confirmed: %s", path_key)
                    else:
                        logger.warning("Excel save probe: doc.Saved is False after save: %s", path_key)
                # For Word/PPT: no reliable Saved property check; trust no-exception
            except Exception as probe_error:  # noqa: BLE001
                logger.warning("Save probe failed for %s: %s", path_key, probe_error)

        # Phase 3: Close
        close_succeeded = False
        try:
            def _close() -> None:
                if app_type == "word":
                    doc.Close(SaveChanges=False)
                elif app_type == "excel":
                    doc.Close(SaveChanges=False)
                elif app_type == "ppt":
                    doc.Close()

            self._retry_on_modal(_close)
            close_succeeded = True
            logger.info("Closed Office file: %s", path_key)
        except Exception as close_error:  # noqa: BLE001
            if save_succeeded:
                # Save already succeeded; close failure is non-critical
                logger.warning("Close failed after successful save for %s: %s", path_key, close_error)
            else:
                # Save failed; try closing with SaveChanges=True as fallback
                logger.warning("Close failed (save also failed) for %s: %s — retrying with SaveChanges=True", path_key, close_error)
                try:
                    if app_type == "word":
                        doc.Close(SaveChanges=True)
                    elif app_type == "excel":
                        doc.Close(SaveChanges=True)
                    elif app_type == "ppt":
                        doc.Close()
                    close_succeeded = True
                    logger.info("Closed Office file with SaveChanges=True fallback: %s", path_key)
                except Exception as fallback_close_error:  # noqa: BLE001
                    logger.warning("Fallback close also failed for %s: %s", path_key, fallback_close_error)
        finally:
            if close_succeeded:
                self._forget_document(file_path)

    def export_pdf(self, file_path: Path, output_path: Path | None = None) -> Path:
        """Export a tracked document to PDF."""
        path_key = _normalize_path_key(file_path)
        doc = self.get_document(file_path, require_active=False)
        app_type = self._doc_types.get(path_key)
        if output_path is None:
            output_path = file_path.with_suffix(".pdf")

        try:
            if app_type == "word":
                doc.SaveAs(str(output_path), FileFormat=PDF_FORMAT_WORD)
            elif app_type == "excel":
                doc.ExportAsFixedFormat(0, str(output_path), Quality=0)
            elif app_type == "ppt":
                doc.SaveAs(str(output_path), FileFormat=PDF_FORMAT_PPT)
            else:
                raise COMOperationError("export_pdf", f"Unsupported app type: {app_type}")
        except Exception as error:  # noqa: BLE001
            raise COMOperationError("export_pdf", str(error)) from error

        if not _path_exists_and_nonempty(output_path):
            raise COMOperationError("export_pdf", f"PDF export did not materialize on disk: {output_path}")

        logger.info("Exported PDF for %s -> %s", path_key, output_path)
        return output_path

    def cleanup(self, force: bool = False) -> dict[str, Any]:
        """Close tracked documents and quit tracked Office apps."""
        result = {"closed_documents": [], "quit_apps": [], "force_killed": []}

        for path_key in list(self._documents.keys()):
            try:
                self.close_document(Path(path_key), save=True)
                result["closed_documents"].append(path_key)
            except Exception as error:  # noqa: BLE001
                logger.warning("Failed to close tracked document %s: %s", path_key, error)

        for app_type, app in list(self._apps.items()):
            try:
                app.Quit()
                result["quit_apps"].append(app_type)
            except Exception as error:  # noqa: BLE001
                logger.warning("Failed to quit Office app %s: %s", app_type, error)
        self._apps.clear()

        if force:
            process_names = {"WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"}
            current_pid = os.getpid()
            for proc in psutil.process_iter(["pid", "name"]):
                try:
                    if proc.info["name"] in process_names and _is_descendant_of(proc.info["pid"], current_pid):
                        proc.kill()
                        result["force_killed"].append(proc.info["name"])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue

        return result

    def status(self) -> dict[str, Any]:
        """Return the currently tracked Office status."""
        open_documents: list[dict[str, str]] = []
        stale_paths: list[str] = []
        for path_key, doc in list(self._documents.items()):
            if self._is_document_alive(doc):
                open_documents.append({
                    "path": path_key,
                    "type": self._doc_types.get(path_key, "unknown"),
                })
            else:
                stale_paths.append(path_key)

        for stale_path in stale_paths:
            self._forget_document(Path(stale_path))

        running_apps = []
        for app_type, app in list(self._apps.items()):
            if self._probe_app(app_type, app):
                running_apps.append(app_type)
            else:
                self._apps.pop(app_type, None)

        return {
            "running_apps": running_apps,
            "open_documents": open_documents,
        }


office_manager = OfficeManager()
