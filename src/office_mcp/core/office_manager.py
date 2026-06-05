"""Office COM 应用生命周期管理."""

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


def _is_descendant_of(pid: int, ancestor_pid: int, max_depth: int = 6) -> bool:
    """判断 pid 是否是 ancestor_pid 的后代进程.

    通过 psutil 查询父进程链, 最多向上追溯 max_depth 层, 避免无限循环.
    """
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

# 常量
WORD_PROGID = "Word.Application"
EXCEL_PROGID = "Excel.Application"
PPT_PROGID = "PowerPoint.Application"

PDF_FORMAT_WORD = 17
PDF_FORMAT_EXCEL = 57
PDF_FORMAT_PPT = 32


class OfficeManager:
    """管理 Office COM 应用实例和已打开文档."""

    def __init__(self):
        self._apps: dict[str, Any] = {}  # progid -> Application
        self._documents: dict[str, Any] = {}  # file_path -> Document/Workbook/Presentation
        self._doc_types: dict[str, str] = {}  # file_path -> app_type
        self._active_file: Path | None = None

    # ---------- App 管理 ----------

    # COM 可重试的错误码 (winerror.h)
    _RETRYABLE_HRESULT = {
        -2147418110,  # RPC_E_SERVERCALL_RETRYLATER
        -2147417848,  # RPC_E_CALL_REJECTED
        -2292178814,  # OLE_E_PROMPTSAVECANCELLED
    }

    def _retry_on_modal(self, func, *args, max_retries: int = 15, retry_interval: float = 1.0):
        """COM 调用被模态对话框阻塞时自动重试 (使用 winerror HRESULT)."""
        for attempt in range(max_retries):
            try:
                return func(*args)
            except Exception as e:
                hresult = getattr(e, 'hresult', None)
                if hresult in self._RETRYABLE_HRESULT and attempt < max_retries - 1:
                    time.sleep(retry_interval)
                    continue
                raise
        raise Exception(f"COM 调用失败，已重试 {max_retries} 次")

    def _get_app(self, app_type: str) -> Any:
        """获取或创建 Office 应用实例."""
        app_type = self._normalize_app_type(app_type)
        if app_type in self._apps:
            return self._apps[app_type]

        try:
            import win32com.client
        except ImportError:
            raise OfficeNotInstalledError()

        progid_map = {
            "word": WORD_PROGID,
            "excel": EXCEL_PROGID,
            "ppt": PPT_PROGID,
        }
        progid = progid_map.get(app_type)
        if not progid:
            raise COMOperationError(f"未知的应用类型: {app_type}")

        def _dispatch():
            app = win32com.client.Dispatch(progid)
            app.Visible = _get_app_visibility(app_type)
            self._apps[app_type] = app
            logger.info(f"已启动 {progid}")
            return app

        try:
            return self._retry_on_modal(_dispatch)
        except Exception as e:
            raise OfficeNotInstalledError() from e

    def _normalize_app_type(self, app_type: str) -> str:
        """Normalize user-facing app type values."""
        normalized = (app_type or "").strip().lower()
        if normalized in {"powerpoint", "power point"}:
            normalized = "ppt"

        if normalized not in {"word", "excel", "ppt"}:
            raise COMOperationError(f"未知的应用类型: {app_type}")

        return normalized

    def _get_app_type_by_path(self, file_path: Path) -> str:
        """根据文件扩展名判断应用类型."""
        ext = file_path.suffix.lower()
        if ext in (".docx", ".doc", ".docm"):
            return "word"
        elif ext in (".xlsx", ".xls", ".xlsm"):
            return "excel"
        elif ext in (".pptx", ".ppt", ".pptm"):
            return "ppt"
        raise COMOperationError(f"不支持的文件类型: {ext}")

    # ---------- 文档管理 ----------

    def _backup_if_needed(self, file_path: Path) -> None:
        """编辑前自动备份."""
        if not settings.backup_before_edit or not file_path.exists():
            return
        timestamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup_name = f"{file_path.name}.office-mcp-backup-{timestamp}{file_path.suffix}"
        backup_path = file_path.parent / backup_name
        try:
            shutil.copy2(file_path, backup_path)
            logger.info(f"已备份: {backup_path}")
        except Exception as e:
            logger.warning(f"备份失败: {e}")

    def open_document(self, file_path: Path) -> Any:
        """打开文档."""
        path_str = str(file_path)
        if path_str in self._documents:
            raise DocumentAlreadyOpenError(path_str)

        app_type = self._get_app_type_by_path(file_path)
        app = self._get_app(app_type)

        self._backup_if_needed(file_path)

        def _open():
            if app_type == "word":
                doc = app.Documents.Open(str(file_path))
            elif app_type == "excel":
                doc = app.Workbooks.Open(str(file_path))
            elif app_type == "ppt":
                doc = app.Presentations.Open(str(file_path))
            else:
                raise COMOperationError(f"不支持的应用类型: {app_type}")
            self._documents[path_str] = doc
            self._doc_types[path_str] = app_type
            logger.info(f"已打开文档: {path_str}")
            return doc

        try:
            return self._retry_on_modal(_open)
        except Exception as e:
            raise COMOperationError("打开文档", str(e)) from e

    def create_document(self, file_path: Path, overwrite: bool = False) -> Any:
        """创建新文档."""
        path_str = str(file_path)
        if path_str in self._documents:
            raise DocumentAlreadyOpenError(path_str)

        if file_path.exists() and not overwrite:
            raise COMOperationError(
                "创建文档", f"文件已存在: {path_str}"
            )

        app_type = self._get_app_type_by_path(file_path)
        app = self._get_app(app_type)

        def _create():
            if app_type == "word":
                doc = app.Documents.Add()
            elif app_type == "excel":
                doc = app.Workbooks.Add()
            elif app_type == "ppt":
                doc = app.Presentations.Add()
            else:
                raise COMOperationError(f"不支持的应用类型: {app_type}")
            self._documents[path_str] = doc
            self._doc_types[path_str] = app_type
            logger.info(f"已创建文档: {path_str}")
            return doc

        try:
            return self._retry_on_modal(_create)
        except Exception as e:
            raise COMOperationError("创建文档", str(e)) from e

    def get_document(self, file_path: Path, require_active: bool = True) -> Any:
        """Get an already-open document."""
        if require_active and self._active_file and str(self._active_file.resolve()) != str(file_path.resolve()):
            raise DocumentNotOpenError(
                f"Document is not active: {file_path}. "
                f"Current active file: {self._active_file}. "
                "Call office_activate first if this tool requires the active document lock."
            )
        path_str = str(file_path)
        doc = self._documents.get(path_str)
        if doc is None:
            raise DocumentNotOpenError(path_str)
        return doc

    def ensure_document(self, file_path: Path, activate: bool = False) -> Any:
        """Open a document on demand and optionally mark it active."""
        path_str = str(file_path)
        if path_str not in self._documents:
            self.open_document(file_path)
        if activate:
            self._active_file = file_path
        return self._documents[path_str]

    def activate_app(self, app_type: str, file_path: Path | None = None) -> dict:
        """激活指定应用/文档，后续操作自动锁定到该文件.

        Args:
            app_type: "word" / "excel" / "ppt"
            file_path: 可选，指定要激活的文件路径
        """
        normalized_app_type = self._normalize_app_type(app_type)

        if file_path:
            inferred_app_type = self._get_app_type_by_path(file_path)
            if normalized_app_type != inferred_app_type:
                raise COMOperationError(
                    "激活应用",
                    f"app_type={normalized_app_type} 与文件类型 {inferred_app_type} 不一致: {file_path}",
                )
            path_str = str(file_path)
            if path_str not in self._documents:
                self.open_document(file_path)
            self._active_file = file_path
        else:
            self._get_app(normalized_app_type)
            self._active_file = None

        return {
            "app_type": normalized_app_type,
            "active_file": str(self._active_file) if self._active_file else None,
            "open_documents": list(self._documents.keys()),
        }

    def close_document(self, file_path: Path, save: bool = True) -> None:
        """关闭文档."""
        path_str = str(file_path)
        doc = self._documents.get(path_str)
        if doc is None:
            raise DocumentNotOpenError(path_str)

        app_type = self._doc_types.get(path_str)

        def _close():
            if save:
                if app_type == "word":
                    doc.SaveAs(path_str)
                elif app_type == "excel":
                    doc.SaveAs(path_str)
                elif app_type == "ppt":
                    doc.SaveAs(path_str)
                logger.info(f"已保存: {path_str}")

            if app_type == "word":
                doc.Close(SaveChanges=False)
            elif app_type == "excel":
                doc.Close(SaveChanges=False)
            elif app_type == "ppt":
                doc.Close()

        try:
            self._retry_on_modal(_close)
        except Exception as e:
            logger.warning(f"关闭文档时出错: {e}")
        finally:
            self._documents.pop(path_str, None)
            self._doc_types.pop(path_str, None)
            if self._active_file == file_path:
                self._active_file = None
            logger.info(f"已关闭文档: {path_str}")

    def export_pdf(self, file_path: Path, output_path: Path | None = None) -> Path:
        """导出文档为 PDF."""
        path_str = str(file_path)
        doc = self.get_document(file_path)
        app_type = self._doc_types.get(path_str)

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
                raise COMOperationError(f"不支持的应用类型: {app_type}")
        except Exception as e:
            raise COMOperationError("导出 PDF", str(e)) from e

        logger.info(f"已导出 PDF: {output_path}")
        return output_path

    # ---------- 清理 ----------

    def cleanup(self, force: bool = False) -> dict[str, Any]:
        """清理 Office 实例."""
        result = {"closed_documents": [], "quit_apps": [], "force_killed": []}

        # 先关闭所有追踪的文档
        for path_str in list(self._documents.keys()):
            try:
                self.close_document(Path(path_str), save=True)
                result["closed_documents"].append(path_str)
            except Exception as e:
                logger.warning(f"关闭文档失败: {path_str} - {e}")

        # 再退出应用
        for app_type, app in list(self._apps.items()):
            try:
                app.Quit()
                result["quit_apps"].append(app_type)
            except Exception as e:
                logger.warning(f"退出应用失败: {app_type} - {e}")
        self._apps.clear()

        if force:
            # 强制终止: 仅杀本进程衍生的 Office 实例 (通过 COM 父进程关系判定)
            # 避免误杀用户手动打开的 Office 窗口 (未保存数据)
            process_names = ["WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"]
            current_pid = os.getpid()
            killed_pids: set[int] = set()
            for proc in psutil.process_iter(["pid", "name", "ppid"]):
                try:
                    if proc.info["name"] in process_names:
                        # 仅当父进程链能追溯到本进程才杀
                        if _is_descendant_of(proc.info["pid"], current_pid):
                            proc.kill()
                            killed_pids.add(proc.info["pid"])
                            result["force_killed"].append(proc.info["name"])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
            if not killed_pids:
                # 没有子进程, 通常意味着用户手动打开的 Office
                logger.warning(
                    "force=True 但未发现本进程派生的 Office 进程, 跳过强杀以保护用户数据"
                )

        logger.info(f"清理完成: {result}")
        return result

    def status(self) -> dict[str, Any]:
        """获取当前状态."""
        return {
            "running_apps": list(self._apps.keys()),
            "open_documents": [
                {
                    "path": path,
                    "type": self._doc_types.get(path, "unknown"),
                }
                for path in self._documents
            ],
        }


# 全局单例
office_manager = OfficeManager()
