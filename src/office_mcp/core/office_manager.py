"""Office COM 应用生命周期管理."""

import datetime
import logging
import shutil
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

    # ---------- App 管理 ----------

    def _get_app(self, app_type: str) -> Any:
        """获取或创建 Office 应用实例."""
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

        try:
            app = win32com.client.Dispatch(progid)
            app.Visible = settings.visible
            self._apps[app_type] = app
            logger.info(f"已启动 {progid}")
            return app
        except Exception as e:
            raise OfficeNotInstalledError() from e

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

        try:
            if app_type == "word":
                doc = app.Documents.Open(str(file_path))
            elif app_type == "excel":
                doc = app.Workbooks.Open(str(file_path))
            elif app_type == "ppt":
                doc = app.Presentations.Open(str(file_path))
            else:
                raise COMOperationError(f"不支持的应用类型: {app_type}")
        except Exception as e:
            raise COMOperationError("打开文档", str(e)) from e

        self._documents[path_str] = doc
        self._doc_types[path_str] = app_type
        logger.info(f"已打开文档: {path_str}")
        return doc

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

        try:
            if app_type == "word":
                doc = app.Documents.Add()
            elif app_type == "excel":
                doc = app.Workbooks.Add()
            elif app_type == "ppt":
                doc = app.Presentations.Add()
            else:
                raise COMOperationError(f"不支持的应用类型: {app_type}")
        except Exception as e:
            raise COMOperationError("创建文档", str(e)) from e

        self._documents[path_str] = doc
        self._doc_types[path_str] = app_type
        logger.info(f"已创建文档: {path_str}")
        return doc

    def get_document(self, file_path: Path) -> Any:
        """获取已打开的文档."""
        path_str = str(file_path)
        doc = self._documents.get(path_str)
        if doc is None:
            raise DocumentNotOpenError(path_str)
        return doc

    def close_document(self, file_path: Path, save: bool = True) -> None:
        """关闭文档."""
        path_str = str(file_path)
        doc = self._documents.get(path_str)
        if doc is None:
            raise DocumentNotOpenError(path_str)

        app_type = self._doc_types.get(path_str)

        try:
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
        except Exception as e:
            logger.warning(f"关闭文档时出错: {e}")
        finally:
            self._documents.pop(path_str, None)
            self._doc_types.pop(path_str, None)
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
            # 强制终止进程
            process_names = ["WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE"]
            for proc in psutil.process_iter(["name"]):
                try:
                    if proc.info["name"] in process_names:
                        proc.kill()
                        result["force_killed"].append(proc.info["name"])
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass

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
