"""Custom exceptions for Office MCP."""

from __future__ import annotations


class OfficeMCPError(Exception):
    """Base exception for Office MCP server errors."""

    def __init__(self, message: str, suggestion: str = ""):
        super().__init__(message)
        self.message = message
        self.suggestion = suggestion

    def __str__(self) -> str:
        if self.suggestion:
            return f"{self.message}\nSuggestion: {self.suggestion}"
        return self.message


class PathNotAllowedError(OfficeMCPError):
    """Raised when a path falls outside the allowed directories."""

    def __init__(self, path: str, allowed_dirs: list[str] | None = None):
        suggestion = (
            "Use an absolute path inside OFFICE_MCP_ALLOWED_DIRS or a detected workspace directory."
        )
        if allowed_dirs:
            suggestion += f" Effective allowed directories: {', '.join(allowed_dirs)}"
        super().__init__(
            message=f"Path is outside the allowed operation scope: {path}",
            suggestion=suggestion,
        )


class DocumentNotOpenError(OfficeMCPError):
    """Raised when the target document is not open."""

    def __init__(self, file_path: str):
        super().__init__(
            message=f"Document is not open: {file_path}",
            suggestion="Open or create the document first, or activate it with office_activate.",
        )


class DocumentAlreadyOpenError(OfficeMCPError):
    """Raised when the target document is already open."""

    def __init__(self, file_path: str):
        super().__init__(
            message=f"Document is already open: {file_path}",
            suggestion="Close the document first if you need to reopen it.",
        )


class COMOperationError(OfficeMCPError):
    """Raised when a COM automation call fails."""

    def __init__(self, operation: str, detail: str = ""):
        message = f"COM operation failed: {operation}"
        if detail:
            message += f" - {detail}"
        super().__init__(
            message=message,
            suggestion="Check that Microsoft Office is installed and responsive, then retry. "
            "If Office is stuck behind a dialog, run office_cleanup before retrying.",
        )


class OfficeNotInstalledError(OfficeMCPError):
    """Raised when Microsoft Office cannot be reached through COM."""

    def __init__(self, detail: str = "", suggestion: str = ""):
        message = "Microsoft Office does not appear to be installed or available through COM"
        if detail:
            message += f" - {detail}"
        super().__init__(
            message=message,
            suggestion=suggestion or "Install Microsoft Office with Word, Excel, and PowerPoint available.",
        )
