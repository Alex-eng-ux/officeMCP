"""自定义异常类."""


class OfficeMCPError(Exception):
    """Office MCP Server 基类异常."""

    def __init__(self, message: str, suggestion: str = ""):
        super().__init__(message)
        self.message = message
        self.suggestion = suggestion

    def __str__(self) -> str:
        if self.suggestion:
            return f"{self.message}\n建议: {self.suggestion}"
        return self.message


class PathNotAllowedError(OfficeMCPError):
    """路径不在允许列表中."""

    def __init__(self, path: str):
        super().__init__(
            message=f"路径不在允许的操作范围内: {path}",
            suggestion="请使用绝对路径，并确保路径在 OFFICE_MCP_ALLOWED_DIRS 配置的目录内。",
        )


class DocumentNotOpenError(OfficeMCPError):
    """文档未打开."""

    def __init__(self, file_path: str):
        super().__init__(
            message=f"文档未打开: {file_path}",
            suggestion="请先使用 open 或 create 工具打开/创建文档。",
        )


class DocumentAlreadyOpenError(OfficeMCPError):
    """文档已被打开."""

    def __init__(self, file_path: str):
        super().__init__(
            message=f"文档已被打开: {file_path}",
            suggestion="如需重新打开，请先关闭该文档。",
        )


class COMOperationError(OfficeMCPError):
    """COM 调用失败."""

    def __init__(self, operation: str, detail: str = ""):
        msg = f"COM 操作失败: {operation}"
        if detail:
            msg += f" - {detail}"
        super().__init__(
            message=msg,
            suggestion="请检查 Office 是否正常运行，或尝试使用 office_cleanup 清理后重试。",
        )


class OfficeNotInstalledError(OfficeMCPError):
    """未检测到 Office 安装."""

    def __init__(self):
        super().__init__(
            message="未检测到 Microsoft Office 安装",
            suggestion="请确保已安装 Microsoft Office (Word, Excel, PowerPoint)。",
        )
