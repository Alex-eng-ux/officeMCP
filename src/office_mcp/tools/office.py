"""通用 Office MCP 工具."""

from mcp.server.fastmcp import FastMCP

from office_mcp.core.errors import OfficeMCPError
from office_mcp.core.office_manager import office_manager


def register_office_tools(mcp: FastMCP) -> None:
    """注册通用 Office 工具."""

    @mcp.tool()
    def office_status() -> dict:
        """查询 Office 应用状态.

        返回当前运行的 Office 应用和已打开的文档列表。
        """
        try:
            return office_manager.status()
        except OfficeMCPError as e:
            return {"error": str(e)}

    @mcp.tool()
    def office_cleanup(force: bool = False) -> dict:
        """清理 Office 实例.

        默认只关闭本 MCP 管理的 Office 实例和文档。
        force=true 时会强制终止所有 Office 进程（可能影响用户手动打开的 Office）。

        Args:
            force: 是否强制终止所有 Office 进程
        """
        try:
            result = office_manager.cleanup(force=force)
            if force:
                result["warning"] = "已强制终止 Office 进程，可能影响用户手动打开的文档"
            return result
        except OfficeMCPError as e:
            return {"error": str(e)}
