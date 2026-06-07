"""Office MCP server entrypoints."""

from __future__ import annotations

import logging

from office_mcp.compat import FastMCP
from office_mcp.tools.excel import register_excel_tools
from office_mcp.tools.office import register_office_tools
from office_mcp.tools.powerpoint import register_ppt_tools
from office_mcp.tools.word import register_word_tools

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)


def _build_server(
    name: str,
    *,
    include_word: bool = False,
    include_excel: bool = False,
    include_ppt: bool = False,
    include_office: bool = False,
) -> FastMCP:
    """Create a FastMCP server with the requested tool families."""
    mcp = FastMCP(name)
    if include_word:
        register_word_tools(mcp)
    if include_excel:
        register_excel_tools(mcp)
    if include_ppt:
        register_ppt_tools(mcp)
    if include_office:
        register_office_tools(mcp)
    return mcp


def build_full_server() -> FastMCP:
    """Build the legacy all-in-one Office MCP server."""
    return _build_server(
        "office-mcp",
        include_word=True,
        include_excel=True,
        include_ppt=True,
        include_office=True,
    )


def build_word_server() -> FastMCP:
    """Build the Word-focused Office MCP server."""
    return _build_server("office-word-mcp", include_word=True, include_office=True)


def build_excel_server() -> FastMCP:
    """Build the Excel-focused Office MCP server."""
    return _build_server("office-excel-mcp", include_excel=True, include_office=True)


def build_powerpoint_server() -> FastMCP:
    """Build the PowerPoint-focused Office MCP server."""
    return _build_server("office-powerpoint-mcp", include_ppt=True, include_office=True)


def main() -> None:
    """Run the legacy all-in-one Office MCP server."""
    build_full_server().run()


def main_word() -> None:
    """Run the Word-focused Office MCP server."""
    build_word_server().run()


def main_excel() -> None:
    """Run the Excel-focused Office MCP server."""
    build_excel_server().run()


def main_powerpoint() -> None:
    """Run the PowerPoint-focused Office MCP server."""
    build_powerpoint_server().run()


if __name__ == "__main__":
    main()
