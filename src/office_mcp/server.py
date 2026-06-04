"""Office MCP Server 入口."""

import logging

from mcp.server.fastmcp import FastMCP

from office_mcp.tools.excel import register_excel_tools
from office_mcp.tools.office import register_office_tools
from office_mcp.tools.powerpoint import register_ppt_tools
from office_mcp.tools.word import register_word_tools

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)

# 创建 FastMCP 服务器
mcp = FastMCP("office-mcp")

# 注册所有工具
register_word_tools(mcp)
register_excel_tools(mcp)
register_ppt_tools(mcp)
register_office_tools(mcp)


def main():
    """启动 MCP 服务器."""
    mcp.run()


if __name__ == "__main__":
    main()
