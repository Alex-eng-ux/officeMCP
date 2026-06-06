"""导出所有 MCP 工具的名称和参数签名到 JSON 文件。"""

import asyncio
import json
import os
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


async def main():
    env = os.environ.copy()
    env["OFFICE_MCP_ALLOWED_DIRS"] = str(Path(__file__).parent.resolve())
    env["OFFICE_MCP_VISIBLE"] = "false"

    server_params = StdioServerParameters(
        command=sys.executable,
        args=["-m", "office_mcp.server"],
        env=env,
    )

    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            tools_result = await session.list_tools()

            tools_info = []
            for t in tools_result.tools:
                params = {}
                if t.inputSchema and "properties" in t.inputSchema:
                    for pname, pval in t.inputSchema["properties"].items():
                        params[pname] = {
                            "type": pval.get("type", "any"),
                            "description": pval.get("description", "")[:80],
                        }
                required = t.inputSchema.get("required", []) if t.inputSchema else []
                tools_info.append({
                    "name": t.name,
                    "description": (t.description or "")[:120],
                    "params": params,
                    "required": required,
                })

            output = Path(__file__).parent / "test_output" / "all_tools.json"
            output.parent.mkdir(exist_ok=True)
            with open(output, "w", encoding="utf-8") as f:
                json.dump(tools_info, f, ensure_ascii=False, indent=2)
            print(f"Exported {len(tools_info)} tools to {output}")


if __name__ == "__main__":
    asyncio.run(main())
