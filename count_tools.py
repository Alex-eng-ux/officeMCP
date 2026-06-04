"""统计工具数量."""
import sys
sys.path.insert(0, 'src')
from office_mcp.server import mcp
tools = mcp._tool_manager._tools
names = sorted(tools.keys())
print(f"总工具数: {len(tools)}")
print(f"Excel: {len([n for n in names if n.startswith('excel_')])}")
print(f"Word: {len([n for n in names if n.startswith('word_')])}")
print(f"PPT: {len([n for n in names if n.startswith('ppt_')])}")
print(f"其他: {len([n for n in names if not n.startswith(('excel_', 'word_', 'ppt_'))])}")
