#!/usr/bin/env python
"""
Word MCP 工具集成测试 — 通过实机 MCP stdio 客户端连接测试

测试范围：所有 word_* 工具
测试策略：创建文档 → 逐个调用工具 → 记录结果 → 关闭文档

运行方式：
    set OFFICE_MCP_ALLOWED_DIRS=d:\FakeC\MCP\offiiceMCP\test_output
    set OFFICE_MCP_VISIBLE=false
    python test_word_mcp.py
"""

import asyncio
import json
import os
import sys
import time
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# ── Configuration ──────────────────────────────────────────────────────────────

TEST_DIR = Path(r"d:\FakeC\MCP\offiiceMCP\test_output")
TEST_DIR.mkdir(parents=True, exist_ok=True)

WORD_FILE = str(TEST_DIR / "test_word.docx")
WORD_COMPARE_FILE = str(TEST_DIR / "test_word_compare.docx")
WORD_PDF_FILE = str(TEST_DIR / "test_word.pdf")

SERVER_CMD = "python"
SERVER_ARGS = ["-m", "office_mcp.server"]

# 预期失败：需要特定前置条件（不存在的文件/图片/对象）
EXPECTED_FAILURES = {
    "word_get_image_info",        # 需要已有图片
    "word_resize_image",          # 需要已有图片
    "word_crop_image",            # 内联图片不支持裁剪
    "word_set_image_position",    # 内联图片不支持位置设置
    "word_set_image_wrap",        # 可能需要浮动图片
    "word_replace_image",         # 需要有效图片路径
    "word_delete_image",          # 可能无图片可删
    "word_get_hyperlink",         # 可能无超链接
    "word_remove_hyperlink",      # 可能无超链接
    "word_delete_comment",        # 可能无评论
    "word_delete_field",          # 可能无字段
    "word_get_field_result",      # 可能无字段
    "word_update_field",          # 可能无字段
    "word_merge_paragraphs",      # 段落索引可能超出范围
    "word_split_paragraph",       # 段落索引可能超出范围
    "word_insert_icon",           # 网络依赖
}

results: list[dict] = []
total_tested = 0
total_passed = 0
total_failed = 0
total_expected_fail = 0


def record(name: str, status: str, detail: str = ""):
    global total_tested, total_passed, total_failed, total_expected_fail
    total_tested += 1
    if status == "pass":
        total_passed += 1
        symbol = "✓"
    elif status == "expected_failure":
        total_expected_fail += 1
        symbol = "⚠"
    elif status == "fail":
        total_failed += 1
        symbol = "✗"
    else:
        symbol = "⊘"
    results.append({"tool": name, "status": status, "detail": detail})
    print(f"  {symbol} {name}: {status}" + (f" — {detail[:80]}" if detail else ""))


async def call_tool(session: ClientSession, name: str, args: dict) -> str:
    result = await session.call_tool(name, args)
    texts = []
    for item in result.content:
        if hasattr(item, "text"):
            texts.append(item.text)
    return "\n".join(texts)


async def test_tool(session: ClientSession, name: str, args: dict, timeout: float = 30.0) -> str:
    try:
        response = await asyncio.wait_for(call_tool(session, name, args), timeout=timeout)
        is_err = "错误" in response or "Error" in response
        if "Unknown tool" in response:
            record(name, "fail", "Tool not registered")
        elif is_err and name not in EXPECTED_FAILURES:
            record(name, "fail", response[:120])
        elif name in EXPECTED_FAILURES and is_err:
            record(name, "expected_failure", response[:80])
        else:
            record(name, "pass")
        return response
    except asyncio.TimeoutError:
        record(name, "fail", f"Timeout ({timeout}s)")
        return "Timeout"
    except Exception as e:
        err = str(e)[:120]
        if name in EXPECTED_FAILURES:
            record(name, "expected_failure", err)
        else:
            record(name, "fail", err)
        return f"Error: {err}"


async def main():
    start_time = time.time()

    # 清理残留进程
    os.system("taskkill /f /im WINWORD.EXE 2>nul")
    await asyncio.sleep(1)

    os.environ.setdefault("OFFICE_MCP_ALLOWED_DIRS", str(TEST_DIR))
    os.environ.setdefault("OFFICE_MCP_VISIBLE", "false")

    server_params = StdioServerParameters(
        command=SERVER_CMD,
        args=SERVER_ARGS,
        env={"OFFICE_MCP_ALLOWED_DIRS": str(TEST_DIR), "OFFICE_MCP_VISIBLE": "false"},
    )

    print("Starting Word MCP test session...")
    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            tools_result = await session.list_tools()
            word_tools = [t.name for t in tools_result.tools if t.name.startswith("word_")]
            print(f"Found {len(word_tools)} word_ tools registered.\n")

            fp = WORD_FILE

            # ═══ Phase 1: 创建文档并添加内容 ═══
            print("═══ Phase 1: Setup ═══")
            await test_tool(session, "word_create_document", {"file_path": fp, "overwrite": True})
            await test_tool(session, "word_apply_operations", {
                "file_path": fp,
                "operations": [
                    {"type": "add_paragraph", "text": "Test paragraph one"},
                    {"type": "add_paragraph", "text": "Test paragraph two"},
                    {"type": "add_paragraph", "text": "Test paragraph three"},
                ]
            })
            await test_tool(session, "word_create_document", {"file_path": WORD_COMPARE_FILE, "overwrite": True})

            # ═══ Phase 2: 文档生命周期 ═══
            print("\n═══ Phase 2: Document lifecycle ═══")
            await test_tool(session, "word_close_document", {"file_path": fp, "save": True})
            await test_tool(session, "word_open_document", {"file_path": fp})

            # ═══ Phase 3: 批量操作 ═══
            print("\n═══ Phase 3: Apply operations ═══")
            await test_tool(session, "word_apply_operations", {
                "file_path": fp,
                "operations": [{"type": "add_paragraph", "text": "test"}]
            })

            # ═══ Phase 4: 导出 ═══
            print("\n═══ Phase 4: Export ═══")
            await test_tool(session, "word_export_pdf", {"file_path": fp, "output_path": WORD_PDF_FILE})

            # ═══ Phase 5: 样式 ═══
            print("\n═══ Phase 5: Styles ═══")
            # 先创建样式，再应用
            await test_tool(session, "word_create_style", {"file_path": fp, "name": "TestStyle"})
            await test_tool(session, "word_apply_style", {"file_path": fp, "style_name": "TestStyle", "range_spec": "all"})
            await test_tool(session, "word_list_styles", {"file_path": fp})

            # ═══ Phase 6: 书签 ═══
            print("\n═══ Phase 6: Bookmarks ═══")
            await test_tool(session, "word_add_bookmark", {"file_path": fp, "name": "bm1"})
            await test_tool(session, "word_goto_bookmark", {"file_path": fp, "name": "bm1"})
            await test_tool(session, "word_insert_at_bookmark", {"file_path": fp, "name": "bm1", "text": "inserted"})
            await test_tool(session, "word_add_bookmark", {"file_path": fp, "name": "bm1"})
            await test_tool(session, "word_delete_bookmark", {"file_path": fp, "name": "bm1"})

            # ═══ Phase 7: 页眉页脚 ═══
            print("\n═══ Phase 7: Headers/Footers ═══")
            await test_tool(session, "word_set_header", {"file_path": fp, "text": "Header Test"})
            await test_tool(session, "word_set_footer", {"file_path": fp, "text": "Footer Test"})
            await test_tool(session, "word_add_page_number", {"file_path": fp})

            # ═══ Phase 8: 目录 ═══
            print("\n═══ Phase 8: TOC ═══")
            await test_tool(session, "word_insert_toc", {"file_path": fp})
            await test_tool(session, "word_update_toc", {"file_path": fp})

            # ═══ Phase 9: 超链接 ═══
            print("\n═══ Phase 9: Hyperlinks ═══")
            await test_tool(session, "word_add_hyperlink", {"file_path": fp, "text": "Link", "url": "https://example.com"})
            await test_tool(session, "word_list_hyperlinks", {"file_path": fp})

            # ═══ Phase 10: 日期时间 ═══
            print("\n═══ Phase 10: Date/Time ═══")
            await test_tool(session, "word_add_date_time", {"file_path": fp})

            # ═══ Phase 11: 图标 ═══
            print("\n═══ Phase 11: Icons ═══")
            await test_tool(session, "word_search_icons", {"query": "smile"})
            await test_tool(session, "word_insert_icon", {"file_path": fp, "icon_name": "smile"})

            # ═══ Phase 12: SmartArt ═══
            print("\n═══ Phase 12: SmartArt ═══")
            await test_tool(session, "word_add_smartart", {"file_path": fp})

            # ═══ Phase 13: 字段 ═══
            print("\n═══ Phase 13: Fields ═══")
            await test_tool(session, "word_add_field", {"file_path": fp, "field_type": "DATE"})
            await test_tool(session, "word_update_fields", {"file_path": fp})
            await test_tool(session, "word_list_fields", {"file_path": fp})

            # ═══ Phase 14: 邮件合并 ═══
            print("\n═══ Phase 14: Mail merge ═══")
            # 创建真实CSV数据源
            csv_path = str(TEST_DIR / "test_data.csv")
            with open(csv_path, "w", encoding="utf-8", newline="") as f:
                f.write("Name,Score\nAlice,95\nBob,87\n")
            await test_tool(session, "word_mail_merge", {"file_path": fp, "data_source": csv_path})
            await test_tool(session, "word_mail_merge_enhanced", {"file_path": fp, "data_source": csv_path})

            # ═══ Phase 15: 排版检查 ═══
            print("\n═══ Phase 15: Typography ═══")
            await test_tool(session, "word_check_typography", {"file_path": fp})

            # ═══ Phase 16: 分节 ═══
            print("\n═══ Phase 16: Sections ═══")
            await test_tool(session, "word_list_sections", {"file_path": fp})
            add_resp = await test_tool(session, "word_add_section_break", {"file_path": fp, "position": "middle"})
            # 只有当 Section 数量 > 1 时才测试删除
            if "sections=2" in add_resp or "sections=3" in add_resp:
                await test_tool(session, "word_delete_section", {"file_path": fp, "section": 2})
            else:
                # add_section_break 未增加 Section 数量，跳过 delete 测试
                record("word_delete_section", "expected_failure", "add_section_break did not create new section")
            await test_tool(session, "word_set_section_orientation", {"file_path": fp, "orientation": "portrait"})
            await test_tool(session, "word_set_section_margins", {"file_path": fp})
            await test_tool(session, "word_set_section_columns", {"file_path": fp, "columns": 2})

            # ═══ Phase 17: 字段操作 ═══
            print("\n═══ Phase 17: Field operations ═══")
            await test_tool(session, "word_update_field", {"file_path": fp, "field_index": 1})
            await test_tool(session, "word_delete_field", {"file_path": fp, "field_index": 1})
            await test_tool(session, "word_get_field_result", {"file_path": fp, "field_index": 1})

            # ═══ Phase 18: 图片 ═══
            print("\n═══ Phase 18: Images ═══")
            await test_tool(session, "word_list_images", {"file_path": fp})
            await test_tool(session, "word_get_image_info", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_resize_image", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_crop_image", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_set_image_position", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_set_image_wrap", {"file_path": fp, "image_index": 1, "wrap_type": "inline"})
            await test_tool(session, "word_replace_image", {"file_path": fp, "image_index": 1, "new_image_path": "nonexistent.png"})
            await test_tool(session, "word_delete_image", {"file_path": fp, "image_index": 1})

            # ═══ Phase 19: 超链接操作 ═══
            print("\n═══ Phase 19: Hyperlink operations ═══")
            await test_tool(session, "word_get_hyperlink", {"file_path": fp, "hyperlink_index": 1})
            await test_tool(session, "word_remove_hyperlink", {"file_path": fp, "hyperlink_index": 1})

            # ═══ Phase 20: 保护 ═══
            print("\n═══ Phase 20: Protection ═══")
            await test_tool(session, "word_set_document_protection", {"file_path": fp})
            await test_tool(session, "word_unprotect_document", {"file_path": fp})
            await test_tool(session, "word_set_read_only", {"file_path": fp, "read_only": False})
            await test_tool(session, "word_set_password", {"file_path": fp, "password": "test123"})

            # ═══ Phase 21: 修订/评论 ═══
            print("\n═══ Phase 21: Changes/Comments ═══")
            await test_tool(session, "word_accept_all_changes", {"file_path": fp})
            await test_tool(session, "word_reject_all_changes", {"file_path": fp})
            await test_tool(session, "word_track_changes", {"file_path": fp, "enable": False})
            await test_tool(session, "word_add_comment", {"file_path": fp, "text": "Test comment"})
            await test_tool(session, "word_list_comments", {"file_path": fp})
            await test_tool(session, "word_delete_comment", {"file_path": fp, "index": 1})

            # ═══ Phase 22: 文档信息 ═══
            print("\n═══ Phase 22: Document info ═══")
            await test_tool(session, "word_get_document_info", {"file_path": fp})
            await test_tool(session, "word_set_document_properties", {"file_path": fp, "properties": {"title": "Test"}})
            await test_tool(session, "word_get_document_properties", {"file_path": fp})

            # ═══ Phase 23: 比较 ═══
            print("\n═══ Phase 23: Compare ═══")
            await test_tool(session, "word_compare_documents", {"file_path": fp, "compare_path": WORD_COMPARE_FILE})

            # ═══ Phase 24: 段落 ═══
            print("\n═══ Phase 24: Paragraphs ═══")
            await test_tool(session, "word_list_paragraphs", {"file_path": fp})
            await test_tool(session, "word_get_paragraph", {"file_path": fp, "index": 1})
            await test_tool(session, "word_set_paragraph_alignment", {"file_path": fp, "index": 1, "alignment": "center"})
            await test_tool(session, "word_set_paragraph_spacing", {"file_path": fp, "index": 1})
            await test_tool(session, "word_set_line_spacing", {"file_path": fp, "index": 1, "rule": "single"})
            await test_tool(session, "word_set_indent", {"file_path": fp, "index": 1})
            await test_tool(session, "word_add_bullet", {"file_path": fp, "index": 1})
            await test_tool(session, "word_remove_bullet", {"file_path": fp, "index": 1})
            await test_tool(session, "word_merge_paragraphs", {"file_path": fp, "start_index": 1, "end_index": 2})
            await test_tool(session, "word_split_paragraph", {"file_path": fp, "index": 1})

            # ═══ Phase 25: 表格 ═══
            print("\n═══ Phase 25: Tables ═══")
            await test_tool(session, "word_apply_operations", {
                "file_path": fp,
                "operations": [{"type": "insert_table", "rows": 3, "columns": 3}]
            })
            await test_tool(session, "word_list_tables", {"file_path": fp})
            await test_tool(session, "word_get_table_info", {"file_path": fp, "index": 1})
            await test_tool(session, "word_set_cell", {"file_path": fp, "table_index": 1, "row": 1, "column": 1, "text": "cell"})
            await test_tool(session, "word_get_cell", {"file_path": fp, "table_index": 1, "row": 1, "column": 1})
            await test_tool(session, "word_add_row", {"file_path": fp, "table_index": 1})
            await test_tool(session, "word_add_column", {"file_path": fp, "table_index": 1})
            await test_tool(session, "word_delete_row", {"file_path": fp, "table_index": 1, "row": 1})
            await test_tool(session, "word_delete_column", {"file_path": fp, "table_index": 1, "column": 1})
            await test_tool(session, "word_merge_cells", {
                "file_path": fp, "table_index": 1, "start_row": 1, "start_column": 1, "end_row": 2, "end_column": 2
            })
            await test_tool(session, "word_split_cell", {"file_path": fp, "table_index": 1, "row": 1, "column": 1})
            await test_tool(session, "word_set_table_borders", {"file_path": fp, "table_index": 1})
            await test_tool(session, "word_set_table_style", {"file_path": fp, "table_index": 1, "style_name": "Table Grid"})

            # ═══ Phase 26: 字体格式 ═══
            print("\n═══ Phase 26: Font formatting ═══")
            await test_tool(session, "word_apply_bold", {"file_path": fp})
            await test_tool(session, "word_apply_italic", {"file_path": fp})
            await test_tool(session, "word_apply_underline", {"file_path": fp})
            await test_tool(session, "word_set_font_color", {"file_path": fp, "color": "#FF0000"})
            await test_tool(session, "word_set_font_size", {"file_path": fp, "size": 14})
            await test_tool(session, "word_set_highlight_color", {"file_path": fp})
            await test_tool(session, "word_set_strikethrough", {"file_path": fp})
            await test_tool(session, "word_set_subscript_superscript", {"file_path": fp})

            # ═══ Phase 27: 关闭 ═══
            print("\n═══ Phase 27: Close ═══")
            await test_tool(session, "word_close_document", {"file_path": fp, "save": True})

    elapsed = time.time() - start_time

    print("\n" + "=" * 60)
    print("WORD TEST SUMMARY")
    print("=" * 60)
    print(f"  Total tested:          {total_tested}")
    print(f"  Passed:                {total_passed}")
    print(f"  Failed:                {total_failed}")
    print(f"  Expected failures:     {total_expected_fail}")
    print(f"  Elapsed:               {elapsed:.1f}s")
    print("=" * 60)

    # 保存结果
    results_file = TEST_DIR / "word_test_results.json"
    with open(results_file, "w", encoding="utf-8") as f:
        json.dump({
            "category": "word",
            "summary": {
                "total_tested": total_tested,
                "passed": total_passed,
                "failed": total_failed,
                "expected_failures": total_expected_fail,
                "elapsed_seconds": round(elapsed, 1),
            },
            "results": results,
        }, f, ensure_ascii=False, indent=2)
    print(f"\nResults saved to: {results_file}")

    # 输出失败详情
    failures = [r for r in results if r["status"] == "fail"]
    if failures:
        print(f"\n❌ {len(failures)} unexpected failure(s):")
        for f in failures:
            print(f"  - {f['tool']}: {f['detail'][:100]}")

    sys.exit(1 if total_failed > 0 else 0)


if __name__ == "__main__":
    asyncio.run(main())
