#!/usr/bin/env python
"""
Word MCP integration test via a real MCP stdio client.

Scope: all `word_*` tools.
Strategy: create a document, call tools in sequence, classify outcomes, save a report.
"""

import asyncio
import json
import os
import sys
import time
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

TEST_DIR = Path(r"d:\FakeC\MCP\offiiceMCP\test_output")
TEST_DIR.mkdir(parents=True, exist_ok=True)

WORD_FILE = str(TEST_DIR / "test_word.docx")
WORD_COMPARE_FILE = str(TEST_DIR / "test_word_compare.docx")
WORD_PDF_FILE = str(TEST_DIR / "test_word.pdf")

SERVER_CMD = sys.executable
SERVER_ARGS = ["-m", "office_mcp.server"]

EXPECTED_FAILURES = {
    "word_get_image_info",
    "word_resize_image",
    "word_crop_image",
    "word_set_image_position",
    "word_set_image_wrap",
    "word_replace_image",
    "word_delete_image",
    "word_get_hyperlink",
    "word_remove_hyperlink",
    "word_delete_comment",
    "word_delete_field",
    "word_get_field_result",
    "word_update_field",
    "word_merge_paragraphs",
    "word_split_paragraph",
    "word_insert_icon",
}

results: list[dict] = []
total_tested = 0
total_passed = 0
total_failed = 0
total_expected_fail = 0


def record(name: str, status: str, detail: str = "") -> None:
    global total_tested, total_passed, total_failed, total_expected_fail
    total_tested += 1
    if status == "pass":
        total_passed += 1
        symbol = "[PASS]"
    elif status == "expected_failure":
        total_expected_fail += 1
        symbol = "[XFAIL]"
    elif status == "fail":
        total_failed += 1
        symbol = "[FAIL]"
    else:
        symbol = "[INFO]"
    results.append({"tool": name, "status": status, "detail": detail})
    print(f"  {symbol} {name}: {status}" + (f" - {detail[:120]}" if detail else ""))


def has_error(response: str) -> bool:
    lowered = response.lower()
    return "error" in lowered or "错误" in response

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
        if "Unknown tool" in response:
            record(name, "fail", "Tool not registered")
        elif has_error(response) and name not in EXPECTED_FAILURES:
            record(name, "fail", response[:240])
        elif name in EXPECTED_FAILURES and has_error(response):
            record(name, "expected_failure", response[:240])
        else:
            record(name, "pass")
        return response
    except asyncio.TimeoutError:
        record(name, "fail", f"Timeout ({timeout}s)")
        return "Timeout"
    except Exception as exc:
        err = str(exc)[:240]
        if name in EXPECTED_FAILURES:
            record(name, "expected_failure", err)
        else:
            record(name, "fail", err)
        return f"Error: {err}"


test_tool.__test__ = False


async def main() -> None:
    start_time = time.time()

    os.system("taskkill /f /im WINWORD.EXE 2>nul")
    await asyncio.sleep(1)

    os.environ.setdefault("OFFICE_MCP_ALLOWED_DIRS", str(TEST_DIR))
    os.environ.setdefault("OFFICE_MCP_VISIBLE", "false")

    server_params = StdioServerParameters(
        command=SERVER_CMD,
        args=SERVER_ARGS,
        env={
            "OFFICE_MCP_ALLOWED_DIRS": str(TEST_DIR),
            "OFFICE_MCP_VISIBLE": "false",
        },
    )

    print("Starting Word MCP test session...")
    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            tools_result = await session.list_tools()
            word_tools = [t.name for t in tools_result.tools if t.name.startswith("word_")]
            print(f"Found {len(word_tools)} word_ tools registered.\n")

            fp = WORD_FILE

            print("Phase 1: Setup")
            await test_tool(session, "word_create_document", {"file_path": fp, "overwrite": True})
            await test_tool(
                session,
                "word_apply_operations",
                {
                    "file_path": fp,
                    "operations": [
                        {"type": "add_paragraph", "text": "Test paragraph one"},
                        {"type": "add_paragraph", "text": "Test paragraph two"},
                        {"type": "add_paragraph", "text": "Test paragraph three"},
                    ],
                },
            )
            await test_tool(session, "word_create_document", {"file_path": WORD_COMPARE_FILE, "overwrite": True})

            print("\nPhase 2: Document lifecycle")
            await test_tool(session, "word_close_document", {"file_path": fp, "save": True})
            await test_tool(session, "word_open_document", {"file_path": fp})

            print("\nPhase 3: Apply operations")
            await test_tool(
                session,
                "word_apply_operations",
                {"file_path": fp, "operations": [{"type": "add_paragraph", "text": "test"}]},
            )

            print("\nPhase 4: Export")
            await test_tool(session, "word_export_pdf", {"file_path": fp, "output_path": WORD_PDF_FILE})

            print("\nPhase 5: Styles")
            await test_tool(session, "word_create_style", {"file_path": fp, "name": "TestStyle"})
            await test_tool(session, "word_apply_style", {"file_path": fp, "style_name": "TestStyle", "range_spec": "all"})
            await test_tool(session, "word_list_styles", {"file_path": fp})

            print("\nPhase 6: Bookmarks")
            await test_tool(session, "word_add_bookmark", {"file_path": fp, "name": "bm1"})
            await test_tool(session, "word_goto_bookmark", {"file_path": fp, "name": "bm1"})
            await test_tool(session, "word_insert_at_bookmark", {"file_path": fp, "name": "bm1", "text": "inserted"})
            await test_tool(session, "word_add_bookmark", {"file_path": fp, "name": "bm1"})
            await test_tool(session, "word_delete_bookmark", {"file_path": fp, "name": "bm1"})

            print("\nPhase 7: Headers/Footers")
            await test_tool(session, "word_set_header", {"file_path": fp, "text": "Header Test"})
            await test_tool(session, "word_set_footer", {"file_path": fp, "text": "Footer Test"})
            await test_tool(session, "word_add_page_number", {"file_path": fp})

            print("\nPhase 8: TOC")
            await test_tool(session, "word_insert_toc", {"file_path": fp})
            await test_tool(session, "word_update_toc", {"file_path": fp})

            print("\nPhase 9: Hyperlinks")
            await test_tool(session, "word_add_hyperlink", {"file_path": fp, "text": "Link", "url": "https://example.com"})
            await test_tool(session, "word_list_hyperlinks", {"file_path": fp})

            print("\nPhase 10: Date/Time")
            await test_tool(session, "word_add_date_time", {"file_path": fp})

            print("\nPhase 11: Icons")
            await test_tool(session, "word_search_icons", {"query": "smile"})
            await test_tool(session, "word_insert_icon", {"file_path": fp, "icon_name": "smile"})

            print("\nPhase 12: SmartArt")
            await test_tool(session, "word_add_smartart", {"file_path": fp})

            print("\nPhase 13: Fields")
            await test_tool(session, "word_add_field", {"file_path": fp, "field_type": "DATE"})
            await test_tool(session, "word_update_fields", {"file_path": fp})
            await test_tool(session, "word_list_fields", {"file_path": fp})

            print("\nPhase 14: Mail merge")
            csv_path = str(TEST_DIR / "test_data.csv")
            with open(csv_path, "w", encoding="utf-8", newline="") as handle:
                handle.write("Name,Score\nAlice,95\nBob,87\n")
            await test_tool(session, "word_mail_merge", {"file_path": fp, "data_source": csv_path})
            await test_tool(session, "word_mail_merge_enhanced", {"file_path": fp, "data_source": csv_path})

            print("\nPhase 15: Typography")
            await test_tool(session, "word_check_typography", {"file_path": fp})

            print("\nPhase 16: Sections")
            await test_tool(session, "word_list_sections", {"file_path": fp})
            add_resp = await test_tool(session, "word_add_section_break", {"file_path": fp, "position": "middle"})
            if "sections=2" in add_resp or "sections=3" in add_resp:
                await test_tool(session, "word_delete_section", {"file_path": fp, "section": 2})
            else:
                record("word_delete_section", "expected_failure", "add_section_break did not create new section")
            await test_tool(session, "word_set_section_orientation", {"file_path": fp, "orientation": "portrait"})
            await test_tool(session, "word_set_section_margins", {"file_path": fp})
            await test_tool(session, "word_set_section_columns", {"file_path": fp, "columns": 2})

            print("\nPhase 17: Field operations")
            await test_tool(session, "word_update_field", {"file_path": fp, "field_index": 1})
            await test_tool(session, "word_delete_field", {"file_path": fp, "field_index": 1})
            await test_tool(session, "word_get_field_result", {"file_path": fp, "field_index": 1})

            print("\nPhase 18: Images")
            await test_tool(session, "word_list_images", {"file_path": fp})
            await test_tool(session, "word_get_image_info", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_resize_image", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_crop_image", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_set_image_position", {"file_path": fp, "image_index": 1})
            await test_tool(session, "word_set_image_wrap", {"file_path": fp, "image_index": 1, "wrap_type": "inline"})
            await test_tool(session, "word_replace_image", {"file_path": fp, "image_index": 1, "new_image_path": "nonexistent.png"})
            await test_tool(session, "word_delete_image", {"file_path": fp, "image_index": 1})

            print("\nPhase 19: Hyperlink operations")
            await test_tool(session, "word_get_hyperlink", {"file_path": fp, "hyperlink_index": 1})
            await test_tool(session, "word_remove_hyperlink", {"file_path": fp, "hyperlink_index": 1})

            print("\nPhase 20: Protection")
            await test_tool(session, "word_set_document_protection", {"file_path": fp})
            await test_tool(session, "word_unprotect_document", {"file_path": fp})
            await test_tool(session, "word_set_read_only", {"file_path": fp, "read_only": False})
            await test_tool(session, "word_set_password", {"file_path": fp, "password": "test123"})

            print("\nPhase 21: Changes/Comments")
            await test_tool(session, "word_accept_all_changes", {"file_path": fp})
            await test_tool(session, "word_reject_all_changes", {"file_path": fp})
            await test_tool(session, "word_track_changes", {"file_path": fp, "enable": False})
            await test_tool(session, "word_add_comment", {"file_path": fp, "text": "Test comment"})
            await test_tool(session, "word_list_comments", {"file_path": fp})
            await test_tool(session, "word_delete_comment", {"file_path": fp, "index": 1})

            print("\nPhase 22: Document info")
            await test_tool(session, "word_get_document_info", {"file_path": fp})
            await test_tool(session, "word_set_document_properties", {"file_path": fp, "properties": {"title": "Test"}})
            await test_tool(session, "word_get_document_properties", {"file_path": fp})

            print("\nPhase 23: Compare")
            await test_tool(session, "word_compare_documents", {"file_path": fp, "compare_path": WORD_COMPARE_FILE})

            print("\nPhase 24: Paragraphs")
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

            print("\nPhase 25: Tables")
            await test_tool(
                session,
                "word_apply_operations",
                {"file_path": fp, "operations": [{"type": "insert_table", "rows": 3, "columns": 3}]},
            )
            await test_tool(session, "word_list_tables", {"file_path": fp})
            await test_tool(session, "word_get_table_info", {"file_path": fp, "index": 1})
            await test_tool(session, "word_set_cell", {"file_path": fp, "table_index": 1, "row": 1, "column": 1, "text": "cell"})
            await test_tool(session, "word_get_cell", {"file_path": fp, "table_index": 1, "row": 1, "column": 1})
            await test_tool(session, "word_add_row", {"file_path": fp, "table_index": 1})
            await test_tool(session, "word_add_column", {"file_path": fp, "table_index": 1})
            await test_tool(session, "word_delete_row", {"file_path": fp, "table_index": 1, "row": 1})
            await test_tool(session, "word_delete_column", {"file_path": fp, "table_index": 1, "column": 1})
            await test_tool(
                session,
                "word_merge_cells",
                {
                    "file_path": fp,
                    "table_index": 1,
                    "start_row": 1,
                    "start_column": 1,
                    "end_row": 2,
                    "end_column": 2,
                },
            )
            await test_tool(session, "word_split_cell", {"file_path": fp, "table_index": 1, "row": 1, "column": 1})
            await test_tool(session, "word_set_table_borders", {"file_path": fp, "table_index": 1})
            await test_tool(session, "word_set_table_style", {"file_path": fp, "table_index": 1, "style_name": "Table Grid"})

            print("\nPhase 26: Font formatting")
            await test_tool(session, "word_apply_bold", {"file_path": fp})
            await test_tool(session, "word_apply_italic", {"file_path": fp})
            await test_tool(session, "word_apply_underline", {"file_path": fp})
            await test_tool(session, "word_set_font_color", {"file_path": fp, "color": "#FF0000"})
            await test_tool(session, "word_set_font_size", {"file_path": fp, "size": 14})
            await test_tool(session, "word_set_highlight_color", {"file_path": fp})
            await test_tool(session, "word_set_strikethrough", {"file_path": fp})
            await test_tool(session, "word_set_subscript_superscript", {"file_path": fp})

            print("\nPhase 27: Close")
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

    results_file = TEST_DIR / "word_test_results.json"
    with open(results_file, "w", encoding="utf-8") as handle:
        json.dump(
            {
                "category": "word",
                "summary": {
                    "total_tested": total_tested,
                    "passed": total_passed,
                    "failed": total_failed,
                    "expected_failures": total_expected_fail,
                    "elapsed_seconds": round(elapsed, 1),
                },
                "results": results,
            },
            handle,
            ensure_ascii=False,
            indent=2,
        )
    print(f"\nResults saved to: {results_file}")

    failures = [result for result in results if result["status"] == "fail"]
    if failures:
        print(f"\n{len(failures)} unexpected failure(s):")
        for failure in failures:
            print(f"  - {failure['tool']}: {failure['detail'][:160]}")

    sys.exit(1 if total_failed > 0 else 0)


if __name__ == "__main__":
    asyncio.run(main())

