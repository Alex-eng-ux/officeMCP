#!/usr/bin/env python
"""
Excel MCP integration test via a real MCP stdio client.

Scope: all `excel_*` tools.
Strategy: create a workbook, call tools in sequence, classify outcomes, and save a report.
"""

import asyncio
import json
import os
import struct
import sys
import time
import zlib
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

# Configuration

TEST_DIR = Path(r"d:\FakeC\MCP\offiiceMCP\test_output")
TEST_DIR.mkdir(parents=True, exist_ok=True)

EXCEL_FILE = str(TEST_DIR / "test_excel.xlsx")
EXCEL_PDF_FILE = str(TEST_DIR / "test_excel.pdf")
EXCEL_EXPORT_FILE = str(TEST_DIR / "test_excel_export.csv")

SERVER_CMD = sys.executable
SERVER_ARGS = ["-m", "office_mcp.server"]

# Expected failures for unstable or unsupported Excel/COM operations.
EXPECTED_FAILURES = {
    # COM-sensitive operations.
    "excel_add_data_validation",
    "excel_add_chart_series",
    # Known unsupported or flaky workflow coverage.
    "excel_add_subtotal",
    "excel_import_data",
    # Operations skipped later because they can trigger blocking COM dialogs.
    "excel_set_open_password",
    "excel_set_write_reservation_password",
    "excel_mark_as_final",
    "excel_recommend_read_only",
    # Save/export variants may still surface workbook state issues.
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
    return "error" in lowered or "exception" in lowered or "traceback" in lowered


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
    except Exception as e:
        err = str(e)[:240]
        if name in EXPECTED_FAILURES:
            record(name, "expected_failure", err)
        else:
            record(name, "fail", err)
        return f"Error: {err}"


test_tool.__test__ = False


async def main():
    start_time = time.time()

    # Clean up leftover Excel processes from prior runs.
    os.system("taskkill /f /im EXCEL.EXE 2>nul")
    await asyncio.sleep(1)

    os.environ.setdefault("OFFICE_MCP_ALLOWED_DIRS", str(TEST_DIR))
    os.environ.setdefault("OFFICE_MCP_VISIBLE", "false")

    server_params = StdioServerParameters(
        command=SERVER_CMD,
        args=SERVER_ARGS,
        env={"OFFICE_MCP_ALLOWED_DIRS": str(TEST_DIR), "OFFICE_MCP_VISIBLE": "false"},
    )

    print("Starting Excel MCP test session...")
    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            await session.initialize()
            tools_result = await session.list_tools()
            excel_tools = [t.name for t in tools_result.tools if t.name.startswith("excel_")]
            print(f"Found {len(excel_tools)} excel_ tools registered.\n")

            fp = EXCEL_FILE
            sh = "Sheet1"

            print("Phase 1: Setup")
            await test_tool(session, "excel_create_workbook", {"file_path": fp, "overwrite": True})
            await test_tool(session, "excel_apply_operations", {
                "file_path": fp,
                "operations": [
                    {"type": "write_cell", "sheet_name": sh, "row": 1, "col": 1, "value": "Name"},
                    {"type": "write_cell", "sheet_name": sh, "row": 1, "col": 2, "value": "Score"},
                    {"type": "write_cell", "sheet_name": sh, "row": 2, "col": 1, "value": "Alice"},
                    {"type": "write_cell", "sheet_name": sh, "row": 2, "col": 2, "value": "95"},
                    {"type": "write_cell", "sheet_name": sh, "row": 3, "col": 1, "value": "Bob"},
                    {"type": "write_cell", "sheet_name": sh, "row": 3, "col": 2, "value": "87"},
                ]
            })

            print("\nPhase 2: Apply operations")
            await test_tool(session, "excel_apply_operations", {
                "file_path": fp,
                "operations": [{"type": "write_cell", "sheet_name": sh, "row": 1, "col": 1, "value": "test"}]
            })

            print("\nPhase 3: Export")
            await test_tool(session, "excel_export_pdf", {"file_path": fp, "output_path": EXCEL_PDF_FILE})

            print("\nPhase 4: Data validation / Conditional format / Pivot / Slicer / Subtotal")
            await test_tool(session, "excel_add_data_validation", {
                "file_path": fp, "sheet": sh, "range": "C1:C10", "validation_type": "list", "formula1": "a,b,c"
            })
            await test_tool(session, "excel_add_conditional_format", {
                "file_path": fp, "sheet": sh, "range": "C1:C10", "condition_type": "cell_value",
                "operator": "greater", "formula1": "0"
            })
            await test_tool(session, "excel_create_pivot_table", {
                "file_path": fp,
                "source_sheet": sh,
                "source_range": "",
                "target_sheet": "Pivot",
                "target_cell": "A3",
                "row_fields": ["Name"],
                "data_fields": {"Score": "average"},
                "pivot_name": "PivotScores",
            })
            await test_tool(session, "excel_add_slicer", {
                "file_path": fp,
                "target_sheet": "Pivot",
                "pivot_sheet": "Pivot",
                "pivot_name": "PivotScores",
                "field_name": "Name",
            })
            await test_tool(session, "excel_add_subtotal", {
                "file_path": fp, "sheet": sh, "range": "A1:B3"
            })

            print("\nPhase 5: Merge / Borders / Named range")
            await test_tool(session, "excel_merge_cells", {"file_path": fp, "sheet": sh, "range": "A1:B1"})
            await test_tool(session, "excel_set_borders", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_add_named_range", {
                "file_path": fp, "name": "TestRange", "refers_to": "=Sheet1!$A$1:$B$3"
            })

            print("\nPhase 6: Import / Export")
            await test_tool(session, "excel_import_data", {"file_path": fp, "csv_file": "nonexistent.csv"})
            await test_tool(session, "excel_export_data", {"file_path": fp, "sheet": sh, "export_path": EXCEL_EXPORT_FILE})

            print("\nPhase 7: Typography")
            await test_tool(session, "excel_check_typography", {"file_path": fp})

            print("\nPhase 8: Worksheets")
            await test_tool(session, "excel_list_worksheets", {"file_path": fp})
            await test_tool(session, "excel_get_worksheet_info", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_copy_worksheet", {"file_path": fp, "sheet": sh})
            # Detect the copied worksheet name from the tool response.
            copy_resp = await call_tool(session, "excel_list_worksheets", {"file_path": fp})
            copied_sheet = None
            import json as _json
            try:
                # Prefer JSON parsing when the tool returns structured data.
                data = _json.loads(copy_resp)
                if isinstance(data, list):
                    for item in data:
                        name = item if isinstance(item, str) else item.get("name", "")
                        if name and name != sh:
                            copied_sheet = name
                            break
                elif isinstance(data, dict):
                    sheets = data.get("worksheets", data.get("sheets", []))
                    for item in sheets:
                        name = item if isinstance(item, str) else item.get("name", "")
                        if name and name != sh:
                            copied_sheet = name
                            break
            except Exception:
                pass
            # Fall back to line parsing for plain-text responses.
            if not copied_sheet:
                for line in copy_resp.split("\n"):
                    line = line.strip().strip('"').strip("'").strip(",")
                    if line and line != sh and "Sheet" in line:
                        copied_sheet = line
                        break
            if not copied_sheet:
                copied_sheet = "Sheet1 (2)"
            await test_tool(session, "excel_delete_worksheet", {"file_path": fp, "sheet": copied_sheet})
            await test_tool(session, "excel_move_worksheet", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_hide_worksheet", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_show_worksheet", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_protect_worksheet", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_unprotect_worksheet", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_tab_color", {"file_path": fp, "sheet": sh, "color": "#FF0000"})

            print("\nPhase 9: Ranges")
            await test_tool(session, "excel_list_used_range", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_clear_range", {"file_path": fp, "sheet": sh, "range": "C1:D1"})
            await test_tool(session, "excel_copy_range", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_paste_range", {"file_path": fp, "sheet": sh, "target_cell": "E1"})
            await test_tool(session, "excel_cut_range", {"file_path": fp, "sheet": sh, "range": "E1:F3"})
            await test_tool(session, "excel_delete_cells", {"file_path": fp, "sheet": sh, "range": "E1:F1"})
            await test_tool(session, "excel_insert_cells", {"file_path": fp, "sheet": sh, "range": "E1:F1"})

            print("\nPhase 10: Row/Column sizing")
            await test_tool(session, "excel_set_row_height", {"file_path": fp, "sheet": sh, "row": 1, "height": 30})
            await test_tool(session, "excel_set_column_width", {"file_path": fp, "sheet": sh, "column": "A", "width": 15})
            await test_tool(session, "excel_hide_rows", {"file_path": fp, "sheet": sh, "rows": "5:5"})

            print("\nPhase 11: Charts")
            # Create a simple chart first so chart tools have something to target.
            await test_tool(session, "excel_apply_operations", {
                "file_path": fp,
                "operations": [{"type": "create_chart", "sheet_name": sh, "chart_type": "column", "data_range": "A1:B3"}]
            })
            await test_tool(session, "excel_list_charts", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_get_chart_info", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_chart_title", {"file_path": fp, "sheet": sh, "chart_index": 1, "title": "Test"})
            await test_tool(session, "excel_set_chart_legend", {"file_path": fp, "sheet": sh, "chart_index": 1})
            await test_tool(session, "excel_add_chart_series", {
                "file_path": fp, "sheet": sh, "chart_index": 1, "series_name": "S1", "values_range": "A1:B3"
            })
            await test_tool(session, "excel_remove_chart_series", {
                "file_path": fp, "sheet": sh, "chart_index": 1, "series_index": 1
            })
            await test_tool(session, "excel_set_chart_axis", {"file_path": fp, "sheet": sh, "chart_index": 1})
            await test_tool(session, "excel_change_chart_type", {
                "file_path": fp, "sheet": sh, "chart_index": 1, "chart_type": "xlBarClustered"
            })
            await test_tool(session, "excel_export_chart", {
                "file_path": fp, "sheet": sh, "chart_index": 1, "output_path": EXCEL_EXPORT_FILE
            })
            await test_tool(session, "excel_delete_chart", {"file_path": fp, "sheet": sh, "chart_index": 1})

            print("\nPhase 12: Font formatting")
            await test_tool(session, "excel_set_font", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_font_bold", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_font_italic", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_font_underline", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_alignment", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_wrap_text", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_indent", {"file_path": fp, "sheet": sh, "range": "A1:B3", "indent": 1})
            await test_tool(session, "excel_set_orientation", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_clear_format", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_copy_format", {
                "file_path": fp, "sheet": sh, "source_range": "A1:B3", "target_range": "C1:D3"
            })

            print("\nPhase 13: Page setup")
            await test_tool(session, "excel_set_page_orientation", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_page_size", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_page_margins", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_header", {"file_path": fp, "sheet": sh, "text": "Header"})
            await test_tool(session, "excel_set_footer", {"file_path": fp, "sheet": sh, "text": "Footer"})
            await test_tool(session, "excel_add_print_title", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_print_area", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_set_page_break", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_set_scale", {"file_path": fp, "sheet": sh, "scale": 100})
            await test_tool(session, "excel_set_fit_to_page", {"file_path": fp, "sheet": sh, "fit_width": 1, "fit_height": 1})

            print("\nPhase 14: Formulas")
            await test_tool(session, "excel_set_array_formula", {
                "file_path": fp, "sheet": sh, "range": "C1:C3", "formula": "=A1:A3*B1:B3"
            })
            await test_tool(session, "excel_evaluate_formula", {"file_path": fp, "sheet": sh, "cell": "A1"})
            await test_tool(session, "excel_replace_formula", {
                "file_path": fp, "sheet": sh, "range": "A1:B3", "find": "A", "replace": "B"
            })
            await test_tool(session, "excel_find_formula_cells", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_convert_to_values", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_get_formula_info", {"file_path": fp, "sheet": sh, "cell": "A1"})

            print("\nPhase 15: Names")
            await test_tool(session, "excel_define_name", {
                "file_path": fp, "name": "MyName", "refers_to": "=Sheet1!$A$1"
            })

            print("\nPhase 16: Tables")
            # Create table on a clean range (D1:E3) to avoid conflicts with merged cells
            await test_tool(session, "excel_apply_operations", {
                "file_path": fp,
                "operations": [
                    {"type": "write_cell", "sheet_name": sh, "row": 1, "col": 4, "value": "Col1"},
                    {"type": "write_cell", "sheet_name": sh, "row": 1, "col": 5, "value": "Col2"},
                    {"type": "write_cell", "sheet_name": sh, "row": 2, "col": 4, "value": "A"},
                    {"type": "write_cell", "sheet_name": sh, "row": 2, "col": 5, "value": "1"},
                    {"type": "write_cell", "sheet_name": sh, "row": 3, "col": 4, "value": "B"},
                    {"type": "write_cell", "sheet_name": sh, "row": 3, "col": 5, "value": "2"},
                ]
            })
            await test_tool(session, "excel_create_table", {"file_path": fp, "sheet": sh, "range": "D1:E3", "table_name": "TestTable"})
            await test_tool(session, "excel_list_tables", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_resize_table", {
                "file_path": fp, "sheet": sh, "table_name": "TestTable", "range": "D1:E5"
            })
            await test_tool(session, "excel_set_table_style", {
                "file_path": fp, "sheet": sh, "table_name": "TestTable", "style_name": "TableStyleMedium2"
            })
            await test_tool(session, "excel_show_table_totals", {
                "file_path": fp, "sheet": sh, "table_name": "TestTable"
            })
            await test_tool(session, "excel_add_table_column", {
                "file_path": fp, "sheet": sh, "table_name": "TestTable", "column_name": "NewCol"
            })
            await test_tool(session, "excel_remove_table_column", {
                "file_path": fp, "sheet": sh, "table_name": "TestTable", "column_name": "NewCol"
            })
            await test_tool(session, "excel_delete_table", {
                "file_path": fp, "sheet": sh, "table_name": "TestTable"
            })

            print("\nPhase 17: Filters / Sort")
            await test_tool(session, "excel_add_auto_filter", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_remove_auto_filter", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_sort_range", {"file_path": fp, "sheet": sh, "range": "A1:B3"})
            await test_tool(session, "excel_advanced_filter", {
                "file_path": fp, "sheet": sh, "range": "A1:B3", "criteria_range": "D1:D2"
            })
            await test_tool(session, "excel_remove_duplicates", {"file_path": fp, "sheet": sh, "range": "A1:B3"})

            print("\nPhase 18: Grouping")
            await test_tool(session, "excel_group_rows", {"file_path": fp, "sheet": sh, "range": "2:3"})
            await test_tool(session, "excel_ungroup_rows", {"file_path": fp, "sheet": sh, "range": "2:3"})
            await test_tool(session, "excel_group_columns", {"file_path": fp, "sheet": sh, "range": "A:B"})
            await test_tool(session, "excel_ungroup_columns", {"file_path": fp, "sheet": sh, "range": "A:B"})

            print("\nPhase 19: Protection")
            await test_tool(session, "excel_protect_workbook", {"file_path": fp})
            await test_tool(session, "excel_unprotect_workbook", {"file_path": fp})

            print("\nPhase 20: Shapes / Images")
            # Build a tiny 10x10 PNG for image insertion tests.
            test_image_path = str(TEST_DIR / "test_image.png")
            try:
                from PIL import Image as PILImage
                img = PILImage.new('RGB', (10, 10), color='red')
                img.save(test_image_path)
            except ImportError:
                # Fallback: write a minimal PNG using struct/zlib only.
                def _create_png(path, w=10, h=10):
                    def _chunk(ctype, data):
                        c = ctype + data
                        return struct.pack('>I', len(data)) + c + struct.pack('>I', zlib.crc32(c) & 0xffffffff)
                    raw = b''
                    for _ in range(h):
                        raw += b'\x00' + b'\xff\x00\x00' * w  # filter byte + red pixels
                    ihdr = struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)
                    with open(path, 'wb') as f:
                        f.write(b'\x89PNG\r\n\x1a\n')
                        f.write(_chunk(b'IHDR', ihdr))
                        f.write(_chunk(b'IDAT', zlib.compress(raw)))
                        f.write(_chunk(b'IEND', b''))
                _create_png(test_image_path)
            await test_tool(session, "excel_add_image", {
                "file_path": fp, "sheet": sh, "image_path": test_image_path
            })
            await test_tool(session, "excel_list_shapes", {"file_path": fp, "sheet": sh})
            await test_tool(session, "excel_delete_shape", {"file_path": fp, "sheet": sh, "index": 1})

            print("\nPhase 21: Comments")
            await test_tool(session, "excel_add_comment", {
                "file_path": fp, "sheet": sh, "cell": "A1", "text": "Test comment"
            })
            await test_tool(session, "excel_delete_comment", {"file_path": fp, "sheet": sh, "cell": "A1"})

            print("\nPhase 22: View")
            await test_tool(session, "excel_set_view_zoom", {"file_path": fp, "sheet": sh, "zoom": 100})
            await test_tool(session, "excel_set_view_gridlines", {"file_path": fp, "sheet": sh, "show": True})
            await test_tool(session, "excel_set_view_headings", {"file_path": fp, "sheet": sh, "show": True})

            print("\nPhase 23: Calculation")
            await test_tool(session, "excel_recalculate", {"file_path": fp})
            await test_tool(session, "excel_set_calculation_mode", {"file_path": fp, "mode": "auto"})
            await test_tool(session, "excel_set_iterative_calc", {"file_path": fp})

            print("\nPhase 24: Destructive protection (skipped: would block COM)")
            # Keep the expected-failure accounting without opening modal COM prompts.
            for name in ["excel_set_open_password", "excel_set_write_reservation_password",
                         "excel_mark_as_final", "excel_recommend_read_only"]:
                record(name, "expected_failure", "Skipped: would block COM with modal dialog")

            print("\nPhase 25: Document lifecycle")
            await test_tool(session, "excel_close_workbook", {"file_path": fp, "save": True})
            await test_tool(session, "excel_open_workbook", {"file_path": fp})
            await test_tool(session, "excel_close_workbook", {"file_path": fp, "save": True})

    elapsed = time.time() - start_time

    print("\n" + "=" * 60)
    print("EXCEL TEST SUMMARY")
    print("=" * 60)
    print(f"  Total tested:          {total_tested}")
    print(f"  Passed:                {total_passed}")
    print(f"  Failed:                {total_failed}")
    print(f"  Expected failures:     {total_expected_fail}")
    print(f"  Elapsed:               {elapsed:.1f}s")
    print("=" * 60)

    results_file = TEST_DIR / "excel_test_results.json"
    with open(results_file, "w", encoding="utf-8") as f:
        json.dump({
            "category": "excel",
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

    failures = [r for r in results if r["status"] == "fail"]
    if failures:
        print(f"\n{len(failures)} unexpected failure(s):")
        for f in failures:
            print(f"  - {f['tool']}: {f['detail'][:100]}")

    sys.exit(1 if total_failed > 0 else 0)


if __name__ == "__main__":
    asyncio.run(main())
