# MCP Multi-Agent Test Report (2026-06-06)

## Scope

This report tracks a coordinated multi-agent validation pass across the Office MCP tool surface:

- Excel
- Word
- PowerPoint

Goals for this round:

1. Run broad tool-coverage tests through the MCP server.
2. Separate expected failures from unexpected failures.
3. Fix bugs that are realistically fixable inside this repository.
4. Rerun affected suites.
5. Leave a concrete handoff for any remaining limitations.

## Test Harness

- Root scripts:
  - `test_excel_mcp.py`
  - `test_word_mcp.py`
  - `test_ppt_mcp.py`
- Output directory:
  - `test_output/`
- Per-suite machine-readable outputs:
  - `test_output/excel_test_results.json`
  - `test_output/word_test_results.json`
  - `test_output/ppt_test_results.json`

## Coordinator Notes

- The three root test scripts currently force:
  - `OFFICE_MCP_ALLOWED_DIRS = test_output`
  - `OFFICE_MCP_VISIBLE = false`
- This makes the broad regression runs isolated and repeatable, but it does not verify the visible Office window workflow.
- Worker ownership for this round was intentionally split by app surface to reduce merge conflicts:
  - Excel worker
  - Word worker
  - PowerPoint worker

## Results Summary

### Excel

- Initial run:
  - The default-shell `python` failed before startup because it did not have the `mcp` dependency.
  - The first valid run also exposed a harness-level console encoding problem.
- Fixes applied:
  - Fixed a real Excel-side implementation bug in `src/office_mcp/operations/excel_ops.py`.
  - `_create_pivot_table()` now passes the Excel `Range` object directly into `PivotCaches.Create` instead of converting it through `.Address(External=True)`.
- Final rerun:
  - Total tested: `111`
  - Passed: `101`
  - Expected failures: `10`
  - Unexpected failures: `0`
- Remaining issues:
  - `excel_create_pivot_table` still expected-fails, but the failure mode improved from a string-call bug to a later `PivotCaches.Create` COM-shape issue.
  - The `EXPECTED_FAILURES` list in `test_excel_mcp.py` is stale and should be cleaned in a follow-up.

### Word

- Initial run:
  - The default-shell `python` failed before startup because it did not have the `mcp` dependency.
  - The first valid run then failed on a harness `UnicodeEncodeError` caused by Unicode status markers in a GBK console.
- Fixes applied:
  - `test_word_mcp.py` was made ASCII-safe and deterministic.
  - The harness now launches the MCP server with `sys.executable` so the server uses the same interpreter as the test process.
- Final rerun:
  - Total tested: `97`
  - Passed: `91`
  - Expected failures: `6`
  - Unexpected failures: `0`
- Remaining issues:
  - Still expected-failing:
    - `word_insert_icon`
    - `word_delete_section`
    - `word_crop_image`
    - `word_set_image_position`
    - `word_replace_image`
    - `word_get_hyperlink`
  - Non-fatal warnings remain around typography regex handling and some table-border COM parameter ranges.

### PowerPoint

- Initial run:
  - The default-shell `python` failed before startup because it did not have the `mcp` dependency.
  - The first valid run failed on harness encoding, and a later rerun exposed a harness `NameError` for `ppt_test_image_path`.
- Fixes applied:
  - `test_ppt_mcp.py` was stabilized for repeatable local execution.
  - Unicode status markers were replaced with ASCII output.
  - The file was normalized so Python could parse it consistently.
  - The swallowed `ppt_test_image_path` assignment was repaired.
  - `EXPECTED_FAILURES` was trimmed to the failures that still reproduce.
- Final rerun:
  - Total tested: `170`
  - Passed: `161`
  - Expected failures: `9`
  - Unexpected failures: `0`
- Remaining issues:
  - Still expected-failing:
    - `ppt_add_section`
    - `ppt_set_theme_preset`
    - `ppt_insert_icon`
    - `ppt_set_chart_data`
    - `ppt_get_chart_data`
    - `ppt_format_chart`
    - `ppt_format_chart_axis`
    - `ppt_set_chart_series`
    - `ppt_change_chart_type`
  - The chart follow-up failures are partly harness-driven because the test is still targeting a non-chart shape after chart creation.

## Cross-Cutting Findings

- All three suites were initially vulnerable to interpreter drift when launched through a generic shell `python`.
- Windows console encoding is still a real failure source for large MCP integration harnesses. ASCII-only progress markers are much safer here.
- In this round, the MCP implementation held up better than the harnesses. Most instability came from harness assumptions, not widespread Office tool failures.

## Candidate Reclassification of Expected Failures

- Excel:
  - `excel_export_data`
  - `excel_advanced_filter`
  - `excel_add_image`
  - `excel_delete_shape`
  - `excel_delete_comment`
- Word:
  - `word_update_field`
  - `word_delete_field`
  - `word_get_field_result`
  - `word_get_image_info`
  - `word_resize_image`
  - `word_set_image_wrap`
  - `word_delete_image`
  - `word_remove_hyperlink`
  - `word_delete_comment`
  - `word_merge_paragraphs`
  - `word_split_paragraph`

## Remaining Limitations

- The broad harnesses still run with `OFFICE_MCP_VISIBLE=false`, so this pass does not validate the visible-window workflow.
- Some remaining failures are true environment or Office-build limitations rather than obvious repository bugs.

## Final Assessment

- This round produced a usable regression baseline across `378` Office MCP tool invocations (`111 + 97 + 170`) with `0` unexpected failures after targeted fixes.
- The most meaningful code fix landed in the Excel implementation, while Word and PowerPoint mainly needed harness hardening.
- The project is in a much stronger state for a second 300+ tool pass, with remaining work now concentrated in stale harness expectations, chart targeting, and a smaller set of environment-dependent limitations.
