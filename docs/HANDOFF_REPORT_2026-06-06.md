# Office MCP Handoff Report

Date: 2026-06-06
Repo: `D:\FakeC\MCP\offiiceMCP`

## 1. Current objective

Drive the Office MCP project to a genuinely landable state based on the latest real-world test reports in `E:\test`, with special focus on:

- Word mail merge reliability
- Excel session stability
- PowerPoint table merge/split correctness

## 2. What has already been changed in this repo

### 2.1 Office app/session lifecycle

Files:

- `src/office_mcp/core/office_manager.py`
- `src/office_mcp/core/errors.py`

Implemented:

- stronger Office app recovery and stale-handle rebinding
- `ensure_document(...)` recovery path for public tool entrypoints
- COM startup logic now attempts:
  - `pythoncom.CoInitialize()`
  - `GetActiveObject(...)`
  - `DispatchEx(...)`
  - fallback `Dispatch(...)`
- dispatch failures no longer collapse into a single misleading "Office not installed" error
- distinction now exists between:
  - true Office/COM registration failure
  - Office busy/dialog/disconnected session
  - generic startup failure

### 2.2 Word mail merge hardening

File:

- `src/office_mcp/operations/word_ops.py`

Implemented:

- staged retry wrapper around mail merge COM calls
- suppression/restoration of:
  - `DisplayAlerts`
  - `ScreenUpdating`
  - `Options.ConfirmConversions`
- explicit Excel mail-merge defaults inferred from workbook XML:
  - first visible worksheet name
  - generated `Connection`
  - generated `SQLStatement`
- additional tests for helper logic

### 2.3 Excel / Word / PowerPoint public recovery surface

Files:

- `src/office_mcp/tools/excel.py`
- `src/office_mcp/tools/word.py`
- `src/office_mcp/tools/powerpoint.py`

Implemented:

- key public entrypoints now use `office_manager.ensure_document(...)`
- PowerPoint merge/split tool entrypoints also now use recovery/activation path

### 2.4 PowerPoint table merge/split behavior

File:

- `src/office_mcp/operations/ppt_ops.py`

Implemented:

- added metadata/tag-based preservation of original cell texts before merge
- split path now attempts to restore original text matrix from stored metadata
- helper functions added for:
  - text extraction
  - text restoration
  - merge metadata persistence
  - region normalization

## 3. Tests added or updated

Files:

- `tests/test_allowed_dirs_config.py`
- `tests/test_office_manager.py`
- `tests/test_excel_tool_resilience.py`
- `tests/test_word_ppt_tool_recovery_surface.py`
- `tests/test_word_mail_merge_defaults.py`
- `tests/test_ppt_table_merge_restore.py`

Latest local validation completed during this session:

- `python -m compileall src tests`
- focused pytest suite: `38 passed`
- tool surface count:
  - Total: `366`
  - Excel: `107`
  - Word: `92`
  - PowerPoint: `164`

## 4. Global MCP install state

The global Codex Office MCP was switched to this repo via editable install.

Observed state before sandbox restrictions returned:

- global venv path:
  - `C:\Users\12279\.codex\mcp-venvs\office-mcp`
- package points to:
  - `D:\FakeC\MCP\offiiceMCP`
- global config path:
  - `C:\Users\12279\.codex\config.toml`

The following visibility env values were configured:

- `OFFICE_MCP_VISIBLE = "true"`
- `OFFICE_MCP_WORD_VISIBLE = "true"`
- `OFFICE_MCP_EXCEL_VISIBLE = "true"`
- `OFFICE_MCP_PPT_VISIBLE = "true"`

Note: that global config file is outside the current writable sandbox, so this report records the state but does not attempt to manage it further.

## 5. Latest real test reports reviewed

Latest reports found under `E:\test`:

- Word:
  - `E:\test\office-mcp-word-test\round4\TEST_REPORT_ROUND4.md`
- Excel:
  - `E:\test\office-mcp-excel-test\round4\office-mcp-excel-round4-stability-report.md`
- PowerPoint:
  - `E:\test\office-mcp-ppt-test\round4\ROUND4_VERIFICATION_REPORT.md`

## 6. Round4 outcome summary

### 6.1 Word

Status: **still blocked**

What improved:

- failures are no longer random timeout/RPC hangs
- `office_cleanup(force=false)` now successfully recovers the Word session

Current failure mode:

- `word_mail_merge_enhanced(...)` now fails consistently with a COM error equivalent to:
  - requested object unavailable

Observed consequences from report:

- no merged output document is created
- failure is reproducible on both minimal and complex inputs
- cleanup can recover the Office session, but not mail merge success
- template files appear to be modified/polluted during failure

Interpretation:

- this is no longer a generic COM session hang
- the likely failure is now in the post-bind / execute / result-document transition
- likely suspects:
  - `MailMerge.Execute(...)`
  - `ActiveDocument` switching assumptions
  - template becoming no longer the expected active/main document
  - Word internal state mutation before result save

### 6.2 Excel

Status: **substantially improved, not fully done**

What improved:

- `excel_create_workbook(...)` now succeeds
- the old misleading "Office not installed" bootstrap failure is no longer the primary issue
- repeated `write/read/save` loops passed
- repeated `list_tables` calls passed
- `office_cleanup(force=false)` successfully restores the session after failure

Current remaining risk:

- `excel_close_workbook(save=true)` times out after 120s
- one earlier batched initialization save also failed once, though later saves passed

Interpretation:

- Excel is no longer blocked at bootstrap
- main remaining instability has moved to save/close lifecycle behavior
- the highest-value remaining Excel work is around close/save sequencing and workbook shutdown

### 6.3 PowerPoint

Status: **partially improved, still not done**

What improved:

- chart type switching was stable in round4
- PDF/PNG export path works
- row/column control remains healthy

Still broken:

- connector binding by named direction sites still fails
- failed connector calls may still create actual connector objects, so return state and real state diverge
- table `merge/split` is still reported as corrupting table data in round4

Important note:

- local helper tests pass, but the real PowerPoint round4 report still says merge/split is not fixed
- that means the current metadata-based text restoration is not sufficient against real PowerPoint COM behavior

## 7. Best current understanding of remaining blockers

### Word blocker

Most likely next debugging target:

- make mail merge non-destructive to the source template
- isolate these stages separately:
  1. bind datasource only
  2. execute merge only
  3. identify output/result doc
  4. save result doc

High-value follow-up ideas:

- create explicit helper stages instead of one `_mail_merge(...)` block
- inspect whether `doc.MailMerge.MainDocumentType` or destination state should be set explicitly
- stop assuming `app.ActiveDocument` is always the produced merged result
- explicitly duplicate/open template in a safer way before merge if Word mutates source state
- add a recovery-safe path that closes any produced transient documents when execute partially succeeds

### Excel blocker

Most likely next debugging target:

- workbook save/close transition

High-value follow-up ideas:

- inspect whether `Save()` vs `SaveAs()` vs `Close(SaveChanges=True)` path assumptions differ
- add staging around:
  - save
  - post-save workbook state probe
  - close
  - post-close state reconciliation
- check whether close should use an activation/rebind path first

### PowerPoint blocker

Most likely next debugging target:

- real COM semantics of table merge/split, not just text preservation

High-value follow-up ideas:

- inspect actual post-merge accessibility of cells in the merged region
- verify whether PowerPoint duplicates merged text across virtual cells after merge
- consider treating merge/split as unsupported for perfect restoration unless a safer reconstruction strategy is implemented
- alternative direction:
  - reconstruct a fresh table shape from captured matrix data instead of relying on COM `Split()`

## 8. Important constraints / caveats

- Current workspace is dirty; do not revert unrelated user changes.
- There is a typo in the repo folder name: `offiiceMCP` with double `ii`. This is intentional in current paths and should be preserved unless the user explicitly asks to rename/move the repo.
- The latest real reports are outside the current writable sandbox (`E:\test`), so they can be read only when environment permissions allow.
- The most important unfinished work is integration-level, not unit-test-level.

## 9. Suggested next execution plan

1. Refactor Word mail merge into explicit stages with stage-specific logging and safer result-doc handling.
2. Add a focused regression test around the non-destructive Word template contract where feasible.
3. Rework Excel save/close shutdown handling.
4. Revisit PowerPoint merge/split with a reconstruction-based strategy if direct COM split cannot preserve structure.
5. Re-run real round5 verification in `E:\test` after each repair wave.

## 10. Bottom line

The project has moved forward materially:

- Excel is much closer to usable.
- Word is no longer failing in the old random way, but mail merge is still not landable.
- PowerPoint improved on charts/exports but table merge/split remains unresolved.

This repo should be treated as **improved but not yet fully landable** until:

- Word mail merge succeeds reliably and does not mutate source templates
- Excel close/save path stops timing out
- PowerPoint merge/split no longer corrupts table structure/data
