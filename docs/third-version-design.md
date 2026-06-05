# Office MCP Third Version Design

## Goal

Turn the current "mostly callable" second version into a genuinely landable third version that is stable across Word, Excel, and PowerPoint in real client-driven sessions.

The user-visible outcome for v3 is:

- Word, Excel, and PowerPoint all open visibly by default unless explicitly configured otherwise
- Office MCP can survive multi-step real-world sessions instead of failing after the first successful action
- path access, activation, and basic lifecycle are predictable across repo/test directories
- mail merge, workbook reads, and presentation edits behave consistently through the public MCP surface

## What round-2 testing says about v2

Based on the reports under `E:\test`:

- Path allowlist issues were materially improved
  - Word test directory can be activated directly
  - Excel files can be created/opened/written under `E:\test`
  - the earlier "path not allowed" blocker no longer appears to be the primary failure mode
- Word is still not landable for mail merge
  - `word_mail_merge` times out
  - `word_mail_merge_enhanced` times out or returns COM/RPC errors
  - output documents are not produced
- Excel is still not landable for longer sessions
  - open/activate/write can work
  - follow-up reads and table listing fail with COM rejection errors
  - recovery attempts can degrade into "Office not available through COM"
- PowerPoint is the healthiest surface, but not fully trustworthy
  - chart data survives chart type changes
  - connector handling is only partially repaired
  - complex merged/split table flows still corrupt logical structure
  - public-path failures shifted away from allowlist issues, but session availability is still inconsistent

## v2 assessment

Second version is worth releasing as a milestone, but not as the final "works reliably" claim.

It is best described as:

- path-guard fixed enough to unblock real directories
- public tool surface more honest and better tested
- degraded import/runtime compatibility improved
- Office session stability still the main blocker

## Third-version priorities

### P0: Session stability layer

Build a dedicated COM session resilience layer instead of scattering retries across individual operations.

Required capabilities:

- deterministic app acquisition strategy
- reconnect logic when COM rejects calls
- explicit recovery after modal-dialog or RPC failure
- per-app health checks before high-risk operations
- better distinction between:
  - Office not installed
  - Office busy
  - Office modal dialog blocked
  - stale COM object
  - document handle lost

Suggested shape:

- add an app-session wrapper in `core/`
- centralize retry/rebind logic there
- stop letting tools directly assume a cached document object is still valid after a previous failure

### P0: Public lifecycle consistency

The public MCP API should follow one simple contract:

- open or activate the target file
- validate that the app/document session is live
- perform the operation
- persist or return a precise failure reason

Needed improvements:

- unify `open`, `activate`, `ensure_document`, and operation-entry behavior
- ensure follow-up calls can recover if the original cached handle is stale
- expose structured diagnostics for "document not open" vs "document lost after COM failure"

### P0: Word mail merge hardening

Word mail merge is the clearest remaining failure cluster.

Third-version work:

- add preflight validation for template/data source accessibility
- detect and log whether Word opens a blocking dialog during merge
- split merge flow into explicit steps:
  - open template
  - bind data source
  - configure destination
  - execute merge
  - resolve result document
  - save/return output
- add bounded operation timing with better timeout diagnostics
- add a smoke script dedicated to mail merge recovery scenarios

Success criteria:

- normal merge completes without timing out
- enhanced merge produces an output file
- SQL/filter merge either works or returns a precise, reproducible reason

### P0: Excel rebind/recovery after initial success

Excel currently demonstrates the classic "first operation succeeds, next COM call dies" pattern.

Third-version work:

- add workbook handle revalidation before every operation batch
- if a workbook handle is stale, reacquire from the live application object
- add explicit "reopen and relink workbook" fallback path
- instrument which specific calls trigger COM rejection

Success criteria:

- create/open/write/read all succeed in one session
- `excel_list_tables` survives after prior mutations
- cleanup and reopen do not degrade into "Office unavailable through COM"

### P1: PowerPoint structural correctness

PowerPoint is already useful, so v3 should focus on correctness instead of only availability.

Third-version work:

- connector binding should normalize site specifications
- merged/split table operations should document and preserve structure more predictably
- PDF export return messaging should match actual output semantics

Success criteria:

- connector APIs either succeed fully or return a clear partial-binding error
- table merge/split roundtrips do not silently drift row count or duplicate content

### P1: Visibility and interactive debugging

v3 should formalize the interactive workflow:

- app-specific visibility defaults
- a documented "interactive mode" config
- optional debug logging for live Office automation

This matters because real-world debugging is much easier when Word/Excel/PPT are visible.

## Proposed architecture changes

### 1. Add a session controller

Suggested module:

- `src/office_mcp/core/session_controller.py`

Responsibilities:

- app acquisition
- document/workbook/presentation rebinding
- health probes
- retry policy
- stale-handle invalidation

### 2. Separate configuration concerns

Current config is already better than v1. For v3, split settings by concern:

- path settings
- visibility settings
- backup settings
- COM resilience settings
  - retry count
  - retry interval
  - health check toggle
  - modal wait timeout

### 3. Add structured diagnostics

Standardize operation failures around categories such as:

- `path_not_allowed`
- `document_not_open`
- `document_handle_stale`
- `office_busy`
- `office_modal_blocked`
- `rpc_unavailable`
- `com_call_rejected`
- `save_failed`

This should improve both MCP responses and test triage.

## Verification strategy for v3

Use the landable workflow gates, not just unit tests.

### Automated

- focused config/path/unit tests
- tool registration contract
- syntax/build/install checks

### Real integration

- Word mail merge smoke script with output assertions
- Excel multi-step session script:
  - create
  - open
  - write
  - read
  - list tables
  - close
  - reopen
- PowerPoint connector and merged-table regression script

### Release gate

v3 should not be called landable until:

- at least one representative Word tool path succeeds end to end
- at least one representative Excel multi-step session succeeds end to end
- at least one representative PowerPoint regression scenario succeeds end to end

## Suggested implementation order

1. session controller foundation
2. Word mail merge stabilization
3. Excel stale-handle recovery
4. PowerPoint structural fixes
5. integration smoke suite and release checklist

## Repo/release positioning

Recommend tagging the current state as a second-version milestone focused on:

- allowlist/path-guard repair
- public-surface cleanup
- compatibility improvements

Then develop v3 as the "stability and real-session" release.
