# Implementation Workflow

This project is easiest to evolve safely when feature work follows a tight implementation and verification loop instead of a one-pass "edit and hope" flow.

The workflow below is the team-default way to ship changes that are meant to be genuinely usable, not just locally plausible.

## Goal

For any non-trivial change, aim to finish with all of the following true:

- the requested behavior is implemented
- the public MCP tool surface exposes that behavior
- the server still starts and registers tools
- focused validation has been run
- any remaining risk is named explicitly

## Default Loop

### 1. Scope the change

Before editing:

- restate the user-visible outcome
- identify the touched tool entry points
- identify the operations layer and core modules involved
- identify the validation commands that should pass afterward

For this repository that usually means checking:

- `src/office_mcp/tools/`
- `src/office_mcp/operations/`
- `src/office_mcp/core/`
- `tests/`
- `.github/workflows/`

### 2. Decompose by ownership

If the task is broad, split work into independent tracks with clear ownership.

Good splits:

- runtime fix vs release/CI fix
- Word vs Excel vs PowerPoint
- implementation vs review/verification

Bad splits:

- multiple workers editing the same large tool file without a clear owner
- review workers duplicating the main implementation path

When multi-agent help is available, one worker should own the file edit and others should inspect risk, packaging, or validation.

### 3. Implement the smallest complete slice

Prefer finishing one vertical slice over partially touching many areas.

Examples:

- tool import path fixed and covered by a focused test
- path guard behavior fixed and verified through public tools
- packaging metadata corrected and validated by wheel install

Keep public signatures stable unless the feature requires a deliberate API change.

### 4. Run code analysis on the changed surface

Do not treat lint as the only review.

Review the changed surface for:

- undefined helper references in public tools
- path validation gaps
- Office app/document activation mismatches
- packaging or entry-point drift
- CI rules that do not match the real state of the repository

For this repo, correctness findings matter more than style debt.

### 5. Stress-test assumptions

Before calling the work done, ask:

- does the user-facing MCP tool actually call the implemented behavior
- does the behavior still work when invoked through the tool layer, not just the operations layer
- does build or installation exercise the updated files
- are relative paths, blocked paths, or missing Office instances handled clearly
- are we claiming support for behavior that we did not actually test

Convert concrete concerns into repair tasks and fix them before moving on.

### 6. Validate in layers

Use a layered validation pass:

1. Focused static or correctness checks
2. Focused tests
3. Packaging and install checks
4. Real MCP connectivity checks

Recommended commands:

```bash
python -m pytest tests
python count_tools.py --min-total 300 --require-prefix excel_ --require-prefix word_ --require-prefix ppt_
python -m build --wheel
```

Recommended runtime checks:

- start the server through `office_mcp.server`
- connect with a real MCP client
- call `office_status`
- call at least one representative tool from the changed area

If a change affects global Codex usage, also verify that the MCP server can be loaded through Codex's configured `mcp_servers` entry.

## Repository-specific Guidance

### Favor realistic release gates

This repository has historical lint debt in large legacy modules, especially PowerPoint.

When updating CI:

- keep correctness checks blocking
- avoid pretending the whole repository is clean when it is not
- gate the changed support files and runtime hotspots first

A "green" workflow should mean "this can ship" rather than "we skipped every meaningful check."

### Protect the public tool layer

Most user-visible regressions in this project show up in `src/office_mcp/tools/`, not in the lower-level operation functions alone.

Common failure modes:

- helper imported in `operations` but not in `tools`
- public tool accepts a path that should be rejected
- activation logic accepts the wrong app/file pairing
- build succeeds but console entry point is broken

Focused tests for the public tool surface are worth more than broad style cleanup.

### Prefer absolute-path examples and tests

This server intentionally rejects relative paths.

Docs, examples, and tests should reflect that contract so users do not learn the wrong usage pattern.

## Done Criteria

A change is ready to land when:

- requested behavior is implemented
- changed MCP tools resolve and register correctly
- focused automated checks pass
- build and install checks pass when relevant
- at least one real MCP call has succeeded for user-facing integration work
- remaining caveats are small, explicit, and non-blocking

## When to stop and escalate

Stop and ask for help only when one of these is true:

- the next step requires credentials or machine state you do not have
- Microsoft Office itself is missing or COM is unavailable
- a global configuration change requires user confirmation
- two valid implementation paths have materially different tradeoffs

Otherwise, keep the loop moving until the project is actually more usable than it was before.
