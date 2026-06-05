"""Inspect and validate the registered Office MCP tool surface."""

from __future__ import annotations

import argparse
import importlib
import sys
from collections.abc import Iterable
from typing import Any

PREFIX_LABELS = {
    "excel_": "Excel",
    "word_": "Word",
    "ppt_": "PowerPoint",
}
OTHER_LABEL = "Other"


def summarize_tool_names(names: Iterable[str]) -> dict[str, int]:
    """Summarize tool names by product prefix."""
    summary = {label: 0 for label in PREFIX_LABELS.values()}
    summary[OTHER_LABEL] = 0
    unique_names = sorted(set(names))

    for name in unique_names:
        for prefix, label in PREFIX_LABELS.items():
            if name.startswith(prefix):
                summary[label] += 1
                break
        else:
            summary[OTHER_LABEL] += 1

    summary["Total"] = len(unique_names)
    return summary


def format_summary(summary: dict[str, int]) -> list[str]:
    """Format a human-readable summary report."""
    return [
        f"Total: {summary['Total']}",
        f"Excel: {summary['Excel']}",
        f"Word: {summary['Word']}",
        f"PowerPoint: {summary['PowerPoint']}",
        f"Other: {summary['Other']}",
    ]


def validate_summary(
    summary: dict[str, int],
    *,
    min_total: int | None = None,
    required_prefixes: Iterable[str] = (),
) -> None:
    """Validate summary thresholds used by CI and local smoke checks."""
    if min_total is not None and summary["Total"] < min_total:
        raise ValueError(
            f"Expected at least {min_total} tools, found {summary['Total']}"
        )

    missing = []
    for prefix in required_prefixes:
        label = PREFIX_LABELS.get(prefix)
        if label is None:
            raise ValueError(f"Unsupported required prefix: {prefix}")
        if summary[label] == 0:
            missing.append(prefix)

    if missing:
        raise ValueError(f"Missing required tool groups: {', '.join(missing)}")


def load_tool_names() -> list[str]:
    """Load registered tool names from the Office MCP server."""
    if "src" not in sys.path:
        sys.path.insert(0, "src")

    server_module = importlib.import_module("office_mcp.server")
    mcp = getattr(server_module, "mcp")

    tools = _resolve_tool_registry(mcp)
    if not isinstance(tools, dict):
        raise RuntimeError("FastMCP tool registry is unavailable")

    return sorted(tools.keys())


def _resolve_tool_registry(mcp: Any) -> dict[str, Any] | None:
    """Resolve the tool registry across likely FastMCP layouts."""
    direct_candidates = [
        getattr(mcp, "tools", None),
        getattr(mcp, "_tools", None),
    ]
    for candidate in direct_candidates:
        if isinstance(candidate, dict):
            return candidate

    tool_manager = getattr(mcp, "tool_manager", None) or getattr(mcp, "_tool_manager", None)
    if tool_manager is None:
        return None

    manager_candidates = [
        getattr(tool_manager, "tools", None),
        getattr(tool_manager, "_tools", None),
    ]
    for candidate in manager_candidates:
        if isinstance(candidate, dict):
            return candidate

    return None


def build_parser() -> argparse.ArgumentParser:
    """Create the CLI parser."""
    parser = argparse.ArgumentParser(
        description="Summarize and validate the registered Office MCP tools."
    )
    parser.add_argument(
        "--min-total",
        type=int,
        default=None,
        help="Fail if fewer than this many tools are registered.",
    )
    parser.add_argument(
        "--require-prefix",
        action="append",
        default=[],
        choices=sorted(PREFIX_LABELS),
        help="Require at least one tool for the given prefix.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    """Run the tool summary CLI."""
    parser = build_parser()
    args = parser.parse_args(argv)

    names = load_tool_names()
    summary = summarize_tool_names(names)
    validate_summary(
        summary,
        min_total=args.min_total,
        required_prefixes=args.require_prefix,
    )

    for line in format_summary(summary):
        print(line)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
