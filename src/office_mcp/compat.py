"""Compatibility helpers for optional runtime dependencies."""

from __future__ import annotations

from typing import Any


class FallbackFastMCP:
    """Minimal FastMCP-compatible registry for local tests and introspection."""

    def __init__(self, name: str):
        self.name = name
        self.tools: dict[str, Any] = {}

    def tool(self):  # noqa: ANN202
        def decorator(func):
            self.tools[func.__name__] = func
            return func

        return decorator

    def run(self) -> None:
        raise RuntimeError(
            "FastMCP runtime dependencies are unavailable. Install the mcp package in a compatible Python environment."
        )


try:
    from mcp.server.fastmcp import FastMCP as FastMCP  # type: ignore
except Exception:  # pragma: no cover - exercised only in degraded environments
    FastMCP = FallbackFastMCP
