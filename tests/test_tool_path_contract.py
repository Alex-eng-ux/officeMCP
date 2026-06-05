from __future__ import annotations

from pathlib import Path

from office_mcp.core import path_guard
from office_mcp.tools.excel import register_excel_tools
from office_mcp.tools.powerpoint import register_ppt_tools
from office_mcp.tools.word import register_word_tools


class FakeMCP:
    def __init__(self) -> None:
        self.tools: dict[str, object] = {}

    def tool(self):  # noqa: ANN202
        def decorator(func):
            self.tools[func.__name__] = func
            return func

        return decorator


def test_public_tools_reject_relative_paths(tmp_path: Path, monkeypatch) -> None:
    allowed_dir = tmp_path.resolve()
    monkeypatch.chdir(allowed_dir)
    monkeypatch.setattr(path_guard, "get_allowed_dirs", lambda: [allowed_dir])

    fake_mcp = FakeMCP()
    register_word_tools(fake_mcp)
    register_excel_tools(fake_mcp)
    register_ppt_tools(fake_mcp)

    word_result = fake_mcp.tools["word_open_document"]("draft.docx")
    assert "draft.docx" in word_result
    assert "OFFICE_MCP_ALLOWED_DIRS" in word_result

    excel_result = fake_mcp.tools["excel_open_workbook"]("draft.xlsx")
    assert "draft.xlsx" in excel_result
    assert "OFFICE_MCP_ALLOWED_DIRS" in excel_result

    ppt_result = fake_mcp.tools["ppt_open_presentation"]("draft.pptx")
    assert isinstance(ppt_result, str)
    assert "draft.pptx" in ppt_result
    assert "OFFICE_MCP_ALLOWED_DIRS" in ppt_result
