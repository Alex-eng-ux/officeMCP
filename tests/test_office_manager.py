from __future__ import annotations

from pathlib import Path

import pytest

from office_mcp.core.errors import COMOperationError
from office_mcp.core.office_manager import OfficeManager


def test_normalize_app_type_accepts_powerpoint_alias() -> None:
    manager = OfficeManager()

    assert manager._normalize_app_type("powerpoint") == "ppt"


def test_activate_app_rejects_mismatched_file_type(tmp_path: Path) -> None:
    manager = OfficeManager()
    ppt_path = tmp_path / "deck.pptx"

    with pytest.raises(COMOperationError, match="不一致"):
        manager.activate_app("word", ppt_path)
