from __future__ import annotations

from pathlib import Path

import office_mcp.tools.powerpoint as powerpoint_tools
import office_mcp.tools.word as word_tools


class FakeMCP:
    def __init__(self) -> None:
        self.tools: dict[str, object] = {}

    def tool(self):  # noqa: ANN202
        def decorator(func):
            self.tools[func.__name__] = func
            return func

        return decorator


def test_word_public_recovery_entrypoints_use_ensure_document(
    tmp_path: Path, monkeypatch
) -> None:
    fake_mcp = FakeMCP()
    word_tools.register_word_tools(fake_mcp)
    document = object()
    calls: list[tuple[Path, bool]] = []

    monkeypatch.setattr(word_tools, "validate_path", lambda raw: tmp_path / raw)
    monkeypatch.setattr(
        word_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or document,
    )
    monkeypatch.setattr(
        word_tools, "apply_word_operations", lambda doc, ops: ["ok", doc is document, ops]
    )
    monkeypatch.setattr(
        word_tools, "_check_typography", lambda doc, op: [{"document_ok": doc is document}]
    )
    monkeypatch.setattr(
        word_tools, "_list_tables", lambda doc, op: [{"document_ok": doc is document}]
    )

    apply_result = fake_mcp.tools["word_apply_operations"]("draft.docx", [{"type": "save"}])
    typo_result = fake_mcp.tools["word_check_typography"]("draft.docx")
    table_result = fake_mcp.tools["word_list_tables"]("draft.docx")

    assert calls == [
        (tmp_path / "draft.docx", True),
        (tmp_path / "draft.docx", False),
        (tmp_path / "draft.docx", False),
    ]
    assert apply_result["results"][1] is True
    assert typo_result["issues"][0]["document_ok"] is True
    assert table_result["tables"][0]["document_ok"] is True


def test_powerpoint_public_recovery_entrypoints_use_ensure_document(
    tmp_path: Path, monkeypatch
) -> None:
    fake_mcp = FakeMCP()
    powerpoint_tools.register_ppt_tools(fake_mcp)
    presentation = object()
    calls: list[tuple[Path, bool]] = []

    monkeypatch.setattr(powerpoint_tools, "validate_path", lambda raw: tmp_path / raw)
    monkeypatch.setattr(
        powerpoint_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or presentation,
    )
    monkeypatch.setattr(
        powerpoint_tools, "apply_ppt_operations", lambda ppt, ops: ["ok", ppt is presentation, ops]
    )
    monkeypatch.setattr(
        powerpoint_tools, "_check_typography", lambda ppt, op: {"presentation_ok": ppt is presentation}
    )

    apply_result = fake_mcp.tools["ppt_apply_operations"]("deck.pptx", [{"type": "save"}])
    typo_result = fake_mcp.tools["ppt_check_typography"]("deck.pptx", 2)

    assert calls == [
        (tmp_path / "deck.pptx", True),
        (tmp_path / "deck.pptx", False),
    ]
    assert apply_result["results"][1] is True
    assert typo_result["presentation_ok"] is True


def test_powerpoint_table_edit_entrypoints_use_ensure_document(
    tmp_path: Path, monkeypatch
) -> None:
    fake_mcp = FakeMCP()
    powerpoint_tools.register_ppt_tools(fake_mcp)
    presentation = object()
    calls: list[tuple[Path, bool]] = []

    monkeypatch.setattr(powerpoint_tools, "validate_path", lambda raw: tmp_path / raw)
    monkeypatch.setattr(
        powerpoint_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or presentation,
    )
    monkeypatch.setattr(
        powerpoint_tools,
        "_ppt_merge_table_cells",
        lambda ppt, op: {"presentation_ok": ppt is presentation, "op": op},
    )
    monkeypatch.setattr(
        powerpoint_tools,
        "_ppt_split_table_cells",
        lambda ppt, op: {"presentation_ok": ppt is presentation, "op": op},
    )

    merge_result = fake_mcp.tools["ppt_merge_table_cells"]("deck.pptx", 1, 1, 1, 1, 2, 2)
    split_result = fake_mcp.tools["ppt_split_table_cells"]("deck.pptx", 1, 1, 1, 1)

    assert calls == [
        (tmp_path / "deck.pptx", True),
        (tmp_path / "deck.pptx", True),
    ]
    assert "已合并单元格" in merge_result
    assert "已拆分单元格" in split_result


def test_powerpoint_media_and_save_entrypoints_use_ensure_document(
    tmp_path: Path, monkeypatch
) -> None:
    fake_mcp = FakeMCP()
    powerpoint_tools.register_ppt_tools(fake_mcp)
    presentation = object()
    calls: list[tuple[Path, bool]] = []

    monkeypatch.setattr(powerpoint_tools, "validate_path", lambda raw: tmp_path / raw)
    monkeypatch.setattr(
        powerpoint_tools.office_manager,
        "ensure_document",
        lambda path, activate=False: calls.append((path, activate)) or presentation,
    )
    monkeypatch.setattr(powerpoint_tools, "_add_video", lambda ppt, op: {"presentation_ok": ppt is presentation, "op": op})
    monkeypatch.setattr(powerpoint_tools, "_add_audio", lambda ppt, op: {"presentation_ok": ppt is presentation, "op": op})
    monkeypatch.setattr(powerpoint_tools, "_set_media_settings", lambda ppt, op: {"presentation_ok": ppt is presentation, "op": op})
    monkeypatch.setattr(powerpoint_tools, "_save_as", lambda ppt, op: {"presentation_ok": ppt is presentation, "op": op})

    video_result = fake_mcp.tools["ppt_add_video"]("deck.pptx", "clip.mp4")
    audio_result = fake_mcp.tools["ppt_add_audio"]("deck.pptx", "music.wav")
    media_result = fake_mcp.tools["ppt_set_media_settings"]("deck.pptx")
    save_result = fake_mcp.tools["ppt_save_as"]("deck.pptx", "deck.pdf", "pdf")

    assert calls == [
        (tmp_path / "deck.pptx", True),
        (tmp_path / "deck.pptx", True),
        (tmp_path / "deck.pptx", True),
        (tmp_path / "deck.pptx", False),
    ]
    assert "slide 1" in video_result
    assert "slide 1" in audio_result
    assert "slide 1" in media_result
    assert "pdf" in save_result
