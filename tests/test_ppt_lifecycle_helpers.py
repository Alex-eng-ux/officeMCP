from __future__ import annotations

from pathlib import Path

from office_mcp.operations import ppt_ops
from office_mcp.operations.ppt_ops import (
    _compare_presentations,
    _create_from_template,
    _merge_presentations,
)


class _FakeSlides:
    def __init__(self, count: int) -> None:
        self.Count = count
        self.pasted_positions: list[int] = []

    def __call__(self, index: int):
        return type("Slide", (), {"Shapes": type("Shapes", (), {"Count": 2})(), "Copy": lambda self: None})()

    def Paste(self, position: int) -> None:  # noqa: N802
        self.pasted_positions.append(position)
        self.Count += 1


class _FakePresentation:
    def __init__(self, count: int) -> None:
        self.Slides = _FakeSlides(count)
        self.closed = False
        self.saved = False
        self.saved_as: list[str] = []
        self.FullName = ""

    def Close(self) -> None:  # noqa: N802
        self.closed = True

    def Save(self) -> None:  # noqa: N802
        self.saved = True

    def SaveAs(self, path: str) -> None:  # noqa: N802
        self.saved_as.append(path)
        self.FullName = path
        Path(path).write_text("ppt", encoding="utf-8")


def test_create_from_template_uses_untitled_nonmodal_open(monkeypatch, tmp_path: Path) -> None:
    template_path = tmp_path / "template.potx"
    template_path.touch()
    output_path = tmp_path / "from-template.pptx"
    open_calls: list[dict[str, object]] = []
    tracked: list[tuple[Path, object, str | None]] = []
    created = _FakePresentation(1)

    class FakePresentations:
        def Open(self, path: str, **kwargs):  # noqa: N802
            open_calls.append({"path": path, **kwargs})
            return created

    app = type("App", (), {"Presentations": FakePresentations()})()
    monkeypatch.setattr(ppt_ops, "validate_path", lambda raw: Path(raw))
    monkeypatch.setattr(
        ppt_ops.office_manager,
        "track_document",
        lambda file_path, doc, app_type=None: tracked.append((file_path, doc, app_type)) or doc,
    )

    result = _create_from_template(app, {"template_path": str(template_path), "output_path": str(output_path)})

    assert result == f"created_from_template: {output_path}"
    assert created.saved_as == [str(output_path)]
    assert tracked == [(output_path, created, "ppt")]
    assert open_calls == [{
        "path": str(template_path),
        "ReadOnly": False,
        "Untitled": True,
        "WithWindow": False,
    }]


def test_compare_presentations_closes_both_internal_sessions(monkeypatch, tmp_path: Path) -> None:
    paths = [tmp_path / "a.pptx", tmp_path / "b.pptx"]
    for path in paths:
        path.touch()
    presentations = [_FakePresentation(2), _FakePresentation(3)]
    open_calls: list[dict[str, object]] = []

    class FakePresentations:
        def Open(self, path: str, **kwargs):  # noqa: N802
            open_calls.append({"path": path, **kwargs})
            return presentations[len(open_calls) - 1]

    app = type("App", (), {"Presentations": FakePresentations()})()
    monkeypatch.setattr(ppt_ops, "validate_path", lambda raw: Path(raw))

    result = _compare_presentations(
        app,
        {
            "presentation1_path": str(paths[0]),
            "presentation2_path": str(paths[1]),
        },
    )

    assert result["status"] == "compared"
    assert presentations[0].closed is True
    assert presentations[1].closed is True
    assert open_calls == [
        {"path": str(paths[0]), "ReadOnly": True, "Untitled": False, "WithWindow": False},
        {"path": str(paths[1]), "ReadOnly": True, "Untitled": False, "WithWindow": False},
    ]


def test_compare_presentations_does_not_close_borrowed_sessions(monkeypatch, tmp_path: Path) -> None:
    paths = [tmp_path / "a.pptx", tmp_path / "b.pptx"]
    for path in paths:
        path.touch()
    borrowed = _FakePresentation(2)
    borrowed.FullName = str(paths[0])
    owned = _FakePresentation(3)
    open_calls: list[dict[str, object]] = []

    class FakePresentations:
        Count = 1

        def __call__(self, index: int):  # noqa: ANN204
            assert index == 1
            return borrowed

        def Open(self, path: str, **kwargs):  # noqa: N802
            open_calls.append({"path": path, **kwargs})
            return owned

    app = type("App", (), {"Presentations": FakePresentations()})()
    monkeypatch.setattr(ppt_ops, "validate_path", lambda raw: Path(raw))

    result = _compare_presentations(
        app,
        {
            "presentation1_path": str(paths[0]),
            "presentation2_path": str(paths[1]),
        },
    )

    assert result["status"] == "compared"
    assert borrowed.closed is False
    assert owned.closed is True
    assert open_calls == [
        {"path": str(paths[1]), "ReadOnly": True, "Untitled": False, "WithWindow": False},
    ]


def test_merge_presentations_saves_and_closes_internal_sessions(monkeypatch, tmp_path: Path) -> None:
    target_path = tmp_path / "target.pptx"
    source_path = tmp_path / "source.pptx"
    target_path.touch()
    source_path.touch()
    target = _FakePresentation(1)
    source = _FakePresentation(2)
    open_calls: list[dict[str, object]] = []

    class FakeSourceSlides(_FakeSlides):
        def __call__(self, index: int):
            return type("Slide", (), {"Copy": lambda self: None})()

    source.Slides = FakeSourceSlides(2)

    class FakePresentations:
        def Open(self, path: str, **kwargs):  # noqa: N802
            open_calls.append({"path": path, **kwargs})
            return target if len(open_calls) == 1 else source

    app = type("App", (), {"Presentations": FakePresentations()})()
    monkeypatch.setattr(ppt_ops, "validate_path", lambda raw: Path(raw))

    result = _merge_presentations(
        app,
        {"target_path": str(target_path), "source_paths": [str(source_path)], "insert_position": -1},
    )

    assert "merge_presentations" in result
    assert target.saved is True
    assert target.closed is True
    assert source.closed is True
    assert target.Slides.pasted_positions == [2, 3]
    assert open_calls == [
        {"path": str(target_path), "ReadOnly": False, "Untitled": False, "WithWindow": False},
        {"path": str(source_path), "ReadOnly": True, "Untitled": False, "WithWindow": False},
    ]


def test_merge_presentations_does_not_close_borrowed_target(monkeypatch, tmp_path: Path) -> None:
    target_path = tmp_path / "target.pptx"
    source_path = tmp_path / "source.pptx"
    target_path.touch()
    source_path.touch()
    target = _FakePresentation(1)
    target.FullName = str(target_path)
    source = _FakePresentation(2)
    open_calls: list[dict[str, object]] = []

    class FakeSourceSlides(_FakeSlides):
        def __call__(self, index: int):
            return type("Slide", (), {"Copy": lambda self: None})()

    source.Slides = FakeSourceSlides(2)

    class FakePresentations:
        Count = 1

        def __call__(self, index: int):  # noqa: ANN204
            assert index == 1
            return target

        def Open(self, path: str, **kwargs):  # noqa: N802
            open_calls.append({"path": path, **kwargs})
            return source

    app = type("App", (), {"Presentations": FakePresentations()})()
    monkeypatch.setattr(ppt_ops, "validate_path", lambda raw: Path(raw))

    result = _merge_presentations(
        app,
        {"target_path": str(target_path), "source_paths": [str(source_path)], "insert_position": -1},
    )

    assert "merge_presentations" in result
    assert target.saved is True
    assert target.closed is False
    assert source.closed is True
    assert open_calls == [
        {"path": str(source_path), "ReadOnly": True, "Untitled": False, "WithWindow": False},
    ]
