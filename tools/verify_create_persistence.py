from pathlib import Path

from office_mcp.core.office_manager import office_manager


def main() -> None:
    base = Path(r"D:\FakeC\MCP\offiiceMCP\demo_fix\direct")
    base.mkdir(parents=True, exist_ok=True)

    targets = [
        ("word", base / "direct-word.docx"),
        ("excel", base / "direct-excel.xlsx"),
    ]

    for _, path in targets:
        if path.exists():
            path.unlink()

    for label, path in targets:
        try:
            office_manager.create_document(path, overwrite=True)
            print(label, "created", path.exists(), str(path))
        except Exception as exc:  # noqa: BLE001
            print(label, "error", repr(exc))


if __name__ == "__main__":
    main()
