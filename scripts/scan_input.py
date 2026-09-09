from __future__ import annotations

import json
from pathlib import Path

try:
    from scripts.utils import ensure_dir
except ModuleNotFoundError:  # pragma: no cover
    from utils import ensure_dir


def scan_input_files(input_dir: str | Path) -> list[str]:
    input_path = Path(input_dir)
    files = []
    for p in input_path.rglob("*.xlsx"):
        if p.name.startswith("~$"):
            continue
        files.append(str(p))
    return sorted(files)


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    files = scan_input_files(repo_root / "Input")
    out_dir = repo_root / "data" / "raw"
    ensure_dir(out_dir)
    (out_dir / "input_files.json").write_text(json.dumps(files, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Detected {len(files)} xlsx files")


if __name__ == "__main__":
    main()
