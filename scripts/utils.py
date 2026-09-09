from __future__ import annotations

from datetime import datetime, date, timezone
from pathlib import Path
from typing import Any
import yaml

UNKNOWN = "UNKNOWN"
MISSING = "MISSING"
NOT_PROVIDED = "NOT PROVIDED"


def load_yaml(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def parse_date(value: Any) -> str:
    if value is None or value == "":
        return MISSING
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%Y.%m.%d"):
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except ValueError:
            continue
    return "INVALID_DATE"


def parse_int(value: Any, default: int = 0) -> int:
    if value is None or value == "":
        return default
    try:
        return int(float(str(value).strip()))
    except ValueError:
        return default


def parse_progress(value: Any) -> float:
    if value is None or value == "":
        return -1.0
    try:
        return float(str(value).replace("%", "").strip())
    except ValueError:
        return -1.0


def split_people(value: Any) -> list[str]:
    if value is None or str(value).strip() == "":
        return [MISSING]
    text = str(value)
    for sep in [";", "|", "、"]:
        text = text.replace(sep, ",")
    people = [p.strip() for p in text.split(",") if p.strip()]
    return people or [MISSING]


def split_dependencies(value: Any) -> list[str]:
    if value is None or str(value).strip() == "":
        return []
    text = str(value)
    for sep in [";", "|", "、"]:
        text = text.replace(sep, ",")
    return [d.strip() for d in text.split(",") if d.strip()]


def bool_from_value(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    text = str(value).strip().lower()
    return text in {"true", "1", "yes", "y"}


def normalize_text(value: Any, missing: str = MISSING) -> str:
    if value is None:
        return missing
    text = str(value).strip()
    return text if text else missing
