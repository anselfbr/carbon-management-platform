from __future__ import annotations
import json, re
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable

MAPPING_PATH = Path(__file__).with_name("template_mapping.json")

@lru_cache(maxsize=1)
def load_mapping() -> dict:
    with MAPPING_PATH.open("r", encoding="utf-8") as fh:
        return json.load(fh)

def normalize_header(value: Any) -> str:
    text = str(value or "").strip().lower().replace("\n", " ").replace("\r", " ")
    text = text.replace("（", "(").replace("）", ")").replace("／", "/")
    text = re.sub(r"\((optional|required|選填|必填)\)", "", text, flags=re.I)
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", text)

def aliases(profile: str, field: str, fallback: Iterable[str] = ()) -> list[str]:
    """Return configured aliases without collapsing display/internal variants.

    Downstream legacy readers do not all use the same normalization rules.
    Therefore ``supplier_name`` and ``Supplier Name (optional)`` must both be
    retained even though the common resolver normalizes them to the same key.
    """
    configured = load_mapping().get("profiles", {}).get(profile, {}).get("fields", {}).get(field, [])
    out: list[str] = []
    seen: set[str] = set()
    for item in [*configured, *list(fallback)]:
        text = str(item or "").strip()
        key = text.casefold()
        if text and key not in seen:
            seen.add(key)
            out.append(text)
    return out

def find_column(header_rows: list[list[Any]], profile: str, field: str, fallback_col: int | None = None) -> int | None:
    keys={normalize_header(x) for x in aliases(profile, field)}
    for row_idx in (1,0):
        if row_idx >= len(header_rows): continue
        for col_idx,value in enumerate(header_rows[row_idx],1):
            if normalize_header(value) in keys: return col_idx
    return fallback_col
