from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Any, Iterable


DEFAULT_DROPDOWN_SHEET_NAMES = ("Dropdown Values", "Dropdown Value", "Dropdown")

# One visible display column and one internal-key column for each supported field.
# Column numbers are 1-based and follow the third-party Raw Material Bulk template.
RAW_MATERIAL_DROPDOWN_FIELDS: dict[str, tuple[int, int]] = {
    "document_type": (1, 2),
    "activity_data_unit": (3, 4),
    "weight_unit": (5, 6),
    "data_source": (7, 8),
    "calculate_transportation_emissions": (13, 14),
    "supplier_name_resolved": (15, 23),
    "supplier_code_resolved": (15, 24),
    "country_area": (22, 25),
}

# Known display aliases from the official English and Chinese Raw Material Bulk
# templates. They are used only to convert an upstream visible value back to its
# language-neutral key before rendering it with the current template.
RAW_MATERIAL_DISPLAY_ALIASES: dict[str, dict[str, str]] = {
    "document_type": {
        "Issued Quantity": "ISSUED_QTY", "領料數量": "ISSUED_QTY",
        "Purchased Quantity": "PURCHASED_QTY", "採購數量": "PURCHASED_QTY",
        "Actual Consumption": "ACTUAL_CONSUMPTION", "實際使用量": "ACTUAL_CONSUMPTION",
        "Bill of Materials (BOM)": "BOM", "物料清單（BOM）": "BOM", "物料清單(BOM)": "BOM",
    },
    "weight_unit": {
        "mg": "MG", "毫克": "MG",
        "g": "G", "公克": "G",
        "kg": "KG", "公斤": "KG",
        "tonnes": "T", "公噸": "T",
        "kt": "KT", "千公噸": "KT",
        "ounce": "OUNCE", "盎司": "OUNCE",
        "pound": "POUND", "磅": "POUND",
        "long ton": "LONG_TON", "長噸": "LONG_TON",
        "short ton": "SHORT_TON", "短噸": "SHORT_TON",
        "dry ton": "TDM", "乾公噸": "TDM",
    },
    "calculate_transportation_emissions": {
        "Yes": "YES", "是": "YES",
        "No": "NO", "否": "NO",
    },
    "country_area": {
        "TBD": "TBD", "待定": "TBD",
    },
}


def normalize_dropdown_value(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", text)


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


@dataclass(frozen=True)
class DropdownFieldMap:
    key_to_display: dict[str, str]
    display_to_key_exact: dict[str, Any]
    display_to_key_normalized: dict[str, Any]


class DropdownValuesCache:
    """Read and cache one complete ``Dropdown Values`` sheet in a single pass.

    The full sheet rows are held in memory once. Field-specific key/display maps
    are then derived from those cached rows, so production loops never re-read
    the workbook and helper-formula cache generation can reuse the same data.
    """

    def __init__(
        self,
        rows: tuple[tuple[Any, ...], ...],
        fields: dict[str, tuple[int, int]] | None = None,
        sheet_name: str = "",
    ) -> None:
        self.rows = rows
        self.fields = dict(fields or RAW_MATERIAL_DROPDOWN_FIELDS)
        self.sheet_name = sheet_name
        self._maps: dict[str, DropdownFieldMap] = {}
        self._alias_to_key: dict[str, dict[str, str]] = {
            field: {normalize_dropdown_value(display): key for display, key in aliases.items()}
            for field, aliases in RAW_MATERIAL_DISPLAY_ALIASES.items()
        }
        self._build_maps()

    @classmethod
    def from_workbook(
        cls,
        workbook,
        fields: dict[str, tuple[int, int]] | None = None,
        sheet_names: Iterable[str] = DEFAULT_DROPDOWN_SHEET_NAMES,
        required: bool = False,
    ) -> "DropdownValuesCache":
        selected = ""
        available = list(getattr(workbook, "sheetnames", []) or [])
        for name in sheet_names:
            if name in available:
                selected = name
                break
        if not selected:
            if required:
                raise ValueError("正式 Raw Material Bulk Template 缺少 Dropdown Values 分頁。")
            return cls((), fields=fields, sheet_name="")

        ws = workbook[selected]
        max_col = max(
            int(getattr(ws, "max_column", 0) or 0),
            max((max(pair) for pair in (fields or RAW_MATERIAL_DROPDOWN_FIELDS).values()), default=1),
        )
        # One and only one worksheet traversal. Keep all rows because supplier and
        # country dropdowns can contain thousands of entries.
        rows = tuple(
            tuple(values)
            for values in ws.iter_rows(min_row=2, min_col=1, max_col=max_col, values_only=True)
        )
        return cls(rows, fields=fields, sheet_name=selected)

    def _build_maps(self) -> None:
        for field, pair in self.fields.items():
            display_col, key_col = int(pair[0]), int(pair[1])
            key_to_display: dict[str, str] = {}
            exact: dict[str, Any] = {}
            normalized: dict[str, Any] = {}
            for row in self.rows:
                display = _text(row[display_col - 1] if display_col <= len(row) else None)
                key_raw = row[key_col - 1] if key_col <= len(row) else None
                key_text = _text(key_raw)
                if not display or not key_text:
                    continue
                key_to_display.setdefault(key_text.upper(), display)
                exact.setdefault(display, key_raw)
                normalized.setdefault(normalize_dropdown_value(display), key_raw)
            self._maps[field] = DropdownFieldMap(key_to_display, exact, normalized)

    def field_map(self, field: str) -> DropdownFieldMap:
        return self._maps.get(field, DropdownFieldMap({}, {}, {}))

    def key_to_display_map(self, field: str) -> dict[str, str]:
        return dict(self.field_map(field).key_to_display)

    def input_to_display_map(self, field: str) -> dict[str, str]:
        """Return a fast map accepting keys plus official English/Chinese labels."""
        data = self.field_map(field)
        result = dict(data.key_to_display)
        for display, key in data.display_to_key_exact.items():
            target = data.key_to_display.get(_text(key).upper())
            if target:
                result.setdefault(_text(display).upper(), target)
        for display, key in RAW_MATERIAL_DISPLAY_ALIASES.get(field, {}).items():
            target = data.key_to_display.get(_text(key).upper())
            if target:
                result.setdefault(_text(display).upper(), target)
        return result

    def display_to_key_maps(self, field: str) -> tuple[dict[str, Any], dict[str, Any]]:
        data = self.field_map(field)
        return dict(data.display_to_key_exact), dict(data.display_to_key_normalized)

    def canonical_key(self, field: str, value: Any, default: str = "") -> str:
        text = _text(value) or _text(default)
        if not text:
            return ""
        data = self.field_map(field)
        upper = text.upper()
        if upper in data.key_to_display:
            return upper
        if text in data.display_to_key_exact:
            return _text(data.display_to_key_exact[text]).upper()
        normalized = normalize_dropdown_value(text)
        if normalized in data.display_to_key_normalized:
            return _text(data.display_to_key_normalized[normalized]).upper()
        return self._alias_to_key.get(field, {}).get(normalized, upper)

    def display(self, field: str, internal_key: Any, default: str = "") -> str:
        original = _text(internal_key) or _text(default)
        if not original:
            return ""
        key = self.canonical_key(field, original)
        return self.field_map(field).key_to_display.get(key, original)

    def key(self, field: str, display_value: Any, default: Any = "") -> Any:
        display = _text(display_value)
        if not display:
            return default
        key = self.canonical_key(field, display)
        return key or default

    def summary(self) -> dict[str, Any]:
        return {
            "sheet_name": self.sheet_name,
            "cached_rows": len(self.rows),
            "fields": {
                name: len(mapping.key_to_display)
                for name, mapping in self._maps.items()
            },
            "read_passes": 1 if self.sheet_name else 0,
        }
