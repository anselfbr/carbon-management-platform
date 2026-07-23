from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
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

# Product Activity / Product Master template dropdown structure.
PRODUCT_DROPDOWN_FIELDS: dict[str, tuple[int, int]] = {
    "product_type": (1, 2),
    "data_source": (3, 4),
    "volume_unit": (5, 6),
    "working_hours_unit": (7, 8),
    "currency": (9, 10),
    "system_boundary": (11, 12),
    "declared_unit": (13, 14),
    "operator": (15, 16),
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

PRODUCT_DISPLAY_ALIASES: dict[str, dict[str, str]] = {
    "product_type": {
        "Target Product": "FINISHED_PRODUCT", "標的產品": "FINISHED_PRODUCT",
        "Finished Product/Co-product": "FINISHED_PRODUCT_OR_CO_PRODUCT",
        "成品／聯產品": "FINISHED_PRODUCT_OR_CO_PRODUCT",
        "Semi-finished Product (Treated as Finished)": "SEMI_FINISHED_AS_FINISHED",
        "半成品（視為成品）": "SEMI_FINISHED_AS_FINISHED",
        "Semi-finished Product": "SEMI_FINISHED_PRODUCT", "半成品": "SEMI_FINISHED_PRODUCT",
        "By-product": "BY_PRODUCT", "副產品": "BY_PRODUCT",
        "Waste/Scrap": "WASTE_SCRAP", "廢棄物／廢料": "WASTE_SCRAP",
        "Indirect Output": "INDIRECT_OUTPUT", "輔助／間接產出": "INDIRECT_OUTPUT",
    },
    "data_source": {
        "None": "NONE", "無": "NONE",
        "SAP": "SAP", "ERP": "ERP", "PLM": "PLM", "MES": "MES",
        "Other": "OTHER", "其他": "OTHER",
    },
    "working_hours_unit": {
        "Hours": "HOURS", "hours": "HOURS", "小時": "HOURS",
        "Minutes": "MINS", "分鐘": "MINS",
        "Seconds": "SEC", "秒": "SEC",
    },
    "system_boundary": {
        "Cradle-to-Gate": "CRADLE_TO_GATE", "搖籃到大門": "CRADLE_TO_GATE",
        "Cradle-to-Grave": "CRADLE_TO_GRAVE", "搖籃到墳墓": "CRADLE_TO_GRAVE",
    },
    "declared_unit": {
        "PC": "PC",
    },
}


def normalize_dropdown_value(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", text)


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def _cell_col_index(ref: str) -> int:
    match = re.match(r"[A-Z]+", ref or "")
    if not match:
        return 0
    value = 0
    for ch in match.group(0):
        value = value * 26 + ord(ch) - 64
    return value


def _xlsx_sheet_paths(path: str | Path) -> dict[str, str]:
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    ns_pkg = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    with zipfile.ZipFile(path, "r") as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relationships: dict[str, str] = {}
        for rel in rels_root.findall(f"{ns_pkg}Relationship"):
            rid = rel.attrib.get("Id")
            target = rel.attrib.get("Target", "")
            if not rid:
                continue
            if target.startswith("/xl/"):
                relationships[rid] = target.lstrip("/")
            elif target.startswith("xl/"):
                relationships[rid] = target
            else:
                relationships[rid] = "xl/" + target.lstrip("/")
        result: dict[str, str] = {}
        sheets = workbook.find(f"{ns_main}sheets")
        if sheets is not None:
            for sheet in sheets.findall(f"{ns_main}sheet"):
                name = sheet.attrib.get("name", "")
                rid = sheet.attrib.get(f"{ns_rel}id")
                if name and rid in relationships:
                    result[name] = relationships[rid]
        return result


def _xlsx_dropdown_rows(
    path: str | Path,
    sheet_names: Iterable[str],
    max_col: int,
) -> tuple[str, tuple[tuple[Any, ...], ...]]:
    ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    sheet_paths = _xlsx_sheet_paths(path)
    selected = next((name for name in sheet_names if name in sheet_paths), "")
    if not selected:
        return "", ()

    with zipfile.ZipFile(path, "r") as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{ns}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{ns}t")))

        root = ET.fromstring(zf.read(sheet_paths[selected]))
        data = root.find(f"{ns}sheetData")
        if data is None:
            return selected, ()

        # Keep every populated column in the worksheet, not only the columns
        # currently used by a resolver field. This makes the cache reusable when
        # the official template adds another dropdown pair without requiring a
        # second worksheet read.
        sparse_rows: list[dict[int, Any]] = []
        observed_max_col = int(max_col or 0)
        for row in data.findall(f"{ns}row"):
            try:
                row_idx = int(row.attrib.get("r", "0") or 0)
            except Exception:
                row_idx = 0
            if row_idx < 2:
                continue
            values: dict[int, Any] = {}
            for cell in row.findall(f"{ns}c"):
                col_idx = _cell_col_index(cell.attrib.get("r", ""))
                if col_idx < 1:
                    continue
                observed_max_col = max(observed_max_col, col_idx)
                typ = cell.attrib.get("t")
                value_node = cell.find(f"{ns}v")
                inline = cell.find(f"{ns}is")
                value: Any = None
                if typ == "s" and value_node is not None and value_node.text is not None:
                    try:
                        value = shared[int(value_node.text)]
                    except (ValueError, IndexError):
                        value = None
                elif typ == "inlineStr" and inline is not None:
                    value = "".join(t.text or "" for t in inline.iter(f"{ns}t"))
                elif value_node is not None:
                    value = value_node.text
                if value is not None:
                    values[col_idx] = value
            if values:
                sparse_rows.append(values)
        rows = tuple(
            tuple(row.get(col_idx) for col_idx in range(1, observed_max_col + 1))
            for row in sparse_rows
        )
        return selected, rows


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
        display_aliases: dict[str, dict[str, str]] | None = None,
    ) -> None:
        self.rows = rows
        self.fields = dict(fields or RAW_MATERIAL_DROPDOWN_FIELDS)
        self.sheet_name = sheet_name
        self.display_aliases = dict(display_aliases or RAW_MATERIAL_DISPLAY_ALIASES)
        self._maps: dict[str, DropdownFieldMap] = {}
        self._alias_to_key: dict[str, dict[str, str]] = {
            field: {normalize_dropdown_value(display): key for display, key in aliases.items()}
            for field, aliases in self.display_aliases.items()
        }
        self._build_maps()

    @classmethod
    def from_workbook(
        cls,
        workbook,
        fields: dict[str, tuple[int, int]] | None = None,
        sheet_names: Iterable[str] = DEFAULT_DROPDOWN_SHEET_NAMES,
        required: bool = False,
        display_aliases: dict[str, dict[str, str]] | None = None,
    ) -> "DropdownValuesCache":
        selected = ""
        available = list(getattr(workbook, "sheetnames", []) or [])
        for name in sheet_names:
            if name in available:
                selected = name
                break
        if not selected:
            if required:
                raise ValueError("正式 Bulk Template 缺少 Dropdown Values 分頁。")
            return cls((), fields=fields, sheet_name="", display_aliases=display_aliases)

        ws = workbook[selected]
        selected_fields = fields or RAW_MATERIAL_DROPDOWN_FIELDS
        max_col = max(
            int(getattr(ws, "max_column", 0) or 0),
            max((max(pair) for pair in selected_fields.values()), default=1),
        )
        # One and only one worksheet traversal. Keep all rows because supplier and
        # country dropdowns can contain thousands of entries.
        rows = tuple(
            tuple(values)
            for values in ws.iter_rows(min_row=2, min_col=1, max_col=max_col, values_only=True)
        )
        return cls(
            rows,
            fields=selected_fields,
            sheet_name=selected,
            display_aliases=display_aliases,
        )

    @classmethod
    def from_xlsx_path(
        cls,
        path: str | Path,
        fields: dict[str, tuple[int, int]] | None = None,
        sheet_names: Iterable[str] = DEFAULT_DROPDOWN_SHEET_NAMES,
        required: bool = False,
        display_aliases: dict[str, dict[str, str]] | None = None,
    ) -> "DropdownValuesCache":
        """Read the dropdown sheet directly from OpenXML without openpyxl save/load.

        This is used by Product Bulk generation, which preserves the official
        template package byte-for-byte except for the target worksheet cells.
        """
        selected_fields = fields or RAW_MATERIAL_DROPDOWN_FIELDS
        max_col = max((max(pair) for pair in selected_fields.values()), default=1)
        selected, rows = _xlsx_dropdown_rows(path, sheet_names, max_col=max_col)
        if not selected and required:
            raise ValueError("正式 Bulk Template 缺少 Dropdown Values 分頁。")
        return cls(
            rows,
            fields=selected_fields,
            sheet_name=selected,
            display_aliases=display_aliases,
        )

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
        for display, key in self.display_aliases.get(field, {}).items():
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
            "cached_columns": max((len(row) for row in self.rows), default=0),
            "fields": {
                name: len(mapping.key_to_display)
                for name, mapping in self._maps.items()
            },
            "read_passes": 1 if self.sheet_name else 0,
        }
