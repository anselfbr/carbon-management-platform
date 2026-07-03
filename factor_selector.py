from __future__ import annotations

import re
import shutil
from pathlib import Path
from typing import Any, Dict, Iterable

import pandas as pd
from openpyxl import load_workbook

ACTIVITY_SHEET_NAME = "Input Sheet Activity Data"
DATA_START_ROW = 3
CCL_SHEET_NAME = "02.料號CCL分類表"
LCIA_SHEET_NAME = "LCIA"

FACTOR_SELECTOR_VERSION = "CMP_MODULE3_STAGE2_20260703_V7"


def _norm(value: Any) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", text)


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value).strip()


def _find_header_row(ws, aliases: Iterable[str], max_scan_rows: int = 30) -> int:
    alias_keys = {_norm(a) for a in aliases}
    for row in range(1, min(ws.max_row, max_scan_rows) + 1):
        values = {_norm(ws.cell(row, col).value) for col in range(1, ws.max_column + 1)}
        if values & alias_keys:
            return row
    return 1


def _find_col(ws, aliases: list[str], header_rows: int = DATA_START_ROW - 1, required: bool = True) -> int | None:
    alias_keys = [_norm(a) for a in aliases if str(a or "").strip()]
    rows = list(range(1, max(1, header_rows) + 1))
    if 2 in rows:
        rows = [2] + [r for r in rows if r != 2]
    for row in rows:
        for col in range(1, ws.max_column + 1):
            if _norm(ws.cell(row, col).value) in alias_keys:
                return col
    if required:
        raise ValueError(f"找不到欄位：{', '.join(aliases)}")
    return None


def _find_col_in_header_row(ws, header_row: int, aliases: list[str], required: bool = True) -> int | None:
    alias_keys = [_norm(a) for a in aliases if str(a or "").strip()]
    for col in range(1, ws.max_column + 1):
        if _norm(ws.cell(header_row, col).value) in alias_keys:
            return col
    if required:
        raise ValueError(f"找不到欄位：{', '.join(aliases)}")
    return None


def _safe_number(value: Any) -> float | None:
    if value is None:
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text or text.lower() in {"nan", "none"}:
        return None
    try:
        return float(text)
    except ValueError:
        return None


def _read_ccl_mapping(ccl_path: str | Path) -> dict[str, Dict[str, Any]]:
    wb = load_workbook(ccl_path, read_only=True, data_only=True)
    sheet_name = CCL_SHEET_NAME if CCL_SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
    ws = wb[sheet_name]
    header_row = _find_header_row(ws, ["Material", "料號", "Material Number"])

    material_col = _find_col_in_header_row(ws, header_row, ["Material", "Material Number", "料號", "物料", "原物料料號"])
    ccl_item_col = _find_col_in_header_row(ws, header_row, ["CCL Item", "CCLItem", "CCL項目", "CCL分類", "Item", "項目"])
    factor_col = _find_col_in_header_row(ws, header_row, ["碳係數", "Emission Factor", "Carbon Factor", "EF", "係數"])
    unit_col = _find_col_in_header_row(ws, header_row, ["單位", "Unit", "Factor Unit", "Emission Factor Unit", "係數單位"], required=False)

    mapping: dict[str, Dict[str, Any]] = {}
    for row in range(header_row + 1, ws.max_row + 1):
        material = _text(ws.cell(row, material_col).value)
        if not material:
            continue
        key = material.upper()
        factor_value = ws.cell(row, factor_col).value
        mapping[key] = {
            "material": material,
            "ccl_item": _text(ws.cell(row, ccl_item_col).value),
            "emission_factor": _safe_number(factor_value) if _safe_number(factor_value) is not None else factor_value,
            "unit": _text(ws.cell(row, unit_col).value) if unit_col else "",
        }
    return mapping


def apply_ccl_factors_to_raw_material_bulk(
    raw_material_bulk_path: str | Path,
    ccl_mapping_path: str | Path,
    output_path: str | Path,
) -> Dict[str, Any]:
    """Fill Module 3 CCL factor fields into a Module 2 raw-material bulk workbook."""
    raw_material_bulk_path = Path(raw_material_bulk_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(raw_material_bulk_path, output_path)

    ccl_map = _read_ccl_mapping(ccl_mapping_path)
    wb = load_workbook(output_path)
    if ACTIVITY_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到分頁：{ACTIVITY_SHEET_NAME}")
    ws = wb[ACTIVITY_SHEET_NAME]

    cols = {
        "material": _find_col(ws, ["Raw Material Code", "Raw Material Number", "Material", "Material Number", "原物料代碼", "料號"]),
        "doc_start": _find_col(ws, ["Doc. Start Date", "Document Start Date", "開始日期"], required=False),
        "factor_name": _find_col(ws, ["Factor Name", "Emission Factor Name", "係數名稱"]),
        "emission_factor": _find_col(ws, ["Emission Factor", "Carbon Factor", "碳係數"]),
        "factor_source": _find_col(ws, ["Factor Source", "Emission Factor Source", "係數來源"], required=False),
        "factor_comment": _find_col(ws, ["Factor Comment", "Emission Factor Comment", "係數備註"], required=False),
        "country": _find_col(ws, ["Country/Area", "Country Area", "Country", "Area", "國家地區"], required=False),
        "enabled_date": _find_col(ws, ["Enabled Date", "Effective Date", "啟用日期"], required=False),
        "data_quality": _find_col(ws, ["Data Quality", "資料品質"], required=False),
        "factor_unit": _find_col(ws, ["Factor Unit", "Emission Factor Unit", "CF Unit", "係數單位"], required=False),
    }

    matched = 0
    unmatched = 0
    written_rows = 0
    for row in range(DATA_START_ROW, ws.max_row + 1):
        material = _text(ws.cell(row, cols["material"]).value)
        if not material:
            continue
        item = ccl_map.get(material.upper())
        if not item:
            unmatched += 1
            continue

        ws.cell(row, cols["factor_name"]).value = item["ccl_item"]
        ws.cell(row, cols["emission_factor"]).value = item["emission_factor"]
        if cols["factor_source"]:
            ws.cell(row, cols["factor_source"]).value = item["ccl_item"]
        if cols["factor_comment"]:
            ws.cell(row, cols["factor_comment"]).value = "無"
        if cols["country"]:
            ws.cell(row, cols["country"]).value = "GLO"
        if cols["enabled_date"]:
            ws.cell(row, cols["enabled_date"]).value = ws.cell(row, cols["doc_start"]).value if cols["doc_start"] else None
            ws.cell(row, cols["enabled_date"]).number_format = "yyyy/mm/dd"
        if cols["data_quality"]:
            ws.cell(row, cols["data_quality"]).value = "SECONDARY"
        if cols["factor_unit"] and item.get("unit"):
            ws.cell(row, cols["factor_unit"]).value = item["unit"]
        matched += 1
        written_rows += 1

    wb.save(output_path)
    return {
        "output_filename": output_path.name,
        "download_url": f"/download/{output_path.name}",
        "ccl_mapping_rows": len(ccl_map),
        "matched_rows": matched,
        "unmatched_rows": unmatched,
        "written_rows": written_rows,
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
    }


def _resolve_lcia_target_column(ws) -> int:
    """Find IPCC 2021 + climate change total excl. biogenic CO2 + GWP100 without fixed column letters."""
    candidates: list[tuple[int, int]] = []
    for col in range(1, ws.max_column + 1):
        method = _text(ws.cell(1, col).value).lower()
        category = _text(ws.cell(2, col).value).lower()
        indicator = _text(ws.cell(3, col).value).lower()
        if "ipcc 2021" not in method:
            continue
        if "climate change: total (excl. biogenic co2)" not in category:
            continue
        if "global warming potential (gwp100)" not in indicator:
            continue
        score = 0
        if "no lt" not in category and "no lt" not in method:
            score += 3
        if "incl. slcfs" not in category:
            score += 2
        candidates.append((score, col))
    if not candidates:
        raise ValueError("LCIA 檔案找不到 IPCC 2021 / climate change: total (excl. biogenic CO2) / GWP100 欄位")
    return sorted(candidates, reverse=True)[0][1]


def _process_type_token(process_type: str | None) -> str:
    value = str(process_type or "all").strip().lower()
    if value in {"production", "production_only"}:
        return "production"
    if value in {"market_for", "market", "production_with_transport"}:
        return "market for"
    return ""


def _search_lcia_file(
    path: str | Path,
    keyword: str,
    source: str,
    limit: int,
    geography: str | None = "all",
    process_type: str | None = "all",
) -> list[Dict[str, Any]]:
    wb = load_workbook(path, read_only=True, data_only=True)
    if LCIA_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"{Path(path).name} 找不到分頁：{LCIA_SHEET_NAME}")
    ws = wb[LCIA_SHEET_NAME]
    value_col = _resolve_lcia_target_column(ws)
    results: list[Dict[str, Any]] = []
    key = keyword.lower().strip()
    geography_key = str(geography or "all").strip()
    process_token = _process_type_token(process_type)
    # iter_rows is much faster than repeated ws.cell calls for the 26k+ row LCIA sheets.
    for values in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
        activity_name = _text(values[1] if len(values) > 1 else "")
        row_geography = _text(values[2] if len(values) > 2 else "")
        if geography_key.lower() != "all" and row_geography != geography_key:
            continue
        if process_token and process_token not in activity_name.lower():
            continue
        searchable = values[:6]
        haystack = " ".join(_text(v).lower() for v in searchable)
        if key not in haystack:
            continue
        results.append({
            "source": source,
            "activity_name": activity_name,
            "geography": row_geography,
            "reference_product_name": _text(values[3] if len(values) > 3 else ""),
            "reference_product_unit": _text(values[4] if len(values) > 4 else ""),
            "ipcc2021_gwp100": values[value_col - 1] if len(values) >= value_col else None,
            "indicator": "IPCC 2021 | climate change: total (excl. biogenic CO2) | global warming potential (GWP100)",
        })
        if len(results) >= limit:
            break
    return results


def collect_factor_library_geographies(*paths: str | Path | None) -> list[str]:
    values: set[str] = set()
    for path in paths:
        if not path or not Path(path).exists():
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        if LCIA_SHEET_NAME not in wb.sheetnames:
            continue
        ws = wb[LCIA_SHEET_NAME]
        for row_values in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=3, max_col=3, values_only=True):
            geo = _text(row_values[0] if row_values else "")
            if geo:
                values.add(geo)
    return sorted(values, key=lambda x: (x != "GLO", x.lower()))


def search_factor_library(
    keyword: str,
    apos_path: str | Path | None,
    cutoff_path: str | Path | None,
    limit: int = 80,
    source: str = "all",
    geography: str = "all",
    process_type: str = "all",
) -> Dict[str, Any]:
    keyword = str(keyword or "").strip()
    if len(keyword) < 2:
        raise ValueError("請輸入至少 2 個字元的關鍵字")
    limit = max(1, min(int(limit or 80), 200))
    selected_source = str(source or "all").strip().lower()
    results: list[Dict[str, Any]] = []
    if selected_source in {"all", "apos"} and apos_path and Path(apos_path).exists():
        results.extend(_search_lcia_file(apos_path, keyword, "APOS", limit, geography, process_type))
    remaining = limit - len(results)
    if remaining > 0 and selected_source in {"all", "cut-off", "cutoff", "cut off"} and cutoff_path and Path(cutoff_path).exists():
        results.extend(_search_lcia_file(cutoff_path, keyword, "Cut-off", remaining, geography, process_type))
    return {
        "keyword": keyword,
        "count": len(results),
        "priority": "APOS first, then Cut-off",
        "filters": {
            "source": source or "all",
            "geography": geography or "all",
            "process_type": process_type or "all",
        },
        "results": results,
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
    }
