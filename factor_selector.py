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

FACTOR_SELECTOR_VERSION = "CMP_MODULE3_STAGE2_20260703_V18"


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



def _resolve_lcia_metadata_columns(ws) -> dict[str, int]:
    """Resolve LCIA metadata columns by row-4 headers instead of fixed column letters."""
    aliases = {
        "activity_name": ["Activity Name"],
        "reference_product_name": ["Reference Product Name", "Reference Product", "Product Name"],
        "geography": ["Geography", "Geographical representativeness"],
        "reference_product_unit": ["Reference Product Unit", "Unit"],
    }
    resolved: dict[str, int] = {}
    header_row = 4
    for key, names in aliases.items():
        alias_keys = {_norm(name) for name in names}
        for col in range(1, ws.max_column + 1):
            if _norm(ws.cell(header_row, col).value) in alias_keys:
                resolved[key] = col
                break
        if key not in resolved and key == "reference_product_name":
            resolved[key] = 4  # ecoinvent LCIA common position; used only as a fallback for keyword search
        elif key not in resolved:
            raise ValueError(f"LCIA 檔案找不到欄位：{', '.join(names)}")
    return resolved

def _process_type_token(process_type: str | None) -> str:
    value = str(process_type or "all").strip().lower()
    if value in {"production", "production_only"}:
        return "production"
    if value in {"market_for", "market", "production_with_transport"}:
        return "market for"
    return ""


def _matches_process_type(activity_lower: str, process_type: str | None) -> bool:
    """Apply process type filtering to Activity Name.

    Production-only results must not include any Activity Name containing
    "market for". This prevents market datasets from appearing when users
    choose 僅生產.
    """
    value = str(process_type or "all").strip().lower()
    activity = str(activity_lower or "").lower()
    if value in {"production", "production_only"}:
        return "production" in activity and "market for" not in activity
    if value in {"market_for", "market", "production_with_transport"}:
        return "market for" in activity
    return True



def _format_emission_factor_unit(unit: Any) -> str:
    raw_unit = _text(unit).strip()
    if not raw_unit:
        return ""
    normalized = raw_unit.lower().replace(" ", "")
    if normalized.startswith("kgco2e/"):
        return raw_unit
    return f"kgCO2e / {raw_unit}"


_SOURCE_DISPLAY_NAMES = {
    "APOS": "Ecoinvent 3.12 APOS",
    "Cut-off": "Ecoinvent 3.12 Cut-off",
}

_LCIA_CACHE: dict[str, Dict[str, Any]] = {}
_FILTER_RESULT_CACHE: dict[tuple, list[int]] = {}
_FILTER_RESULT_CACHE_MAX = 256


def _source_display_name(source: str) -> str:
    return _SOURCE_DISPLAY_NAMES.get(source, source)


def _keyword_word_text(value: Any) -> str:
    """Normalize searchable text for fast whole-word / whole-phrase matching."""
    text = _text(value).lower()
    tokens = re.findall(r"[a-z0-9]+|[\u4e00-\u9fff]+", text, flags=re.I)
    return " " + " ".join(tokens) + " "


def _load_lcia_cache(path: str | Path, source: str) -> Dict[str, Any]:
    """Load LCIA rows once into memory so repeated keyword searches do not re-read Excel."""
    path = Path(path)
    cache_key = str(path.resolve())
    stat = path.stat()
    cached = _LCIA_CACHE.get(cache_key)
    if cached and cached.get("mtime") == stat.st_mtime and cached.get("size") == stat.st_size:
        return cached

    wb = load_workbook(path, read_only=True, data_only=True)
    if LCIA_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"{path.name} 找不到分頁：{LCIA_SHEET_NAME}")
    ws = wb[LCIA_SHEET_NAME]
    value_col = _resolve_lcia_target_column(ws)
    meta_cols = _resolve_lcia_metadata_columns(ws)

    rows: list[Dict[str, Any]] = []
    geographies: set[str] = set()
    display_source = _source_display_name(source)
    for values in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
        activity_name = _text(values[meta_cols["activity_name"] - 1] if len(values) >= meta_cols["activity_name"] else "")
        row_geography = _text(values[meta_cols["geography"] - 1] if len(values) >= meta_cols["geography"] else "")
        reference_product_name = _text(values[meta_cols["reference_product_name"] - 1] if len(values) >= meta_cols["reference_product_name"] else "")
        ref_unit = _text(values[meta_cols["reference_product_unit"] - 1] if len(values) >= meta_cols["reference_product_unit"] else "")
        factor_value = values[value_col - 1] if len(values) >= value_col else None
        if row_geography:
            geographies.add(row_geography)

        search_word_text = _keyword_word_text(activity_name + " " + reference_product_name)
        activity_lower = activity_name.lower()
        rows.append({
            "source": display_source,
            "source_key": source,
            "activity_name": activity_name,
            "geography": row_geography,
            "emission_factor": factor_value,
            "reference_product_unit": ref_unit,
            "emission_factor_unit": _format_emission_factor_unit(ref_unit),
            "reference_product_name": reference_product_name,
            "ipcc2021_gwp100": factor_value,
            "indicator": "IPCC 2021 | climate change: total (excl. biogenic CO2) | global warming potential (GWP100)",
            "_activity_lower": activity_lower,
            "_has_market_for": "market for" in activity_lower,
            "_has_production": "production" in activity_lower,
            "_search_word_text": search_word_text,
        })

    # Database file changed or was first loaded; remove old filter-result cache entries for this path.
    for key in list(_FILTER_RESULT_CACHE.keys()):
        if key and key[0] == cache_key:
            _FILTER_RESULT_CACHE.pop(key, None)

    cached = {
        "cache_key": cache_key,
        "mtime": stat.st_mtime,
        "size": stat.st_size,
        "source": source,
        "rows": rows,
        "geographies": geographies,
    }
    _LCIA_CACHE[cache_key] = cached
    return cached


def _keyword_matches_activity_or_reference(keyword: str, search_word_text: str) -> bool:
    """Match keyword only against Activity Name and Reference Product Name.

    English/alphanumeric input is matched by complete words or complete phrases.
    Example: "resin" matches "epoxy resin" but not "resinous"; "market for"
    matches the phrase. Non-English input continues to use containment.
    """
    key = str(keyword or "").strip().lower()
    if not key:
        return True
    if re.fullmatch(r"[a-z0-9][a-z0-9\s\-_/.,()+]*", key, flags=re.I):
        key_text = _keyword_word_text(key)
        return key_text in str(search_word_text or "")
    return key in str(search_word_text or "").lower()


def _filter_cache_key(cache: Dict[str, Any], keyword: str, geography: str | None, process_type: str | None) -> tuple:
    return (
        cache.get("cache_key"),
        cache.get("mtime"),
        cache.get("size"),
        str(keyword or "").strip().lower(),
        str(geography or "all").strip(),
        str(process_type or "all").strip().lower(),
    )


def _get_matching_indices(cache: Dict[str, Any], keyword: str, geography: str | None, process_type: str | None) -> list[int]:
    key = _filter_cache_key(cache, keyword, geography, process_type)
    cached = _FILTER_RESULT_CACHE.get(key)
    if cached is not None:
        return cached

    geography_key = str(geography or "all").strip()
    geography_all = geography_key.lower() == "all"
    process_value = str(process_type or "all").strip().lower()
    keyword_key = str(keyword or "").strip().lower()

    indices: list[int] = []
    for idx, row in enumerate(cache["rows"]):
        if not geography_all and row.get("geography") != geography_key:
            continue

        if process_value in {"production", "production_only"}:
            # 僅生產：只允許 production，且只要 Activity Name 含 market for 就排除。
            if row.get("_has_market_for") or not row.get("_has_production"):
                continue
        elif process_value in {"market_for", "market", "production_with_transport"}:
            if not row.get("_has_market_for"):
                continue
        elif not _matches_process_type(row.get("_activity_lower", ""), process_type):
            continue

        if not _keyword_matches_activity_or_reference(keyword_key, row.get("_search_word_text", "")):
            continue
        indices.append(idx)

    if len(_FILTER_RESULT_CACHE) >= _FILTER_RESULT_CACHE_MAX:
        _FILTER_RESULT_CACHE.pop(next(iter(_FILTER_RESULT_CACHE)), None)
    _FILTER_RESULT_CACHE[key] = indices
    return indices


def _search_lcia_file(
    path: str | Path,
    keyword: str,
    source: str,
    offset: int = 0,
    limit: int = 10,
    geography: str | None = "all",
    process_type: str | None = "all",
) -> tuple[list[Dict[str, Any]], int]:
    cache = _load_lcia_cache(path, source)
    indices = _get_matching_indices(cache, keyword, geography, process_type)
    total_count = len(indices)
    if limit <= 0:
        return [], total_count

    rows = cache["rows"]
    selected_indices = indices[max(0, offset): max(0, offset) + limit]
    results: list[Dict[str, Any]] = []
    for idx in selected_indices:
        row = rows[idx]
        clean_row = {k: v for k, v in row.items() if not k.startswith("_") and k != "source_key"}
        results.append(clean_row)
    return results, total_count


def collect_factor_library_geographies(*paths: str | Path | None) -> list[str]:
    """Return the common geography filters shown in the UI.

    The database can contain many Geography values, but the factor library UI is
    intentionally limited to the common review filters requested for CMP:
    GLO, RoW, and RER.
    """
    return ["GLO", "RoW", "RER"]

def search_factor_library(
    keyword: str,
    apos_path: str | Path | None,
    cutoff_path: str | Path | None,
    limit: int = 10,
    source: str = "all",
    geography: str = "all",
    process_type: str = "all",
    page: int = 1,
    page_size: int = 10,
) -> Dict[str, Any]:
    keyword = str(keyword or "").strip()
    if len(keyword) < 2:
        raise ValueError("請輸入至少 2 個字元的關鍵字")
    page_size = int(page_size or limit or 10)
    if page_size not in {10, 20, 50}:
        page_size = 10
    page = max(1, int(page or 1))
    selected_source = str(source or "all").strip().lower()
    start = (page - 1) * page_size

    results: list[Dict[str, Any]] = []
    total_count = 0

    apos_count = 0
    cutoff_count = 0

    if selected_source in {"all", "apos"} and apos_path and Path(apos_path).exists():
        # APOS is always prioritized. Fetch only the requested slice, not all matches.
        apos_results, apos_count = _search_lcia_file(
            apos_path, keyword, "APOS", offset=start, limit=page_size, geography=geography, process_type=process_type
        )
        results.extend(apos_results)
        total_count += apos_count

    if selected_source in {"all", "cut-off", "cutoff", "cut off"} and cutoff_path and Path(cutoff_path).exists():
        if selected_source == "all":
            # Page across APOS first, then Cut-off.
            if start < apos_count:
                cutoff_offset = 0
                cutoff_limit = max(0, page_size - len(results))
            else:
                cutoff_offset = start - apos_count
                cutoff_limit = page_size
        else:
            cutoff_offset = start
            cutoff_limit = page_size
        cutoff_results, cutoff_count = _search_lcia_file(
            cutoff_path, keyword, "Cut-off", offset=cutoff_offset, limit=cutoff_limit, geography=geography, process_type=process_type
        )
        results.extend(cutoff_results[:max(0, page_size - len(results))])
        total_count += cutoff_count

    total_pages = max(1, (total_count + page_size - 1) // page_size) if total_count else 0
    if total_pages and page > total_pages:
        page = total_pages

    return {
        "keyword": keyword,
        "count": len(results),
        "total_count": total_count,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "priority": "APOS first, then Cut-off",
        "filters": {
            "source": source or "all",
            "geography": geography or "all",
            "process_type": process_type or "all",
        },
        "results": results,
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
    }
