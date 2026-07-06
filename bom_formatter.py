from __future__ import annotations

import gc
import re
import shutil
import tempfile
import zipfile
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from openpyxl import load_workbook


ACTIVITY_SHEET_NAME = "Input Sheet Activity Data"
RAW_MATERIAL_SHEET_NAME = "Input Sheet Raw Material"
DATA_START_ROW = 3
BOM_FORMATTER_VERSION = "CMP_V17_6_MODULE2_HYBRID_UI_MEMORY_OPT"


DEFAULT_MAPPING = {
    "parent_col": "Parent Node",
    "component_col": "Component",
    "qty_col": "CS03 Qty",
    "unit_col": "CS03 UoM",
    "description_col": "Component Description",
    "material_group_col": "Material group",
    "valid_from_col": "BOM Valid From",
    "net_weight_col": "Net weight",
    "gross_weight_col": "Gross weight",
    "weight_uom_col": "Weight UoM",
    "altitem_group_col": "Altitem group",
    "usage_probability_col": "Usage probability%",
}


def _normalize_col(value: Any) -> str:
    return str(value or "").strip().replace("\n", " ").replace("\r", " ")


def _resolve_mapping(mapping: dict[str, str | None] | None) -> dict[str, str]:
    mapping = mapping or {}
    resolved = {}
    for key, default_value in DEFAULT_MAPPING.items():
        value = str(mapping.get(key) or "").strip()
        resolved[key] = value if value else default_value
    return resolved


def _find_column(df: pd.DataFrame, column_name: str) -> str:
    target = _normalize_col(column_name).lower()
    normalized = {_normalize_col(c).lower(): c for c in df.columns}
    if target in normalized:
        return normalized[target]
    raise ValueError(f"找不到 BOM 欄位：{column_name}")


def _find_optional_column(df: pd.DataFrame, column_name: str) -> str | None:
    try:
        return _find_column(df, column_name)
    except ValueError:
        return None


def _safe_text(value: Any) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _safe_number(value: Any) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value).strip()
    if text.upper() in ["", "NAN", "NONE"]:
        return 0.0
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return 0.0


def _normalize_altitem_group(value: Any) -> str:
    """Normalize SAP Altitem group values for grouping alternative materials.

    Excel sometimes reads group values as 1.0 instead of 1.  Empty, zero,
    and non-numeric values are treated as no alternative group.
    """
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.upper() in ["", "NAN", "NONE"]:
        return ""
    try:
        number = float(text.replace(",", ""))
    except ValueError:
        return text
    if number == 0:
        return ""
    if number.is_integer():
        return str(int(number))
    return str(number)


def _usage_probability_ratio(value: Any) -> float | None:
    """Convert Usage probability% to a multiplier.

    Supports both SAP-style percent values (50 -> 0.5) and ratio values
    already stored as decimals (0.5 -> 0.5). Blank/invalid values return
    None so normal BOM quantity remains unchanged.
    """
    if pd.isna(value):
        return None
    text = str(value).strip().replace("%", "")
    if text.upper() in ["", "NAN", "NONE"]:
        return None
    try:
        number = float(text.replace(",", ""))
    except ValueError:
        return None
    if number < 0:
        return None
    return number / 100.0 if number > 1 else number


def _apply_altitem_usage_probability(df: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Apply alternative-item usage probability to CS03 quantity.

    Rule: within the same parent material, rows with the same non-empty
    Altitem group are alternative materials. For such duplicated groups,
    effective quantity = CS03 Qty × Usage probability%.

    Memory V1: operate on the already-normalized BOM frame in-place.
    """
    work = df
    if "_altitem_group" not in work.columns:
        work["_altitem_group"] = ""
    if "_usage_probability_ratio" not in work.columns:
        work["_usage_probability_ratio"] = None

    alt_mask = work["_altitem_group"].astype(str).str.strip() != ""
    if not alt_mask.any():
        work["_qty_original"] = work["_qty"]
        work["_qty_adjusted_by_altitem"] = False
        return work, {
            "altitem_rows": 0,
            "altitem_groups": 0,
            "altitem_adjusted_rows": 0,
            "altitem_probability_missing_rows": 0,
        }

    group_keys = ["_parent", "_altitem_group"]
    group_sizes = work.groupby(group_keys, dropna=False)["_component"].transform("count")
    duplicated_alt_group = alt_mask & (group_sizes > 1)
    has_probability = work["_usage_probability_ratio"].apply(lambda x: x is not None)
    apply_mask = duplicated_alt_group & has_probability

    work["_qty_original"] = work["_qty"]
    work.loc[apply_mask, "_qty"] = work.loc[apply_mask, "_qty"] * work.loc[apply_mask, "_usage_probability_ratio"].astype(float)
    work["_qty_adjusted_by_altitem"] = apply_mask

    alt_groups = (
        work.loc[duplicated_alt_group, group_keys]
        .drop_duplicates()
        .shape[0]
    )
    missing_probability_rows = int((duplicated_alt_group & ~has_probability).sum())
    return work, {
        "altitem_rows": int(duplicated_alt_group.sum()),
        "altitem_groups": int(alt_groups),
        "altitem_adjusted_rows": int(apply_mask.sum()),
        "altitem_probability_missing_rows": missing_probability_rows,
        "altitem_rule": "Effective CS03 Qty = CS03 Qty × Usage probability% for duplicated Altitem group within the same Parent Node.",
    }


_XML_NUMERIC_CHAR_REF_RE = re.compile(rb"&#(?:x([0-9A-Fa-f]+)|([0-9]+));")
_XML_INVALID_ASCII_CONTROL_RE = re.compile(rb"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def _is_valid_xml_char_number(codepoint: int) -> bool:
    """Return True if the code point is valid in XML 1.0."""
    return (
        codepoint in (0x09, 0x0A, 0x0D)
        or 0x20 <= codepoint <= 0xD7FF
        or 0xE000 <= codepoint <= 0xFFFD
        or 0x10000 <= codepoint <= 0x10FFFF
    )


def _clean_invalid_xml_bytes(data: bytes) -> bytes:
    """Remove invalid XML 1.0 character references/control bytes from XLSX XML parts.

    Some SAP-exported Excel files contain strings such as ``&#11;`` or
    ``&#x0B;`` inside sharedStrings.xml. openpyxl rejects those XML files with
    "reference to invalid character number". Removing only invalid XML
    characters keeps the workbook readable without changing normal BOM values.
    """

    def replace_numeric_ref(match: re.Match[bytes]) -> bytes:
        raw_hex, raw_dec = match.groups()
        try:
            codepoint = int(raw_hex, 16) if raw_hex is not None else int(raw_dec, 10)
        except Exception:
            return b""
        return match.group(0) if _is_valid_xml_char_number(codepoint) else b""

    data = _XML_NUMERIC_CHAR_REF_RE.sub(replace_numeric_ref, data)
    data = _XML_INVALID_ASCII_CONTROL_RE.sub(b"", data)
    return data


def _repair_xlsx_invalid_xml(source_path: str | Path, repaired_path: str | Path) -> Path:
    """Create a temporary XLSX copy with invalid XML characters removed."""
    source_path = Path(source_path)
    repaired_path = Path(repaired_path)
    with zipfile.ZipFile(source_path, "r") as zin, zipfile.ZipFile(repaired_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.lower().endswith((".xml", ".rels")):
                data = _clean_invalid_xml_bytes(data)
            zout.writestr(item, data)
    return repaired_path


def _read_excel_first_sheet(path: str | Path) -> pd.DataFrame:
    """Read the first sheet from a sanitized XLSX copy.

    SAP/exported BOM workbooks can contain invalid XML numeric character
    references such as ``&#11;`` in sharedStrings.xml. openpyxl fails before
    pandas can build the DataFrame. Therefore we always sanitize XML parts into
    a temporary workbook first, then read that repaired workbook.
    """
    path = Path(path)
    with tempfile.TemporaryDirectory() as tmp_dir:
        repaired_path = Path(tmp_dir) / f"repaired_{path.name}"
        try:
            _repair_xlsx_invalid_xml(path, repaired_path)
            return pd.read_excel(repaired_path, sheet_name=0, dtype=object)
        except Exception as repaired_exc:
            # Fall back to direct read only if the repair/copy itself was the issue.
            # If direct read also fails, return the clearer repaired-read error.
            try:
                return pd.read_excel(path, sheet_name=0, dtype=object)
            except Exception:
                raise ValueError(f"Excel XML 修復後仍無法讀取：{repaired_exc}") from repaired_exc


def _date_from_value(value: Any) -> date:
    if pd.isna(value) or value in [None, ""]:
        return date(datetime.now().year, 1, 1)
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return date(datetime.now().year, 1, 1)
    return parsed.date()


def _year_start(d: date) -> date:
    return date(d.year, 1, 1)

def _year_end(d: date) -> date:
    return date(d.year, 12, 31)


def _clear_target_cells(ws, start_row: int, columns: list[int]) -> None:
    max_row = max(ws.max_row, start_row)
    for row_idx in range(start_row, max_row + 1):
        for col_idx in columns:
            ws.cell(row_idx, col_idx).value = None

def _normalize_template_header(value: Any) -> str:
    """Normalize a bulk-template header for resilient column matching."""
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", "", text)


def _find_template_columns(ws, aliases: list[str], header_rows: int = DATA_START_ROW - 1) -> list[int]:
    """Find worksheet columns by exact header text in the uploaded bulk template.

    V14.2 rule: Module 2 output follows the user's uploaded template instead
    of a historical fixed column order. Matching is exact after normalization to
    avoid accidental matches such as Raw Material Name -> Raw Material Code.
    """
    alias_keys = [_normalize_template_header(a) for a in aliases if str(a or "").strip()]
    if not alias_keys:
        return []

    found: list[int] = []
    max_header_row = max(1, int(header_rows or 1))
    row_order = list(range(1, max_header_row + 1))
    if 2 in row_order:
        row_order = [2] + [r for r in row_order if r != 2]

    for row in row_order:
        for col in range(1, ws.max_column + 1):
            header_key = _normalize_template_header(ws.cell(row, col).value)
            if header_key and header_key in alias_keys and col not in found:
                found.append(col)
    return found


def _find_template_column(ws, aliases: list[str], fallback_col: int | None = None) -> int:
    cols = _find_template_columns(ws, aliases)
    if cols:
        return int(cols[0])
    if fallback_col is not None:
        return int(fallback_col)
    raise ValueError(f"Bulk template 缺少必要欄位：{', '.join(aliases)}")


def _write_template_value(ws, row_idx: int, col_idx: int | None, value: Any) -> None:
    if col_idx:
        ws.cell(row_idx, int(col_idx)).value = value


def _clear_template_columns(ws, start_row: int, columns: list[int]) -> None:
    unique_columns = sorted({int(c) for c in columns if c})
    if unique_columns:
        _clear_target_cells(ws, start_row, unique_columns)

RAW_MATERIAL_NAME_ALIASES = ["raw_material_name", "Raw Material Name", "原物料名稱", "原料名稱"]
RAW_MATERIAL_CODE_ALIASES = ["raw_material_code", "Raw Material Code", "Raw Material ID", "Raw Material Number", "原物料代碼", "原料代碼"]
RAW_MATERIAL_DESC_ALIASES = ["raw_material_description", "Raw Material Description (Optional)", "Raw Material Description", "Description", "原物料描述", "品名"]
DOC_START_DATE_ALIASES = ["doc_start_date", "Doc. Start Date", "Document Start Date", "開始日期"]
DOC_END_DATE_ALIASES = ["doc_end_date", "Doc. End Date", "Document End Date", "結束日期"]
DOCUMENT_TYPE_ALIASES = ["document_type", "Document Type", "文件類型"]
DOCUMENT_NUMBER_ALIASES = ["document_number", "Document Number (optional)", "Document Number", "文件號碼"]
USAGE_ALIASES = ["usage", "Usage", "用量"]
ACTIVITY_DATA_UNIT_ALIASES = ["activity_data_unit", "Activity Data Unit", "活動數據單位", "單位"]
DATA_SOURCE_ALIASES = ["data_source", "Data Source", "資料來源"]
DATA_SOURCE_OTHER_ALIASES = ["data_source_other", "Data Source Other", "其他資料來源"]
TRANSPORT_ORIGIN_ALIASES = ["transportation_origin", "Transportation Origin", "運輸起點"]
TRANSPORT_DESTINATION_ALIASES = ["transportation_destination", "Transportation Destination", "運輸終點"]
SUPPLIER_NAME_ALIASES = ["supplier_name", "Supplier Name (optional)", "Supplier Name", "供應商名稱"]
PRODUCT_LINK_ALIASES = ["allocated_target_product_service", "Allocated Target Product/Service", "Target Product", "Product Code", "Product Name", "產品代碼", "產品名稱"]
COMMENT_ALIASES = ["comment", "Comment (optional)", "Comment", "備註"]
MATERIAL_GROUP_ALIASES = ["material_group", "Material Group", "Material group", "物料群組"]
NET_WEIGHT_ALIASES = ["net_weight", "Net Weight (optional)", "Net Weight", "Net weight", "淨重"]
GROSS_WEIGHT_ALIASES = ["gross_weight", "Gross Weight (optional)", "Gross Weight", "Gross weight", "毛重"]
WEIGHT_UNIT_ALIASES = ["weight_unit", "Weight Unit (optional)", "Weight Unit", "Weight UoM", "Weight UOM", "重量單位"]


def _sheet_has_dropdown_label(wb, dropdown_column_name: str, label: str) -> bool:
    """Check whether a visible dropdown label exists in the template.

    The user's bulk template has visible input columns and hidden key/helper
    columns. For Document Type, the visible value is "Bill of Materials (BOM)"
    while the helper formula converts it to key "BOM". Writing the key into
    the visible column causes the helper formula to fail.
    """
    if "Dropdown Values" not in wb.sheetnames:
        return False
    ws = wb["Dropdown Values"]
    target_header = _normalize_template_header(dropdown_column_name)
    target_label = str(label or "").strip()
    if not target_header or not target_label:
        return False

    candidate_cols: list[int] = []
    for col in range(1, ws.max_column + 1):
        if _normalize_template_header(ws.cell(1, col).value) == target_header:
            candidate_cols.append(col)
    for col in candidate_cols:
        for row in range(2, min(ws.max_row, 20000) + 1):
            if str(ws.cell(row, col).value or "").strip() == target_label:
                return True
    return False


def _document_type_for_template(wb) -> str:
    """Return the visible Document Type value expected by this template."""
    preferred_label = "Bill of Materials (BOM)"
    return preferred_label if _sheet_has_dropdown_label(wb, "Document Type", preferred_label) else "BOM"

def _read_bom(bom_path: str | Path, mapping: dict[str, str | None] | None = None) -> tuple[pd.DataFrame, dict[str, str]]:
    df = _read_excel_first_sheet(bom_path)
    m = _resolve_mapping(mapping)

    parent_col = _find_column(df, m["parent_col"])
    component_col = _find_column(df, m["component_col"])
    qty_col = _find_column(df, m["qty_col"])
    unit_col = _find_column(df, m["unit_col"])
    description_col = _find_optional_column(df, m["description_col"])
    material_group_col = _find_optional_column(df, m["material_group_col"])
    valid_from_col = _find_optional_column(df, m["valid_from_col"])
    net_weight_col = _find_optional_column(df, m["net_weight_col"])
    gross_weight_col = _find_optional_column(df, m["gross_weight_col"])
    weight_uom_col = _find_optional_column(df, m["weight_uom_col"])
    altitem_group_col = _find_optional_column(df, m["altitem_group_col"])
    usage_probability_col = _find_optional_column(df, m["usage_probability_col"])

    # Memory V1: keep only the columns required by Module 2 before creating
    # normalized helper columns. This prevents carrying unused SAP columns
    # through BOM explosion / structure / roll-up.
    required_source_cols = [
        parent_col, component_col, qty_col, unit_col, description_col, material_group_col,
        valid_from_col, net_weight_col, gross_weight_col, weight_uom_col,
        altitem_group_col, usage_probability_col,
    ]
    required_source_cols = list(dict.fromkeys([c for c in required_source_cols if c]))
    df = df.loc[:, required_source_cols]
    df["_parent"] = df[parent_col].apply(_safe_text)
    df["_component"] = df[component_col].apply(_safe_text)
    df["_qty"] = df[qty_col].apply(_safe_number)
    df["_uom"] = df[unit_col].apply(_safe_text)
    df["_description"] = df[description_col].apply(_safe_text) if description_col else ""
    df["_material_group"] = df[material_group_col].apply(_safe_text) if material_group_col else ""
    df["_valid_from"] = df[valid_from_col].apply(_date_from_value) if valid_from_col else date(datetime.now().year, 1, 1)
    df["_net_weight"] = df[net_weight_col].apply(_safe_number) if net_weight_col else ""
    df["_gross_weight"] = df[gross_weight_col].apply(_safe_number) if gross_weight_col else ""
    df["_weight_uom"] = df[weight_uom_col].apply(_safe_text) if weight_uom_col else ""
    df["_altitem_group"] = df[altitem_group_col].apply(_normalize_altitem_group) if altitem_group_col else ""
    df["_usage_probability_ratio"] = df[usage_probability_col].apply(_usage_probability_ratio) if usage_probability_col else None

    df = df.loc[(df["_parent"] != "") & (df["_component"] != "")].reset_index(drop=True)
    df, altitem_summary = _apply_altitem_usage_probability(df)

    used_columns = {
        "parent_col": parent_col,
        "component_col": component_col,
        "qty_col": qty_col,
        "unit_col": unit_col,
        "description_col": description_col or "",
        "material_group_col": material_group_col or "",
        "valid_from_col": valid_from_col or "",
        "net_weight_col": net_weight_col or "",
        "gross_weight_col": gross_weight_col or "",
        "weight_uom_col": weight_uom_col or "",
        "altitem_group_col": altitem_group_col or "",
        "usage_probability_col": usage_probability_col or "",
        **altitem_summary,
    }
    return df, used_columns




def _as_bom_path_list(bom_path: str | Path | list[str | Path] | tuple[str | Path, ...]) -> list[Path]:
    """Normalize single or multiple BOM paths while preserving backward compatibility."""
    if isinstance(bom_path, (list, tuple)):
        return [Path(p) for p in bom_path]
    return [Path(bom_path)]


def _read_boms(
    bom_paths: str | Path | list[str | Path] | tuple[str | Path, ...],
    mapping: dict[str, str | None] | None = None,
) -> tuple[pd.DataFrame, dict[str, Any]]:
    """Read and merge one or multiple standard BOM Excel files.

    Only fully identical normalized BOM rows are removed. Rows with different
    quantities, units, descriptions, material groups, or valid-from dates are
    preserved to avoid accidental data loss.
    """
    paths = _as_bom_path_list(bom_paths)
    if not paths:
        raise ValueError("請至少上傳一個 Standard BOM Excel 檔案")

    frames: list[pd.DataFrame] = []
    used_columns: dict[str, str] | None = None
    source_rows: list[dict[str, Any]] = []
    errors: list[str] = []

    for path in paths:
        try:
            df, cols = _read_bom(path, mapping=mapping)
            df["_source_file"] = path.name
            frames.append(df)
            used_columns = used_columns or cols
            source_rows.append({"filename": path.name, "rows": int(len(df))})
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")

    if errors:
        raise ValueError("；".join(errors))
    if not frames:
        raise ValueError("沒有可處理的 BOM 資料")

    merged = pd.concat(frames, ignore_index=True)
    before_dedup = int(len(merged))
    dedup_subset = ["_parent", "_component", "_qty", "_uom", "_description", "_material_group", "_valid_from", "_net_weight", "_gross_weight", "_weight_uom", "_altitem_group", "_usage_probability_ratio"]
    merged = merged.drop_duplicates(subset=dedup_subset, keep="first").reset_index(drop=True)
    after_dedup = int(len(merged))

    used = dict(used_columns or {})
    used["bom_files"] = int(len(paths))
    used["bom_rows_before_dedup"] = before_dedup
    used["bom_rows_after_dedup"] = after_dedup
    used["bom_duplicate_rows_removed"] = before_dedup - after_dedup
    used["bom_source_files"] = source_rows
    return merged, used

def _exclude_zero_usage_rows(exploded: pd.DataFrame) -> tuple[pd.DataFrame, int]:
    """Exclude true zero-usage raw material rows for Raw Material Bulk output.

    Rule: rows whose final calculated usage equals exactly 0 are not exported
    to raw_material_activity_data_bulk. Non-zero decimal usages are preserved.
    """
    if exploded is None or exploded.empty or "usage" not in exploded.columns:
        return exploded.copy() if isinstance(exploded, pd.DataFrame) else pd.DataFrame(), 0

    work = exploded.copy()
    usage_numeric = pd.to_numeric(work["usage"], errors="coerce").fillna(0.0)
    keep_mask = usage_numeric != 0
    excluded_rows = int((~keep_mask).sum())
    if excluded_rows:
        work = work.loc[keep_mask].copy().reset_index(drop=True)
    return work, excluded_rows



def _exclude_zero_total_working_hour_target_rows(
    exploded: pd.DataFrame,
    total_hour_by_material: dict[str, float] | None,
) -> tuple[pd.DataFrame, int]:
    """Exclude raw material rows whose target product has zero total working hours.

    Rule: if a Target Product is found in the Step1/BOM working-hour roll-up map
    and its Total Annual Working Hour equals 0, all raw material activity rows
    under that Target Product are excluded from raw_material_activity_data_bulk.
    Products not found in the map are kept to avoid silently dropping rows when
    Step1 data is incomplete.
    """
    if exploded is None or exploded.empty or "target_product" not in exploded.columns:
        return exploded.copy() if isinstance(exploded, pd.DataFrame) else pd.DataFrame(), 0
    if not total_hour_by_material:
        return exploded.copy(), 0

    work = exploded.copy()
    target_keys = work["target_product"].apply(_normalize_material_key)
    known_zero_targets = {
        str(k).strip().upper()
        for k, v in total_hour_by_material.items()
        if str(k or "").strip() and float(v or 0.0) == 0.0
    }
    drop_mask = target_keys.isin(known_zero_targets)
    excluded_rows = int(drop_mask.sum())
    if excluded_rows:
        work = work.loc[~drop_mask].copy().reset_index(drop=True)
    return work, excluded_rows


def _calculate_total_working_hour_by_target(
    step1_output_path: str | Path,
    bom_df: pd.DataFrame,
) -> tuple[dict[str, float], Dict[str, Any]]:
    """Calculate Target Product total annual working hours including semi-finished roll-up.

    Total Annual Working Hour = direct annual working hour of the target product
    + annual working-hour contribution from semi-finished components in the BOM.
    This mirrors generate_working_hour_rollup_file but returns a lightweight map
    for filtering Raw Material Bulk rows.
    """
    step1_output_path = Path(step1_output_path)
    try:
        step1_df = pd.read_excel(step1_output_path, sheet_name=STEP1_SOURCE_SHEET_NAME, dtype=object)
    except Exception:
        step1_df = pd.read_excel(step1_output_path, sheet_name=0, dtype=object)

    material_col = _find_step1_column(step1_df, ["Material Number", "Material", "Product Material Number"])
    qty_col = _find_step1_optional_column(step1_df, ["年度生產量", "Annual Quantity", "Delivered quantity"])
    hour_col = _find_step1_optional_column(step1_df, ["年度總工時", "Total working hours", "Selected Hours", "Total Hours", "Working Hours"])
    if not hour_col:
        return {}, {
            "working_hour_filter_applied": False,
            "working_hour_filter_reason": "Step1 Output 找不到年度總工時欄位，未套用 Total Annual Working Hour = 0 過濾。",
            "zero_total_working_hour_targets": 0,
        }

    work = step1_df.copy()
    work["_material_key"] = work[material_col].apply(_normalize_material_key)
    work["_annual_qty"] = work[qty_col].apply(_safe_number) if qty_col else 0.0
    work["_direct_hour"] = work[hour_col].apply(_safe_number)
    work = work[work["_material_key"] != ""].copy()

    material_totals = work.groupby(["_material_key"], dropna=False, as_index=False).agg({"_annual_qty": "sum", "_direct_hour": "sum"})
    qty_by_material: dict[str, float] = {}
    direct_by_material: dict[str, float] = {}
    hour_per_pc_by_material: dict[str, float] = {}
    for _, r in material_totals.iterrows():
        material = str(r["_material_key"] or "").strip().upper()
        qty = float(r["_annual_qty"] or 0.0)
        hours = float(r["_direct_hour"] or 0.0)
        qty_by_material[material] = qty
        direct_by_material[material] = hours
        hour_per_pc_by_material[material] = hours / qty if qty else 0.0

    structure, _structure_summary = _explode_bom_structure(bom_df)
    semi_by_target: dict[str, float] = {}
    if not structure.empty:
        b = structure.copy()
        b["_target_key"] = b["Target Product"].apply(_normalize_material_key)
        b["_component_key"] = b["Component"].apply(_normalize_material_key)
        b["_accumulated_qty"] = b["Accumulated Quantity"].apply(_safe_number)
        b = b[b["Is Semi-finished"].astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1"])].copy()
        for _, edge in b.iterrows():
            target = str(edge["_target_key"] or "").strip().upper()
            semi = str(edge["_component_key"] or "").strip().upper()
            target_qty = float(qty_by_material.get(target, 0.0) or 0.0)
            acc_qty = float(edge["_accumulated_qty"] or 0.0)
            semi_hr_pc = float(hour_per_pc_by_material.get(semi, 0.0) or 0.0)
            contribution = target_qty * acc_qty * semi_hr_pc
            if contribution:
                semi_by_target[target] = semi_by_target.get(target, 0.0) + contribution

    total_by_target: dict[str, float] = {}
    for material in sorted(set(direct_by_material) | set(semi_by_target)):
        total_by_target[material] = float(direct_by_material.get(material, 0.0) or 0.0) + float(semi_by_target.get(material, 0.0) or 0.0)

    zero_targets = [m for m, h in total_by_target.items() if float(h or 0.0) == 0.0]
    return total_by_target, {
        "working_hour_filter_applied": True,
        "working_hour_filter_rule": "Exclude all Raw Material Activity rows when Target Product Total Annual Working Hour, including semi-finished roll-up, equals 0.",
        "working_hour_mapped_targets": int(len(total_by_target)),
        "zero_total_working_hour_targets": int(len(zero_targets)),
    }

def _explode_bom(df: pd.DataFrame) -> tuple[pd.DataFrame, Dict[str, Any]]:
    parent_set = set(df["_parent"].dropna().astype(str))
    component_set = set(df["_component"].dropna().astype(str))
    semi_finished_set = parent_set.intersection(component_set)

    roots = sorted(parent_set - component_set)
    if not roots:
        roots = sorted(parent_set)

    children: dict[str, list[dict[str, Any]]] = defaultdict(list)

    for _, r in df.iterrows():
        row = {
            "parent": r["_parent"],
            "component": r["_component"],
            "qty": r["_qty"],
            "uom": r["_uom"],
            "description": r["_description"],
            "material_group": r["_material_group"],
            "valid_from": r["_valid_from"],
            "net_weight": r.get("_net_weight", ""),
            "gross_weight": r.get("_gross_weight", ""),
            "weight_uom": r.get("_weight_uom", ""),
        }
        children[row["parent"]].append(row)

    output_rows: list[dict[str, Any]] = []
    cycle_count = 0

    for root in roots:
        stack: list[tuple[str, float, int, list[str]]] = [(root, 1.0, 0, [root])]

        while stack:
            current_parent, accumulated_qty, level, path = stack.pop()

            for child in children.get(current_parent, []):
                component = child["component"]
                qty = child["qty"]
                next_qty = accumulated_qty * qty
                next_level = level + 1

                if component in path:
                    cycle_count += 1
                    continue

                if component in semi_finished_set:
                    stack.append((component, next_qty, next_level, path + [component]))
                else:
                    output_rows.append({
                        "target_product": root,
                        "raw_material": component,
                        "usage": next_qty,
                        "unit": child["uom"],
                        "description": child["description"],
                        "material_group": child["material_group"],
                        "valid_from": child["valid_from"],
                        "net_weight": child.get("net_weight", ""),
                        "gross_weight": child.get("gross_weight", ""),
                        "weight_uom": child.get("weight_uom", ""),
                        "level": next_level,
                    })

    exploded = pd.DataFrame(output_rows)
    if exploded.empty:
        exploded = pd.DataFrame(columns=[
            "target_product", "raw_material", "usage", "unit", "description", "material_group", "valid_from", "net_weight", "gross_weight", "weight_uom", "level"
        ])
    else:
        exploded = (
            exploded.groupby(["target_product", "raw_material", "unit"], dropna=False, as_index=False)
            .agg({
                "usage": "sum",
                "description": "first",
                "material_group": "first",
                "valid_from": "first",
                "net_weight": "first",
                "gross_weight": "first",
                "weight_uom": "first",
                "level": "max",
            })
            .sort_values(["target_product", "raw_material"])
            .reset_index(drop=True)
        )

    summary = {
        "products": len(roots),
        "semi_finished": len(semi_finished_set),
        "raw_materials": int(exploded["raw_material"].nunique()) if not exploded.empty else 0,
        "activity_rows": int(len(exploded)),
        "max_level": int(exploded["level"].max()) if not exploded.empty else 0,
        "cycles_skipped": cycle_count,
    }
    return exploded, summary



def _write_raw_material_bulk_from_exploded(
    exploded: pd.DataFrame,
    raw_material_template_path: str | Path,
    output_path: str | Path,
) -> Dict[str, Any]:
    """Write Raw Material Bulk workbook from an already exploded BOM DataFrame.

    This keeps Module 2 template-driven: data is written by header name, not
    fixed Excel column positions. It is reused by the all-site export and the
    Production Site split export.

    V14.9: true zero-usage raw material rows are excluded from Bulk output;
    non-zero decimal usages are preserved.
    """
    exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)

    raw_material_template_path = Path(raw_material_template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(raw_material_template_path, output_path)

    wb = load_workbook(output_path)

    if ACTIVITY_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 raw material bulk 分頁：{ACTIVITY_SHEET_NAME}")
    if RAW_MATERIAL_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 raw material bulk 分頁：{RAW_MATERIAL_SHEET_NAME}")

    activity_ws = wb[ACTIVITY_SHEET_NAME]
    raw_ws = wb[RAW_MATERIAL_SHEET_NAME]

    activity_cols = {
        "raw_name": _find_template_column(activity_ws, RAW_MATERIAL_NAME_ALIASES, 1),
        "raw_code": _find_template_column(activity_ws, RAW_MATERIAL_CODE_ALIASES, 2),
        "start_date": _find_template_column(activity_ws, DOC_START_DATE_ALIASES, 3),
        "end_date": _find_template_column(activity_ws, DOC_END_DATE_ALIASES, 4),
        "document_type": _find_template_column(activity_ws, DOCUMENT_TYPE_ALIASES, 5),
        "document_number": _find_template_column(activity_ws, DOCUMENT_NUMBER_ALIASES, 6),
        "usage": _find_template_column(activity_ws, USAGE_ALIASES, 7),
        "unit": _find_template_column(activity_ws, ACTIVITY_DATA_UNIT_ALIASES, 8),
        "data_source": _find_template_column(activity_ws, DATA_SOURCE_ALIASES, 12),
        "data_source_other": _find_template_column(activity_ws, DATA_SOURCE_OTHER_ALIASES, 13),
        "transport_origin": _find_template_column(activity_ws, TRANSPORT_ORIGIN_ALIASES, 15),
        "transport_destination": _find_template_column(activity_ws, TRANSPORT_DESTINATION_ALIASES, 16),
        "supplier_name": _find_template_column(activity_ws, SUPPLIER_NAME_ALIASES, 14),
        "target_product": _find_template_column(activity_ws, PRODUCT_LINK_ALIASES, 17),
        "comment": _find_template_column(activity_ws, COMMENT_ALIASES, 18),
        "material_group": _find_template_column(activity_ws, MATERIAL_GROUP_ALIASES, 19),
        "net_weight": (_find_template_columns(activity_ws, NET_WEIGHT_ALIASES) or [None])[0],
        "gross_weight": (_find_template_columns(activity_ws, GROSS_WEIGHT_ALIASES) or [None])[0],
        "weight_unit": (_find_template_columns(activity_ws, WEIGHT_UNIT_ALIASES) or [None])[0],
    }
    raw_cols = {
        "raw_name": _find_template_column(raw_ws, RAW_MATERIAL_NAME_ALIASES, 1),
        "raw_code": _find_template_column(raw_ws, RAW_MATERIAL_CODE_ALIASES, 2),
        "description": _find_template_column(raw_ws, RAW_MATERIAL_DESC_ALIASES, 6),
    }

    document_type_value = _document_type_for_template(wb)

    _clear_template_columns(activity_ws, DATA_START_ROW, list(activity_cols.values()))
    _clear_template_columns(raw_ws, DATA_START_ROW, list(raw_cols.values()))

    row_idx = DATA_START_ROW
    for _, r in exploded.iterrows():
        valid_from = r["valid_from"]
        if not isinstance(valid_from, date):
            valid_from = _date_from_value(valid_from)

        raw_material = r["raw_material"]
        target_product = r["target_product"]
        usage_value = float(r["usage"]) if not pd.isna(r["usage"]) else 0

        _write_template_value(activity_ws, row_idx, activity_cols["raw_name"], raw_material)
        _write_template_value(activity_ws, row_idx, activity_cols["raw_code"], raw_material)
        _write_template_value(activity_ws, row_idx, activity_cols["start_date"], _year_start(valid_from))
        _write_template_value(activity_ws, row_idx, activity_cols["end_date"], _year_end(valid_from))
        _write_template_value(activity_ws, row_idx, activity_cols["document_type"], document_type_value)
        _write_template_value(activity_ws, row_idx, activity_cols["document_number"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["usage"], usage_value)
        _write_template_value(activity_ws, row_idx, activity_cols["unit"], r["unit"])
        _write_template_value(activity_ws, row_idx, activity_cols["data_source"], "SAP")
        _write_template_value(activity_ws, row_idx, activity_cols["data_source_other"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["transport_origin"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["transport_destination"], r.get("transport_destination", ""))
        _write_template_value(activity_ws, row_idx, activity_cols["supplier_name"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["target_product"], target_product)
        _write_template_value(activity_ws, row_idx, activity_cols["comment"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["material_group"], r["material_group"])
        _write_template_value(activity_ws, row_idx, activity_cols.get("net_weight"), r.get("net_weight", ""))
        _write_template_value(activity_ws, row_idx, activity_cols.get("gross_weight"), r.get("gross_weight", ""))
        _write_template_value(activity_ws, row_idx, activity_cols.get("weight_unit"), r.get("weight_uom", ""))

        activity_ws.cell(row_idx, activity_cols["start_date"]).number_format = "yyyy/mm/dd"
        activity_ws.cell(row_idx, activity_cols["end_date"]).number_format = "yyyy/mm/dd"
        row_idx += 1

    raw_unique = (
        exploded.sort_values(["raw_material"])
        .drop_duplicates(subset=["raw_material"])
        [["raw_material", "description"]]
        if not exploded.empty else pd.DataFrame(columns=["raw_material", "description"])
    )

    row_idx = DATA_START_ROW
    for _, r in raw_unique.iterrows():
        raw_material = r["raw_material"]
        description = r["description"]

        _write_template_value(raw_ws, row_idx, raw_cols["raw_name"], raw_material)
        _write_template_value(raw_ws, row_idx, raw_cols["raw_code"], raw_material)
        _write_template_value(raw_ws, row_idx, raw_cols["description"], description)
        row_idx += 1

    wb.save(output_path)

    return {
        "output_filename": output_path.name,
        "activity_template_columns": activity_cols,
        "raw_material_template_columns": raw_cols,
        "activity_rows": int(len(exploded)),
        "raw_materials": int(exploded["raw_material"].nunique()) if not exploded.empty else 0,
        "zero_usage_rows_excluded": int(zero_usage_rows_excluded),
    }


def generate_raw_material_bulk_file(
    bom_path: str | Path | list[str | Path] | tuple[str | Path, ...],
    raw_material_template_path: str | Path,
    output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
) -> Dict[str, Any]:
    """Generate one Raw Material Bulk workbook for all target products."""
    output_path = Path(output_path)

    bom_df, used_columns = _read_boms(bom_path, mapping=mapping)
    exploded, summary = _explode_bom(bom_df)
    exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)
    summary["zero_usage_rows_excluded"] = int(zero_usage_rows_excluded)
    summary["activity_rows"] = int(len(exploded))
    summary["raw_materials"] = int(exploded["raw_material"].nunique()) if not exploded.empty else 0

    write_summary = _write_raw_material_bulk_from_exploded(
        exploded=exploded,
        raw_material_template_path=raw_material_template_path,
        output_path=output_path,
    )
    summary.update(write_summary)

    summary["used_columns"] = used_columns
    summary["bom_files"] = int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1
    summary["bom_rows_before_dedup"] = int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_rows_after_dedup"] = int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_duplicate_rows_removed"] = int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0
    return summary


def _sanitize_filename_part(value: Any, fallback: str = "Unassigned") -> str:
    text = str(value or "").strip()
    if not text:
        text = fallback
    text = re.sub(r"[\\/:*?\"<>|]+", "_", text)
    text = re.sub(r"\s+", "_", text)
    return text[:80] or fallback


def _read_step1_product_master_maps(step1_output_path: str | Path) -> tuple[dict[str, str], Dict[str, Any]]:
    """Read Step1 output and return product master-data maps.

    Module 2 uses Step1 output only for product-level master data:
    - Material Number / Target Product -> Production Site for split export.

    Material Group is a raw-material attribute and must remain sourced from
    the Standard BOM Material group column. Step1 output must not overwrite it.
    """
    step1_output_path = Path(step1_output_path)
    try:
        df = pd.read_excel(step1_output_path, sheet_name="Plant_Material年度產量", dtype=object)
    except Exception:
        df = pd.read_excel(step1_output_path, sheet_name=0, dtype=object)

    material_col = _find_step1_column(df, ["Material Number", "Material", "Product Material Number"])
    site_col = _find_step1_column(df, ["Production Site", "production site", "生產廠區", "廠區", "廠別"])

    work = df.copy()
    work["_material_key"] = work[material_col].apply(_normalize_material_key)
    work["_production_site"] = work[site_col].apply(_safe_text)
    work = work[work["_material_key"] != ""].copy()

    site_map: dict[str, str] = {}
    duplicate_site_conflicts: list[str] = []

    for material, group in work.groupby("_material_key", dropna=False):
        sites = sorted({str(x).strip() for x in group["_production_site"] if str(x).strip()})

        if sites:
            site_map[str(material)] = sites[0]
            if len(sites) > 1:
                duplicate_site_conflicts.append(f"{material}: {', '.join(sites)}")

    return site_map, {
        "step1_rows": int(len(df)),
        "step1_mapped_materials": int(len(site_map)),
        "step1_site_conflicts": duplicate_site_conflicts,
        "material_group_source": "Standard BOM",
        "transportation_destination_source": "Step1 Production Site",
    }


def generate_raw_material_bulk_files_by_site_zip(
    bom_path: str | Path | list[str | Path] | tuple[str | Path, ...],
    raw_material_template_path: str | Path,
    output_dir: str | Path,
    token: str,
    step1_output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
) -> Dict[str, Any]:
    """Generate one Raw Material Bulk workbook per Production Site and ZIP them.

    Split source:
    - BOM explosion gives Target Product -> Raw Material usage.
    - Step1 output gives Material Number / Target Product -> Production Site.
    - Rows whose target product is not found in Step1 are exported under
      "Unassigned" so they are visible instead of silently dropped.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    bom_df, used_columns = _read_boms(bom_path, mapping=mapping)
    exploded, base_summary = _explode_bom(bom_df)
    exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)

    total_hour_by_target, working_hour_summary = _calculate_total_working_hour_by_target(
        step1_output_path=step1_output_path,
        bom_df=bom_df,
    )
    exploded, zero_total_working_hour_rows_excluded = _exclude_zero_total_working_hour_target_rows(
        exploded=exploded,
        total_hour_by_material=total_hour_by_target,
    )

    base_summary["zero_usage_rows_excluded"] = int(zero_usage_rows_excluded)
    base_summary["zero_total_working_hour_rows_excluded"] = int(zero_total_working_hour_rows_excluded)
    base_summary["activity_rows"] = int(len(exploded))
    base_summary["raw_materials"] = int(exploded["raw_material"].nunique()) if not exploded.empty else 0
    base_summary.update(working_hour_summary)

    site_map, step1_summary = _read_step1_product_master_maps(step1_output_path)

    work = exploded.copy()
    if work.empty:
        work["_production_site"] = ""
    else:
        work["_target_key"] = work["target_product"].apply(_normalize_material_key)
        work["_production_site"] = work["_target_key"].map(site_map).fillna("Unassigned")
        work["_production_site"] = work["_production_site"].apply(lambda x: str(x or "").strip() or "Unassigned")
        # V14.6: Transportation Destination in Raw Material Bulk follows Step1 Production Site.
        # Step1 Output is the product master-data source; Standard BOM only provides material structure and usage.
        work["transport_destination"] = work["_production_site"]

        # V14.7: Material Group is a raw-material attribute from Standard BOM.
        # Do not overwrite it with Step1 product classification fields.
        if "material_group" not in work.columns:
            work["material_group"] = ""

    site_values = sorted({str(x).strip() or "Unassigned" for x in work["_production_site"].tolist()}) if not work.empty else ["Unassigned"]

    generated_files: list[dict[str, Any]] = []
    zip_filename = f"raw_material_activity_data_bulk_by_site_{token}.zip"
    zip_path = output_dir / zip_filename

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for site in site_values:
            site_df = work[work["_production_site"] == site].copy()
            site_df = site_df.drop(columns=["_target_key", "_production_site"], errors="ignore")
            safe_site = _sanitize_filename_part(site)
            file_path = output_dir / f"raw_material_activity_data_bulk_{safe_site}_{token}.xlsx"

            write_summary = _write_raw_material_bulk_from_exploded(
                exploded=site_df,
                raw_material_template_path=raw_material_template_path,
                output_path=file_path,
            )
            zf.write(file_path, arcname=file_path.name)
            generated_files.append({
                "production_site": site,
                "filename": file_path.name,
                "activity_rows": int(write_summary.get("activity_rows", 0)),
                "raw_materials": int(write_summary.get("raw_materials", 0)),
            })

    unassigned_rows = int((work["_production_site"] == "Unassigned").sum()) if not work.empty else 0

    summary = dict(base_summary)
    summary.update({
        "output_filename": zip_filename,
        "download_url": f"/download/{zip_filename}",
        "split_by_production_site": True,
        "production_site_files": generated_files,
        "production_site_count": int(len(site_values)),
        "unassigned_rows": unassigned_rows,
        "used_columns": used_columns,
        "bom_files": int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1,
        "bom_rows_before_dedup": int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0,
        "bom_rows_after_dedup": int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0,
        "bom_duplicate_rows_removed": int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0,
    })
    summary.update(step1_summary)
    return summary


def _explode_bom_structure(df: pd.DataFrame) -> tuple[pd.DataFrame, Dict[str, Any]]:
    """Create normalized multi-level BOM structure for Step 2 working-hour roll-up."""
    parent_set = set(df["_parent"].dropna().astype(str))
    component_set = set(df["_component"].dropna().astype(str))
    semi_finished_set = parent_set.intersection(component_set)
    roots = sorted(parent_set - component_set) or sorted(parent_set)
    children: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for _, r in df.iterrows():
        children[r["_parent"]].append({
            "parent": r["_parent"], "component": r["_component"], "qty": r["_qty"],
            "uom": r["_uom"], "description": r["_description"],
            "material_group": r["_material_group"], "valid_from": r["_valid_from"],
            "net_weight": r.get("_net_weight", ""), "gross_weight": r.get("_gross_weight", ""),
            "weight_uom": r.get("_weight_uom", ""),
        })
    rows: list[dict[str, Any]] = []
    cycle_count = 0
    for root in roots:
        stack: list[tuple[str, float, int, list[str]]] = [(root, 1.0, 0, [root])]
        while stack:
            current_parent, accumulated_qty, level, path = stack.pop()
            for child in children.get(current_parent, []):
                component = child["component"]
                next_qty = accumulated_qty * child["qty"]
                next_level = level + 1
                is_semi = component in semi_finished_set
                if component in path:
                    cycle_count += 1
                    continue
                rows.append({
                    "Target Product": root,
                    "Parent Material": current_parent,
                    "Component": component,
                    "Quantity Per Parent": child["qty"],
                    "Accumulated Quantity": next_qty,
                    "Unit": child["uom"],
                    "Component Description": child["description"],
                    "Material Group": child["material_group"],
                    "Valid From": child["valid_from"],
                    "Net weight": child.get("net_weight", ""),
                    "Gross weight": child.get("gross_weight", ""),
                    "Weight UoM": child.get("weight_uom", ""),
                    "Level": next_level,
                    "Is Semi-finished": "Y" if is_semi else "N",
                })
                if is_semi:
                    stack.append((component, next_qty, next_level, path + [component]))
    structure = pd.DataFrame(rows)
    if structure.empty:
        structure = pd.DataFrame(columns=["Target Product", "Parent Material", "Component", "Quantity Per Parent", "Accumulated Quantity", "Unit", "Component Description", "Material Group", "Valid From", "Net weight", "Gross weight", "Weight UoM", "Level", "Is Semi-finished"])
    summary = {"products": len(roots), "semi_finished": len(semi_finished_set), "structure_rows": int(len(structure)), "max_level": int(structure["Level"].max()) if not structure.empty else 0, "cycles_skipped": cycle_count}
    return structure, summary


def export_bom_structure_file(
    bom_path: str | Path | list[str | Path] | tuple[str | Path, ...],
    output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
) -> Dict[str, Any]:
    """Export latest normalized BOM structure for Step 2 semi-finished working-hour roll-up."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    bom_df, used_columns = _read_boms(bom_path, mapping=mapping)
    structure, summary = _explode_bom_structure(bom_df)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        structure.to_excel(writer, index=False, sheet_name="BOM Structure")
        ws = writer.book["BOM Structure"]
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = 12
            letter = col[0].column_letter
            for cell in col[:1000]:
                max_len = max(max_len, len(str(cell.value or "")) + 2)
            ws.column_dimensions[letter].width = min(max_len, 45)
    summary["output_filename"] = output_path.name
    summary["used_columns"] = used_columns
    summary["bom_files"] = int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1
    summary["bom_rows_before_dedup"] = int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_rows_after_dedup"] = int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_duplicate_rows_removed"] = int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0
    return summary


# =========================================================
# V12 Working Hour Roll-up
# Step1 Output + BOM Structure -> working_hour_rollup.xlsx
# =========================================================

STEP1_SOURCE_SHEET_NAME = "Plant_Material年度產量"


def _find_step1_column(df: pd.DataFrame, candidates: list[str]) -> str:
    normalized = {_normalize_col(c).lower(): c for c in df.columns}
    for name in candidates:
        key = _normalize_col(name).lower()
        if key in normalized:
            return normalized[key]
    raise ValueError(f"找不到 Step1 欄位：{', '.join(candidates)}")


def _find_step1_optional_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    try:
        return _find_step1_column(df, candidates)
    except ValueError:
        return None


def _normalize_material_key(value: Any) -> str:
    return str(value or "").strip().upper()


def _read_bom_structure_file(path: str | Path) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="BOM Structure", dtype=object)
    except ValueError:
        return pd.read_excel(path, sheet_name=0, dtype=object)


def generate_working_hour_rollup_file(
    step1_output_path: str | Path,
    bom_structure_path: str | Path,
    output_path: str | Path,
) -> Dict[str, Any]:
    """Generate auditable working-hour roll-up workbook.

    Summary.Total Annual Working Hour is the source used by Step2 when
    Include Semi-finished Working Hour is selected.
    """
    step1_output_path = Path(step1_output_path)
    bom_structure_path = Path(bom_structure_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    step1_df = pd.read_excel(step1_output_path, sheet_name=STEP1_SOURCE_SHEET_NAME, dtype=object)
    bom_df = _read_bom_structure_file(bom_structure_path)
    if bom_df.empty:
        raise ValueError("BOM Structure is empty. Please complete Module 2 → BOM Expansion first.")

    material_col = _find_step1_column(step1_df, ["Material Number"])
    qty_col = _find_step1_column(step1_df, ["年度生產量", "Annual Quantity", "Delivered quantity"])
    hour_col = _find_step1_optional_column(step1_df, ["年度總工時", "Total working hours", "Selected Hours", "Total Hours", "Working Hours"])
    if not hour_col:
        raise ValueError("Step1 Output 找不到年度總工時欄位，無法產生 working_hour_rollup.xlsx")

    plant_col = _find_step1_optional_column(step1_df, ["Plant"])
    site_col = _find_step1_optional_column(step1_df, ["Production Site", "production site", "生產廠區", "廠區", "廠別"])
    type_col = _find_step1_optional_column(step1_df, ["產品類型", "Product Type"])
    wip_col = _find_step1_optional_column(step1_df, ["Is_WIP", "Is WIP", "WIP"])

    work = step1_df.copy()
    work["_material_key"] = work[material_col].apply(_normalize_material_key)
    work["_annual_qty"] = work[qty_col].apply(_safe_number)
    work["_direct_hour"] = work[hour_col].apply(_safe_number)
    work["_plant"] = work[plant_col].apply(_safe_text) if plant_col else ""
    work["_production_site"] = work[site_col].apply(_safe_text) if site_col else ""
    work["_product_type"] = work[type_col].apply(_safe_text) if type_col else ""
    work["_is_wip"] = work[wip_col].apply(_safe_text) if wip_col else ""

    summary_base = work.groupby(
        ["_material_key", "_plant", "_production_site", "_product_type", "_is_wip"],
        dropna=False,
        as_index=False,
    ).agg({"_annual_qty": "sum", "_direct_hour": "sum"})

    parent_values = set(bom_df["Parent Material"].dropna().astype(str).str.strip().str.upper()) if "Parent Material" in bom_df.columns else set()
    component_values = set(bom_df["Component"].dropna().astype(str).str.strip().str.upper()) if "Component" in bom_df.columns else set()
    semi_materials = parent_values.intersection(component_values)

    material_totals = work.groupby(["_material_key"], dropna=False, as_index=False).agg({"_annual_qty": "sum", "_direct_hour": "sum"})
    qty_by_material = {}
    direct_by_material = {}
    hour_per_pc_by_material = {}
    for _, r in material_totals.iterrows():
        material = str(r["_material_key"] or "").strip()
        qty = float(r["_annual_qty"] or 0)
        hours = float(r["_direct_hour"] or 0)
        qty_by_material[material] = qty
        direct_by_material[material] = hours
        hour_per_pc_by_material[material] = hours / qty if qty else 0.0

    semi_hour_per_pc_rows = [{
        "Semi Material": m,
        "Semi Annual Qty": qty_by_material.get(m, 0.0),
        "Semi Direct Annual Working Hour": direct_by_material.get(m, 0.0),
        "Semi Direct Hour per PC": hour_per_pc_by_material.get(m, 0.0),
    } for m in sorted(semi_materials)]
    semi_hour_per_pc_df = pd.DataFrame(semi_hour_per_pc_rows)

    target_col = _find_step1_optional_column(bom_df, ["Target Product", "target_product"])
    parent_col = _find_step1_optional_column(bom_df, ["Parent Material", "parent_material"])
    comp_col = _find_step1_optional_column(bom_df, ["Component", "component"])
    acc_col = _find_step1_optional_column(bom_df, ["Accumulated Quantity", "usage", "Quantity"])
    level_col = _find_step1_optional_column(bom_df, ["Level", "level"])
    semi_flag_col = _find_step1_optional_column(bom_df, ["Is Semi-finished", "Is Semi", "semi_finished"])
    if not target_col or not comp_col or not acc_col:
        raise ValueError("BOM Structure 缺少 Target Product、Component 或 Accumulated Quantity 欄位")

    b = bom_df.copy()
    b["_target_key"] = b[target_col].apply(_normalize_material_key)
    b["_component_key"] = b[comp_col].apply(_normalize_material_key)
    b["_accumulated_qty"] = b[acc_col].apply(_safe_number)
    if semi_flag_col:
        b = b[b[semi_flag_col].astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1"])].copy()
    else:
        b = b[b["_component_key"].isin(semi_materials)].copy()

    detail_rows = []
    semi_by_key: dict[tuple[str, str, str], float] = {}
    for _, target_row in summary_base.iterrows():
        target = str(target_row["_material_key"] or "").strip()
        plant = _safe_text(target_row["_plant"])
        site = _safe_text(target_row["_production_site"])
        target_qty = float(target_row["_annual_qty"] or 0)
        for _, edge in b[b["_target_key"] == target].iterrows():
            semi = str(edge["_component_key"] or "").strip()
            acc_qty = float(edge["_accumulated_qty"] or 0)
            semi_hr_pc = float(hour_per_pc_by_material.get(semi, 0.0) or 0.0)
            contrib_pc = acc_qty * semi_hr_pc
            contrib_annual = target_qty * contrib_pc
            if contrib_annual:
                key = (target, plant, site)
                semi_by_key[key] = semi_by_key.get(key, 0.0) + contrib_annual
            detail_rows.append({
                "Target Product": target,
                "Plant": plant,
                "Production Site": site,
                "Target Annual Qty": target_qty,
                "Parent Material": edge.get(parent_col, "") if parent_col else "",
                "Semi Material": semi,
                "BOM Accumulated Qty": acc_qty,
                "Semi Direct Hour per PC": semi_hr_pc,
                "Semi Hour Contribution per PC": contrib_pc,
                "Semi Annual Working Hour Contribution": contrib_annual,
                "Level": edge.get(level_col, "") if level_col else "",
            })

    detail_df = pd.DataFrame(detail_rows)
    if detail_df.empty:
        detail_df = pd.DataFrame(columns=["Target Product", "Plant", "Production Site", "Target Annual Qty", "Parent Material", "Semi Material", "BOM Accumulated Qty", "Semi Direct Hour per PC", "Semi Hour Contribution per PC", "Semi Annual Working Hour Contribution", "Level"])

    summary_rows = []
    for _, r in summary_base.iterrows():
        material = str(r["_material_key"] or "").strip()
        plant = _safe_text(r["_plant"])
        site = _safe_text(r["_production_site"])
        qty = float(r["_annual_qty"] or 0)
        direct = float(r["_direct_hour"] or 0)
        semi = float(semi_by_key.get((material, plant, site), 0.0) or 0.0)
        total = direct + semi
        summary_rows.append({
            "Material Number": material,
            "Plant": plant,
            "Production Site": site,
            "Product Type": _safe_text(r["_product_type"]),
            "Is_WIP": _safe_text(r["_is_wip"]),
            "Annual Qty": qty,
            "Direct Annual Working Hour": direct,
            "Semi Annual Working Hour": semi,
            "Total Annual Working Hour": total,
            "Direct Hour per PC": direct / qty if qty else 0.0,
            "Semi Hour per PC": semi / qty if qty else 0.0,
            "Total Hour per PC": total / qty if qty else 0.0,
        })
    summary_df = pd.DataFrame(summary_rows).sort_values(["Plant", "Production Site", "Material Number"]).reset_index(drop=True)

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        detail_df.to_excel(writer, index=False, sheet_name="Roll-up Detail")
        semi_hour_per_pc_df.to_excel(writer, index=False, sheet_name="Semi Hour per PC")
        bom_df.to_excel(writer, index=False, sheet_name="BOM Structure")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            for col in sheet.columns:
                max_len = 12
                letter = col[0].column_letter
                for cell in col[:1000]:
                    max_len = max(max_len, len(str(cell.value or "")) + 2)
                sheet.column_dimensions[letter].width = min(max_len, 48)

    return {
        "output_filename": output_path.name,
        "summary_rows": int(len(summary_df)),
        "detail_rows": int(len(detail_df)),
        "semi_materials": int(len(semi_hour_per_pc_df)),
        "total_direct_hours": float(summary_df["Direct Annual Working Hour"].sum()) if not summary_df.empty else 0.0,
        "total_semi_hours": float(summary_df["Semi Annual Working Hour"].sum()) if not summary_df.empty else 0.0,
        "total_hours": float(summary_df["Total Annual Working Hour"].sum()) if not summary_df.empty else 0.0,
    }

# =========================================================
# Module 2 Supplier Master + Supplier Bulk Export
# Clean implementation based on official baseline.
# =========================================================

SUPPLIER_BULK_SHEET_NAME = "Input Sheet"
SUPPLIER_BULK_TEMPLATE_FILENAME = "supplier_bulk_create_template_v1.xlsx"

SUPPLIER_MATERIAL_ALIASES = [
    "Raw Material Code", "Raw Material Number", "Material", "Material Number", "Component", "Component Number",
    "物料", "料號", "原物料代碼", "原料代碼", "元件料號",
]
SUPPLIER_VENDOR_ALIASES = [
    "Vendor", "Vender", "Vendor Code", "Vendor Number", "Supplier Vendor", "供應商代碼", "供應商編號", "廠商代碼",
]
SUPPLIER_ADDRESS_ALIASES = [
    "Supplier Address", "Supplier Address 1", "Supplier Address1", "Supplier Address Line1",
    "Supplier Address (English)", "Supplier Address (Local)", "Supplier Addr", "Supplier_Address",
    "供應商地址", "廠商地址",
]
SUPPLIER_BULK_NAME_ALIASES = ["Supplier Name", "Supplier Name (optional)", "Supplier Name(optional)", "供應商名稱"]
SUPPLIER_BULK_CODE_ALIASES = ["Supplier Code", "Supplier Code (optional)", "Vendor", "Vendor Code", "供應商代碼"]
SUPPLIER_BULK_COUNTRY_ALIASES = ["Country/Area", "Country / Area", "Country", "Country Area", "國家/地區", "國家"]
SUPPLIER_BULK_ADDRESS_ALIASES = ["Supplier Address", "Supplier Address (optional)", "Supplier Address1", "供應商地址"]
SUPPLIER_BULK_UNIT_ALIASES = ["Unit Name", "Unit", "Transportation Destination", "Production Site", "單位名稱", "廠區"]

# Broaden raw material template header aliases while keeping old behavior.
TRANSPORT_ORIGIN_ALIASES = list(dict.fromkeys(TRANSPORT_ORIGIN_ALIASES + [
    "Transportation Origin (optional)", "Transport Origin", "Origin", "運輸起點(optional)",
]))
TRANSPORT_DESTINATION_ALIASES = list(dict.fromkeys(TRANSPORT_DESTINATION_ALIASES + [
    "Transportation Destination (optional)", "Transport Destination", "Destination", "運輸終點(optional)",
]))
SUPPLIER_NAME_ALIASES = list(dict.fromkeys(SUPPLIER_NAME_ALIASES + [
    "Supplier Name(optional)", "Supplier Name Optional", "Supplier Name（optional）",
]))


def _normalize_vendor_code(value: Any) -> str:
    text = _safe_text(value).upper()
    if text.endswith(".0"):
        text = text[:-2]
    text = re.sub(r"\s+", "", text)
    return text


def _supplier_header_key(value: Any) -> str:
    return re.sub(r"[^0-9A-Z]+", "", str(value or "").upper())


def _find_any_dataframe_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized = {_normalize_template_header(c): c for c in df.columns}
    for alias in aliases:
        key = _normalize_template_header(alias)
        if key in normalized:
            return normalized[key]
    return None


def _find_supplier_col_by_rule(df: pd.DataFrame, kind: str) -> str | None:
    keyed = [(_supplier_header_key(c), c) for c in df.columns]
    if kind == "material":
        for key, col in keyed:
            if "MATERIAL" in key and not any(x in key for x in ["GROUP", "DESC", "DESCRIPTION"]):
                return col
        return _find_any_dataframe_column(df, SUPPLIER_MATERIAL_ALIASES)
    if kind == "vendor":
        for key, col in keyed:
            if key.endswith("VENDOR") or key.endswith("VENDER") or "VENDORCODE" in key or "VENDORNUMBER" in key:
                return col
        return _find_any_dataframe_column(df, SUPPLIER_VENDOR_ALIASES)
    if kind == "vendor_name":
        for key, col in keyed:
            if "VENDORNAME" in key or "SUPPLIERNAME" in key or "SEARCHTERM" in key:
                return col
        return None
    if kind == "country":
        for key, col in keyed:
            if key.endswith("COUNTRY") or "COUNTRYAREA" in key or key == "COUNTRYREGION":
                return col
        return None
    if kind == "city":
        for key, col in keyed:
            if key.endswith("CITY") or key == "CITY":
                return col
        return None
    if kind == "street":
        for key, col in keyed:
            if "STREET" in key or "ADDRESSLINE" in key:
                return col
        return None
    if kind == "incoterms2":
        for key, col in keyed:
            if "INCOTERMS" in key and ("PART2" in key or key.endswith("2")):
                return col
        return None
    return None


def _find_supplier_address_column(df: pd.DataFrame) -> str | None:
    exact = _find_any_dataframe_column(df, SUPPLIER_ADDRESS_ALIASES)
    if exact:
        return exact
    for col in df.columns:
        key = _normalize_template_header(col)
        if "supplier" in key and "address" in key:
            return col
    return None


def _build_supplier_address(row: pd.Series, address_col: str | None, country_col: str | None, city_col: str | None, street_col: str | None, incoterms2_col: str | None) -> str:
    address = _safe_text(row.get(address_col)) if address_col else ""
    if address:
        return address
    parts: list[str] = []
    for col in [country_col, city_col, street_col]:
        text = _safe_text(row.get(col)) if col else ""
        if text and text not in parts:
            parts.append(text)
    if parts:
        return " ".join(parts)
    return _safe_text(row.get(incoterms2_col)) if incoterms2_col else ""


def _supplier_record_score(record: dict[str, str]) -> int:
    score = 0
    if record.get("country_area"):
        score += 1000
    if record.get("supplier_address"):
        score += min(len(record.get("supplier_address") or ""), 500)
    if record.get("supplier_master_name"):
        score += 50
    return score


def _read_supplier_files(supplier_paths: list[str | Path] | tuple[str | Path, ...] | None) -> tuple[dict[str, list[dict[str, str]]], Dict[str, Any]]:
    """Read one or many supplier masters and normalize to Material -> suppliers.

    A/B supplier formats are supported. If the same Material+Vendor appears in
    multiple uploaded files, the record with richer address information wins.
    """
    if not supplier_paths:
        return {}, {"supplier_files": 0, "supplier_rows": 0, "supplier_mapped_materials": 0, "supplier_mapped_suppliers": 0, "supplier_skipped_files": []}

    by_material_vendor: dict[tuple[str, str], dict[str, str]] = {}
    total_rows = 0
    skipped_files: list[str] = []

    for path in supplier_paths:
        path = Path(path)
        try:
            df = pd.read_excel(path, sheet_name=0, dtype=object)
        except Exception as exc:
            skipped_files.append(f"{path.name}: {exc}")
            continue

        material_col = _find_supplier_col_by_rule(df, "material")
        vendor_col = _find_supplier_col_by_rule(df, "vendor")
        vendor_name_col = _find_supplier_col_by_rule(df, "vendor_name")
        address_col = _find_supplier_address_column(df)
        country_col = _find_supplier_col_by_rule(df, "country")
        city_col = _find_supplier_col_by_rule(df, "city")
        street_col = _find_supplier_col_by_rule(df, "street")
        incoterms2_col = _find_supplier_col_by_rule(df, "incoterms2")

        if not material_col or not vendor_col:
            skipped_files.append(f"{path.name}: missing material/vendor column")
            continue

        total_rows += int(len(df))
        for _, row in df.iterrows():
            material_key = _normalize_material_key(row.get(material_col))
            vendor_code = _normalize_vendor_code(row.get(vendor_col))
            if not material_key or not vendor_code:
                continue
            vendor_name = _safe_text(row.get(vendor_name_col)) if vendor_name_col else ""
            country = _safe_text(row.get(country_col)) if country_col else ""
            supplier_address = _build_supplier_address(row, address_col, country_col, city_col, street_col, incoterms2_col)
            candidate = {
                "vendor_code": vendor_code,
                "supplier_code": vendor_code,
                "supplier_master_name": vendor_name,
                "country_area": country,
                "supplier_address": supplier_address,
                "transport_origin": supplier_address,
                "source_file": path.name,
            }
            key = (material_key, vendor_code)
            current = by_material_vendor.get(key)
            if current is None or _supplier_record_score(candidate) > _supplier_record_score(current):
                by_material_vendor[key] = candidate

    records: dict[str, list[dict[str, str]]] = {}
    for (material_key, _vendor_code), record in by_material_vendor.items():
        records.setdefault(material_key, []).append(record)
    for material_key in records:
        records[material_key].sort(key=lambda r: r.get("vendor_code", ""))

    return records, {
        "supplier_files": int(len(supplier_paths)),
        "supplier_rows": int(total_rows),
        "supplier_mapped_materials": int(len(records)),
        "supplier_mapped_suppliers": int(sum(len(v) for v in records.values())),
        "supplier_skipped_files": skipped_files,
    }



def _extract_site_tbc_supplier_map_from_raw_template(wb) -> dict[str, dict[str, str]]:
    """Read site-specific TBC supplier rows from raw material template Dropdown Values.

    Expected Supplier Name (optional) pattern:
        <Transportation Destination>_TBC - TBC
    Example:
        常州廠(A9)-IPS_TBC - TBC

    The Transportation Origin is read from the same row in Dropdown Values.
    """
    if "Dropdown Values" not in wb.sheetnames:
        return {}

    ws = wb["Dropdown Values"]
    supplier_cols = _find_template_columns(
        ws,
        ["Supplier Name (optional)", "Supplier Name(optional)", "Supplier Name", "supplier_name", "供應商名稱"],
        header_rows=10,
    )
    origin_cols = _find_template_columns(
        ws,
        ["Transportation Origin", "Transportation Origin (optional)", "Transport Origin", "transportation_origin", "運輸起點"],
        header_rows=10,
    )
    if not supplier_cols:
        return {}

    supplier_col = int(supplier_cols[0])
    origin_col = int(origin_cols[0]) if origin_cols else None
    tbc_map: dict[str, dict[str, str]] = {}

    for row_idx in range(2, ws.max_row + 1):
        supplier_name = _safe_text(ws.cell(row_idx, supplier_col).value)
        if not supplier_name or "TBC" not in supplier_name.upper():
            continue

        # Preferred exact expandable pattern: <site>_TBC - TBC
        m = re.match(r"^(.*?)\s*_\s*TBC\s*-\s*TBC\s*$", supplier_name, flags=re.IGNORECASE)
        if not m:
            continue
        site_name = _safe_text(m.group(1))
        if not site_name:
            continue

        origin = _safe_text(ws.cell(row_idx, origin_col).value) if origin_col else ""
        item = {
            "supplier_name": supplier_name,
            "transport_origin": origin,
            "supplier_code": "TBC",
            "supplier_master_name": _supplier_name_from_option(supplier_name),
            "supplier_country_area": "",
            "supplier_address": origin,
        }
        tbc_map[_normalize_template_header(site_name)] = item
        tbc_map[re.sub(r"\s+", "", site_name).upper()] = item

    return tbc_map


def _select_tbc_supplier_for_destination(
    tbc_supplier_map: dict[str, dict[str, str]],
    transportation_destination: Any,
) -> dict[str, str] | None:
    destination = _safe_text(transportation_destination)
    if not destination:
        return None
    keys = [
        _normalize_template_header(destination),
        re.sub(r"\s+", "", destination).upper(),
    ]
    for key in keys:
        if key and key in tbc_supplier_map:
            return tbc_supplier_map[key]
    return None


def _extract_supplier_name_options_from_raw_template(wb) -> list[str]:
    """Read all supplier-name candidates from raw_materials_&_activity_data_template_v1.

    Source of truth: sheet "Dropdown Values", column "Supplier Name (optional)".
    """
    values: list[str] = []
    seen: set[str] = set()

    def add(value: Any) -> None:
        text = _safe_text(value)
        if text and text not in seen:
            seen.add(text)
            values.append(text)

    if "Dropdown Values" not in wb.sheetnames:
        return values
    ws = wb["Dropdown Values"]
    cols = _find_template_columns(
        ws,
        ["Supplier Name (optional)", "Supplier Name(optional)", "Supplier Name", "supplier_name", "供應商名稱"],
        header_rows=10,
    )
    if not cols:
        return values
    col_idx = int(cols[0])
    for row_idx in range(2, ws.max_row + 1):
        add(ws.cell(row_idx, col_idx).value)
    return values


def _supplier_name_from_option(option: Any) -> str:
    text = _safe_text(option)
    if "_" in text:
        tail = text.split("_", 1)[1].strip()
        tail = re.sub(r"\s+-\s+.*$", "", tail).strip()
        return tail or text
    return text


def _select_supplier_name_option(options: list[str], destination: Any, vendor_code: Any) -> str:
    dest = _safe_text(destination)
    dest_compact = re.sub(r"\s+", "", dest).upper()
    vendor = _normalize_vendor_code(vendor_code)
    if not vendor:
        return ""
    candidates: list[str] = []
    for option in options or []:
        text = _safe_text(option)
        compact = re.sub(r"\s+", "", text).upper()
        if vendor not in compact:
            continue
        option_dest = text.split("_", 1)[0].strip() if "_" in text else ""
        option_dest_compact = re.sub(r"\s+", "", option_dest).upper()
        has_dest = bool(dest_compact) and (
            dest_compact == option_dest_compact or dest_compact in compact or (option_dest_compact and option_dest_compact in dest_compact)
        )
        if has_dest:
            candidates.append(text)
    if not candidates:
        return ""
    suffix_matches = [c for c in candidates if re.search(r"-\s*[^-]*" + re.escape(vendor) + r"\s*$", c, flags=re.I)]
    return suffix_matches[0] if suffix_matches else candidates[0]


def _apply_supplier_mapping_to_exploded(
    exploded: pd.DataFrame,
    supplier_map: dict[str, list[dict[str, str]]],
    supplier_options: list[str],
    tbc_supplier_map: dict[str, dict[str, str]] | None = None,
) -> tuple[pd.DataFrame, Dict[str, Any]]:
    work = exploded.copy()
    for col in ["supplier_name", "transport_origin", "supplier_code", "supplier_master_name", "supplier_country_area", "supplier_address"]:
        if col not in work.columns:
            work[col] = ""

    matched_source_rows = 0
    expanded_rows = 0
    supplier_name_matched = 0
    supplier_name_missing = 0
    tbc_fallback_rows = 0
    output_rows: list[dict[str, Any]] = []

    if work.empty:
        return work, {
            "supplier_matched_rows": 0,
            "supplier_expanded_rows": 0,
            "supplier_name_matched_rows": 0,
            "supplier_name_missing_rows": 0,
            "supplier_dropdown_matched_rows": 0,
            "supplier_dropdown_missing_rows": 0,
            "tbc_fallback_rows": 0,
        }

    tbc_supplier_map = tbc_supplier_map or {}

    for _, row in work.iterrows():
        original = row.to_dict()
        raw_key = _normalize_material_key(row.get("raw_material"))
        destination = _safe_text(row.get("transport_destination"))
        suppliers = supplier_map.get(raw_key) or []
        if not suppliers:
            # No material/vendor match: use the same-site TBC supplier from the raw material template.
            tbc_supplier = _select_tbc_supplier_for_destination(tbc_supplier_map, destination)
            if tbc_supplier:
                fallback_row = dict(original)
                fallback_row["transport_destination"] = destination
                fallback_row["supplier_name"] = tbc_supplier.get("supplier_name", "")
                fallback_row["transport_origin"] = tbc_supplier.get("transport_origin", "")
                fallback_row["supplier_code"] = tbc_supplier.get("supplier_code", "TBC")
                fallback_row["supplier_master_name"] = tbc_supplier.get("supplier_master_name", "")
                fallback_row["supplier_country_area"] = tbc_supplier.get("supplier_country_area", "")
                fallback_row["supplier_address"] = tbc_supplier.get("supplier_address", tbc_supplier.get("transport_origin", ""))
                supplier_name_matched += 1 if fallback_row.get("supplier_name") else 0
                supplier_name_missing += 0 if fallback_row.get("supplier_name") else 1
                tbc_fallback_rows += 1
                output_rows.append(fallback_row)
            else:
                output_rows.append(original)
            continue
        matched_source_rows += 1
        for info in suppliers:
            new_row = dict(original)
            # Supplier logic only reads destination. It never clears or overwrites it.
            new_row["transport_destination"] = destination
            supplier_address = info.get("supplier_address", "") or info.get("transport_origin", "")
            supplier_code = info.get("supplier_code", "") or info.get("vendor_code", "")
            supplier_name = info.get("supplier_master_name", "").strip()
            if not supplier_name:
                supplier_name = _select_supplier_name_option(supplier_options, destination, supplier_code)
            new_row["transport_origin"] = supplier_address
            new_row["supplier_code"] = supplier_code
            new_row["supplier_master_name"] = info.get("supplier_master_name", "") or _supplier_name_from_option(supplier_name)
            new_row["supplier_country_area"] = info.get("country_area", "")
            new_row["supplier_address"] = supplier_address
            new_row["supplier_name"] = supplier_name
            if supplier_name:
                supplier_name_matched += 1
            else:
                supplier_name_missing += 1
            output_rows.append(new_row)
            expanded_rows += 1

    ordered_columns = list(dict.fromkeys(list(work.columns) + ["supplier_code", "supplier_master_name", "supplier_country_area", "supplier_address"]))
    result = pd.DataFrame(output_rows) if output_rows else work
    for col in ordered_columns:
        if col not in result.columns:
            result[col] = ""
    return result[ordered_columns], {
        "supplier_matched_rows": int(matched_source_rows),
        "supplier_expanded_rows": int(expanded_rows),
        "supplier_name_matched_rows": int(supplier_name_matched),
        "supplier_name_missing_rows": int(supplier_name_missing),
        # Backward-compatible summary keys for the current UI text.
        "supplier_dropdown_matched_rows": int(supplier_name_matched),
        "supplier_dropdown_missing_rows": int(supplier_name_missing),
        "tbc_fallback_rows": int(tbc_fallback_rows),
    }


def _first_text(row: pd.Series, names: list[str]) -> str:
    for name in names:
        if name in row.index:
            text = _safe_text(row.get(name))
            if text:
                return text
    return ""


def _write_supplier_bulk_create_file(expanded_with_suppliers: pd.DataFrame, supplier_bulk_template_path: str | Path, output_path: str | Path) -> Dict[str, Any]:
    supplier_bulk_template_path = Path(supplier_bulk_template_path)
    output_path = Path(output_path)
    if expanded_with_suppliers is None or expanded_with_suppliers.empty:
        return {"supplier_bulk_rows": 0, "supplier_bulk_filename": "", "supplier_bulk_download_url": ""}
    if not supplier_bulk_template_path.exists():
        return {"supplier_bulk_rows": 0, "supplier_bulk_filename": "", "supplier_bulk_download_url": "", "supplier_bulk_error": f"找不到內建供應商 Bulk Template：{supplier_bulk_template_path.name}"}

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(supplier_bulk_template_path, output_path)
    wb = load_workbook(output_path)
    ws = wb[SUPPLIER_BULK_SHEET_NAME] if SUPPLIER_BULK_SHEET_NAME in wb.sheetnames else wb[wb.sheetnames[0]]
    cols = {
        "supplier_name": _find_template_column(ws, SUPPLIER_BULK_NAME_ALIASES, 1),
        "supplier_code": _find_template_column(ws, SUPPLIER_BULK_CODE_ALIASES, 2),
        "country_area": _find_template_column(ws, SUPPLIER_BULK_COUNTRY_ALIASES, 3),
        "supplier_address": _find_template_column(ws, SUPPLIER_BULK_ADDRESS_ALIASES, 4),
        "unit_name": _find_template_column(ws, SUPPLIER_BULK_UNIT_ALIASES, 5),
    }
    _clear_template_columns(ws, DATA_START_ROW, list(cols.values()))

    rows: list[dict[str, str]] = []
    seen: set[tuple[str, str, str, str, str]] = set()
    for _, row in expanded_with_suppliers.iterrows():
        supplier_code = _normalize_vendor_code(_first_text(row, ["supplier_code", "vendor_code", "Vendor", "Vender", "Supplier Code"]))
        if not supplier_code:
            continue
        unit_name = _first_text(row, ["transport_destination", "Transportation Destination", "transportation_destination", "production_site", "Production Site", "Unit Name"])
        supplier_name = _first_text(row, ["supplier_name", "Supplier Name", "Supplier Name (optional)"])
        if not supplier_name:
            supplier_name = _first_text(row, ["supplier_master_name", "Vendor Name", "Search Term"])
        country_area = _first_text(row, ["supplier_country_area", "Country/Area", "country_area", "Country"])
        supplier_address = _first_text(row, ["supplier_address", "Supplier Address", "transport_origin", "Transportation Origin"])
        key = (supplier_name, supplier_code, country_area, supplier_address, unit_name)
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "supplier_name": supplier_name,
            "supplier_code": supplier_code,
            "country_area": country_area,
            "supplier_address": supplier_address,
            "unit_name": unit_name,
        })

    row_idx = DATA_START_ROW
    for row in rows:
        _write_template_value(ws, row_idx, cols["supplier_name"], row["supplier_name"])
        _write_template_value(ws, row_idx, cols["supplier_code"], row["supplier_code"])
        _write_template_value(ws, row_idx, cols["country_area"], row["country_area"])
        _write_template_value(ws, row_idx, cols["supplier_address"], row["supplier_address"])
        _write_template_value(ws, row_idx, cols["unit_name"], row["unit_name"])
        row_idx += 1

    try:
        wb.save(output_path)
    finally:
        wb.close()
    return {
        "supplier_bulk_rows": int(len(rows)),
        "supplier_bulk_filename": output_path.name,
        "supplier_bulk_download_url": f"/download/{output_path.name}",
        "supplier_bulk_template_columns": cols,
    }


def _write_raw_material_bulk_from_exploded(
    exploded: pd.DataFrame,
    raw_material_template_path: str | Path,
    output_path: str | Path,
    supplier_map: dict[str, list[dict[str, str]]] | None = None,
    return_expanded: bool = False,
) -> Dict[str, Any] | tuple[Dict[str, Any], pd.DataFrame]:
    exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)

    raw_material_template_path = Path(raw_material_template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(raw_material_template_path, output_path)

    wb = load_workbook(output_path)
    if ACTIVITY_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 raw material bulk 分頁：{ACTIVITY_SHEET_NAME}")
    if RAW_MATERIAL_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 raw material bulk 分頁：{RAW_MATERIAL_SHEET_NAME}")

    activity_ws = wb[ACTIVITY_SHEET_NAME]
    raw_ws = wb[RAW_MATERIAL_SHEET_NAME]
    activity_cols = {
        "raw_name": _find_template_column(activity_ws, RAW_MATERIAL_NAME_ALIASES, 1),
        "raw_code": _find_template_column(activity_ws, RAW_MATERIAL_CODE_ALIASES, 2),
        "start_date": _find_template_column(activity_ws, DOC_START_DATE_ALIASES, 3),
        "end_date": _find_template_column(activity_ws, DOC_END_DATE_ALIASES, 4),
        "document_type": _find_template_column(activity_ws, DOCUMENT_TYPE_ALIASES, 5),
        "document_number": _find_template_column(activity_ws, DOCUMENT_NUMBER_ALIASES, 6),
        "usage": _find_template_column(activity_ws, USAGE_ALIASES, 7),
        "unit": _find_template_column(activity_ws, ACTIVITY_DATA_UNIT_ALIASES, 8),
        "data_source": _find_template_column(activity_ws, DATA_SOURCE_ALIASES, 12),
        "data_source_other": _find_template_column(activity_ws, DATA_SOURCE_OTHER_ALIASES, 13),
        "supplier_name": _find_template_column(activity_ws, SUPPLIER_NAME_ALIASES, 14),
        "transport_origin": _find_template_column(activity_ws, TRANSPORT_ORIGIN_ALIASES, 15),
        "transport_destination": _find_template_column(activity_ws, TRANSPORT_DESTINATION_ALIASES, 16),
        "target_product": _find_template_column(activity_ws, PRODUCT_LINK_ALIASES, 17),
        "comment": _find_template_column(activity_ws, COMMENT_ALIASES, 18),
        "material_group": _find_template_column(activity_ws, MATERIAL_GROUP_ALIASES, 19),
        "net_weight": (_find_template_columns(activity_ws, NET_WEIGHT_ALIASES) or [None])[0],
        "gross_weight": (_find_template_columns(activity_ws, GROSS_WEIGHT_ALIASES) or [None])[0],
        "weight_unit": (_find_template_columns(activity_ws, WEIGHT_UNIT_ALIASES) or [None])[0],
    }
    raw_cols = {
        "raw_name": _find_template_column(raw_ws, RAW_MATERIAL_NAME_ALIASES, 1),
        "raw_code": _find_template_column(raw_ws, RAW_MATERIAL_CODE_ALIASES, 2),
        "description": _find_template_column(raw_ws, RAW_MATERIAL_DESC_ALIASES, 6),
    }

    document_type_value = _document_type_for_template(wb)
    supplier_options = _extract_supplier_name_options_from_raw_template(wb)
    tbc_supplier_map = _extract_site_tbc_supplier_map_from_raw_template(wb)
    expanded, supplier_write_summary = _apply_supplier_mapping_to_exploded(
        exploded,
        supplier_map or {},
        supplier_options,
        tbc_supplier_map=tbc_supplier_map,
    )

    _clear_template_columns(activity_ws, DATA_START_ROW, list(activity_cols.values()))
    _clear_template_columns(raw_ws, DATA_START_ROW, list(raw_cols.values()))

    row_idx = DATA_START_ROW
    for _, r in expanded.iterrows():
        valid_from = r["valid_from"]
        if not isinstance(valid_from, date):
            valid_from = _date_from_value(valid_from)
        usage_value = float(r["usage"]) if not pd.isna(r["usage"]) else 0
        _write_template_value(activity_ws, row_idx, activity_cols["raw_name"], r["raw_material"])
        _write_template_value(activity_ws, row_idx, activity_cols["raw_code"], r["raw_material"])
        _write_template_value(activity_ws, row_idx, activity_cols["start_date"], _year_start(valid_from))
        _write_template_value(activity_ws, row_idx, activity_cols["end_date"], _year_end(valid_from))
        _write_template_value(activity_ws, row_idx, activity_cols["document_type"], document_type_value)
        _write_template_value(activity_ws, row_idx, activity_cols["document_number"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["usage"], usage_value)
        _write_template_value(activity_ws, row_idx, activity_cols["unit"], r["unit"])
        _write_template_value(activity_ws, row_idx, activity_cols["data_source"], "SAP")
        _write_template_value(activity_ws, row_idx, activity_cols["data_source_other"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["supplier_name"], r.get("supplier_name", ""))
        _write_template_value(activity_ws, row_idx, activity_cols["transport_origin"], r.get("transport_origin", ""))
        _write_template_value(activity_ws, row_idx, activity_cols["transport_destination"], r.get("transport_destination", ""))
        _write_template_value(activity_ws, row_idx, activity_cols["target_product"], r["target_product"])
        _write_template_value(activity_ws, row_idx, activity_cols["comment"], "")
        _write_template_value(activity_ws, row_idx, activity_cols["material_group"], r["material_group"])
        _write_template_value(activity_ws, row_idx, activity_cols.get("net_weight"), r.get("net_weight", ""))
        _write_template_value(activity_ws, row_idx, activity_cols.get("gross_weight"), r.get("gross_weight", ""))
        _write_template_value(activity_ws, row_idx, activity_cols.get("weight_unit"), r.get("weight_uom", ""))
        activity_ws.cell(row_idx, activity_cols["start_date"]).number_format = "yyyy/mm/dd"
        activity_ws.cell(row_idx, activity_cols["end_date"]).number_format = "yyyy/mm/dd"
        row_idx += 1

    raw_unique = expanded.sort_values(["raw_material"]).drop_duplicates(subset=["raw_material"])[["raw_material", "description"]] if not expanded.empty else pd.DataFrame(columns=["raw_material", "description"])
    row_idx = DATA_START_ROW
    for _, r in raw_unique.iterrows():
        _write_template_value(raw_ws, row_idx, raw_cols["raw_name"], r["raw_material"])
        _write_template_value(raw_ws, row_idx, raw_cols["raw_code"], r["raw_material"])
        _write_template_value(raw_ws, row_idx, raw_cols["description"], r["description"])
        row_idx += 1

    try:
        wb.save(output_path)
    finally:
        wb.close()
    result = {
        "output_filename": output_path.name,
        "activity_template_columns": activity_cols,
        "raw_material_template_columns": raw_cols,
        "activity_rows": int(len(expanded)),
        "raw_materials": int(expanded["raw_material"].nunique()) if not expanded.empty else 0,
        "zero_usage_rows_excluded": int(zero_usage_rows_excluded),
        "supplier_name_options": int(len(supplier_options)),
        "site_tbc_supplier_count": int(len(tbc_supplier_map)),
    }
    result.update(supplier_write_summary)
    if return_expanded:
        return result, expanded
    return result


def generate_raw_material_bulk_file(
    bom_path: str | Path | list[str | Path] | tuple[str | Path, ...],
    raw_material_template_path: str | Path,
    output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
    supplier_paths: list[str | Path] | tuple[str | Path, ...] | None = None,
    supplier_bulk_template_path: str | Path | None = None,
    supplier_bulk_output_path: str | Path | None = None,
) -> Dict[str, Any]:
    output_path = Path(output_path)
    bom_df, used_columns = _read_boms(bom_path, mapping=mapping)
    supplier_map, supplier_summary = _read_supplier_files(supplier_paths)
    exploded, summary = _explode_bom(bom_df)
    exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)
    summary["zero_usage_rows_excluded"] = int(zero_usage_rows_excluded)
    summary["activity_rows"] = int(len(exploded))
    summary["raw_materials"] = int(exploded["raw_material"].nunique()) if not exploded.empty else 0
    write_summary, expanded = _write_raw_material_bulk_from_exploded(
        exploded=exploded,
        raw_material_template_path=raw_material_template_path,
        output_path=output_path,
        supplier_map=supplier_map,
        return_expanded=True,
    )
    summary.update(write_summary)
    summary.update(supplier_summary)
    if supplier_bulk_template_path and supplier_bulk_output_path:
        summary.update(_write_supplier_bulk_create_file(expanded, supplier_bulk_template_path, supplier_bulk_output_path))
    summary["used_columns"] = used_columns
    summary["bom_files"] = int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1
    summary["bom_rows_before_dedup"] = int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_rows_after_dedup"] = int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_duplicate_rows_removed"] = int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0
    return summary


def generate_raw_material_bulk_files_by_site_zip(
    bom_path: str | Path | list[str | Path] | tuple[str | Path, ...],
    raw_material_template_path: str | Path,
    output_dir: str | Path,
    token: str,
    step1_output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
    supplier_paths: list[str | Path] | tuple[str | Path, ...] | None = None,
    supplier_bulk_template_path: str | Path | None = None,
    supplier_bulk_output_path: str | Path | None = None,
) -> Dict[str, Any]:
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    bom_df, used_columns = _read_boms(bom_path, mapping=mapping)
    supplier_map, supplier_summary = _read_supplier_files(supplier_paths)
    exploded, base_summary = _explode_bom(bom_df)
    exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)
    total_hour_by_target, working_hour_summary = _calculate_total_working_hour_by_target(step1_output_path=step1_output_path, bom_df=bom_df)
    exploded, zero_total_working_hour_rows_excluded = _exclude_zero_total_working_hour_target_rows(exploded=exploded, total_hour_by_material=total_hour_by_target)
    base_summary["zero_usage_rows_excluded"] = int(zero_usage_rows_excluded)
    base_summary["zero_total_working_hour_rows_excluded"] = int(zero_total_working_hour_rows_excluded)
    base_summary["activity_rows"] = int(len(exploded))
    base_summary["raw_materials"] = int(exploded["raw_material"].nunique()) if not exploded.empty else 0
    base_summary.update(working_hour_summary)

    site_map, step1_summary = _read_step1_product_master_maps(step1_output_path)
    work = exploded.copy()
    if work.empty:
        work["_production_site"] = ""
    else:
        work["_target_key"] = work["target_product"].apply(_normalize_material_key)
        work["_production_site"] = work["_target_key"].map(site_map).fillna("Unassigned")
        work["_production_site"] = work["_production_site"].apply(lambda x: str(x or "").strip() or "Unassigned")
        # Original official logic: Transportation Destination follows Step1 Production Site.
        work["transport_destination"] = work["_production_site"]
        if "material_group" not in work.columns:
            work["material_group"] = ""

    site_values = sorted({str(x).strip() or "Unassigned" for x in work["_production_site"].tolist()}) if not work.empty else ["Unassigned"]
    generated_files: list[dict[str, Any]] = []
    expanded_all: list[pd.DataFrame] = []
    supplier_matched_total = 0
    supplier_expanded_total = 0
    supplier_name_matched_total = 0
    supplier_name_missing_total = 0
    supplier_options_total = 0
    zip_filename = f"raw_material_activity_data_bulk_by_site_{token}.zip"
    zip_path = output_dir / zip_filename

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for site in site_values:
            site_df = work[work["_production_site"] == site].copy().drop(columns=["_target_key", "_production_site"], errors="ignore")
            safe_site = _sanitize_filename_part(site)
            file_path = output_dir / f"raw_material_activity_data_bulk_{safe_site}_{token}.xlsx"
            write_summary, expanded_site = _write_raw_material_bulk_from_exploded(
                exploded=site_df,
                raw_material_template_path=raw_material_template_path,
                output_path=file_path,
                supplier_map=supplier_map,
                return_expanded=True,
            )
            expanded_all.append(expanded_site)
            supplier_matched_total += int(write_summary.get("supplier_matched_rows", 0))
            supplier_expanded_total += int(write_summary.get("supplier_expanded_rows", 0))
            supplier_name_matched_total += int(write_summary.get("supplier_name_matched_rows", 0))
            supplier_name_missing_total += int(write_summary.get("supplier_name_missing_rows", 0))
            supplier_options_total = max(supplier_options_total, int(write_summary.get("supplier_name_options", 0)))
            zf.write(file_path, arcname=file_path.name)
            generated_files.append({
                "production_site": site,
                "filename": file_path.name,
                "activity_rows": int(write_summary.get("activity_rows", 0)),
                "raw_materials": int(write_summary.get("raw_materials", 0)),
            })

    combined_expanded = pd.concat(expanded_all, ignore_index=True) if expanded_all else pd.DataFrame()
    supplier_bulk_summary = {}
    if supplier_bulk_template_path and supplier_bulk_output_path:
        supplier_bulk_summary = _write_supplier_bulk_create_file(combined_expanded, supplier_bulk_template_path, supplier_bulk_output_path)

    unassigned_rows = int((work["_production_site"] == "Unassigned").sum()) if not work.empty else 0
    summary = dict(base_summary)
    summary.update({
        "output_filename": zip_filename,
        "download_url": f"/download/{zip_filename}",
        "split_by_production_site": True,
        "production_site_files": generated_files,
        "production_site_count": int(len(site_values)),
        "unassigned_rows": unassigned_rows,
        "used_columns": used_columns,
        "bom_files": int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1,
        "bom_rows_before_dedup": int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0,
        "bom_rows_after_dedup": int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0,
        "bom_duplicate_rows_removed": int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0,
        "supplier_matched_rows": int(supplier_matched_total),
        "supplier_expanded_rows": int(supplier_expanded_total),
        "supplier_name_matched_rows": int(supplier_name_matched_total),
        "supplier_name_missing_rows": int(supplier_name_missing_total),
        "supplier_dropdown_matched_rows": int(supplier_name_matched_total),
        "supplier_dropdown_missing_rows": int(supplier_name_missing_total),
        "supplier_name_options": int(supplier_options_total),
    })
    summary.update(step1_summary)
    summary.update(supplier_summary)
    summary.update(supplier_bulk_summary)
    return summary


BOM_FORMATTER_VERSION = "CMP_V17_6_MODULE2_HYBRID_UI_MEMORY_OPT"


# =========================================================
# Module 2 Memory Optimization V1 helpers
# Read Standard BOM once, reuse the same normalized DataFrame for:
# 1) Raw Material Bulk by site ZIP
# 2) BOM Structure latest file
# 3) Working Hour Roll-up file
# =========================================================

def export_bom_structure_file_from_dataframe(
    bom_df: pd.DataFrame,
    used_columns: dict[str, Any],
    output_path: str | Path,
) -> Dict[str, Any]:
    """Export BOM Structure using an already-loaded normalized BOM DataFrame."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    structure, summary = _explode_bom_structure(bom_df)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        structure.to_excel(writer, index=False, sheet_name="BOM Structure")
        ws = writer.book["BOM Structure"]
        ws.freeze_panes = "A2"
        for col in ws.columns:
            max_len = 12
            letter = col[0].column_letter
            for cell in col[:1000]:
                max_len = max(max_len, len(str(cell.value or "")) + 2)
            ws.column_dimensions[letter].width = min(max_len, 45)
    summary["output_filename"] = output_path.name
    summary["used_columns"] = used_columns
    summary["bom_files"] = int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1
    summary["bom_rows_before_dedup"] = int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_rows_after_dedup"] = int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0
    summary["bom_duplicate_rows_removed"] = int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0
    return summary


def generate_module2_outputs_memory_optimized(
    bom_path: str | Path | list[str | Path] | tuple[str | Path, ...],
    raw_material_template_path: str | Path,
    output_dir: str | Path,
    token: str,
    step1_output_path: str | Path,
    bom_structure_output_path: str | Path,
    working_hour_rollup_output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
    supplier_paths: list[str | Path] | tuple[str | Path, ...] | None = None,
    supplier_bulk_template_path: str | Path | None = None,
    supplier_bulk_output_path: str | Path | None = None,
) -> Dict[str, Any]:
    """Generate all Module 2 outputs while loading Standard BOM only once."""
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    bom_df, used_columns = _read_boms(bom_path, mapping=mapping)
    try:
        supplier_map, supplier_summary = _read_supplier_files(supplier_paths)
        exploded, base_summary = _explode_bom(bom_df)
        exploded, zero_usage_rows_excluded = _exclude_zero_usage_rows(exploded)

        total_hour_by_target, working_hour_summary = _calculate_total_working_hour_by_target(
            step1_output_path=step1_output_path,
            bom_df=bom_df,
        )
        exploded, zero_total_working_hour_rows_excluded = _exclude_zero_total_working_hour_target_rows(
            exploded=exploded,
            total_hour_by_material=total_hour_by_target,
        )
        base_summary["zero_usage_rows_excluded"] = int(zero_usage_rows_excluded)
        base_summary["zero_total_working_hour_rows_excluded"] = int(zero_total_working_hour_rows_excluded)
        base_summary["activity_rows"] = int(len(exploded))
        base_summary["raw_materials"] = int(exploded["raw_material"].nunique()) if not exploded.empty else 0
        base_summary.update(working_hour_summary)

        site_map, step1_summary = _read_step1_product_master_maps(step1_output_path)
        if exploded.empty:
            work = exploded.assign(_production_site="")
        else:
            work = exploded
            work["_target_key"] = work["target_product"].apply(_normalize_material_key)
            work["_production_site"] = work["_target_key"].map(site_map).fillna("Unassigned")
            work["_production_site"] = work["_production_site"].apply(lambda x: str(x or "").strip() or "Unassigned")
            work["transport_destination"] = work["_production_site"]
            if "material_group" not in work.columns:
                work["material_group"] = ""

        site_values = sorted({str(x).strip() or "Unassigned" for x in work["_production_site"].tolist()}) if not work.empty else ["Unassigned"]
        generated_files: list[dict[str, Any]] = []
        expanded_all: list[pd.DataFrame] = []
        supplier_matched_total = supplier_expanded_total = 0
        supplier_name_matched_total = supplier_name_missing_total = 0
        supplier_options_total = 0
        zip_filename = f"raw_material_activity_data_bulk_by_site_{token}.zip"
        zip_path = output_dir / zip_filename

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for site in site_values:
                site_mask = work["_production_site"].eq(site)
                site_df = work.loc[site_mask].drop(columns=["_target_key", "_production_site"], errors="ignore")
                safe_site = _sanitize_filename_part(site)
                file_path = output_dir / f"raw_material_activity_data_bulk_{safe_site}_{token}.xlsx"
                write_summary, expanded_site = _write_raw_material_bulk_from_exploded(
                    exploded=site_df,
                    raw_material_template_path=raw_material_template_path,
                    output_path=file_path,
                    supplier_map=supplier_map,
                    return_expanded=True,
                )
                if supplier_bulk_template_path and supplier_bulk_output_path and expanded_site is not None and not expanded_site.empty:
                    expanded_all.append(expanded_site)
                supplier_matched_total += int(write_summary.get("supplier_matched_rows", 0))
                supplier_expanded_total += int(write_summary.get("supplier_expanded_rows", 0))
                supplier_name_matched_total += int(write_summary.get("supplier_name_matched_rows", 0))
                supplier_name_missing_total += int(write_summary.get("supplier_name_missing_rows", 0))
                supplier_options_total = max(supplier_options_total, int(write_summary.get("supplier_name_options", 0)))
                zf.write(file_path, arcname=file_path.name)
                generated_files.append({
                    "production_site": site,
                    "filename": file_path.name,
                    "activity_rows": int(write_summary.get("activity_rows", 0)),
                    "raw_materials": int(write_summary.get("raw_materials", 0)),
                    "supplier_matched_rows": int(write_summary.get("supplier_matched_rows", 0)),
                    "supplier_expanded_rows": int(write_summary.get("supplier_expanded_rows", 0)),
                })
                del site_df, expanded_site
                gc.collect()

        supplier_bulk_summary: Dict[str, Any] = {}
        if supplier_bulk_template_path and supplier_bulk_output_path and expanded_all:
            supplier_bulk_df = pd.concat(expanded_all, ignore_index=True)
            supplier_bulk_summary = _write_supplier_bulk_create_file(
                supplier_bulk_df,
                supplier_bulk_template_path,
                supplier_bulk_output_path,
            )
            del supplier_bulk_df, expanded_all
            gc.collect()
        elif supplier_bulk_template_path and supplier_bulk_output_path:
            supplier_bulk_summary = {"supplier_bulk_rows": 0, "supplier_bulk_filename": "", "supplier_bulk_download_url": ""}

        bom_structure_summary = export_bom_structure_file_from_dataframe(
            bom_df=bom_df,
            used_columns=used_columns,
            output_path=bom_structure_output_path,
        )
        rollup_summary = generate_working_hour_rollup_file(
            step1_output_path=step1_output_path,
            bom_structure_path=bom_structure_output_path,
            output_path=working_hour_rollup_output_path,
        )

        unassigned_rows = int((work["_production_site"] == "Unassigned").sum()) if not work.empty else 0
        summary = dict(base_summary)
        summary.update({
            "output_filename": zip_filename,
            "download_url": f"/download/{zip_filename}",
            "split_by_production_site": True,
            "production_site_files": generated_files,
            "production_site_count": int(len(site_values)),
            "unassigned_rows": unassigned_rows,
            "used_columns": used_columns,
            "bom_files": int(used_columns.get("bom_files", 1)) if isinstance(used_columns, dict) else 1,
            "bom_rows_before_dedup": int(used_columns.get("bom_rows_before_dedup", 0)) if isinstance(used_columns, dict) else 0,
            "bom_rows_after_dedup": int(used_columns.get("bom_rows_after_dedup", 0)) if isinstance(used_columns, dict) else 0,
            "bom_duplicate_rows_removed": int(used_columns.get("bom_duplicate_rows_removed", 0)) if isinstance(used_columns, dict) else 0,
            "supplier_matched_rows": int(supplier_matched_total),
            "supplier_expanded_rows": int(supplier_expanded_total),
            "supplier_name_matched_rows": int(supplier_name_matched_total),
            "supplier_name_missing_rows": int(supplier_name_missing_total),
            "supplier_dropdown_matched_rows": int(supplier_name_matched_total),
            "supplier_dropdown_missing_rows": int(supplier_name_missing_total),
            "supplier_name_options": int(supplier_options_total),
            "module2_memory_optimized": True,
            "memory_optimization_version": BOM_FORMATTER_VERSION,
        })
        summary.update(supplier_summary)
        summary.update(step1_summary)
        summary.update(supplier_bulk_summary)
        summary["bom_structure_latest"] = Path(bom_structure_output_path).name
        summary["bom_structure_rows"] = int(bom_structure_summary.get("structure_rows", 0))
        summary["bom_structure_download_url"] = f"/download/{Path(bom_structure_output_path).name}"
        summary["working_hour_rollup_filename"] = Path(working_hour_rollup_output_path).name
        summary["working_hour_rollup_download_url"] = f"/download/{Path(working_hour_rollup_output_path).name}"
        summary["working_hour_rollup_latest"] = Path(working_hour_rollup_output_path).name
        summary["working_hour_rollup_latest_download_url"] = f"/download/{Path(working_hour_rollup_output_path).name}"
        summary["working_hour_rollup_rows"] = int(rollup_summary.get("summary_rows", 0))
        summary["working_hour_rollup_detail_rows"] = int(rollup_summary.get("detail_rows", 0))
        summary["working_hour_rollup_total_direct_hours"] = float(rollup_summary.get("total_direct_hours", 0))
        summary["working_hour_rollup_total_semi_hours"] = float(rollup_summary.get("total_semi_hours", 0))
        summary["working_hour_rollup_total_hours"] = float(rollup_summary.get("total_hours", 0))
        return summary
    finally:
        for name in ["bom_df", "exploded", "work", "supplier_map"]:
            if name in locals():
                try:
                    del locals()[name]
                except Exception:
                    pass
        gc.collect()
