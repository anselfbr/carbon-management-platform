from __future__ import annotations

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
BOM_FORMATTER_VERSION = "CMP_V14_7_STANDARD_BOM_MATERIAL_GROUP"


DEFAULT_MAPPING = {
    "parent_col": "Parent Node",
    "component_col": "Component",
    "qty_col": "CS03 Qty",
    "unit_col": "CS03 UoM",
    "description_col": "Component Description",
    "material_group_col": "Material group",
    "valid_from_col": "BOM Valid From",
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

    df = df.copy()
    df["_parent"] = df[parent_col].apply(_safe_text)
    df["_component"] = df[component_col].apply(_safe_text)
    df["_qty"] = df[qty_col].apply(_safe_number)
    df["_uom"] = df[unit_col].apply(_safe_text)
    df["_description"] = df[description_col].apply(_safe_text) if description_col else ""
    df["_material_group"] = df[material_group_col].apply(_safe_text) if material_group_col else ""
    df["_valid_from"] = df[valid_from_col].apply(_date_from_value) if valid_from_col else date(datetime.now().year, 1, 1)

    df = df[(df["_parent"] != "") & (df["_component"] != "")].copy()

    used_columns = {
        "parent_col": parent_col,
        "component_col": component_col,
        "qty_col": qty_col,
        "unit_col": unit_col,
        "description_col": description_col or "",
        "material_group_col": material_group_col or "",
        "valid_from_col": valid_from_col or "",
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
            part = df.copy()
            part["_source_file"] = path.name
            frames.append(part)
            used_columns = used_columns or cols
            source_rows.append({"filename": path.name, "rows": int(len(part))})
        except Exception as exc:
            errors.append(f"{path.name}: {exc}")

    if errors:
        raise ValueError("；".join(errors))
    if not frames:
        raise ValueError("沒有可處理的 BOM 資料")

    merged = pd.concat(frames, ignore_index=True)
    before_dedup = int(len(merged))
    dedup_subset = ["_parent", "_component", "_qty", "_uom", "_description", "_material_group", "_valid_from"]
    merged = merged.drop_duplicates(subset=dedup_subset, keep="first").reset_index(drop=True)
    after_dedup = int(len(merged))

    used = dict(used_columns or {})
    used["bom_files"] = int(len(paths))
    used["bom_rows_before_dedup"] = before_dedup
    used["bom_rows_after_dedup"] = after_dedup
    used["bom_duplicate_rows_removed"] = before_dedup - after_dedup
    used["bom_source_files"] = source_rows
    return merged, used

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
                        "level": next_level,
                    })

    exploded = pd.DataFrame(output_rows)
    if exploded.empty:
        exploded = pd.DataFrame(columns=[
            "target_product", "raw_material", "usage", "unit", "description", "material_group", "valid_from", "level"
        ])
    else:
        exploded = (
            exploded.groupby(["target_product", "raw_material", "unit"], dropna=False, as_index=False)
            .agg({
                "usage": "sum",
                "description": "first",
                "material_group": "first",
                "valid_from": "first",
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
    """
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
        _write_template_value(activity_ws, row_idx, activity_cols["start_date"], valid_from)
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

    Module 2 uses the Step1 output as product master data:
    - Material Number / Target Product -> Production Site for split export.
    - Material Number / Target Product -> Material Group for Raw Material Bulk.

    Material Group intentionally comes from Step1 output, not the Standard BOM
    Material group column, because Step1 is the controlled classification result.
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
        material_groups = sorted({str(x).strip() for x in group["_step1_material_group"] if str(x).strip()})

        if sites:
            site_map[str(material)] = sites[0]
            if len(sites) > 1:
                duplicate_site_conflicts.append(f"{material}: {', '.join(sites)}")

    return site_map, {
        "step1_rows": int(len(df)),
        "step1_mapped_materials": int(len(site_map)),
                "step1_site_conflicts": duplicate_site_conflicts,
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

    site_values = sorted({str(x).strip() or "Unassigned" for x in work["_production_site"].tolist()}) if not work.empty else ["Unassigned"]

    generated_files: list[dict[str, Any]] = []
    zip_filename = f"raw_material_activity_data_bulk_by_site_{token}.zip"
    zip_path = output_dir / zip_filename

    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for site in site_values:
            site_df = work[work["_production_site"] == site].copy()
            site_df = site_df.drop(columns=["_target_key", "_production_site", "_step1_material_group"], errors="ignore")
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
                    "Level": next_level,
                    "Is Semi-finished": "Y" if is_semi else "N",
                })
                if is_semi:
                    stack.append((component, next_qty, next_level, path + [component]))
    structure = pd.DataFrame(rows)
    if structure.empty:
        structure = pd.DataFrame(columns=["Target Product", "Parent Material", "Component", "Quantity Per Parent", "Accumulated Quantity", "Unit", "Component Description", "Material Group", "Valid From", "Level", "Is Semi-finished"])
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
