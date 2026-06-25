from __future__ import annotations

import shutil
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from openpyxl import load_workbook


ACTIVITY_SHEET_NAME = "Input Sheet Activity Data"
RAW_MATERIAL_SHEET_NAME = "Input Sheet Raw Material"
DATA_START_ROW = 3


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


def _read_bom(bom_path: str | Path, mapping: dict[str, str | None] | None = None) -> tuple[pd.DataFrame, dict[str, str]]:
    df = pd.read_excel(bom_path, sheet_name=0, dtype=object)
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


def generate_raw_material_bulk_file(
    bom_path: str | Path,
    raw_material_template_path: str | Path,
    output_path: str | Path,
    mapping: dict[str, str | None] | None = None,
) -> Dict[str, Any]:
    bom_path = Path(bom_path)
    raw_material_template_path = Path(raw_material_template_path)
    output_path = Path(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(raw_material_template_path, output_path)

    bom_df, used_columns = _read_bom(bom_path, mapping=mapping)
    exploded, summary = _explode_bom(bom_df)

    wb = load_workbook(output_path)

    if ACTIVITY_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 raw material bulk 分頁：{ACTIVITY_SHEET_NAME}")
    if RAW_MATERIAL_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 raw material bulk 分頁：{RAW_MATERIAL_SHEET_NAME}")

    activity_ws = wb[ACTIVITY_SHEET_NAME]
    raw_ws = wb[RAW_MATERIAL_SHEET_NAME]

    _clear_target_cells(activity_ws, DATA_START_ROW, columns=[1, 2, 3, 4, 6, 7, 11, 16])
    _clear_target_cells(raw_ws, DATA_START_ROW, columns=[1, 2, 6])

    row_idx = DATA_START_ROW
    for _, r in exploded.iterrows():
        valid_from = r["valid_from"]
        if not isinstance(valid_from, date):
            valid_from = _date_from_value(valid_from)

        activity_ws.cell(row_idx, 1).value = r["raw_material"]
        activity_ws.cell(row_idx, 2).value = valid_from
        activity_ws.cell(row_idx, 3).value = _year_end(valid_from)
        activity_ws.cell(row_idx, 4).value = "BOM"
        activity_ws.cell(row_idx, 6).value = float(r["usage"]) if not pd.isna(r["usage"]) else 0
        activity_ws.cell(row_idx, 7).value = r["unit"]
        activity_ws.cell(row_idx, 11).value = "SAP"
        activity_ws.cell(row_idx, 16).value = r["target_product"]

        activity_ws.cell(row_idx, 2).number_format = "yyyy/mm/dd"
        activity_ws.cell(row_idx, 3).number_format = "yyyy/mm/dd"
        row_idx += 1

    raw_unique = (
        exploded.sort_values(["raw_material"])
        .drop_duplicates(subset=["raw_material"])
        [["raw_material", "description"]]
        if not exploded.empty else pd.DataFrame(columns=["raw_material", "description"])
    )

    row_idx = DATA_START_ROW
    for _, r in raw_unique.iterrows():
        raw_ws.cell(row_idx, 1).value = r["raw_material"]
        raw_ws.cell(row_idx, 2).value = r["raw_material"]
        raw_ws.cell(row_idx, 6).value = r["description"]
        row_idx += 1

    wb.save(output_path)

    summary["output_filename"] = output_path.name
    summary["used_columns"] = used_columns
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


def export_bom_structure_file(bom_path: str | Path, output_path: str | Path, mapping: dict[str, str | None] | None = None) -> Dict[str, Any]:
    """Export latest normalized BOM structure for Step 2 semi-finished working-hour roll-up."""
    bom_path = Path(bom_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    bom_df, used_columns = _read_bom(bom_path, mapping=mapping)
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
    return summary
