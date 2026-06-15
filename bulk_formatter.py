from __future__ import annotations

import shutil
from datetime import date
from pathlib import Path
from typing import Any, Dict

import pandas as pd
from openpyxl import load_workbook


SOURCE_SHEET_NAME = "Plant_Material年度產量"
ACTIVITY_SHEET_NAME = "Input Sheet Activity Data"
PRODUCTS_SHEET_NAME = "Input Sheet Products"

# Bulk template row 1 = system field key
# Bulk template row 2 = display header
# Data starts from row 3
DATA_START_ROW = 3


def _normalize_col(value: Any) -> str:
    return str(value or "").strip().replace("\n", " ").replace("\r", " ")


def _normalize_header(value: Any) -> str:
    return (
        str(value or "")
        .strip()
        .replace("\n", " ")
        .replace("\r", " ")
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .lower()
    )


def _find_column(df: pd.DataFrame, candidates: list[str]) -> str:
    normalized = {_normalize_col(c).lower(): c for c in df.columns}
    for name in candidates:
        key = _normalize_col(name).lower()
        if key in normalized:
            return normalized[key]
    raise ValueError(f"找不到必要欄位：{', '.join(candidates)}")


def _find_optional_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    try:
        return _find_column(df, candidates)
    except ValueError:
        return None


def _find_excel_column(ws, candidates: list[str], header_rows: tuple[int, ...] = (1, 2)) -> int | None:
    candidate_keys = {_normalize_header(c) for c in candidates if str(c or "").strip()}
    if not candidate_keys:
        return None

    for row_idx in header_rows:
        for col_idx in range(1, ws.max_column + 1):
            cell_key = _normalize_header(ws.cell(row_idx, col_idx).value)
            if cell_key in candidate_keys:
                return col_idx

    for row_idx in header_rows:
        for col_idx in range(1, ws.max_column + 1):
            cell_key = _normalize_header(ws.cell(row_idx, col_idx).value)
            if not cell_key:
                continue
            for target in candidate_keys:
                if target and (target in cell_key or cell_key in target):
                    return col_idx

    return None


def _is_wip(product_type: Any, is_wip: Any = None) -> bool:
    product_type_text = str(product_type or "").strip().upper()
    is_wip_text = str(is_wip or "").strip().upper()
    return product_type_text == "WIP" or is_wip_text in ["Y", "YES", "TRUE", "1"]


def _production_site(product_type: Any) -> str:
    product_type_text = str(product_type or "").strip().upper()
    if product_type_text == "NB":
        return "常州廠(A2)-IPS"
    if product_type_text == "TP":
        return "常州廠(A9)-IPS"
    return "石碣廠-IPS"


def _as_year(value: Any) -> int:
    if pd.isna(value):
        raise ValueError("Year 欄位有空白值")
    if hasattr(value, "year"):
        return int(value.year)
    return int(str(value).strip()[:4])


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


def _clear_target_cells(ws, start_row: int, columns: list[int]) -> None:
    max_row = ws.max_row
    for row_idx in range(start_row, max_row + 1):
        for col_idx in columns:
            if col_idx:
                ws.cell(row_idx, col_idx).value = None


def generate_product_activity_bulk_file(
    step1_output_path: str | Path,
    bulk_template_path: str | Path,
    output_path: str | Path,
) -> Dict[str, Any]:
    """
    Step 2：Step 1 Output + Product Activity Bulk Template -> Formatted Bulk File

    - 從 Step 1 Output 的 Plant_Material年度產量 分頁抓「年度工時」
    - 寫入 Bulk Template 的 Input Sheet Activity Data 分頁「工時」欄位
    - 「工時單位」欄位固定填「小時」
    - Input Sheet Products 分頁「產品 ID」欄位 = Material Number
    """

    step1_output_path = Path(step1_output_path)
    bulk_template_path = Path(bulk_template_path)
    output_path = Path(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    shutil.copy2(bulk_template_path, output_path)

    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)

    year_col = _find_column(df, ["Year"])
    material_col = _find_column(df, ["Material Number"])
    material_desc_col = _find_optional_column(df, ["Material description", "Material Description", "產品描述", "品名"])
    product_type_col = _find_column(df, ["產品類型", "Product Type"])
    qty_col = _find_column(df, ["年度生產量", "Annual Quantity", "Delivered quantity"])
    labor_hours_col = _find_optional_column(df, ["年度工時", "Selected Hours", "Total Hours", "Annual Labor Hours", "Labor Hours"])

    is_wip_col = None
    for candidate in ["Is_WIP", "Is WIP", "WIP"]:
        try:
            is_wip_col = _find_column(df, [candidate])
            break
        except ValueError:
            pass

    wb = load_workbook(output_path)

    if ACTIVITY_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 bulk template 分頁：{ACTIVITY_SHEET_NAME}")

    if PRODUCTS_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 bulk template 分頁：{PRODUCTS_SHEET_NAME}")

    activity_ws = wb[ACTIVITY_SHEET_NAME]
    products_ws = wb[PRODUCTS_SHEET_NAME]

    activity_labor_hours_col = _find_excel_column(activity_ws, [
        "工時", "生產工時", "年度工時", "Labor Hours", "Labor Hour", "Working Hours", "Hours", "Total Hours"
    ])
    activity_labor_unit_col = _find_excel_column(activity_ws, [
        "工時單位", "Labor Unit", "Labor Hours Unit", "Hour Unit", "Hours Unit", "Unit of Labor Hours"
    ])
    product_id_col = _find_excel_column(products_ws, [
        "產品 ID", "產品ID", "Product ID", "ProductID", "Product Id"
    ])

    activity_clear_cols = [1, 2, 3, 4, 5, 6, 7, 8]
    for col in [activity_labor_hours_col, activity_labor_unit_col]:
        if col and col not in activity_clear_cols:
            activity_clear_cols.append(col)

    products_clear_cols = [1, 3, 4, 6]
    if product_id_col and product_id_col not in products_clear_cols:
        products_clear_cols.append(product_id_col)

    _clear_target_cells(activity_ws, DATA_START_ROW, columns=activity_clear_cols)
    _clear_target_cells(products_ws, DATA_START_ROW, columns=products_clear_cols)

    activity_row = DATA_START_ROW
    products_row = DATA_START_ROW
    excluded_wip_rows = 0
    skipped_blank_rows = 0
    labor_hours_written = 0
    labor_unit_written = 0
    product_id_written = 0

    for _, row in df.iterrows():
        product_name = row.get(material_col)
        if pd.isna(product_name) or str(product_name).strip() == "":
            skipped_blank_rows += 1
            continue

        product_type = row.get(product_type_col)
        is_wip = row.get(is_wip_col) if is_wip_col else None

        if _is_wip(product_type, is_wip):
            excluded_wip_rows += 1
            continue

        year = _as_year(row.get(year_col))
        qty = row.get(qty_col)
        product_name = str(product_name).strip()
        product_description = _safe_text(row.get(material_desc_col)) if material_desc_col else ""
        labor_hours = _safe_number(row.get(labor_hours_col)) if labor_hours_col else 0.0

        activity_ws.cell(activity_row, 1).value = product_name
        activity_ws.cell(activity_row, 2).value = date(year, 1, 1)
        activity_ws.cell(activity_row, 3).value = date(year, 12, 31)
        activity_ws.cell(activity_row, 4).value = "Target Product"
        activity_ws.cell(activity_row, 5).value = _production_site(product_type)
        activity_ws.cell(activity_row, 6).value = qty
        activity_ws.cell(activity_row, 7).value = "SAP"
        activity_ws.cell(activity_row, 8).value = None

        if activity_labor_hours_col:
            activity_ws.cell(activity_row, activity_labor_hours_col).value = labor_hours
            labor_hours_written += 1

        if activity_labor_unit_col:
            activity_ws.cell(activity_row, activity_labor_unit_col).value = "小時"
            labor_unit_written += 1

        activity_ws.cell(activity_row, 2).number_format = "yyyy/mm/dd"
        activity_ws.cell(activity_row, 3).number_format = "yyyy/mm/dd"

        products_ws.cell(products_row, 1).value = product_name
        products_ws.cell(products_row, 3).value = product_description
        products_ws.cell(products_row, 4).value = "Cradle-to-Gate"
        products_ws.cell(products_row, 6).value = "PC"

        if product_id_col:
            products_ws.cell(products_row, product_id_col).value = product_name
            product_id_written += 1

        activity_row += 1
        products_row += 1

    wb.save(output_path)

    return {
        "source_rows": int(len(df)),
        "activity_rows": int(activity_row - DATA_START_ROW),
        "product_rows": int(products_row - DATA_START_ROW),
        "excluded_wip_rows": int(excluded_wip_rows),
        "skipped_blank_rows": int(skipped_blank_rows),
        "output_filename": output_path.name,
        "template_copy_mode": True,
        "product_description_from_material_description": True,
        "labor_hours_source_column": labor_hours_col or "",
        "labor_hours_template_column": int(activity_labor_hours_col or 0),
        "labor_unit_template_column": int(activity_labor_unit_col or 0),
        "product_id_template_column": int(product_id_col or 0),
        "labor_hours_written": int(labor_hours_written),
        "labor_unit_written": int(labor_unit_written),
        "product_id_written": int(product_id_written),
    }
