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


def _clear_target_cells(ws, start_row: int, columns: list[int]) -> None:
    """
    只清除指定欄位的儲存格內容，不重建 sheet、不刪除列欄、不改格式。
    這樣可以保留原始 bulk template 的格式、資料驗證、凍結窗格、欄寬、隱藏分頁等設定。
    """
    max_row = ws.max_row
    for row_idx in range(start_row, max_row + 1):
        for col_idx in columns:
            ws.cell(row_idx, col_idx).value = None


def generate_product_activity_bulk_file(
    step1_output_path: str | Path,
    bulk_template_path: str | Path,
    output_path: str | Path,
) -> Dict[str, Any]:
    """
    方案 1：直接複製原始 Bulk Template，再只覆蓋指定分頁的指定儲存格內容。

    不重新建立 Workbook。
    不新增 / 刪除分頁。
    不破壞原始 template 的樣式、欄寬、資料驗證、隱藏分頁與公式。
    原本 bulk template 的下拉選單會保留；前提是 template 的 Data Validation
    原本就涵蓋寫入的列數範圍。
    """

    step1_output_path = Path(step1_output_path)
    bulk_template_path = Path(bulk_template_path)
    output_path = Path(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 重要：先完整複製原始 template，再在複製檔上寫入資料
    shutil.copy2(bulk_template_path, output_path)

    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)

    year_col = _find_column(df, ["Year"])
    material_col = _find_column(df, ["Material Number"])
    material_desc_col = _find_optional_column(df, ["Material description", "Material Description", "產品描述", "品名"])
    product_type_col = _find_column(df, ["產品類型", "Product Type"])
    qty_col = _find_column(df, ["年度生產量", "Annual Quantity", "Delivered quantity"])
    labor_hours_col = _find_optional_column(df, ["年度總工時", "年度工時", "Selected Hours", "Total working hours", "Total Hours", "Annual Labor Hours", "Labor Hours"])

    is_wip_col = None
    for candidate in ["Is_WIP", "Is WIP", "WIP"]:
        try:
            is_wip_col = _find_column(df, [candidate])
            break
        except ValueError:
            pass

    # 直接開啟複製後的檔案，只寫入兩個指定分頁
    wb = load_workbook(output_path)

    if ACTIVITY_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 bulk template 分頁：{ACTIVITY_SHEET_NAME}")

    if PRODUCTS_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"找不到 bulk template 分頁：{PRODUCTS_SHEET_NAME}")

    activity_ws = wb[ACTIVITY_SHEET_NAME]
    products_ws = wb[PRODUCTS_SHEET_NAME]

    activity_labor_hours_col = _find_excel_column(activity_ws, ["工時", "生產工時", "年度總工時", "年度工時", "Labor Hours", "Working Hours", "Hours", "Total working hours", "Total Hours"])
    activity_labor_unit_col = _find_excel_column(activity_ws, ["工時單位", "Labor Unit", "Hour Unit", "Hours Unit", "Unit of Labor Hours"])
    product_id_col = _find_excel_column(products_ws, ["產品 ID", "產品ID", "Product ID", "ProductID", "Product Id"])

    # 只清除要寫入的欄位內容，不碰格式/驗證/公式
    # Activity Data: A:H
    activity_clear_cols = [1, 2, 3, 4, 5, 6, 7, 8]
    for col in [activity_labor_hours_col, activity_labor_unit_col]:
        if col and col not in activity_clear_cols:
            activity_clear_cols.append(col)
    _clear_target_cells(activity_ws, DATA_START_ROW, columns=activity_clear_cols)
    # Products: A, C, D, F
    # A 欄 Product Name 直接寫入實際值，不使用公式。
    # C 欄 Product Description 新增由 Step1 Output 的 Material Description 帶入
    products_clear_cols = [1, 3, 4, 6]
    if product_id_col and product_id_col not in products_clear_cols:
        products_clear_cols.append(product_id_col)
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

        # 依目前規則：WIP 不寫入 bulk file
        if _is_wip(product_type, is_wip):
            excluded_wip_rows += 1
            continue

        year = _as_year(row.get(year_col))
        qty = row.get(qty_col)
        product_name = str(product_name).strip()
        product_description = _safe_text(row.get(material_desc_col)) if material_desc_col else ""
        labor_hours = _safe_number(row.get(labor_hours_col)) if labor_hours_col else 0.0

        # 分頁 1：Input Sheet Activity Data
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

        # 分頁 2：Input Sheet Products
        # A欄 Product Name 直接寫入實際值，不使用公式，避免 Excel 陣列公式 {} 或 Spill 問題。
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
