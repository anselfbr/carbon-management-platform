from __future__ import annotations

import shutil
import tempfile
import re
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


def _sanitize_filename(value: Any) -> str:
    text = str(value or "Unknown").strip()
    text = re.sub(r"[\\/:*?\"<>|]+", "_", text)
    text = re.sub(r"\s+", "_", text)
    return text[:80] or "Unknown"



def _normalize_col(value: Any) -> str:
    return str(value or "").strip().replace("\n", " ").replace("\r", " ")


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

def _normalize_header(value: Any) -> str:
    return str(value or "").strip().replace("\n", " ").replace("\r", " ").lower()


def _find_excel_column(ws, candidates: list[str], header_rows: tuple[int, ...] = (1, 2)) -> int | None:
    """Find a column in an Excel template by row 1 system key or row 2 display header."""
    candidate_keys = {_normalize_header(c) for c in candidates if str(c or "").strip()}
    for row_idx in header_rows:
        for col_idx in range(1, ws.max_column + 1):
            if _normalize_header(ws.cell(row_idx, col_idx).value) in candidate_keys:
                return col_idx
    return None


def _is_wip(product_type: Any, is_wip: Any = None) -> bool:
    product_type_text = str(product_type or "").strip().upper()
    is_wip_text = str(is_wip or "").strip().upper()
    return product_type_text == "WIP" or is_wip_text in ["Y", "YES", "TRUE", "1"]


def _production_site(product_type: Any) -> str:
    """Rule Master only mode.

    Production Site should come from Step1 output / rule_master.csv.
    Do not infer Production Site from Product Type in Step2.
    """
    return ""


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

    # Production Site is optional in old Step1 outputs.
    # If present, Step2 writes it into Activity Data and can split files by site.
    production_site_col = _find_optional_column(
        df,
        ["Production Site", "production site", "生產廠區", "廠區", "廠別"],
    )

    year_col = _find_column(df, ["Year"])
    material_col = _find_column(df, ["Material Number"])
    material_desc_col = _find_optional_column(df, ["Material description", "Material Description", "產品描述", "品名"])
    product_type_col = _find_column(df, ["產品類型", "Product Type"])
    qty_col = _find_column(df, ["年度生產量", "Annual Quantity", "Delivered quantity"])
    labor_hours_col = _find_optional_column(
        df,
        ["年度總工時", "Total working hours", "Selected Hours", "Total Hours", "Working Hours"],
    )

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

    # Optional working-hour target columns in bulk template.
    # Row 1 system key and row 2 display header are both searched.
    activity_labor_hours_col = _find_excel_column(
        activity_ws,
        [
            "Working Hour (optional)", "Working Hours (optional)",
            "Working Hour", "Working Hours",
            "年度總工時", "Total working hours", "Total Hours",
            "Labor Hours", "生產工時", "工時", "Hours"
        ],
    )
    activity_labor_hours_unit_col = _find_excel_column(
        activity_ws,
        [
            "Working Hours Unit (optional)", "Working Hour Unit (optional)",
            "Working Hours Unit", "Working Hour Unit",
            "工時單位", "生產工時單位", "Hours Unit", "Hour Unit"
        ],
    )

    # 只清除要寫入的欄位內容，不碰格式/驗證/公式
    # Activity Data: A:H + optional working-hour columns
    activity_clear_cols = [1, 2, 3, 4, 5, 6, 7, 8]
    for col_idx in [activity_labor_hours_col, activity_labor_hours_unit_col]:
        if col_idx and col_idx not in activity_clear_cols:
            activity_clear_cols.append(col_idx)
    _clear_target_cells(activity_ws, DATA_START_ROW, columns=activity_clear_cols)
    # Products: A, C, D, F
    # C 欄 Product Description 新增由 Step1 Output 的 Material Description 帶入
    _clear_target_cells(products_ws, DATA_START_ROW, columns=[1, 3, 4, 6])

    activity_row = DATA_START_ROW
    products_row = DATA_START_ROW
    excluded_wip_rows = 0
    skipped_blank_rows = 0
    excluded_zero_labor_hour_rows = 0

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
        labor_hours = row.get(labor_hours_col) if labor_hours_col else None
        if labor_hours_col:
            labor_hours = pd.to_numeric(labor_hours, errors="coerce")
            labor_hours = 0 if pd.isna(labor_hours) else float(labor_hours)

            # 年度總工時等於 0 的產品不寫入 Bulk。
            # 0.01、0.001 等任何非 0 小數仍需保留並寫入。
            if labor_hours == 0:
                excluded_zero_labor_hour_rows += 1
                continue

        product_name = str(product_name).strip()
        product_description = _safe_text(row.get(material_desc_col)) if material_desc_col else ""

        production_site = _safe_text(row.get(production_site_col)) if production_site_col else ""
        if not production_site:
            # Rule Master only mode: keep blank if Step1 did not provide Production Site.
            production_site = ""

        # 分頁 1：Input Sheet Activity Data
        activity_ws.cell(activity_row, 1).value = product_name
        activity_ws.cell(activity_row, 2).value = date(year, 1, 1)
        activity_ws.cell(activity_row, 3).value = date(year, 12, 31)
        activity_ws.cell(activity_row, 4).value = "Target Product"
        activity_ws.cell(activity_row, 5).value = production_site
        activity_ws.cell(activity_row, 6).value = qty
        activity_ws.cell(activity_row, 7).value = "SAP"
        activity_ws.cell(activity_row, 8).value = None

        if activity_labor_hours_col and labor_hours_col:
            activity_ws.cell(activity_row, activity_labor_hours_col).value = labor_hours

        if activity_labor_hours_unit_col and labor_hours_col:
            activity_ws.cell(activity_row, activity_labor_hours_unit_col).value = "小時"

        activity_ws.cell(activity_row, 2).number_format = "yyyy/mm/dd"
        activity_ws.cell(activity_row, 3).number_format = "yyyy/mm/dd"

        # 分頁 2：Input Sheet Products
        # A欄 Product Name 保留公式邏輯，直接引用分頁1的 A欄 Product Name。
        # 這樣不會把原本 template 的設計改成純文字。
        products_ws.cell(products_row, 1).value = f"='{ACTIVITY_SHEET_NAME}'!A{activity_row}"
        products_ws.cell(products_row, 3).value = product_description
        products_ws.cell(products_row, 4).value = "Cradle-to-Gate"
        products_ws.cell(products_row, 6).value = "PC"

        activity_row += 1
        products_row += 1

    wb.save(output_path)

    return {
        "source_rows": int(len(df)),
        "activity_rows": int(activity_row - DATA_START_ROW),
        "product_rows": int(products_row - DATA_START_ROW),
        "excluded_wip_rows": int(excluded_wip_rows),
        "skipped_blank_rows": int(skipped_blank_rows),
        "excluded_zero_labor_hour_rows": int(excluded_zero_labor_hour_rows),
        "output_filename": output_path.name,
        "template_copy_mode": True,
        "product_description_from_material_description": True,
        "labor_hours_from_step1": bool(labor_hours_col),
        "labor_hours_written_to_bulk": bool(labor_hours_col and activity_labor_hours_col),
        "labor_hours_template_column": int(activity_labor_hours_col) if activity_labor_hours_col else None,
        "labor_hours_unit_written_to_bulk": bool(labor_hours_col and activity_labor_hours_unit_col),
        "labor_hours_unit_template_column": int(activity_labor_hours_unit_col) if activity_labor_hours_unit_col else None,
    }



def generate_product_activity_bulk_files_by_site(
    step1_output_path: str | Path,
    bulk_template_path: str | Path,
    output_dir: str | Path,
    token: str | None = None,
) -> Dict[str, Any]:
    """Generate one or multiple bulk files according to Production Site.

    If Step1 output contains multiple Production Site values, this function generates
    one xlsx per site. Files are not zipped; the API returns multiple download links.
    """
    step1_output_path = Path(step1_output_path)
    bulk_template_path = Path(bulk_template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    token = token or "bulk"

    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)
    production_site_col = _find_optional_column(df, ["Production Site", "production site", "生產廠區", "廠區", "廠別"])

    if production_site_col is None:
        output_path = output_dir / f"formatted_product_activity_data_bulk_create_{token}.xlsx"
        summary = generate_product_activity_bulk_file(step1_output_path, bulk_template_path, output_path)
        return {
            "split_by_production_site": False,
            "production_site_count": 1,
            "files": [{
                "production_site": "ALL",
                "filename": output_path.name,
                "download_url": f"/download/{output_path.name}",
                "summary": summary,
            }],
            **summary,
        }

    df[production_site_col] = df[production_site_col].fillna("").astype(str).str.strip()
    sites = [s for s in sorted(df[production_site_col].unique()) if s]

    if len(sites) <= 1:
        site = sites[0] if sites else "ALL"
        output_path = output_dir / f"formatted_product_activity_data_bulk_create_{_sanitize_filename(site)}_{token}.xlsx"
        summary = generate_product_activity_bulk_file(step1_output_path, bulk_template_path, output_path)
        return {
            "split_by_production_site": False,
            "production_site_count": len(sites) or 1,
            "files": [{
                "production_site": site,
                "filename": output_path.name,
                "download_url": f"/download/{output_path.name}",
                "summary": summary,
            }],
            **summary,
        }

    files: list[Dict[str, Any]] = []
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)
        for site in sites:
            site_df = df[df[production_site_col] == site].copy()
            safe_site = _sanitize_filename(site)
            temp_step1 = tmpdir_path / f"step1_{safe_site}.xlsx"
            output_path = output_dir / f"formatted_product_activity_data_bulk_create_{safe_site}_{token}.xlsx"

            with pd.ExcelWriter(temp_step1, engine="openpyxl") as writer:
                site_df.to_excel(writer, index=False, sheet_name=SOURCE_SHEET_NAME)

            summary = generate_product_activity_bulk_file(temp_step1, bulk_template_path, output_path)
            files.append({
                "production_site": site,
                "filename": output_path.name,
                "download_url": f"/download/{output_path.name}",
                "summary": summary,
            })

    total_activity_rows = sum(int(f["summary"].get("activity_rows", 0)) for f in files)
    total_product_rows = sum(int(f["summary"].get("product_rows", 0)) for f in files)
    total_excluded_wip = sum(int(f["summary"].get("excluded_wip_rows", 0)) for f in files)

    return {
        "split_by_production_site": True,
        "production_site_count": len(files),
        "activity_rows": total_activity_rows,
        "product_rows": total_product_rows,
        "excluded_wip_rows": total_excluded_wip,
        "files": files,
    }



def generate_product_activity_bulk_files_by_site_zip(
    step1_output_path: str | Path,
    bulk_template_path: str | Path,
    output_dir: str | Path,
    token: str | None = None,
) -> Dict[str, Any]:
    """Generate Bulk files by Production Site and package them into one ZIP."""
    import zipfile

    step1_output_path = Path(step1_output_path)
    bulk_template_path = Path(bulk_template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    token = token or "bulk"

    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)
    production_site_col = _find_optional_column(
        df,
        ["Production Site", "production site", "生產廠區", "廠區", "廠別"],
    )

    generated_files: list[Dict[str, Any]] = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        if production_site_col is None:
            output_path = tmpdir_path / f"formatted_product_activity_data_bulk_create_ALL_{token}.xlsx"
            summary = generate_product_activity_bulk_file(step1_output_path, bulk_template_path, output_path)
            generated_files.append({
                "production_site": "ALL",
                "filename": output_path.name,
                "path": output_path,
                "summary": summary,
            })
        else:
            df[production_site_col] = df[production_site_col].fillna("").astype(str).str.strip()
            sites = [s for s in sorted(df[production_site_col].unique()) if s]

            if not sites:
                sites = ["ALL"]
                df[production_site_col] = "ALL"

            for site in sites:
                site_df = df[df[production_site_col] == site].copy()
                safe_site = _sanitize_filename(site)
                temp_step1 = tmpdir_path / f"step1_{safe_site}_{token}.xlsx"
                output_path = tmpdir_path / f"formatted_product_activity_data_bulk_create_{safe_site}_{token}.xlsx"

                with pd.ExcelWriter(temp_step1, engine="openpyxl") as writer:
                    site_df.to_excel(writer, index=False, sheet_name=SOURCE_SHEET_NAME)

                summary = generate_product_activity_bulk_file(temp_step1, bulk_template_path, output_path)
                generated_files.append({
                    "production_site": site,
                    "filename": output_path.name,
                    "path": output_path,
                    "summary": summary,
                })

        zip_name = f"formatted_product_activity_data_bulk_by_production_site_{token}.zip"
        zip_path = output_dir / zip_name

        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for item in generated_files:
                z.write(item["path"], item["filename"])

    files_summary = [
        {
            "production_site": item["production_site"],
            "filename": item["filename"],
            "activity_rows": int(item["summary"].get("activity_rows", 0)),
            "product_rows": int(item["summary"].get("product_rows", 0)),
            "excluded_wip_rows": int(item["summary"].get("excluded_wip_rows", 0)),
        }
        for item in generated_files
    ]

    return {
        "split_by_production_site": len(generated_files) > 1,
        "production_site_count": len(generated_files),
        "activity_rows": sum(item["activity_rows"] for item in files_summary),
        "product_rows": sum(item["product_rows"] for item in files_summary),
        "excluded_wip_rows": sum(item["excluded_wip_rows"] for item in files_summary),
        "files": files_summary,
        "output_filename": zip_name,
        "download_url": f"/download/{zip_name}",
        "zip_output": True,
    }
