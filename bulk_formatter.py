from __future__ import annotations

import shutil
import tempfile
import re
from datetime import date
from pathlib import Path
from typing import Any, Dict, Callable

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



def _safe_number(value: Any) -> float:
    if pd.isna(value):
        return 0.0
    text = str(value or "").strip()
    if text.upper() in ["", "NAN", "NONE"]:
        return 0.0
    try:
        return float(text.replace(",", ""))
    except ValueError:
        return 0.0


def _normalize_material(value: Any) -> str:
    return str(value or "").strip().upper()


def normalize_working_hour_source(value: Any) -> str:
    text = str(value or "direct").strip().lower().replace("-", "_").replace(" ", "_")
    if text in {"include_semi", "semi", "semi_finished", "include_semifinished", "rollup", "rolled_up", "total"}:
        return "include_semi"
    return "direct"




def _read_working_hour_rollup(working_hour_rollup_path: str | Path | None) -> pd.DataFrame:
    if not working_hour_rollup_path:
        return pd.DataFrame()
    path = Path(working_hour_rollup_path)
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="Summary", dtype=object)
    except ValueError:
        return pd.read_excel(path, sheet_name=0, dtype=object)


def _build_total_hour_lookup_from_rollup(working_hour_rollup_path: str | Path | None) -> tuple[dict[tuple[str, str], float], dict[str, float]]:
    rollup = _read_working_hour_rollup(working_hour_rollup_path)
    if rollup.empty:
        return {}, {}
    material_col = _find_optional_column(rollup, ["Material Number", "Material", "Product"])
    site_col = _find_optional_column(rollup, ["Production Site", "production site", "生產廠區", "廠區", "廠別"])
    total_col = _find_optional_column(rollup, ["Total Annual Working Hour", "Total Working Hour", "Total Hour", "Total Hours"])
    if not material_col or not total_col:
        return {}, {}
    work = rollup.copy()
    work["_material_key"] = work[material_col].apply(_normalize_material)
    work["_site_key"] = work[site_col].apply(lambda x: str(x or "").strip().upper()) if site_col else ""
    work["_total_hour"] = work[total_col].apply(_safe_number)
    by_site: dict[tuple[str, str], float] = {}
    for _, r in work.groupby(["_site_key", "_material_key"], dropna=False, as_index=False)["_total_hour"].sum().iterrows():
        site = str(r["_site_key"] or "").strip()
        material = str(r["_material_key"] or "").strip()
        if material:
            by_site[(site, material)] = float(r["_total_hour"] or 0)
    by_mat: dict[str, float] = {}
    for _, r in work.groupby(["_material_key"], dropna=False, as_index=False)["_total_hour"].sum().iterrows():
        material = str(r["_material_key"] or "").strip()
        if material:
            by_mat[material] = float(r["_total_hour"] or 0)
    return by_site, by_mat


def _read_latest_bom_structure(bom_structure_path: str | Path | None) -> pd.DataFrame:
    if not bom_structure_path:
        return pd.DataFrame()
    path = Path(bom_structure_path)
    if not path.exists():
        return pd.DataFrame()
    try:
        return pd.read_excel(path, sheet_name="BOM Structure", dtype=object)
    except ValueError:
        return pd.read_excel(path, sheet_name=0, dtype=object)


def _build_semi_hour_per_product(step1_df: pd.DataFrame, material_col: str, qty_col: str, direct_hour_col: str | None, bom_structure_path: str | Path | None) -> dict[str, float]:
    if not direct_hour_col:
        return {}
    bom_df = _read_latest_bom_structure(bom_structure_path)
    if bom_df.empty:
        return {}
    target_col = _find_optional_column(bom_df, ["Target Product", "target_product"])
    component_col = _find_optional_column(bom_df, ["Component", "component"])
    accumulated_qty_col = _find_optional_column(bom_df, ["Accumulated Quantity", "usage", "Quantity"])
    semi_col = _find_optional_column(bom_df, ["Is Semi-finished", "Is Semi", "semi_finished"])
    if not target_col or not component_col or not accumulated_qty_col:
        return {}
    annual = step1_df.copy()
    annual["_material_key"] = annual[material_col].apply(_normalize_material)
    annual["_annual_qty"] = annual[qty_col].apply(_safe_number)
    annual["_direct_hours"] = annual[direct_hour_col].apply(_safe_number)
    grouped = annual.groupby("_material_key", dropna=False, as_index=False).agg({"_annual_qty":"sum", "_direct_hours":"sum"})
    hour_per_pc, qty_by_material = {}, {}
    for _, r in grouped.iterrows():
        material = str(r["_material_key"] or "").strip()
        qty = float(r["_annual_qty"] or 0)
        hours = float(r["_direct_hours"] or 0)
        if material and qty != 0:
            hour_per_pc[material] = hours / qty
            qty_by_material[material] = qty
    if not hour_per_pc:
        return {}
    bom = bom_df.copy()
    bom["_target_key"] = bom[target_col].apply(_normalize_material)
    bom["_component_key"] = bom[component_col].apply(_normalize_material)
    bom["_accumulated_qty"] = bom[accumulated_qty_col].apply(_safe_number)
    if semi_col:
        bom = bom[bom[semi_col].astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1"])].copy()
    else:
        bom = bom[bom["_component_key"].isin(hour_per_pc.keys())].copy()
    semi_per_pc_by_target: dict[str, float] = {}
    for _, r in bom.iterrows():
        target = str(r["_target_key"] or "").strip()
        component = str(r["_component_key"] or "").strip()
        if not target or not component:
            continue
        contribution = float(r["_accumulated_qty"] or 0) * float(hour_per_pc.get(component, 0) or 0)
        if contribution:
            semi_per_pc_by_target[target] = semi_per_pc_by_target.get(target, 0.0) + contribution
    return {target: per_pc * float(qty_by_material.get(target, 0) or 0) for target, per_pc in semi_per_pc_by_target.items() if float(qty_by_material.get(target, 0) or 0)}


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
    working_hour_source: str = "direct",
    bom_structure_path: str | Path | None = None,
    working_hour_rollup_path: str | Path | None = None,
    progress_callback: Callable[..., None] | None = None,
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

    def report(step: str, processed: int = 0, total: int = 0, progress: int = 0, **extra: Any) -> None:
        if progress_callback is None:
            return
        progress_callback(
            step,
            processed=int(processed or 0),
            total=int(total or 0),
            progress=max(0, min(100, int(progress or 0))),
            **extra,
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    # 重要：先完整複製原始 template，再在複製檔上寫入資料
    shutil.copy2(bulk_template_path, output_path)

    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)
    source_total_rows = int(len(df))
    report("讀取產品活動來源資料", source_total_rows, source_total_rows, 12)

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

    working_hour_source = normalize_working_hour_source(working_hour_source)
    semi_hour_by_material: dict[str, float] = {}
    rollup_by_site_material: dict[tuple[str, str], float] = {}
    rollup_by_material: dict[str, float] = {}
    if working_hour_source == "include_semi":
        if working_hour_rollup_path is not None and Path(working_hour_rollup_path).exists():
            rollup_by_site_material, rollup_by_material = _build_total_hour_lookup_from_rollup(working_hour_rollup_path)
        elif bom_structure_path is not None and Path(bom_structure_path).exists():
            semi_hour_by_material = _build_semi_hour_per_product(
                step1_df=df,
                material_col=material_col,
                qty_col=qty_col,
                direct_hour_col=labor_hours_col,
                bom_structure_path=bom_structure_path,
            )
        else:
            raise ValueError("No Working Hour Roll-up result found. Please complete Module 2 → BOM Expansion with Step 1 Output first, then return to Step 2 to generate Product Activity Data Bulk.")

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
    source_valid_rows = 0
    duplicated_product_rows_merged = 0

    # Step2 consolidation:
    # Third-party bulk upload validates Product Name uniqueness. Therefore,
    # before writing to the template, consolidate rows by finished product
    # Material Number / Product Name. Quantity and working hours are summed.
    consolidated_rows: dict[str, Dict[str, Any]] = {}

    for source_index, (_, row) in enumerate(df.iterrows(), start=1):
        if source_index == source_total_rows or source_index % 1000 == 0:
            consolidation_progress = 12 + int((source_index / max(1, source_total_rows)) * 28)
            report("整理產品活動來源資料", source_index, source_total_rows, consolidation_progress)
        raw_product_name = row.get(material_col)
        if pd.isna(raw_product_name) or str(raw_product_name).strip() == "":
            skipped_blank_rows += 1
            continue

        product_type = row.get(product_type_col)
        is_wip = row.get(is_wip_col) if is_wip_col else None

        # 依目前規則：WIP 不寫入 bulk file
        if _is_wip(product_type, is_wip):
            excluded_wip_rows += 1
            continue

        product_name = str(raw_product_name).strip()
        product_key = _normalize_material(product_name)
        if not product_key:
            skipped_blank_rows += 1
            continue

        year = _as_year(row.get(year_col))
        qty = _safe_number(row.get(qty_col))
        product_description = _safe_text(row.get(material_desc_col)) if material_desc_col else ""
        production_site = _safe_text(row.get(production_site_col)) if production_site_col else ""

        direct_labor_hours = 0.0
        if labor_hours_col:
            raw_labor_hours = pd.to_numeric(row.get(labor_hours_col), errors="coerce")
            direct_labor_hours = 0 if pd.isna(raw_labor_hours) else float(raw_labor_hours)

        if product_key not in consolidated_rows:
            consolidated_rows[product_key] = {
                "product_name": product_name,
                "year": year,
                "qty": 0.0,
                "product_description": product_description,
                "production_site": production_site,
                "direct_labor_hours": 0.0,
                "source_row_count": 0,
            }
        else:
            duplicated_product_rows_merged += 1

        item = consolidated_rows[product_key]
        item["qty"] = float(item.get("qty", 0) or 0) + qty
        item["direct_labor_hours"] = float(item.get("direct_labor_hours", 0) or 0) + direct_labor_hours
        item["source_row_count"] = int(item.get("source_row_count", 0) or 0) + 1

        # Keep the first non-empty descriptive fields.
        if not item.get("product_description") and product_description:
            item["product_description"] = product_description
        if not item.get("production_site") and production_site:
            item["production_site"] = production_site
        if not item.get("year") and year:
            item["year"] = year

        source_valid_rows += 1

    rows_to_write: list[Dict[str, Any]] = []
    consolidated_total = int(len(consolidated_rows))
    for consolidated_index, (material_key, item) in enumerate(consolidated_rows.items(), start=1):
        if consolidated_index == consolidated_total or consolidated_index % 1000 == 0:
            validation_progress = 40 + int((consolidated_index / max(1, consolidated_total)) * 12)
            report("檢查產品活動資料", consolidated_index, consolidated_total, validation_progress)
        production_site = str(item.get("production_site") or "").strip()
        direct_labor_hours = float(item.get("direct_labor_hours", 0) or 0)
        labor_hours = None

        if labor_hours_col:
            if working_hour_source == "include_semi":
                site_key = str(production_site or "").strip().upper()
                if rollup_by_site_material or rollup_by_material:
                    if (site_key, material_key) in rollup_by_site_material:
                        labor_hours = float(rollup_by_site_material.get((site_key, material_key), 0) or 0)
                    elif material_key in rollup_by_material:
                        labor_hours = float(rollup_by_material.get(material_key, 0) or 0)
                    else:
                        labor_hours = direct_labor_hours
                else:
                    semi_labor_hours = float(semi_hour_by_material.get(material_key, 0) or 0)
                    labor_hours = direct_labor_hours + semi_labor_hours
            else:
                labor_hours = direct_labor_hours

            # 年度總工時等於 0 的產品不寫入 Bulk。
            # 先合併再排除，避免同產品多筆有正負/空值時誤判。
            if labor_hours == 0:
                excluded_zero_labor_hour_rows += int(item.get("source_row_count", 1) or 1)
                continue

        rows_to_write.append({
            "product_name": str(item.get("product_name") or "").strip(),
            "year": int(item.get("year") or date.today().year),
            "qty": float(item.get("qty", 0) or 0),
            "product_description": str(item.get("product_description") or "").strip(),
            "production_site": production_site,
            "labor_hours": labor_hours,
        })

    rows_total = int(len(rows_to_write))
    for write_index, item in enumerate(rows_to_write, start=1):
        if write_index == rows_total or write_index % 1000 == 0:
            write_progress = 52 + int((write_index / max(1, rows_total)) * 38)
            report("寫入 Product Activity Bulk", write_index, rows_total, write_progress)
        product_name = item["product_name"]
        year = int(item["year"])
        qty = item["qty"]
        product_description = item["product_description"]
        production_site = item["production_site"]
        labor_hours = item["labor_hours"]

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

    report("儲存 Product Activity Bulk", rows_total, rows_total, 95)
    wb.save(output_path)
    report("Product Activity Bulk 已完成", rows_total, rows_total, 100)

    return {
        "source_rows": int(len(df)),
        "activity_rows": int(activity_row - DATA_START_ROW),
        "product_rows": int(products_row - DATA_START_ROW),
        "excluded_wip_rows": int(excluded_wip_rows),
        "skipped_blank_rows": int(skipped_blank_rows),
        "excluded_zero_labor_hour_rows": int(excluded_zero_labor_hour_rows),
        "step2_product_name_consolidation_enabled": True,
        "valid_rows_before_consolidation": int(source_valid_rows),
        "unique_product_names_after_consolidation": int(activity_row - DATA_START_ROW),
        "duplicated_product_rows_merged": int(duplicated_product_rows_merged),
        "output_filename": output_path.name,
        "template_copy_mode": True,
        "product_description_from_material_description": True,
        "labor_hours_from_step1": bool(labor_hours_col),
        "labor_hours_written_to_bulk": bool(labor_hours_col and activity_labor_hours_col),
        "labor_hours_template_column": int(activity_labor_hours_col) if activity_labor_hours_col else None,
        "labor_hours_unit_written_to_bulk": bool(labor_hours_col and activity_labor_hours_unit_col),
        "labor_hours_unit_template_column": int(activity_labor_hours_unit_col) if activity_labor_hours_unit_col else None,
        "working_hour_source": working_hour_source,
        "working_hour_rollup_used": bool(rollup_by_site_material or rollup_by_material),
        "semi_finished_working_hour_products": int(len(semi_hour_by_material)) if isinstance(semi_hour_by_material, dict) else int(len(rollup_by_material)),
    }



def generate_product_activity_bulk_files_by_site(
    step1_output_path: str | Path,
    bulk_template_path: str | Path,
    output_dir: str | Path,
    token: str | None = None,
    working_hour_source: str = "direct",
    bom_structure_path: str | Path | None = None,
    working_hour_rollup_path: str | Path | None = None,
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
        summary = generate_product_activity_bulk_file(step1_output_path, bulk_template_path, output_path, working_hour_source=working_hour_source, bom_structure_path=bom_structure_path, working_hour_rollup_path=working_hour_rollup_path)
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
        summary = generate_product_activity_bulk_file(step1_output_path, bulk_template_path, output_path, working_hour_source=working_hour_source, bom_structure_path=bom_structure_path, working_hour_rollup_path=working_hour_rollup_path)
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

            summary = generate_product_activity_bulk_file(temp_step1, bulk_template_path, output_path, working_hour_source=working_hour_source, bom_structure_path=bom_structure_path, working_hour_rollup_path=working_hour_rollup_path)
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
    working_hour_source: str = "direct",
    bom_structure_path: str | Path | None = None,
    working_hour_rollup_path: str | Path | None = None,
    progress_callback: Callable[..., None] | None = None,
) -> Dict[str, Any]:
    """Generate Bulk files by Production Site and package them into one ZIP."""
    import zipfile

    step1_output_path = Path(step1_output_path)
    bulk_template_path = Path(bulk_template_path)
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    token = token or "bulk"

    def report(step: str, processed: int = 0, total: int = 0, progress: int = 0, **extra: Any) -> None:
        if progress_callback is None:
            return
        progress_callback(
            step,
            processed=int(processed or 0),
            total=int(total or 0),
            progress=max(0, min(100, int(progress or 0))),
            **extra,
        )

    report("讀取產品活動來源檔", 0, 0, 2)
    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)
    source_total_rows = int(len(df))
    report("讀取產品活動來源檔", source_total_rows, source_total_rows, 8)
    production_site_col = _find_optional_column(
        df,
        ["Production Site", "production site", "生產廠區", "廠區", "廠別"],
    )

    generated_files: list[Dict[str, Any]] = []

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir_path = Path(tmpdir)

        if production_site_col is None:
            output_path = tmpdir_path / f"formatted_product_activity_data_bulk_create_ALL_{token}.xlsx"
            summary = generate_product_activity_bulk_file(
                step1_output_path,
                bulk_template_path,
                output_path,
                working_hour_source=working_hour_source,
                bom_structure_path=bom_structure_path,
                working_hour_rollup_path=working_hour_rollup_path,
                progress_callback=lambda step, processed=0, total=0, progress=0, **extra: report(
                    step, processed, total, 8 + int(progress * 0.82), production_site="ALL", **extra
                ),
            )
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

            site_count = max(1, len(sites))
            for site_index, site in enumerate(sites):
                site_df = df[df[production_site_col] == site].copy()
                safe_site = _sanitize_filename(site)
                temp_step1 = tmpdir_path / f"step1_{safe_site}_{token}.xlsx"
                output_path = tmpdir_path / f"formatted_product_activity_data_bulk_create_{safe_site}_{token}.xlsx"

                with pd.ExcelWriter(temp_step1, engine="openpyxl") as writer:
                    site_df.to_excel(writer, index=False, sheet_name=SOURCE_SHEET_NAME)

                stage_start = 8 + int((site_index / site_count) * 82)
                stage_span = 82 / site_count
                summary = generate_product_activity_bulk_file(
                    temp_step1,
                    bulk_template_path,
                    output_path,
                    working_hour_source=working_hour_source,
                    bom_structure_path=bom_structure_path,
                    working_hour_rollup_path=working_hour_rollup_path,
                    progress_callback=lambda step, processed=0, total=0, progress=0, _site=site, _start=stage_start, _span=stage_span, **extra: report(
                        f"{_site}｜{step}", processed, total, int(_start + (progress / 100) * _span), production_site=_site, **extra
                    ),
                )
                generated_files.append({
                    "production_site": site,
                    "filename": output_path.name,
                    "path": output_path,
                    "summary": summary,
                })

        zip_name = f"formatted_product_activity_data_bulk_by_production_site_{token}.zip"
        zip_path = output_dir / zip_name

        report("封裝 Product Activity Bulk ZIP", 0, len(generated_files), 92)
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
            for file_index, item in enumerate(generated_files, start=1):
                z.write(item["path"], item["filename"])
                report("封裝 Product Activity Bulk ZIP", file_index, len(generated_files), 92 + int((file_index / max(1, len(generated_files))) * 7))

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

    total_activity_rows = sum(item["activity_rows"] for item in files_summary)
    report("Product Activity Bulk ZIP 已完成", total_activity_rows, total_activity_rows, 100)

    return {
        "split_by_production_site": len(generated_files) > 1,
        "production_site_count": len(generated_files),
        "activity_rows": total_activity_rows,
        "product_rows": sum(item["product_rows"] for item in files_summary),
        "excluded_wip_rows": sum(item["excluded_wip_rows"] for item in files_summary),
        "files": files_summary,
        "output_filename": zip_name,
        "download_url": f"/download/{zip_name}",
        "zip_output": True,
        "working_hour_source": normalize_working_hour_source(working_hour_source),
        "bom_structure_used": bool(bom_structure_path and Path(bom_structure_path).exists()),
        "working_hour_rollup_used": bool(working_hour_rollup_path and Path(working_hour_rollup_path).exists()),
    }
