from __future__ import annotations

import shutil
import tempfile
import re
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
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



# ---------------------------------------------------------------------------
# XLSX XML writer helpers
# ---------------------------------------------------------------------------
# Some third-party bulk uploaders validate the original workbook structure and
# external workbook references in the official template. Saving with openpyxl can
# rewrite formulas such as '[1]Input Sheet Activity Data'!A3 into ordinary local
# references. The helpers below update only worksheet XML cell contents inside a
# copied template, so workbook relationships, external links, hidden sheets,
# defined names and the original formula reference syntax remain intact.

_XLSX_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XLSX_MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_XLSX_X14AC_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
_XLSX_XR_NS = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
_XLSX_XR2_NS = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"
_XLSX_XR3_NS = "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"

# Keep Office namespace prefixes stable when serialising worksheet XML.
# If ElementTree rewrites these prefixes to ns1/ns2 while mc:Ignorable still
# contains x14ac/xr/xr2/xr3, Excel reports worksheet XML errors.
ET.register_namespace("", _XLSX_MAIN_NS)
ET.register_namespace("r", _XLSX_REL_NS)
ET.register_namespace("mc", _XLSX_MC_NS)
ET.register_namespace("x14ac", _XLSX_X14AC_NS)
ET.register_namespace("xr", _XLSX_XR_NS)
ET.register_namespace("xr2", _XLSX_XR2_NS)
ET.register_namespace("xr3", _XLSX_XR3_NS)


def _col_to_letter(col_idx: int) -> str:
    text = ""
    while col_idx:
        col_idx, rem = divmod(col_idx - 1, 26)
        text = chr(65 + rem) + text
    return text


def _cell_ref(row_idx: int, col_idx: int) -> str:
    return f"{_col_to_letter(col_idx)}{row_idx}"


def _split_cell_ref(ref: str) -> tuple[int, int]:
    m = re.match(r"^([A-Z]+)(\d+)$", ref or "")
    if not m:
        return 0, 0
    col_text, row_text = m.groups()
    col = 0
    for ch in col_text:
        col = col * 26 + (ord(ch) - 64)
    return int(row_text), col


def _excel_date_serial(d: date) -> int:
    # Excel 1900 date system serial. 1899-12-30 matches openpyxl/Excel behavior.
    return (d - date(1899, 12, 30)).days


def _find_sheet_xml_paths(xlsx_path: Path, sheet_names: list[str]) -> dict[str, str]:
    ns = {"m": _XLSX_MAIN_NS, "r": _XLSX_REL_NS, "pr": _PACKAGE_REL_NS}
    with zipfile.ZipFile(xlsx_path, "r") as z:
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        rels_root = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid_to_target = {}
    for rel in rels_root.findall("pr:Relationship", ns):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        rel_type = rel.attrib.get("Type", "")
        if rid and rel_type.endswith("/worksheet"):
            target = target.lstrip("/")
            if not target.startswith("xl/"):
                target = "xl/" + target
            rid_to_target[rid] = target
    result = {}
    for sheet in wb_root.findall("m:sheets/m:sheet", ns):
        name = sheet.attrib.get("name")
        rid = sheet.attrib.get(f"{{{_XLSX_REL_NS}}}id")
        if name in sheet_names and rid in rid_to_target:
            result[name] = rid_to_target[rid]
    missing = [name for name in sheet_names if name not in result]
    if missing:
        raise ValueError(f"找不到 bulk template 分頁：{', '.join(missing)}")
    return result


def _row_sort_key(row_el: ET.Element) -> int:
    try:
        return int(row_el.attrib.get("r", "0"))
    except ValueError:
        return 0


def _cell_sort_key(cell_el: ET.Element) -> tuple[int, int]:
    return _split_cell_ref(cell_el.attrib.get("r", ""))


def _get_or_create_row(sheet_data: ET.Element, row_idx: int) -> ET.Element:
    ns_tag = f"{{{_XLSX_MAIN_NS}}}row"
    for row_el in sheet_data.findall(ns_tag):
        if int(row_el.attrib.get("r", "0")) == row_idx:
            return row_el
    row_el = ET.Element(ns_tag, {"r": str(row_idx)})
    sheet_data.append(row_el)
    sheet_data[:] = sorted(list(sheet_data), key=_row_sort_key)
    return row_el


def _get_cell_style_template(row_el: ET.Element, col_idx: int) -> dict[str, str]:
    ref = _cell_ref(int(row_el.attrib.get("r", "0")), col_idx)
    for cell in row_el.findall(f"{{{_XLSX_MAIN_NS}}}c"):
        if cell.attrib.get("r") == ref:
            return {k: v for k, v in cell.attrib.items() if k in {"s", "cm", "vm", "ph"}}
    return {}


def _get_or_create_cell(row_el: ET.Element, row_idx: int, col_idx: int, style_template: dict[str, str] | None = None) -> ET.Element:
    ref = _cell_ref(row_idx, col_idx)
    ns_tag = f"{{{_XLSX_MAIN_NS}}}c"
    for cell in row_el.findall(ns_tag):
        if cell.attrib.get("r") == ref:
            return cell
    attrib = {"r": ref}
    if style_template:
        attrib.update(style_template)
    cell = ET.Element(ns_tag, attrib)
    row_el.append(cell)
    row_el[:] = sorted(list(row_el), key=_cell_sort_key)
    return cell


def _clear_cell(cell: ET.Element, keep_formula: bool = False) -> None:
    for child in list(cell):
        local = child.tag.split("}")[-1]
        if keep_formula and local == "f":
            continue
        cell.remove(child)
    if not keep_formula:
        cell.attrib.pop("t", None)


def _set_inline_string(cell: ET.Element, value: Any) -> None:
    _clear_cell(cell)
    cell.attrib["t"] = "inlineStr"
    is_el = ET.SubElement(cell, f"{{{_XLSX_MAIN_NS}}}is")
    t_el = ET.SubElement(is_el, f"{{{_XLSX_MAIN_NS}}}t")
    text = "" if value is None else str(value)
    if text.startswith(" ") or text.endswith(" "):
        t_el.attrib["{http://www.w3.org/XML/1998/namespace}space"] = "preserve"
    t_el.text = text


def _set_number(cell: ET.Element, value: Any) -> None:
    _clear_cell(cell)
    cell.attrib.pop("t", None)
    v_el = ET.SubElement(cell, f"{{{_XLSX_MAIN_NS}}}v")
    v_el.text = str(value)


def _set_formula(cell: ET.Element, formula_without_equal: str) -> None:
    _clear_cell(cell)
    f_el = ET.SubElement(cell, f"{{{_XLSX_MAIN_NS}}}f")
    f_el.text = formula_without_equal


def _clear_columns_from_row(root: ET.Element, start_row: int, columns: set[int], preserve_formula_cols: set[int] | None = None) -> None:
    preserve_formula_cols = preserve_formula_cols or set()
    for row_el in root.findall(f".//{{{_XLSX_MAIN_NS}}}sheetData/{{{_XLSX_MAIN_NS}}}row"):
        row_idx = int(row_el.attrib.get("r", "0"))
        if row_idx < start_row:
            continue
        for cell in row_el.findall(f"{{{_XLSX_MAIN_NS}}}c"):
            _, col_idx = _split_cell_ref(cell.attrib.get("r", ""))
            if col_idx in columns:
                _clear_cell(cell, keep_formula=col_idx in preserve_formula_cols)



def _serialize_worksheet_xml(root: ET.Element) -> bytes:
    """Serialise worksheet XML while keeping Excel-required namespace declarations.

    Important bug fix:
    ElementTree writes an XML declaration before the worksheet root. We must add
    missing namespace declarations to the <worksheet ...> start tag, not to the
    XML declaration. Adding xmlns:* before the declaration's closing ?> corrupts
    sheet1.xml / sheet2.xml and Excel reports: line 1, column 38, expected '>'.
    """
    data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    text = data.decode("utf-8")

    worksheet_start = text.find("<worksheet")
    if worksheet_start == -1:
        return data
    worksheet_end = text.find(">", worksheet_start)
    if worksheet_end == -1:
        return data

    worksheet_tag = text[worksheet_start:worksheet_end]
    required = {
        "r": _XLSX_REL_NS,
        "x14ac": _XLSX_X14AC_NS,
        "xr": _XLSX_XR_NS,
        "xr2": _XLSX_XR2_NS,
        "xr3": _XLSX_XR3_NS,
    }
    additions = []
    for prefix, uri in required.items():
        if f"xmlns:{prefix}=" not in worksheet_tag:
            additions.append(f' xmlns:{prefix}="{uri}"')
    if additions:
        text = text[:worksheet_end] + "".join(additions) + text[worksheet_end:]

    # Match Excel's usual XML declaration format closely.
    text = text.replace("<?xml version='1.0' encoding='utf-8'?>", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>', 1)
    return text.encode("utf-8")


def _write_bulk_template_xml(output_path: Path, rows: list[Dict[str, Any]], activity_labor_hours_col: int | None, activity_labor_hours_unit_col: int | None) -> None:
    sheet_paths = _find_sheet_xml_paths(output_path, [ACTIVITY_SHEET_NAME, PRODUCTS_SHEET_NAME])
    activity_xml = sheet_paths[ACTIVITY_SHEET_NAME]
    products_xml = sheet_paths[PRODUCTS_SHEET_NAME]

    with zipfile.ZipFile(output_path, "r") as zin:
        contents = {name: zin.read(name) for name in zin.namelist()}

    activity_root = ET.fromstring(contents[activity_xml])
    products_root = ET.fromstring(contents[products_xml])
    activity_sheet_data = activity_root.find(f"{{{_XLSX_MAIN_NS}}}sheetData")
    products_sheet_data = products_root.find(f"{{{_XLSX_MAIN_NS}}}sheetData")
    if activity_sheet_data is None or products_sheet_data is None:
        raise ValueError("Bulk template worksheet XML 結構異常，找不到 sheetData。")

    activity_clear_cols = {1, 2, 3, 4, 5, 6, 7, 8}
    if activity_labor_hours_col:
        activity_clear_cols.add(activity_labor_hours_col)
    if activity_labor_hours_unit_col:
        activity_clear_cols.add(activity_labor_hours_unit_col)
    _clear_columns_from_row(activity_root, DATA_START_ROW, activity_clear_cols)
    # Keep Product Sheet A-column formulas/external references. Only clear writable descriptive columns.
    _clear_columns_from_row(products_root, DATA_START_ROW, {3, 4, 6})

    # Reuse styles from row 3 cells when a target cell has to be created.
    activity_row3 = _get_or_create_row(activity_sheet_data, DATA_START_ROW)
    products_row3 = _get_or_create_row(products_sheet_data, DATA_START_ROW)
    activity_styles = {c: _get_cell_style_template(activity_row3, c) for c in activity_clear_cols}
    products_styles = {c: _get_cell_style_template(products_row3, c) for c in {1, 3, 4, 6}}

    for offset, item in enumerate(rows):
        row_idx = DATA_START_ROW + offset
        arow = _get_or_create_row(activity_sheet_data, row_idx)
        prow = _get_or_create_row(products_sheet_data, row_idx)

        def acell(col: int) -> ET.Element:
            return _get_or_create_cell(arow, row_idx, col, activity_styles.get(col))

        def pcell(col: int) -> ET.Element:
            return _get_or_create_cell(prow, row_idx, col, products_styles.get(col))

        year = int(item.get("year") or date.today().year)
        _set_inline_string(acell(1), item.get("product_name", ""))
        _set_number(acell(2), _excel_date_serial(date(year, 1, 1)))
        _set_number(acell(3), _excel_date_serial(date(year, 12, 31)))
        _set_inline_string(acell(4), "Target Product")
        _set_inline_string(acell(5), item.get("production_site", ""))
        _set_number(acell(6), item.get("qty", 0))
        _set_inline_string(acell(7), "SAP")
        _clear_cell(acell(8))
        if activity_labor_hours_col and item.get("labor_hours") is not None:
            _set_number(acell(activity_labor_hours_col), item.get("labor_hours", 0))
        if activity_labor_hours_unit_col and item.get("labor_hours") is not None:
            _set_inline_string(acell(activity_labor_hours_unit_col), "小時")

        # Preserve the official external workbook reference syntax in Product Sheet formulas.
        _set_formula(pcell(1), f"'[1]{ACTIVITY_SHEET_NAME}'!A{row_idx}")
        _set_inline_string(pcell(3), item.get("product_description", ""))
        _set_inline_string(pcell(4), "Cradle-to-Gate")
        _set_inline_string(pcell(6), "PC")

    contents[activity_xml] = _serialize_worksheet_xml(activity_root)
    contents[products_xml] = _serialize_worksheet_xml(products_root)

    tmp_path = output_path.with_suffix(output_path.suffix + ".tmp")
    with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in contents.items():
            zout.writestr(name, data)
    tmp_path.replace(output_path)


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
) -> Dict[str, Any]:
    """
    XML-safe Step 2 generator.

    The output file is created by copying the official Bulk Template and then
    updating only cell values in the two input sheets at the worksheet-XML level.
    This avoids openpyxl save-side effects that can remove the official template's
    external reference prefix, for example:
        '[1]Input Sheet Activity Data'!A3
        '[1]Dropdown Values'!...
    """

    step1_output_path = Path(step1_output_path)
    bulk_template_path = Path(bulk_template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # Copy the official template first. Do not save it through openpyxl.
    shutil.copy2(bulk_template_path, output_path)

    df = pd.read_excel(step1_output_path, sheet_name=SOURCE_SHEET_NAME, dtype=object)

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

    # Read-only inspection of template headers only. The copied workbook is never
    # saved by openpyxl, so original XML relationships remain untouched.
    wb_header = load_workbook(output_path, read_only=True, data_only=False)
    if ACTIVITY_SHEET_NAME not in wb_header.sheetnames:
        raise ValueError(f"找不到 bulk template 分頁：{ACTIVITY_SHEET_NAME}")
    if PRODUCTS_SHEET_NAME not in wb_header.sheetnames:
        raise ValueError(f"找不到 bulk template 分頁：{PRODUCTS_SHEET_NAME}")
    activity_ws = wb_header[ACTIVITY_SHEET_NAME]
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
    wb_header.close()

    excluded_wip_rows = 0
    skipped_blank_rows = 0
    excluded_zero_labor_hour_rows = 0
    consolidated_rows: dict[str, Dict[str, Any]] = {}
    source_valid_rows = 0
    duplicated_product_rows_merged = 0

    for _, row in df.iterrows():
        raw_product_name = row.get(material_col)
        if pd.isna(raw_product_name) or str(raw_product_name).strip() == "":
            skipped_blank_rows += 1
            continue

        product_type = row.get(product_type_col)
        is_wip = row.get(is_wip_col) if is_wip_col else None
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
        if not item.get("product_description") and product_description:
            item["product_description"] = product_description
        if not item.get("production_site") and production_site:
            item["production_site"] = production_site
        source_valid_rows += 1

    rows_to_write: list[Dict[str, Any]] = []
    for material_key, item in consolidated_rows.items():
        direct_labor_hours = float(item.get("direct_labor_hours", 0) or 0)
        labor_hours = None
        if labor_hours_col:
            if working_hour_source == "include_semi":
                production_site = str(item.get("production_site") or "").strip()
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
            if labor_hours == 0:
                excluded_zero_labor_hour_rows += int(item.get("source_row_count", 1) or 1)
                continue

        rows_to_write.append({
            "product_name": str(item.get("product_name") or "").strip(),
            "year": int(item.get("year") or date.today().year),
            "qty": float(item.get("qty", 0) or 0),
            "product_description": str(item.get("product_description") or "").strip(),
            "production_site": str(item.get("production_site") or "").strip(),
            "labor_hours": labor_hours,
        })

    _write_bulk_template_xml(output_path, rows_to_write, activity_labor_hours_col, activity_labor_hours_unit_col)

    return {
        "source_rows": int(len(df)),
        "activity_rows": int(len(rows_to_write)),
        "product_rows": int(len(rows_to_write)),
        "excluded_wip_rows": int(excluded_wip_rows),
        "skipped_blank_rows": int(skipped_blank_rows),
        "excluded_zero_labor_hour_rows": int(excluded_zero_labor_hour_rows),
        "step2_product_name_consolidation_enabled": True,
        "valid_rows_before_consolidation": int(source_valid_rows),
        "unique_product_names_after_consolidation": int(len(rows_to_write)),
        "duplicated_product_rows_merged": int(duplicated_product_rows_merged),
        "output_filename": output_path.name,
        "template_copy_mode": True,
        "template_xml_safe_write_mode": True,
        "preserve_external_references": True,
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
            summary = generate_product_activity_bulk_file(step1_output_path, bulk_template_path, output_path, working_hour_source=working_hour_source, bom_structure_path=bom_structure_path, working_hour_rollup_path=working_hour_rollup_path)
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

                summary = generate_product_activity_bulk_file(temp_step1, bulk_template_path, output_path, working_hour_source=working_hour_source, bom_structure_path=bom_structure_path, working_hour_rollup_path=working_hour_rollup_path)
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
        "working_hour_source": normalize_working_hour_source(working_hour_source),
        "bom_structure_used": bool(bom_structure_path and Path(bom_structure_path).exists()),
        "working_hour_rollup_used": bool(working_hour_rollup_path and Path(working_hour_rollup_path).exists()),
    }
