from __future__ import annotations

import re
import shutil
import tempfile
import zipfile
import time
from pathlib import Path
from typing import Any, Callable, Dict, Iterable

import pandas as pd
from openpyxl import load_workbook, Workbook
from bom_formatter import _rewrite_xlsx_sheets_preserve_package

ACTIVITY_SHEET_NAME = "Input Sheet Activity Data"
DATA_START_ROW = 3
CCL_SHEET_NAME = "02.料號CCL分類表"
LCIA_SHEET_NAME = "LCIA"

FACTOR_SELECTOR_VERSION = "CMP_MODULE3_FULL_TEMPLATE_PRESERVE_LIGHTWEIGHT_20260708"


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




def _normalize_material_key(value: Any) -> str:
    text = _text(value).strip().upper()
    if text.endswith(".0"):
        text = text[:-2]
    text = text.replace("\t", "").replace("\n", "").replace("\r", "")
    return text.strip()


def _emit_progress(
    callback: Callable[..., None] | None,
    progress: int,
    step: str,
    remaining_seconds: int | None = None,
    processed_rows: int | None = None,
    total_rows: int | None = None,
) -> None:
    if not callback:
        return
    try:
        callback(progress, step, remaining_seconds, processed_rows=processed_rows, total_rows=total_rows)
    except TypeError:
        try:
            callback(progress, step, remaining_seconds)
        except TypeError:
            callback(progress, step)


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


def _read_ccl_mapping(ccl_path: str | Path, progress_callback: Callable[..., None] | None = None) -> dict[str, Dict[str, Any]]:
    _emit_progress(progress_callback, 10, "讀取 CCL 係數組配表", 30)
    wb = load_workbook(ccl_path, read_only=True, data_only=True)
    try:
        sheet_name = CCL_SHEET_NAME if CCL_SHEET_NAME in wb.sheetnames else wb.sheetnames[0]
        ws = wb[sheet_name]
        header_row = _find_header_row(ws, ["Material"])

        # CCL 對照表正式固定欄位：Material / CCL Item / 碳係數 / 係數單位
        material_col = _find_col_in_header_row(ws, header_row, ["Material"])
        ccl_item_col = _find_col_in_header_row(ws, header_row, ["CCL Item"])
        factor_name_col = _find_col_in_header_row(ws, header_row, ["係數名稱"], required=False)
        factor_col = _find_col_in_header_row(ws, header_row, ["碳係數"])
        unit_col = _find_col_in_header_row(ws, header_row, ["係數單位"], required=False)

        mapping: dict[str, Dict[str, Any]] = {}
        total_rows = max(1, ws.max_row - header_row)
        start_time = time.perf_counter()
        value_rows = ws.iter_rows(
            min_row=header_row + 1,
            max_row=ws.max_row,
            values_only=True,
        )
        for idx, values in enumerate(value_rows, start=1):
            material = _text(values[material_col - 1] if len(values) >= material_col else None)
            if not material:
                continue
            key = _normalize_material_key(material)
            if not key:
                continue
            factor_value = values[factor_col - 1] if len(values) >= factor_col else None
            unit_value = values[unit_col - 1] if unit_col and len(values) >= unit_col else ""
            safe_factor = _safe_number(factor_value)
            mapping[key] = {
                "material": material,
                "ccl_item": _text(values[ccl_item_col - 1] if len(values) >= ccl_item_col else None),
                "factor_name": _text(values[factor_name_col - 1] if factor_name_col and len(values) >= factor_name_col else ""),
                "emission_factor": safe_factor if safe_factor is not None else factor_value,
                "unit": _text(unit_value),
            }
            if idx == 1 or idx % 2000 == 0:
                elapsed = max(0.001, time.perf_counter() - start_time)
                rate = idx / elapsed
                remaining = int(max(1, (total_rows - idx) / rate + 12)) if rate > 0 else 30
                _emit_progress(progress_callback, 10 + int(min(20, idx / total_rows * 20)), "建立 CCL Material → Factor 對應索引", remaining)
        _emit_progress(progress_callback, 32, f"CCL 索引建立完成，共 {len(mapping):,} 筆", 25)
        return mapping
    finally:
        wb.close()

def _first_header_rows(ws, header_row_count: int = DATA_START_ROW - 1) -> list[list[Any]]:
    rows: list[list[Any]] = []
    max_col = max(1, int(getattr(ws, "max_column", 1) or 1))
    for values in ws.iter_rows(min_row=1, max_row=header_row_count, values_only=True):
        row = list(values or [])
        if len(row) < max_col:
            row.extend([None] * (max_col - len(row)))
        rows.append(row)
    while len(rows) < header_row_count:
        rows.append([None] * max_col)
    return rows


def _find_col_from_header_rows(header_rows: list[list[Any]], aliases: list[str], required: bool = True) -> int | None:
    alias_keys = {_norm(a) for a in aliases if str(a or "").strip()}
    for row_idx in list(range(min(len(header_rows), DATA_START_ROW - 1))):
        for col_idx, value in enumerate(header_rows[row_idx], start=1):
            if _norm(value) in alias_keys:
                return col_idx
    if required:
        raise ValueError(f"找不到欄位：{', '.join(aliases)}")
    return None


def _ensure_output_col(header_rows: list[list[Any]], aliases: list[str], preferred_header: str) -> int:
    col = _find_col_from_header_rows(header_rows, aliases, required=False)
    if col:
        return col
    max_len = max((len(r) for r in header_rows), default=0)
    new_col = max_len + 1
    for row in header_rows:
        if len(row) < new_col:
            row.extend([None] * (new_col - len(row)))
    if not header_rows:
        header_rows.append([None] * new_col)
    header_rows[0][new_col - 1] = preferred_header
    return new_col


def _pad_row(row: list[Any], width: int) -> list[Any]:
    if len(row) < width:
        row.extend([None] * (width - len(row)))
    elif len(row) > width:
        row = row[:width]
    return row


def _copy_non_activity_sheet_streaming(src_ws, dst_wb: Workbook) -> int:
    dst_ws = dst_wb.create_sheet(title=src_ws.title)
    rows = 0
    for values in src_ws.iter_rows(values_only=True):
        dst_ws.append(list(values or []))
        rows += 1
    return rows


def apply_ccl_factors_to_raw_material_bulk(
    raw_material_bulk_path: str | Path,
    ccl_mapping_path: str | Path,
    output_path: str | Path,
    progress_callback: Callable[..., None] | None = None,
    ccl_map: Dict[str, Dict[str, Any]] | None = None,
) -> Dict[str, Any]:
    """Fill Module 3 CCL factor fields while preserving the full Raw Material Bulk template.

    This full-lightweight mode copies the XLSX package and replaces only Sheet1
    data rows by streaming OpenXML. Sheet2, Sheet3, Sheet4, workbook relationships,
    styles, dropdowns, validations, formulas, hidden sheets and sheet order are
    retained byte-for-byte wherever not targeted.
    """
    perf_start = time.perf_counter()
    perf: dict[str, float] = {}
    raw_material_bulk_path = Path(raw_material_bulk_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    t0 = time.perf_counter()
    if ccl_map is None:
        ccl_map = _read_ccl_mapping(ccl_mapping_path, progress_callback=progress_callback)
    perf["read_ccl_and_build_dict"] = time.perf_counter() - t0

    t0 = time.perf_counter()
    _emit_progress(progress_callback, 34, "以全模板保留低記憶體模式開啟 raw material bulk", 60, 0, None)
    wb_meta = load_workbook(raw_material_bulk_path, read_only=True, data_only=False)
    try:
        if ACTIVITY_SHEET_NAME not in wb_meta.sheetnames:
            raise ValueError(f"找不到分頁：{ACTIVITY_SHEET_NAME}")
        ws_meta = wb_meta[ACTIVITY_SHEET_NAME]
        total_activity_rows = max(0, int(ws_meta.max_row or 0) - DATA_START_ROW + 1)
        header_rows = _first_header_rows(ws_meta, DATA_START_ROW - 1)
        original_width = max((len(r) for r in header_rows), default=0)
        headers_modified = False

        def ensure_output_col(aliases: list[str], key_header: str, label_header: str) -> int:
            nonlocal headers_modified
            col = _find_col_from_header_rows(header_rows, aliases, required=False)
            if col:
                return int(col)
            new_col = max((len(r) for r in header_rows), default=0) + 1
            for row in header_rows:
                if len(row) < new_col:
                    row.extend([None] * (new_col - len(row)))
            while len(header_rows) < DATA_START_ROW - 1:
                header_rows.append([None] * new_col)
            header_rows[0][new_col - 1] = key_header
            if len(header_rows) >= 2:
                header_rows[1][new_col - 1] = label_header
            headers_modified = True
            return new_col

        cols = {
            "material": _find_col_from_header_rows(header_rows, ["Raw Material Code", "raw_material_code", "Material", "Raw Material Name", "raw_material_name"]),
            "doc_start": _find_col_from_header_rows(header_rows, ["Doc. Start Date", "Document Start Date", "doc_start_date"], required=False),
            "factor_name": ensure_output_col(["Factor Name", "factor_name", "係數名稱", "CCL Item"], "factor_name", "Factor Name"),
            "emission_factor": ensure_output_col(["Emission Factor", "emission_factor", "碳係數"], "emission_factor", "Emission Factor"),
            "factor_source": ensure_output_col(["Factor Source", "factor_source", "係數來源"], "factor_source", "Factor Source"),
            "factor_comment": ensure_output_col(["Factor Comment", "factor_comment", "係數備註"], "factor_comment", "Factor Comment"),
            "country": ensure_output_col(["Country/Area", "Country / Area", "country_area", "Geography", "地區"], "country_area", "Country/Area"),
            "enabled_date": ensure_output_col(["Enabled Date", "activation_date", "enabled_date", "生效日期"], "activation_date", "Enabled Date"),
            "data_quality": ensure_output_col(["Data Quality", "data_quality", "資料品質"], "data_quality", "Data Quality"),
        }
        output_width = max(max(len(r) for r in header_rows), max(c for c in cols.values() if c))
    finally:
        wb_meta.close()
    perf["open_readonly_workbook"] = time.perf_counter() - t0

    counters = {
        "matched": 0,
        "unmatched": 0,
        "written_rows": 0,
        "non_empty_material_rows": 0,
        "idx": 0,
    }

    def activity_rows_factory():
        wb = load_workbook(raw_material_bulk_path, read_only=True, data_only=False)
        try:
            src_ws = wb[ACTIVITY_SHEET_NAME]
            t_stream = time.perf_counter()
            for idx, values in enumerate(src_ws.iter_rows(min_row=DATA_START_ROW, max_row=src_ws.max_row, values_only=True), start=1):
                row_values = _pad_row(list(values or []), output_width)
                material = _text(row_values[cols["material"] - 1] if len(row_values) >= cols["material"] else None)
                if material:
                    counters["non_empty_material_rows"] += 1
                    item = ccl_map.get(_normalize_material_key(material))
                    if item:
                        row_values[cols["factor_name"] - 1] = item.get("factor_name") or item.get("ccl_item") or ""
                        row_values[cols["emission_factor"] - 1] = item.get("emission_factor")
                        row_values[cols["factor_source"] - 1] = "Ecoinvent"
                        row_values[cols["factor_comment"] - 1] = "CCLibrary"
                        row_values[cols["country"] - 1] = "GLO"
                        row_values[cols["enabled_date"] - 1] = row_values[cols["doc_start"] - 1] if cols.get("doc_start") else ""
                        row_values[cols["data_quality"] - 1] = "SECONDARY"
                        counters["matched"] += 1
                        counters["written_rows"] += 1
                    else:
                        counters["unmatched"] += 1
                counters["idx"] = idx
                if idx == 1 or idx % 1000 == 0:
                    elapsed = max(0.001, time.perf_counter() - t_stream)
                    rate = idx / elapsed
                    remaining = int(max(1, (total_activity_rows - idx) / rate + 15)) if rate > 0 and total_activity_rows else 60
                    progress = 40 + int(min(48, idx / max(1, total_activity_rows) * 48))
                    _emit_progress(
                        progress_callback,
                        progress,
                        "比對原物料並串流寫入 CCL 係數欄位",
                        remaining,
                        idx,
                        total_activity_rows,
                    )
                yield row_values
        finally:
            wb.close()

    t0 = time.perf_counter()
    _emit_progress(progress_callback, 38, "複製完整 Raw Material Bulk Template 並替換係數資料列", 45, 0, total_activity_rows)
    counts = _rewrite_xlsx_sheets_preserve_package(
        raw_material_bulk_path,
        output_path,
        {
            ACTIVITY_SHEET_NAME: {
                "rows_factory": activity_rows_factory,
                "width": output_width,
                "expected_rows": int(total_activity_rows),
                "header_rows_values": header_rows if headers_modified else None,
            }
        },
    )
    perf["stream_map_write_preserve_package"] = time.perf_counter() - t0

    total_time = time.perf_counter() - perf_start
    perf["total"] = total_time
    _emit_progress(progress_callback, 100, "CCL 係數對應完成", 0, total_activity_rows, total_activity_rows)

    return {
        "output_filename": output_path.name,
        "download_url": f"/download/{output_path.name}",
        "ccl_mapping_rows": len(ccl_map),
        "matched_rows": int(counters["matched"]),
        "unmatched_rows": int(counters["unmatched"]),
        "written_rows": int(counters["written_rows"]),
        "total_rows": int(counters["non_empty_material_rows"]),
        "activity_rows": int(counts.get(ACTIVITY_SHEET_NAME, total_activity_rows)),
        "copied_other_sheet_rows": "preserved_byte_for_byte",
        "performance_seconds": {k: round(v, 3) for k, v in perf.items()},
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
        "large_dataset_mode": True,
        "template_strategy": "full-template-preserve OpenXML streaming; Sheet1 data rows replaced only; Sheet2/3/4 preserved",
    }



def _is_raw_material_bulk_zip_member(filename: str) -> bool:
    """Return True only for Raw Material Bulk workbooks inside Module 2 ZIP packages.

    Module 2C packages may also contain supplier_bulk_create workbooks. Those are
    not Raw Material Activity Data bulk files and must not be filled with factors.
    """
    name = Path(filename).name.lower()
    if name.startswith("~$") or not name.endswith((".xlsx", ".xlsm", ".xls")):
        return False
    if "supplier_bulk" in name or "supplier_create" in name or "supplier-bulk" in name:
        return False
    raw_tokens = (
        "raw_material",
        "raw-material",
        "raw materials",
        "raw_materials",
        "activity_data_bulk",
        "activity data bulk",
    )
    return any(token in name for token in raw_tokens)


def apply_ccl_factors_to_raw_material_bulk_package(
    raw_material_bulk_path: str | Path,
    ccl_mapping_path: str | Path,
    output_path: str | Path,
    progress_callback: Callable[..., None] | None = None,
) -> Dict[str, Any]:
    """Apply CCL factors to a Module 2 raw-material bulk Excel or ZIP package.

    ZIP input is processed one workbook at a time. Each workbook is streamed from
    the extracted file to a write-only XLSX, then immediately added to the output
    ZIP. Output files are not accumulated in memory.
    """
    raw_material_bulk_path = Path(raw_material_bulk_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    if raw_material_bulk_path.suffix.lower() != ".zip":
        return apply_ccl_factors_to_raw_material_bulk(
            raw_material_bulk_path,
            ccl_mapping_path,
            output_path,
            progress_callback=progress_callback,
        )

    _emit_progress(progress_callback, 2, "讀取 Module 2 ZIP 內原物料 Bulk 檔案", 60, 0, None)
    with zipfile.ZipFile(raw_material_bulk_path, "r") as zin:
        all_excel_members = [
            info for info in zin.infolist()
            if not info.is_dir()
            and not Path(info.filename).name.startswith("~$")
            and Path(info.filename).suffix.lower() in {".xlsx", ".xlsm", ".xls"}
        ]
        excel_members = [info for info in all_excel_members if _is_raw_material_bulk_zip_member(info.filename)]
        skipped_excel_members = [Path(info.filename).name for info in all_excel_members if info not in excel_members]
        if not excel_members:
            raise ValueError("Module 2 ZIP 內找不到原物料 Bulk Excel 檔案；已排除 supplier_bulk_create 等非原物料 Bulk 檔。")

        ccl_map = _read_ccl_mapping(ccl_mapping_path, progress_callback=progress_callback)
        totals = {
            "ccl_mapping_rows": len(ccl_map),
            "matched_rows": 0,
            "unmatched_rows": 0,
            "written_rows": 0,
            "total_rows": 0,
            "activity_rows": 0,
        }
        processed_files: list[dict[str, Any]] = []

        with tempfile.TemporaryDirectory(prefix="cmp_module3_zip_") as tmpdir:
            tmpdir_path = Path(tmpdir)
            total_files = len(excel_members)
            with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as zout:
                for file_idx, info in enumerate(excel_members, start=1):
                    original_name = Path(info.filename).name
                    input_file = tmpdir_path / f"input_{file_idx}_{original_name}"
                    filled_name = f"factor_filled_{original_name}"
                    filled_file = tmpdir_path / filled_name
                    with zin.open(info, "r") as src, input_file.open("wb") as dst:
                        shutil.copyfileobj(src, dst, length=1024 * 1024)

                    base_pct = 10 + int((file_idx - 1) / max(1, total_files) * 80)
                    _emit_progress(progress_callback, base_pct, f"處理 ZIP 內第 {file_idx}/{total_files} 個 Bulk：{original_name}", 60, 0, None)

                    def nested_progress(p: int, step: str, remaining_seconds: int | None = None, *, processed_rows: int | None = None, total_rows: int | None = None) -> None:
                        # Map each file's internal 34-100% progress into its share of 10-90% total progress.
                        file_span = 80 / max(1, total_files)
                        normalized = max(0, min(1, (int(p) - 34) / 66)) if p >= 34 else 0
                        package_progress = int(10 + ((file_idx - 1) + normalized) * file_span)
                        display_step = f"{step}: {original_name}"
                        _emit_progress(progress_callback, package_progress, display_step, remaining_seconds, processed_rows, total_rows)

                    summary = apply_ccl_factors_to_raw_material_bulk(
                        input_file,
                        ccl_mapping_path,
                        filled_file,
                        progress_callback=nested_progress,
                        ccl_map=ccl_map,
                    )
                    zout.write(filled_file, arcname=filled_name)
                    processed_files.append({
                        "filename": original_name,
                        "output_filename": filled_name,
                        "matched_rows": summary.get("matched_rows", 0),
                        "unmatched_rows": summary.get("unmatched_rows", 0),
                        "written_rows": summary.get("written_rows", 0),
                        "total_rows": summary.get("total_rows", 0),
                        "activity_rows": summary.get("activity_rows", 0),
                    })
                    for key in ["matched_rows", "unmatched_rows", "written_rows", "total_rows", "activity_rows"]:
                        totals[key] += int(summary.get(key, 0) or 0)

    _emit_progress(progress_callback, 100, "ZIP 內全部 Bulk 係數對應完成", 0, totals.get("activity_rows", 0), totals.get("activity_rows", 0))
    return {
        "output_filename": output_path.name,
        "download_url": f"/download/{output_path.name}",
        "input_package_filename": raw_material_bulk_path.name,
        "processed_file_count": len(processed_files),
        "processed_files": processed_files,
        "skipped_non_raw_material_bulk_files": skipped_excel_members,
        "raw_material_bulk_file_filter": "include raw_material/activity_data_bulk workbooks; exclude supplier_bulk_create workbooks",
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
        "large_dataset_mode": True,
        "template_strategy": "write_only workbook; formal bulk columns; no full-template load",
        **totals,
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

def _process_type_mode(process_type: str | None) -> str:
    """Normalize process type filter.

    Business rule:
    - production_with_transport: Activity Name must start with ``market for``.
    - production_only: all non-``market for`` activities are treated as production only.
    - all: no process-type filtering.
    """
    value = str(process_type or "all").strip().lower()
    if value in {"market_for", "market", "production_with_transport"}:
        return "market_for"
    if value in {"production", "production_only"}:
        return "production_only"
    return "all"


def _matches_process_type(activity_name: str, process_type: str | None) -> bool:
    activity = str(activity_name or "").strip().lower()
    mode = _process_type_mode(process_type)
    if mode == "market_for":
        return activity.startswith("market for")
    if mode == "production_only":
        return not activity.startswith("market for")
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


def _source_display_name(source: str) -> str:
    return _SOURCE_DISPLAY_NAMES.get(source, source)


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
        # Keyword search is intentionally limited to Activity Name and Reference Product Name only.
        activity_searchable = activity_name.lower()
        reference_searchable = reference_product_name.lower()
        searchable = f"{activity_searchable} {reference_searchable}"
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
            "_activity_lower": activity_name.lower(),
            "_activity_searchable": activity_searchable,
            "_reference_searchable": reference_searchable,
            "_searchable": searchable,
        })

    cached = {
        "mtime": stat.st_mtime,
        "size": stat.st_size,
        "source": source,
        "rows": rows,
        "geographies": geographies,
    }
    _LCIA_CACHE[cache_key] = cached
    return cached



def _keyword_matches_activity_or_reference(keyword: str, searchable: str) -> bool:
    """Match keyword against Activity Name and Reference Product Name.

    English / alphanumeric keywords are matched as complete words or complete
    phrases, so a query such as "tin" will not match "coating" merely because
    the letters appear inside another word. Non-English keywords continue to use
    a direct containment check because word boundaries are less reliable.
    """
    key = str(keyword or "").strip().lower()
    text = str(searchable or "").lower()
    if not key:
        return True
    if re.fullmatch(r"[a-z0-9][a-z0-9\s\-_/.,()+]*", key, flags=re.I):
        # Treat separators and punctuation as boundaries, but do not match inside
        # another English/alphanumeric token.
        escaped = re.escape(key)
        escaped = escaped.replace(r"\ ", r"\s+")
        pattern = rf"(?<![a-z0-9]){escaped}(?![a-z0-9])"
        return re.search(pattern, text, flags=re.I) is not None
    return key in text

def _search_lcia_file(
    path: str | Path,
    keyword: str,
    source: str,
    limit: int,
    geography: str | None = "all",
    process_type: str | None = "all",
    activity_name_keyword: str | None = "",
    reference_product_keyword: str | None = "",
) -> tuple[list[Dict[str, Any]], int]:
    cache = _load_lcia_cache(path, source)
    results: list[Dict[str, Any]] = []
    total_count = 0
    key = keyword.lower().strip()
    activity_key = str(activity_name_keyword or "").lower().strip()
    reference_key = str(reference_product_keyword or "").lower().strip()
    geography_key = str(geography or "all").strip()
    for row in cache["rows"]:
        if geography_key.lower() != "all" and row.get("geography") != geography_key:
            continue
        if not _matches_process_type(row.get("_activity_lower", ""), process_type):
            continue
        if activity_key and not _keyword_matches_activity_or_reference(activity_key, row.get("_activity_searchable", "")):
            continue
        if reference_key and not _keyword_matches_activity_or_reference(reference_key, row.get("_reference_searchable", "")):
            continue
        if not activity_key and not reference_key and not _keyword_matches_activity_or_reference(key, row.get("_searchable", "")):
            continue
        total_count += 1
        if len(results) < limit:
            clean_row = {k: v for k, v in row.items() if not k.startswith("_") and k != "source_key"}
            results.append(clean_row)
    return results, total_count



def preload_factor_libraries(apos_path: str | Path | None, cutoff_path: str | Path | None) -> Dict[str, Any]:
    """Preload APOS and Cut-off LCIA workbooks into memory at application startup.

    This keeps the first user search from paying the Excel read cost. Missing
    files are skipped so local development can still start without the databases.
    """
    loaded: list[str] = []
    skipped: list[str] = []
    errors: list[dict[str, str]] = []
    for path, source in ((apos_path, "APOS"), (cutoff_path, "Cut-off")):
        if not path:
            skipped.append(source)
            continue
        p = Path(path)
        if not p.exists():
            skipped.append(source)
            continue
        try:
            _load_lcia_cache(p, source)
            loaded.append(source)
        except Exception as exc:  # keep startup robust
            errors.append({"source": source, "message": str(exc)})
    return {
        "loaded": loaded,
        "skipped": skipped,
        "errors": errors,
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
    }

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
    activity_name_keyword: str | None = "",
    reference_product_keyword: str | None = "",
) -> Dict[str, Any]:
    keyword = str(keyword or "").strip()
    activity_name_keyword = str(activity_name_keyword or "").strip()
    reference_product_keyword = str(reference_product_keyword or "").strip()
    if len(activity_name_keyword) < 2 and len(reference_product_keyword) < 2 and len(keyword) < 2:
        raise ValueError("請至少在 Activity Name 或 Reference Product Name 輸入 2 個字元")
    page_size = int(page_size or limit or 10)
    if page_size not in {10, 20, 50}:
        page_size = 10
    page = max(1, int(page or 1))
    # Load enough rows to slice the requested page after APOS -> Cut-off priority ordering.
    fetch_limit = page * page_size
    selected_source = str(source or "all").strip().lower()
    ordered_results: list[Dict[str, Any]] = []
    total_count = 0
    if selected_source in {"all", "apos"} and apos_path and Path(apos_path).exists():
        apos_results, apos_count = _search_lcia_file(apos_path, keyword, "APOS", fetch_limit, geography, process_type, activity_name_keyword, reference_product_keyword)
        ordered_results.extend(apos_results)
        total_count += apos_count
    remaining_fetch = max(0, fetch_limit - len(ordered_results))
    if remaining_fetch > 0 and selected_source in {"all", "cut-off", "cutoff", "cut off"} and cutoff_path and Path(cutoff_path).exists():
        cutoff_results, cutoff_count = _search_lcia_file(cutoff_path, keyword, "Cut-off", remaining_fetch, geography, process_type, activity_name_keyword, reference_product_keyword)
        ordered_results.extend(cutoff_results)
        total_count += cutoff_count
    elif selected_source in {"all", "cut-off", "cutoff", "cut off"} and cutoff_path and Path(cutoff_path).exists():
        # Count Cut-off matches even when the requested page is fully occupied by APOS results.
        _, cutoff_count = _search_lcia_file(cutoff_path, keyword, "Cut-off", 0, geography, process_type, activity_name_keyword, reference_product_keyword)
        total_count += cutoff_count
    start = (page - 1) * page_size
    end = start + page_size
    results = ordered_results[start:end]
    total_pages = max(1, (total_count + page_size - 1) // page_size) if total_count else 0
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
            "activity_name_keyword": activity_name_keyword,
            "reference_product_keyword": reference_product_keyword,
        },
        "results": results,
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
    }


# =========================================================
# Module 3 Factor Library Search - SQLite index implementation
# Version 1: keep existing UI/API, replace Excel cache search with SQLite FTS
# =========================================================
import sqlite3
from contextlib import closing

FACTOR_SELECTOR_VERSION = "CMP_MODULE3_FULL_TEMPLATE_PRESERVE_LIGHTWEIGHT_20260708"
FACTOR_DB_FILENAME = "factors.db"
FACTOR_DB_SCHEMA_VERSION = "20260704_v1"


def _resolve_factor_excel_path(path: str | Path | None, source: str) -> Path | None:
    """Resolve factor workbook path, including ZIP-safe filenames such as (#U9867#U554f)."""
    if path:
        p = Path(path)
        if p.exists():
            return p
        base = p.parent if p.parent.exists() else Path(__file__).resolve().parent / "data" / "factor_library"
    else:
        base = Path(__file__).resolve().parent / "data" / "factor_library"
    if not base.exists():
        return None
    source_key = str(source or "").lower()
    patterns = ["*.xlsx", "*.xlsm", "*.xls"]
    for pattern in patterns:
        for candidate in sorted(base.glob(pattern)):
            name = candidate.name.lower()
            if source_key == "apos" and "apos" in name and "lcia" in name:
                return candidate
            if source_key in {"cut-off", "cutoff", "cut off"} and ("cut-off" in name or "cutoff" in name) and "lcia" in name:
                return candidate
    return None


def _factor_db_path(*paths: str | Path | None) -> Path:
    for path in paths:
        if path:
            parent = Path(path).parent
            if parent.exists():
                return parent / FACTOR_DB_FILENAME
    return Path(__file__).resolve().parent / "data" / "factor_library" / FACTOR_DB_FILENAME


def _excel_signature(path: Path | None) -> str:
    if not path or not path.exists():
        return "missing"
    stat = path.stat()
    return f"{path.name}|{stat.st_size}|{int(stat.st_mtime)}"


def _connect_factor_db(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA temp_store=MEMORY")
    return conn


def _init_factor_db(conn: sqlite3.Connection) -> None:
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS factor_meta (
            key TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS factor_library (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            source TEXT NOT NULL,
            source_key TEXT NOT NULL,
            activity_name TEXT,
            geography TEXT,
            emission_factor REAL,
            emission_factor_text TEXT,
            reference_product_unit TEXT,
            emission_factor_unit TEXT,
            reference_product_name TEXT,
            ipcc2021_gwp100 REAL,
            ipcc2021_gwp100_text TEXT,
            indicator TEXT
        );
        CREATE VIRTUAL TABLE IF NOT EXISTS factor_library_fts USING fts5(
            activity_name,
            reference_product_name,
            content='factor_library',
            content_rowid='id'
        );
        CREATE INDEX IF NOT EXISTS idx_factor_source_geo ON factor_library(source_key, geography);
        CREATE INDEX IF NOT EXISTS idx_factor_activity_lower ON factor_library(activity_name);
        """
    )
    conn.commit()


def _get_meta(conn: sqlite3.Connection, key: str) -> str | None:
    row = conn.execute("SELECT value FROM factor_meta WHERE key = ?", (key,)).fetchone()
    return row["value"] if row else None


def _set_meta(conn: sqlite3.Connection, key: str, value: str) -> None:
    conn.execute(
        "INSERT INTO factor_meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )


def _coerce_float_or_none(value: Any) -> float | None:
    number = _safe_number(value)
    return float(number) if number is not None else None


def _insert_lcia_rows_to_db(conn: sqlite3.Connection, path: Path, source: str) -> int:
    wb = load_workbook(path, read_only=True, data_only=True)
    if LCIA_SHEET_NAME not in wb.sheetnames:
        raise ValueError(f"{path.name} 找不到分頁：{LCIA_SHEET_NAME}")
    ws = wb[LCIA_SHEET_NAME]
    value_col = _resolve_lcia_target_column(ws)
    meta_cols = _resolve_lcia_metadata_columns(ws)
    display_source = _source_display_name(source)
    source_key = "apos" if source.lower() == "apos" else "cut-off"
    indicator = "IPCC 2021 | climate change: total (excl. biogenic CO2) | global warming potential (GWP100)"

    batch: list[tuple[Any, ...]] = []
    inserted = 0
    for values in ws.iter_rows(min_row=5, max_row=ws.max_row, values_only=True):
        activity_name = _text(values[meta_cols["activity_name"] - 1] if len(values) >= meta_cols["activity_name"] else "")
        reference_product_name = _text(values[meta_cols["reference_product_name"] - 1] if len(values) >= meta_cols["reference_product_name"] else "")
        if not activity_name and not reference_product_name:
            continue
        row_geography = _text(values[meta_cols["geography"] - 1] if len(values) >= meta_cols["geography"] else "")
        ref_unit = _text(values[meta_cols["reference_product_unit"] - 1] if len(values) >= meta_cols["reference_product_unit"] else "")
        raw_factor = values[value_col - 1] if len(values) >= value_col else None
        factor_number = _coerce_float_or_none(raw_factor)
        factor_text = _text(raw_factor)
        batch.append((
            display_source,
            source_key,
            activity_name,
            row_geography,
            factor_number,
            factor_text,
            ref_unit,
            _format_emission_factor_unit(ref_unit),
            reference_product_name,
            factor_number,
            factor_text,
            indicator,
        ))
        if len(batch) >= 1000:
            conn.executemany(
                """
                INSERT INTO factor_library(
                    source, source_key, activity_name, geography, emission_factor, emission_factor_text,
                    reference_product_unit, emission_factor_unit, reference_product_name,
                    ipcc2021_gwp100, ipcc2021_gwp100_text, indicator
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                batch,
            )
            inserted += len(batch)
            batch.clear()
    if batch:
        conn.executemany(
            """
            INSERT INTO factor_library(
                source, source_key, activity_name, geography, emission_factor, emission_factor_text,
                reference_product_unit, emission_factor_unit, reference_product_name,
                ipcc2021_gwp100, ipcc2021_gwp100_text, indicator
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            batch,
        )
        inserted += len(batch)
    wb.close()
    return inserted


def ensure_factor_database(apos_path: str | Path | None, cutoff_path: str | Path | None, force_rebuild: bool = False) -> Dict[str, Any]:
    """Build or reuse SQLite factor index. Excel is read only when the index is missing/outdated."""
    apos = _resolve_factor_excel_path(apos_path, "APOS")
    cutoff = _resolve_factor_excel_path(cutoff_path, "Cut-off")
    db_path = _factor_db_path(apos or apos_path, cutoff or cutoff_path)
    signature = f"schema={FACTOR_DB_SCHEMA_VERSION};apos={_excel_signature(apos)};cutoff={_excel_signature(cutoff)}"

    with closing(_connect_factor_db(db_path)) as conn:
        _init_factor_db(conn)
        existing_signature = _get_meta(conn, "signature")
        if not force_rebuild and existing_signature == signature:
            total = conn.execute("SELECT COUNT(*) AS c FROM factor_library").fetchone()["c"]
            return {
                "db_path": str(db_path),
                "rebuilt": False,
                "rows": int(total),
                "apos_path": str(apos) if apos else "",
                "cutoff_path": str(cutoff) if cutoff else "",
                "factor_selector_version": FACTOR_SELECTOR_VERSION,
            }

        conn.execute("DELETE FROM factor_library_fts")
        conn.execute("DELETE FROM factor_library")
        conn.execute("DELETE FROM factor_meta")
        rows_by_source: dict[str, int] = {}
        if apos and apos.exists():
            rows_by_source["APOS"] = _insert_lcia_rows_to_db(conn, apos, "APOS")
        if cutoff and cutoff.exists():
            rows_by_source["Cut-off"] = _insert_lcia_rows_to_db(conn, cutoff, "Cut-off")
        conn.execute(
            "INSERT INTO factor_library_fts(rowid, activity_name, reference_product_name) "
            "SELECT id, activity_name, reference_product_name FROM factor_library"
        )
        _set_meta(conn, "schema_version", FACTOR_DB_SCHEMA_VERSION)
        _set_meta(conn, "signature", signature)
        _set_meta(conn, "apos_path", str(apos) if apos else "")
        _set_meta(conn, "cutoff_path", str(cutoff) if cutoff else "")
        conn.commit()
        total = conn.execute("SELECT COUNT(*) AS c FROM factor_library").fetchone()["c"]
        return {
            "db_path": str(db_path),
            "rebuilt": True,
            "rows": int(total),
            "rows_by_source": rows_by_source,
            "apos_path": str(apos) if apos else "",
            "cutoff_path": str(cutoff) if cutoff else "",
            "factor_selector_version": FACTOR_SELECTOR_VERSION,
        }


def preload_factor_libraries(apos_path: str | Path | None, cutoff_path: str | Path | None) -> Dict[str, Any]:
    """Compatibility name used by main.py startup. Builds SQLite index without keeping Excel rows in memory."""
    return ensure_factor_database(apos_path, cutoff_path, force_rebuild=False)


def collect_factor_library_geographies(*paths: str | Path | None) -> list[str]:
    return ["GLO", "RoW", "RER"]


def _fts_term_query(text: str) -> str:
    """Convert user input into a conservative FTS5 AND query."""
    raw = str(text or "").strip().lower()
    if not raw:
        return ""
    tokens = re.findall(r"[a-z0-9\u4e00-\u9fff]{2,}", raw, flags=re.I)
    return " ".join(tokens[:8])


def _sqlite_source_keys(source: str) -> list[str]:
    selected = str(source or "all").strip().lower()
    if selected == "apos":
        return ["apos"]
    if selected in {"cut-off", "cutoff", "cut off"}:
        return ["cut-off"]
    return ["apos", "cut-off"]


def _build_sql_conditions(
    source_keys: list[str],
    geography: str,
    process_type: str,
    keyword: str,
    activity_name_keyword: str,
    reference_product_keyword: str,
) -> tuple[str, list[Any], str, list[Any]]:
    where = []
    params: list[Any] = []
    fts_terms: list[str] = []

    placeholders = ",".join("?" for _ in source_keys)
    where.append(f"fl.source_key IN ({placeholders})")
    params.extend(source_keys)

    geography_key = str(geography or "all").strip()
    if geography_key.lower() != "all":
        where.append("fl.geography = ?")
        params.append(geography_key)

    mode = _process_type_mode(process_type)
    if mode == "market_for":
        where.append("LOWER(fl.activity_name) LIKE 'market for%'")
    elif mode == "production_only":
        # Updated business rule: production-only means Activity Name contains production.
        where.append("LOWER(fl.activity_name) LIKE '%production%'")

    activity_query = _fts_term_query(activity_name_keyword)
    reference_query = _fts_term_query(reference_product_keyword)
    keyword_query = _fts_term_query(keyword)
    if activity_query:
        fts_terms.append(f"activity_name : ({activity_query})")
    if reference_query:
        fts_terms.append(f"reference_product_name : ({reference_query})")
    if not activity_query and not reference_query and keyword_query:
        fts_terms.append(keyword_query)

    join = ""
    if fts_terms:
        join = "JOIN factor_library_fts fts ON fts.rowid = fl.id"
        where.append("factor_library_fts MATCH ?")
        params.append(" ".join(fts_terms))

    return " AND ".join(where), params, join, params


def _row_to_result(row: sqlite3.Row) -> Dict[str, Any]:
    factor_value = row["emission_factor"]
    if factor_value is None:
        factor_value = row["emission_factor_text"]
    return {
        "source": row["source"],
        "activity_name": row["activity_name"] or "",
        "geography": row["geography"] or "",
        "emission_factor": factor_value,
        "reference_product_unit": row["reference_product_unit"] or "",
        "emission_factor_unit": row["emission_factor_unit"] or "",
        "reference_product_name": row["reference_product_name"] or "",
        "ipcc2021_gwp100": row["ipcc2021_gwp100"] if row["ipcc2021_gwp100"] is not None else row["ipcc2021_gwp100_text"],
        "indicator": row["indicator"] or "",
    }


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
    activity_name_keyword: str | None = "",
    reference_product_keyword: str | None = "",
) -> Dict[str, Any]:
    keyword = str(keyword or "").strip()
    activity_name_keyword = str(activity_name_keyword or "").strip()
    reference_product_keyword = str(reference_product_keyword or "").strip()
    if len(activity_name_keyword) < 2 and len(reference_product_keyword) < 2 and len(keyword) < 2:
        raise ValueError("請至少在關鍵字查詢或名稱查詢輸入 2 個字元")

    page_size = int(page_size or limit or 10)
    if page_size not in {10, 20, 50}:
        page_size = 10
    page = max(1, int(page or 1))
    offset = (page - 1) * page_size

    db_summary = ensure_factor_database(apos_path, cutoff_path, force_rebuild=False)
    db_path = Path(db_summary["db_path"])
    source_keys = _sqlite_source_keys(source)
    where_sql, params, join_sql, _ = _build_sql_conditions(
        source_keys,
        geography,
        process_type,
        keyword,
        activity_name_keyword,
        reference_product_keyword,
    )

    with closing(_connect_factor_db(db_path)) as conn:
        count_sql = f"SELECT COUNT(*) AS c FROM factor_library fl {join_sql} WHERE {where_sql}"
        total_count = int(conn.execute(count_sql, params).fetchone()["c"])
        rows = conn.execute(
            f"""
            SELECT fl.*
            FROM factor_library fl
            {join_sql}
            WHERE {where_sql}
            ORDER BY CASE fl.source_key WHEN 'apos' THEN 0 ELSE 1 END, fl.id
            LIMIT ? OFFSET ?
            """,
            params + [page_size, offset],
        ).fetchall()

    results = [_row_to_result(row) for row in rows]
    total_pages = max(1, (total_count + page_size - 1) // page_size) if total_count else 0
    return {
        "keyword": keyword,
        "count": len(results),
        "total_count": total_count,
        "page": page,
        "page_size": page_size,
        "total_pages": total_pages,
        "priority": "SQLite index: APOS first, then Cut-off",
        "filters": {
            "source": source or "all",
            "geography": geography or "all",
            "process_type": process_type or "all",
            "activity_name_keyword": activity_name_keyword,
            "reference_product_keyword": reference_product_keyword,
        },
        "results": results,
        "factor_db": {
            "rows": db_summary.get("rows"),
            "rebuilt": db_summary.get("rebuilt"),
        },
        "factor_selector_version": FACTOR_SELECTOR_VERSION,
    }
