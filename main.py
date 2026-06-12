
from __future__ import annotations

import re
import uuid
import traceback
from pathlib import Path
from typing import Optional

import pandas as pd
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

app = FastAPI(title="年度產品產量與分類平台", version="3.0.0")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

# =========================================================
# v3 分類優先順序：
# 1. Material Number Rule（使用者自訂成品料號規則）
# 2. Naming Rule（Prefix 命名規則）
# 3. Product Series Master（產品系列.xlsx 轉入）
# 4. WIP（以上都找不到時，視為半品）
# =========================================================

NB_CUSTOMERS = {
    "1": "HP", "2": "Dell", "3": "Lenovo", "4": "Others", "5": "ASUS",
    "6": "Microsoft", "7": "Acer", "8": "LG", "9": "Focal/歌泰",
    "0": "Incubation", "A": "Framework", "M": "MSI",
}

DT_CUSTOMERS = {
    "00": "Google", "12": "Cooler Master", "20": "Old Pickup", "21": "HP",
    "27": "Goldtouch", "33": "Microsoft", "34": "TUL", "36": "TG",
    "38": "Logitech", "39": "Logitech", "54": "Fujitsu", "55": "Japan Others",
    "60": "Roccat", "62": "Corsair", "65": "SteelSeries", "66": "Distribution",
    "67": "Razer", "68": "LG", "69": "HyperX", "70": "Samsung",
    "71": "Glorious", "72": "Onward", "80": "UI", "81": "Dell",
    "83": "Massdrop", "86": "Nytec", "88": "Lenovo", "89": "Taiwan Others",
    "90": "LITEON", "93": "ASUS", "95": "Cherry", "96": "Acer", "98": "Acer Gadget",
}

GENERIC_WORDS = {
    "FG", "ASSY", "ASSEMBLY", "MODULE", "MOD", "BL", "BLANK", "NEW", "EURO",
    "NOEURO", "NO", "PCBA", "KB", "KEYBOARD", "TOUCH", "PAD", "CH",
}
SERIES_RE = re.compile(r"^([A-Z]{2}[A-Z0-9]*\d[A-Z0-9]*)")

REQUIRED_ALIASES = {
    "order": ["Order"],
    "plant": ["Plant", "Plant code"],
    "material_number": ["Material Number", "Material", "Product Material Number"],
    "material_description": ["Material description", "Material Description"],
    "delivered_quantity": ["Delivered quantity (GMEIN)", "Delivered quantity", "Delivered Quantity"],
    "finish_date": ["Actual finish date", "Actual Finish Date", "Finish date", "Actual finish"],
}


def ensure_master_files() -> None:
    rule_path = DATA_DIR / "rule_master.csv"
    material_rule_path = DATA_DIR / "material_rule_master.csv"
    product_series_path = DATA_DIR / "product_series_master.csv"

    if not rule_path.exists():
        rule_path.write_text(
            "Prefix,Product Type,Customer Code Logic,Enabled\n"
            "SN,NB,NB_3RD_CHAR,Y\n"
            "SP,TP,NB_3RD_CHAR,Y\n"
            "SM,DT Mouse,DT_3_4_CHAR,Y\n"
            "SK,DT Keyboard,DT_3_4_CHAR,Y\n",
            encoding="utf-8-sig",
        )

    if not material_rule_path.exists():
        material_rule_path.write_text(
            "Material Number,Product Type,Customer,Enabled\n",
            encoding="utf-8-sig",
        )

    if not product_series_path.exists():
        product_series_path.write_text(
            "Plant,Product series,產品類型,客戶代碼,客戶名稱\n",
            encoding="utf-8-sig",
        )


ensure_master_files()


def read_csv_flexible(path: Path, columns: list[str]) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns)
    try:
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    except UnicodeDecodeError:
        df = pd.read_csv(path, dtype=str, encoding="big5", errors="ignore").fillna("")
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    return df[columns].copy()


def load_rule_master() -> pd.DataFrame:
    df = read_csv_flexible(DATA_DIR / "rule_master.csv", ["Prefix", "Product Type", "Customer Code Logic", "Enabled"])
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    df["Prefix"] = df["Prefix"].str.upper()
    df["Enabled"] = df["Enabled"].str.upper().replace("", "Y")
    return df[df["Enabled"].isin(["Y", "YES", "TRUE", "1"])]


def load_material_rule_master() -> pd.DataFrame:
    df = read_csv_flexible(DATA_DIR / "material_rule_master.csv", ["Material Number", "Product Type", "Customer", "Enabled"])
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip()
    df["Material Number"] = df["Material Number"].str.upper()
    df["Enabled"] = df["Enabled"].str.upper().replace("", "Y")
    return df[df["Enabled"].isin(["Y", "YES", "TRUE", "1"])]


def load_product_series_master() -> pd.DataFrame:
    path = DATA_DIR / "product_series_master.csv"
    if not path.exists():
        return pd.DataFrame(columns=["Plant", "Product series", "產品類型", "客戶代碼", "客戶名稱"])
    df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    for col in ["Plant", "Product series", "產品類型", "客戶代碼", "客戶名稱"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()
    df["Plant"] = df["Plant"].str.replace(r"\.0$", "", regex=True)
    df["Product series"] = df["Product series"].str.upper().str.replace(r"\s+", "", regex=True)
    return df


def build_masters() -> dict:
    rule_master = load_rule_master()
    material_master = load_material_rule_master()
    series_master = load_product_series_master()

    material_by_number = {}
    for _, row in material_master.iterrows():
        mn = row["Material Number"]
        if mn:
            material_by_number[mn] = row

    prefix_rules = []
    for _, row in rule_master.iterrows():
        prefix = row["Prefix"]
        if prefix:
            prefix_rules.append(row)
    prefix_rules.sort(key=lambda r: len(str(r["Prefix"])), reverse=True)

    by_plant_series = {}
    by_series = {}
    for _, row in series_master.iterrows():
        series = row.get("Product series", "")
        plant = str(row.get("Plant", "") or "").replace(".0", "")
        if series:
            by_plant_series[(plant, series)] = row
            if series not in by_series:
                by_series[series] = row

    return {
        "material_by_number": material_by_number,
        "prefix_rules": prefix_rules,
        "by_plant_series": by_plant_series,
        "by_series": by_series,
    }


def find_col(df: pd.DataFrame, aliases: list[str]) -> Optional[str]:
    normalized = {str(c).strip().lower(): c for c in df.columns}
    for alias in aliases:
        key = alias.strip().lower()
        if key in normalized:
            return normalized[key]
    return None


def parse_product_series(description: object) -> tuple[str, str]:
    if pd.isna(description):
        return "", "Material description 空白"
    text = str(description).upper().replace("_", ",").replace(";", ",")
    parts = text.split(",")
    for idx, part in enumerate(parts, start=1):
        compact = re.sub(r"\s+", "", part)
        compact = re.sub(r"[^A-Z0-9].*$", "", compact)
        if not compact or compact in GENERIC_WORDS:
            continue
        match = SERIES_RE.match(compact)
        if match:
            series = match.group(1)
            return series, "第1段有效產品系列" if idx == 1 else f"前段非系列，取第{idx}段有效產品系列"
    return "", "未找到符合產品系列格式的英數碼"


def get_customer(series: str, logic: str) -> tuple[str, str]:
    series = str(series or "").upper()
    logic = str(logic or "").upper().strip()
    if logic == "NB_3RD_CHAR" and len(series) >= 3:
        code = series[2]
        return code, NB_CUSTOMERS.get(code, "")
    if logic == "DT_3_4_CHAR" and len(series) >= 4:
        code = series[2:4]
        return code, DT_CUSTOMERS.get(code, "")
    return "", ""


def classify_by_material_rule(material_number: object, masters: dict) -> dict:
    mn = str(material_number or "").upper().strip()
    row = masters["material_by_number"].get(mn)
    if row is None:
        return {}
    product_type = str(row.get("Product Type", "") or "").strip()
    customer = str(row.get("Customer", "") or "").strip()
    return {
        "產品類型": product_type,
        "客戶代碼": "",
        "客戶名稱": customer,
        "判斷來源": "Material Number Rule",
        "規則判定結果": "符合" if product_type else "待補產品分類",
        "命中規則": f"Material Number={mn}",
    }


def classify_by_naming_rule(series: str, masters: dict) -> dict:
    series = str(series or "").upper().strip()
    if not series:
        return {}
    for row in masters["prefix_rules"]:
        prefix = str(row.get("Prefix", "") or "").upper().strip()
        if prefix and series.startswith(prefix):
            product_type = str(row.get("Product Type", "") or "").strip()
            logic = str(row.get("Customer Code Logic", "") or "").strip()
            code, customer = get_customer(series, logic)
            return {
                "產品類型": product_type,
                "客戶代碼": code,
                "客戶名稱": customer,
                "判斷來源": "Naming Rule",
                "規則判定結果": "符合" if product_type else "待補產品分類",
                "命中規則": f"Prefix={prefix}",
            }
    return {}


def classify_by_series_master(plant: object, series: str, masters: dict) -> dict:
    plant_str = str(plant or "").strip().replace(".0", "")
    series = str(series or "").upper().strip()
    if not series:
        return {}
    row = masters["by_plant_series"].get((plant_str, series)) or masters["by_series"].get(series)
    if row is None:
        return {}
    product_type = str(row.get("產品類型", "") or "").strip()
    code = str(row.get("客戶代碼", "") or "").strip()
    customer = str(row.get("客戶名稱", "") or "").strip()
    return {
        "產品類型": product_type,
        "客戶代碼": code,
        "客戶名稱": customer,
        "判斷來源": "Product Series Master",
        "規則判定結果": "符合" if product_type else "待補產品分類",
        "命中規則": f"Product series={series}",
    }


def classify(material_number: object, series: str, plant: object, masters: dict) -> dict:
    for fn in (
        lambda: classify_by_material_rule(material_number, masters),
        lambda: classify_by_naming_rule(series, masters),
        lambda: classify_by_series_master(plant, series, masters),
    ):
        result = fn()
        if result and result.get("產品類型"):
            result["Is_WIP"] = "Y" if result.get("產品類型") == "WIP" else "N"
            return result

    # 不在分類規則 / master 中者視為半品 WIP
    return {
        "產品類型": "WIP",
        "客戶代碼": "",
        "客戶名稱": "",
        "判斷來源": "WIP",
        "規則判定結果": "WIP",
        "命中規則": "No rule matched → WIP",
        "Is_WIP": "Y",
    }


def process_file(path: Path, year: Optional[int]) -> tuple[Path, dict]:
    masters = build_masters()
    df = pd.read_excel(path, dtype=str)
    cols = {key: find_col(df, aliases) for key, aliases in REQUIRED_ALIASES.items()}
    missing = [key for key, col in cols.items() if col is None]
    if missing:
        raise ValueError(f"缺少必要欄位：{', '.join(missing)}")

    out = pd.DataFrame()
    out["Order"] = df[cols["order"]]
    out["Plant"] = df[cols["plant"]].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
    out["Material Number"] = df[cols["material_number"]].astype(str).str.strip()
    out["Material description"] = df[cols["material_description"]]
    out["Delivered quantity"] = pd.to_numeric(df[cols["delivered_quantity"]], errors="coerce").fillna(0)
    out["Actual finish date"] = pd.to_datetime(df[cols["finish_date"]], errors="coerce")
    out["Year"] = out["Actual finish date"].dt.year

    if year:
        out = out[out["Year"] == int(year)].copy()

    parsed = out["Material description"].apply(parse_product_series)
    out["Product series"] = parsed.apply(lambda x: x[0])
    out["解析說明"] = parsed.apply(lambda x: x[1])

    classified = out.apply(lambda r: classify(r["Material Number"], r["Product series"], r["Plant"], masters), axis=1)
    out["產品類型"] = classified.apply(lambda x: x.get("產品類型", ""))
    out["客戶代碼"] = classified.apply(lambda x: x.get("客戶代碼", ""))
    out["客戶名稱"] = classified.apply(lambda x: x.get("客戶名稱", ""))
    out["判斷來源"] = classified.apply(lambda x: x.get("判斷來源", ""))
    out["規則判定結果"] = classified.apply(lambda x: x.get("規則判定結果", ""))
    out["命中規則"] = classified.apply(lambda x: x.get("命中規則", ""))
    out["Is_WIP"] = classified.apply(lambda x: x.get("Is_WIP", "N"))

    group_cols = [
        "Year", "Plant", "Material Number", "Material description", "Product series",
        "產品類型", "客戶代碼", "客戶名稱", "判斷來源", "Is_WIP"
    ]
    annual = (
        out.groupby(group_cols, dropna=False, as_index=False)["Delivered quantity"]
        .sum()
        .rename(columns={"Delivered quantity": "年度生產量"})
        .sort_values(["Plant", "Material Number"])
    )

    type_summary = (
        out.groupby(["Year", "Plant", "產品類型", "Is_WIP"], dropna=False, as_index=False)["Delivered quantity"]
        .sum()
        .rename(columns={"Delivered quantity": "年度生產量"})
        .sort_values(["Plant", "產品類型"])
    )

    customer_summary = (
        out.groupby(["Year", "Plant", "客戶名稱"], dropna=False, as_index=False)["Delivered quantity"]
        .sum()
        .rename(columns={"Delivered quantity": "年度生產量"})
        .sort_values(["Plant", "客戶名稱"])
    )

    source_summary = (
        out.groupby(["判斷來源", "規則判定結果"], dropna=False, as_index=False)["Delivered quantity"]
        .agg(筆數="count", 生產量="sum")
    )

    wip = out[out["Is_WIP"] == "Y"].copy()

    file_id = uuid.uuid4().hex[:10]
    output_path = OUTPUT_DIR / f"年度產品產量與分類結果_v3_{year or 'ALL'}_{file_id}.xlsx"
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="工單明細_已分類")
        annual.to_excel(writer, index=False, sheet_name="Plant_Material年度產量")
        type_summary.to_excel(writer, index=False, sheet_name="Plant_產品類型年度產量")
        customer_summary.to_excel(writer, index=False, sheet_name="Plant_客戶年度產量")
        source_summary.to_excel(writer, index=False, sheet_name="判斷來源摘要")
        wip.to_excel(writer, index=False, sheet_name="WIP清單")
        for sheet in writer.book.worksheets:
            sheet.freeze_panes = "A2"
            for col in sheet.columns:
                max_len = 12
                letter = col[0].column_letter
                for cell in col[:1000]:
                    max_len = max(max_len, len(str(cell.value or "")) + 2)
                sheet.column_dimensions[letter].width = min(max_len, 45)

    summary = {
        "rows": int(len(out)),
        "annual_rows": int(len(annual)),
        "total_qty": float(out["Delivered quantity"].sum()),
        "wip_rows": int(len(wip)),
        "output_filename": output_path.name,
        "year": year or "ALL",
    }
    return output_path, summary


def save_uploaded_rule(file_path: Path) -> int:
    """匯入使用者維護的 Material Number Rule。支援 xlsx / csv。"""
    if file_path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        df = pd.read_excel(file_path, dtype=str).fillna("")
    else:
        df = pd.read_csv(file_path, dtype=str, encoding="utf-8-sig").fillna("")

    # 支援中英文欄位名稱
    rename_map = {}
    for c in df.columns:
        key = str(c).strip().lower()
        if key in ["material number", "成品料號", "料號"]:
            rename_map[c] = "Material Number"
        elif key in ["product type", "產品分類", "產品類型"]:
            rename_map[c] = "Product Type"
        elif key in ["customer", "客戶", "客戶名稱"]:
            rename_map[c] = "Customer"
        elif key in ["enabled", "啟用"]:
            rename_map[c] = "Enabled"
    df = df.rename(columns=rename_map)
    for col in ["Material Number", "Product Type", "Customer", "Enabled"]:
        if col not in df.columns:
            df[col] = ""
    df = df[["Material Number", "Product Type", "Customer", "Enabled"]].copy()
    df["Material Number"] = df["Material Number"].astype(str).str.strip().str.upper()
    df["Product Type"] = df["Product Type"].astype(str).str.strip()
    df["Customer"] = df["Customer"].astype(str).str.strip()
    df["Enabled"] = df["Enabled"].astype(str).str.strip().replace("", "Y")
    df = df[(df["Material Number"] != "") & (df["Product Type"] != "")]
    df = df.drop_duplicates(subset=["Material Number"], keep="last")
    df.to_csv(DATA_DIR / "material_rule_master.csv", index=False, encoding="utf-8-sig")
    return len(df)


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/process")
async def process(file: UploadFile = File(...), year: Optional[int] = Form(None)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "請上傳 Excel 檔案"}, status_code=400)
    saved = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    content = await file.read()
    saved.write_bytes(content)
    try:
        output_path, summary = process_file(saved, year)
    except Exception as exc:
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)
    return {"ok": True, "summary": summary, "download_url": f"/download/{output_path.name}"}


@app.post("/upload-material-rule")
async def upload_material_rule(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
        return JSONResponse({"ok": False, "message": "請上傳 Excel 或 CSV 檔案"}, status_code=400)
    saved = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    saved.write_bytes(await file.read())
    try:
        count = save_uploaded_rule(saved)
    except Exception as exc:
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)
    return {"ok": True, "count": count}


@app.get("/download/{filename}")
def download(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        return JSONResponse({"ok": False, "message": "檔案不存在"}, status_code=404)
    return FileResponse(path, filename=filename, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.get("/download-rule-master")
def download_rule_master():
    """Download Rule Master used by the current UI button."""
    path = DATA_DIR / "rule_master.csv"
    if not path.exists():
        ensure_master_files()
    return FileResponse(
        path,
        filename="rule_master.csv",
        media_type="text/csv; charset=utf-8",
    )


@app.get("/download-product-series-master")
def download_product_series_master():
    """Download Product Series Master used by the current UI button."""
    path = DATA_DIR / "product_series_master.csv"
    if not path.exists():
        ensure_master_files()
    return FileResponse(
        path,
        filename="product_series_master.csv",
        media_type="text/csv; charset=utf-8",
    )


@app.get("/download-material-rule")
def download_material_rule():
    path = DATA_DIR / "material_rule_master.csv"
    if not path.exists():
        ensure_master_files()
    return FileResponse(
        path,
        filename="material_rule_master.csv",
        media_type="text/csv; charset=utf-8",
    )


@app.get("/download-prefix-rule")
def download_prefix_rule():
    # Backward-compatible alias for older frontend versions.
    return download_rule_master()

# =========================================================
# Step 2 · Batch Data Formatting
# Step1 Output + Bulk Template -> Formatted Bulk File
# 方案 1：直接複製原始 Bulk Template，只覆蓋指定分頁儲存格內容
# =========================================================
from bulk_formatter import generate_product_activity_bulk_file
from bom_formatter import generate_raw_material_bulk_file


@app.post("/generate-bulk-file")
async def generate_bulk_file(
    step1_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    if not step1_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "Step 1 Output 請上傳 Excel 檔案"}, status_code=400)

    if not template_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "Bulk Template 請上傳 Excel 檔案"}, status_code=400)

    token = uuid.uuid4().hex[:10]

    step1_path = UPLOAD_DIR / f"step1_output_{token}_{step1_file.filename}"
    template_path = UPLOAD_DIR / f"bulk_template_{token}_{template_file.filename}"
    output_path = OUTPUT_DIR / f"formatted_product_activity_data_bulk_create_{token}.xlsx"

    step1_path.write_bytes(await step1_file.read())
    template_path.write_bytes(await template_file.read())

    try:
        summary = generate_product_activity_bulk_file(
            step1_output_path=step1_path,
            bulk_template_path=template_path,
            output_path=output_path,
        )
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)

    return {
        "ok": True,
        "message": "Bulk file generated successfully.",
        "summary": summary,
        "download_url": f"/download/{output_path.name}",
    }

# =========================================================
# Module 2 · BOM Expansion
# Standard BOM + Raw Material Bulk Template -> Raw Material Bulk
# =========================================================
@app.post("/process-bom-expansion")
async def process_bom_expansion(
    bom_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
    parent_col: str = Form(""),
    component_col: str = Form(""),
    qty_col: str = Form(""),
    unit_col: str = Form(""),
    description_col: str = Form(""),
    material_group_col: str = Form(""),
    valid_from_col: str = Form(""),
):
    if not bom_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse(
            {"ok": False, "message": "Standard BOM 請上傳 Excel 檔案"},
            status_code=400,
        )

    if not template_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse(
            {"ok": False, "message": "Raw Material Bulk Template 請上傳 Excel 檔案"},
            status_code=400,
        )

    token = uuid.uuid4().hex[:10]

    bom_path = UPLOAD_DIR / f"standard_bom_{token}_{bom_file.filename}"
    template_path = UPLOAD_DIR / f"raw_material_template_{token}_{template_file.filename}"
    output_path = OUTPUT_DIR / f"raw_material_activity_data_bulk_{token}.xlsx"

    bom_path.write_bytes(await bom_file.read())
    template_path.write_bytes(await template_file.read())

    mapping = {
        "parent_col": parent_col,
        "component_col": component_col,
        "qty_col": qty_col,
        "unit_col": unit_col,
        "description_col": description_col,
        "material_group_col": material_group_col,
        "valid_from_col": valid_from_col,
    }

    try:
        summary = generate_raw_material_bulk_file(
            bom_path=bom_path,
            raw_material_template_path=template_path,
            output_path=output_path,
            mapping=mapping,
        )
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(
            {"ok": False, "message": str(exc)},
            status_code=400,
        )

    return {
        "ok": True,
        "message": "BOM Expansion completed successfully.",
        "summary": summary,
        "download_url": f"/download/{output_path.name}",
    }

