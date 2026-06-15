from __future__ import annotations

import re
import traceback
import uuid
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

app = FastAPI(title="Annual Output Platform v6", version="6.0.0")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

from bulk_formatter import generate_product_activity_bulk_file

# =========================================================
# v6 分類邏輯：
# 1. rule_master.csv 依 Priority 由小到大判斷
#    - Material Number Exact
#    - Material Number Prefix
#    - Description Contains
#    - Series Prefix
#    - Default
# 2. 若 rule_master 未命中，再查 product_series_master.csv
# 3. 若都未命中，寫 WIP
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
SERIES_RE = re.compile(r"([A-Z]{2}[A-Z0-9]*\d[A-Z0-9]*)")

# Product Series 抽取用：遇到這些描述詞時要截斷，避免把 Assy/Touchpad/Module 等描述文字吃進系列名稱
SERIES_STOP_WORDS = [
    "ASSY", "ASSEMBLY", "MODULE", "TOUCHPAD", "TOUCH", "HAPTIC", "FORCE",
    "KEYBOARD", "MOUSE", "TRIBUTO", "LIGHTGRAY", "NOKEYCAP", "NOKEYCA",
    "NOKEY", "BLANK", "NEW", "EURO", "NOEURO",
]

REQUIRED_ALIASES = {
    "order": ["Order"],
    "plant": ["Plant", "Plant code"],
    "material_number": ["Material Number", "Material", "Product Material Number"],
    "material_description": ["Material description", "Material Description"],
    "delivered_quantity": ["Delivered quantity (GMEIN)", "Delivered quantity", "Delivered Quantity"],
    "finish_date": ["Actual finish date", "Actual Finish Date", "Finish date", "Actual finish"],
}

LABOR_ALIASES = {
    "order": ["Order Number", "Order", "Production Order", "Process order"],
    "plant": ["Plant", "Plant code"],
    "material_number": ["Material Number", "Material", "Product Material Number"],
    "labor_hr": ["Labor HR.Act", "Labor HR Act", "Labor HR.Act.", "Labor HR", "Labor Hour"],
    "foh_others": ["FOH-Others.Act", "FOH Others.Act", "FOH-Others Act", "FOH Others Act", "FOH-Others.Act.", "FOH Others"],
}

VALID_LABOR_MODES = {"labor_hr", "foh", "both"}

RULE_COLUMNS = [
    "Priority", "Rule Type", "Key", "Product Type", "Customer",
    "Customer Code Logic", "Is_WIP", "Enabled",
]

DEFAULT_RULE_MASTER = (
    "Priority,Rule Type,Key,Product Type,Customer,Customer Code Logic,Is_WIP,Enabled\n"
    "1,Material Number Prefix,851-,WIP,,,Y,Y\n"
    "2,Material Number Prefix,852-,WIP,,,Y,Y\n"
    "10,Series Prefix,SN,NB,,NB_3RD_CHAR,N,Y\n"
    "11,Series Prefix,FU,NB,,NB_3RD_CHAR,N,Y\n"
    "12,Series Prefix,SP,TP,,NB_3RD_CHAR,N,Y\n"
    "13,Series Prefix,SM,DT Mouse,,DT_3_4_CHAR,N,Y\n"
    "14,Series Prefix,SA,DT Accessory,,DT_3_4_CHAR,N,Y\n"
    "15,Description Contains,RECEIVER,DT Dongle,,,N,Y\n"
    "16,Series Prefix,SK,DT Keyboard,,DT_3_4_CHAR,N,Y\n"
    "17,Series Prefix,SB,DT Keyboard+Mouse,,DT_3_4_CHAR,N,Y\n"
    "18,Series Prefix,ST,DT Tablet Keyboard,,DT_3_4_CHAR,N,Y\n"
    "19,Description Contains,TOUCH PAD MODULE,TP,,,N,Y\n"
    "20,Description Contains,TOUCHPAD MODULE,TP,,,N,Y\n"
    "21,Series Prefix,SCMC,WIP,,,Y,Y\n"
    "90,Description Contains,ASSY,WIP,,,Y,Y\n"
    "999,Default,*,WIP,,,Y,Y\n"
)


def ensure_master_files() -> None:
    rule_path = DATA_DIR / "rule_master.csv"
    series_path = DATA_DIR / "product_series_master.csv"
    if not rule_path.exists():
        rule_path.write_text(DEFAULT_RULE_MASTER, encoding="utf-8-sig")
    if not series_path.exists():
        series_path.write_text("Plant,Product series,產品類型,客戶代碼,客戶名稱\n", encoding="utf-8-sig")


ensure_master_files()


def read_csv_flexible(path: Path, columns: Optional[list[str]] = None) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns or [])
    try:
        df = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
    except UnicodeDecodeError:
        df = pd.read_csv(path, dtype=str, encoding="big5").fillna("")
    if columns is not None:
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        df = df[columns].copy()
    return df


def load_rule_master() -> pd.DataFrame:
    df = read_csv_flexible(DATA_DIR / "rule_master.csv", RULE_COLUMNS)
    for c in RULE_COLUMNS:
        df[c] = df[c].astype(str).str.strip()
    df["Enabled"] = df["Enabled"].str.upper().replace("", "Y")
    df["Rule Type"] = df["Rule Type"].str.strip()
    df["Key"] = df["Key"].str.upper().str.strip()
    df["Priority_num"] = pd.to_numeric(df["Priority"], errors="coerce").fillna(9999)
    df = df[df["Enabled"].isin(["Y", "YES", "TRUE", "1"])]
    return df.sort_values(["Priority_num", "Rule Type", "Key"], kind="stable").reset_index(drop=True)


def load_product_series_master() -> pd.DataFrame:
    path = DATA_DIR / "product_series_master.csv"
    df = read_csv_flexible(path)
    for col in ["Plant", "Product series", "產品類型", "客戶代碼", "客戶名稱"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()
    df["Plant"] = df["Plant"].str.replace(r"\.0$", "", regex=True)
    df["Product series"] = df["Product series"].str.upper().str.replace(r"\s+", "", regex=True)
    return df[["Plant", "Product series", "產品類型", "客戶代碼", "客戶名稱"]].copy()


def build_masters() -> dict:
    rule_master = load_rule_master()
    series_master = load_product_series_master()

    by_plant_series: dict[tuple[str, str], pd.Series] = {}
    by_series: dict[str, pd.Series] = {}
    for _, row in series_master.iterrows():
        series = str(row.get("Product series", "") or "").upper().strip()
        plant = str(row.get("Plant", "") or "").replace(".0", "").strip()
        if series:
            by_plant_series[(plant, series)] = row
            if series not in by_series:
                by_series[series] = row

    return {
        "rules": rule_master,
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


def normalize_order_key(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.lower() in ["", "nan", "none", "nat"]:
        return ""
    text = re.sub(r"\.0$", "", text)
    text = re.sub(r"\s+", "", text)
    if text.isdigit():
        text = text.lstrip("0") or "0"
    return text


def get_series_prefixes(masters: Optional[dict] = None) -> list[str]:
    """從 rule_master.csv 動態讀取 Series Prefix，讓新增 SN/SP/SM/SK/FU... 不用再改 Python。"""
    prefixes: list[str] = []
    try:
        rules = masters.get("rules") if masters else load_rule_master()
        if isinstance(rules, pd.DataFrame):
            for _, row in rules.iterrows():
                rule_type = normalize_rule_type(row.get("Rule Type", ""))
                key = str(row.get("Key", "") or "").upper().strip()
                if rule_type in ["series prefix", "product series prefix", "產品系列前綴", "系列前綴"] and key and key not in ["*", "DEFAULT"]:
                    prefixes.append(re.escape(key))
    except Exception:
        prefixes = []

    # fallback：避免 rule_master 空白時完全抓不到
    if not prefixes:
        prefixes = ["SN", "SP", "SM", "SK", "SB", "ST", "SA", "FU"]

    return sorted(set(prefixes), key=len, reverse=True)


def trim_series_candidate(candidate: str) -> str:
    """把候選字串在描述詞前截斷，例如 SP2B20XF0ASSYHAPTIC... -> SP2B20XF0。"""
    candidate = re.sub(r"[^A-Z0-9].*$", "", str(candidate or "").upper())
    cut_positions = [pos for word in SERIES_STOP_WORDS if (pos := candidate.find(word)) > 0]
    if cut_positions:
        candidate = candidate[:min(cut_positions)]
    return candidate


def _series_candidate_from_text(text: str, pattern: re.Pattern) -> Optional[str]:
    """在一段文字中尋找第一個符合 Series Prefix 的候選值，並套用描述詞截斷。"""
    if not text:
        return None

    # 移除空白與非英數符號，讓 CHSN4396BL1 / AssyCH SN5372BL 都能被抓到。
    # 標點符號的優先判斷在 parse_product_series() 的第一階段完成，這裡只處理單一區段。
    compact_text = re.sub(r"[^A-Z0-9]+", "", str(text).upper())

    for match in pattern.finditer(compact_text):
        candidate = trim_series_candidate(match.group(0))
        if not candidate or candidate in GENERIC_WORDS:
            continue
        if not re.search(r"\d", candidate):
            continue
        if len(candidate) < 5:
            continue
        return candidate

    return None


def parse_product_series(description: object, masters: Optional[dict] = None) -> tuple[str, str]:
    """
    v6 Product Series 抽取邏輯：
    1. 保留標點符號判斷：先依逗號 / 分號 / 底線切段，在每個區段內搜尋 Series。
    2. 若標點切段找不到，再用整段 Material description 做 Regex Prefix Search。
    3. Prefix 由 rule_master.csv 的 Series Prefix 動態讀取，不用寫死在 Python。
    4. 遇到 ASSY / TOUCHPAD / HAPTIC / MODULE / BLANK / NEW 等描述詞自動截斷。

    可處理：
    - AssyCH SN5372BL,110K,JPBlank,nokeycapNEW -> SN5372BL
    - Assy,CHSN4396BL1,UK,106KBlank,nokeycaNEW -> SN4396BL1
    - SP2B20XF0AssyHapticForce Touchpad module -> SP2B20XF0
    - SP2B20XF0Assy Touchpad TributoLightGray -> SP2B20XF0
    """
    if pd.isna(description):
        return "", "Material description 空白"

    raw_text = str(description).upper()
    prefixes = get_series_prefixes(masters)
    prefix_pattern = "|".join(prefixes)
    pattern = re.compile(rf"({prefix_pattern})[A-Z0-9]{{3,40}}")

    # 第一階段：保留標點符號作為重要邊界，先分段判斷。
    # 這可以避免 SN5372BL,110K 被抓成 SN5372BL110K。
    segments = [seg for seg in re.split(r"[,;_]+", raw_text) if str(seg).strip()]
    for idx, segment in enumerate(segments, start=1):
        candidate = _series_candidate_from_text(segment, pattern)
        if candidate:
            if idx == 1:
                return candidate, "標點切段解析產品系列"
            return candidate, f"標點切段解析產品系列：取第{idx}段"

    # 第二階段：若資料完全沒有標點，改用整段搜尋。
    candidate = _series_candidate_from_text(raw_text, pattern)
    if candidate:
        return candidate, "全文 Regex Prefix Search 解析產品系列"

    # fallback：若 rule_master 沒有涵蓋 prefix，仍保留舊邏輯，但同樣套用截斷
    parts = raw_text.replace("_", ",").replace(";", ",").split(",")
    for idx, part in enumerate(parts, start=1):
        compact = re.sub(r"\s+", "", part)
        compact = trim_series_candidate(compact)
        if not compact or compact in GENERIC_WORDS:
            continue
        match = SERIES_RE.match(compact)
        if match:
            series = trim_series_candidate(match.group(0))
            if series and re.search(r"\d", series) and len(series) >= 5:
                return series, "fallback：逗號分段解析產品系列" if idx == 1 else f"fallback：前段非系列，取第{idx}段"

    return "", "未找到符合產品系列格式的英數碼"


def get_customer(series: str, logic: str, customer_override: str = "") -> tuple[str, str]:
    if customer_override:
        return "", customer_override
    series = str(series or "").upper()
    logic = str(logic or "").upper().strip()
    if logic == "NB_3RD_CHAR" and len(series) >= 3:
        code = series[2]
        return code, NB_CUSTOMERS.get(code, "")
    if logic == "DT_3_4_CHAR" and len(series) >= 4:
        code = series[2:4]
        return code, DT_CUSTOMERS.get(code, "")
    return "", ""


def normalize_rule_type(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def rule_matches(rule_type: str, key: str, material_number: str, description: str, series: str) -> bool:
    rt = normalize_rule_type(rule_type)
    key_u = str(key or "").upper().strip()
    if not key_u:
        return False

    # 支援英文與常見中文寫法
    if rt in ["material number exact", "material exact", "成品料號", "成品料號完全符合"]:
        return material_number == key_u
    if rt in ["material number prefix", "material prefix", "成品料號前綴", "料號前綴"]:
        return material_number.startswith(key_u)
    if rt in ["description contains", "material description contains", "描述包含", "品名包含"]:
        return key_u in description
    if rt in ["series prefix", "product series prefix", "產品系列前綴", "系列前綴"]:
        return series.startswith(key_u)
    if rt in ["series exact", "product series exact", "產品系列", "產品系列完全符合"]:
        return series == key_u
    if rt in ["default", "預設"]:
        return key_u in ["*", "DEFAULT", ""]
    return False


def classify_by_rule_master(material_number: object, description: object, series: str, masters: dict) -> dict:
    material_number_u = str(material_number or "").upper().strip()
    description_u = str(description or "").upper().strip()
    series_u = str(series or "").upper().strip()

    rules: pd.DataFrame = masters["rules"]
    for _, row in rules.iterrows():
        rule_type = str(row.get("Rule Type", "") or "").strip()
        key = str(row.get("Key", "") or "").strip()

        # Default 規則只作為文件與下載範本保留；實際 Default WIP 放在 Product Series Master 之後。
        if normalize_rule_type(rule_type) in ["default", "預設"]:
            continue

        if not rule_matches(rule_type, key, material_number_u, description_u, series_u):
            continue

        product_type = str(row.get("Product Type", "") or "").strip()
        customer_override = str(row.get("Customer", "") or "").strip()
        logic = str(row.get("Customer Code Logic", "") or "").strip()
        code, customer = get_customer(series_u, logic, customer_override)
        is_wip = str(row.get("Is_WIP", "") or "").upper().strip()
        if not is_wip:
            is_wip = "Y" if product_type.upper() == "WIP" else "N"

        return {
            "產品類型": product_type,
            "客戶代碼": code,
            "客戶名稱": customer,
            "判斷來源": "Rule Master",
            "規則判定結果": "符合" if product_type else "待補產品分類",
            "命中規則": f"{rule_type}={key}",
            "Is_WIP": "Y" if is_wip in ["Y", "YES", "TRUE", "1"] else "N",
        }
    return {}


def classify_by_series_master(plant: object, series: str, masters: dict) -> dict:
    plant_str = str(plant or "").strip().replace(".0", "")
    series_u = str(series or "").upper().strip()
    if not series_u:
        return {}

    # 注意：不能使用 row_a or row_b，Pandas Series 不能直接做布林判斷
    row = masters["by_plant_series"].get((plant_str, series_u))
    if row is None:
        row = masters["by_series"].get(series_u)
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
        "命中規則": f"Product series={series_u}",
        "Is_WIP": "Y" if product_type.upper() == "WIP" else "N",
    }


def classify(material_number: object, description: object, series: str, plant: object, masters: dict) -> dict:
    result = classify_by_rule_master(material_number, description, series, masters)
    if result.get("產品類型"):
        return result

    result = classify_by_series_master(plant, series, masters)
    if result.get("產品類型"):
        return result

    return {
        "產品類型": "WIP",
        "客戶代碼": "",
        "客戶名稱": "",
        "判斷來源": "Default WIP",
        "規則判定結果": "WIP",
        "命中規則": "No rule matched → WIP",
        "Is_WIP": "Y",
    }


def load_production_dataframe(paths: list[Path]) -> pd.DataFrame:
    """Load one or multiple production quantity work order files."""
    frames: list[pd.DataFrame] = []
    errors: list[str] = []

    for path in paths:
        df = pd.read_excel(path, dtype=str)
        cols = {key: find_col(df, aliases) for key, aliases in REQUIRED_ALIASES.items()}
        missing = [key for key, col in cols.items() if col is None]
        if missing:
            errors.append(f"{path.name} 缺少必要欄位：{', '.join(missing)}")
            continue

        part = pd.DataFrame(index=df.index)
        part["Source file"] = path.name
        part["Order"] = df[cols["order"]]
        part["Order Merge Key"] = part["Order"].apply(normalize_order_key)
        part["Plant"] = df[cols["plant"]].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        part["Material Number"] = df[cols["material_number"]].astype(str).str.strip()
        part["Material description"] = df[cols["material_description"]]
        part["Delivered quantity"] = pd.to_numeric(df[cols["delivered_quantity"]], errors="coerce").fillna(0)
        part["Actual finish date"] = pd.to_datetime(df[cols["finish_date"]], errors="coerce")
        part["Year"] = part["Actual finish date"].dt.year
        frames.append(part)

    if errors:
        raise ValueError("；".join(errors))
    if not frames:
        raise ValueError("沒有可處理的生產數量工單資料")

    return pd.concat(frames, ignore_index=True)


def load_labor_dataframe(paths: list[Path], labor_mode: str = "both") -> pd.DataFrame:
    """Load one or multiple production labor work order files."""
    columns = [
        "Order Merge Key", "Order", "Plant", "Material Number",
        "Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"
    ]
    if not paths:
        return pd.DataFrame(columns=columns)

    mode = str(labor_mode or "both").strip().lower()
    if mode not in VALID_LABOR_MODES:
        mode = "both"

    frames: list[pd.DataFrame] = []
    errors: list[str] = []

    for path in paths:
        df = pd.read_excel(path, dtype=str)
        cols = {key: find_col(df, aliases) for key, aliases in LABOR_ALIASES.items()}

        missing = [key for key in ["order", "labor_hr", "foh_others"] if cols.get(key) is None]
        if missing:
            errors.append(f"{path.name} 缺少必要工時欄位：{', '.join(missing)}")
            continue

        part = pd.DataFrame(index=df.index)
        part["Labor Source file"] = path.name
        part["Order"] = df[cols["order"]].astype(str).str.strip()
        part["Order Merge Key"] = part["Order"].apply(normalize_order_key)

        if cols.get("plant") is not None:
            part["Plant"] = df[cols["plant"]].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
        else:
            part["Plant"] = ""

        if cols.get("material_number") is not None:
            part["Material Number"] = df[cols["material_number"]].astype(str).str.strip()
        else:
            part["Material Number"] = ""

        part["Labor HR.Act"] = pd.to_numeric(df[cols["labor_hr"]], errors="coerce").fillna(0)
        part["FOH-Others.Act"] = pd.to_numeric(df[cols["foh_others"]], errors="coerce").fillna(0)

        if mode == "labor_hr":
            part["Selected Hours"] = part["Labor HR.Act"]
        elif mode == "foh":
            part["Selected Hours"] = part["FOH-Others.Act"]
        else:
            part["Selected Hours"] = part["Labor HR.Act"] + part["FOH-Others.Act"]

        frames.append(part)

    if errors:
        raise ValueError("；".join(errors))
    if not frames:
        return pd.DataFrame(columns=columns)

    labor = pd.concat(frames, ignore_index=True)
    labor = labor[(labor["Order Merge Key"] != "") | (labor["Material Number"] != "")].copy()

    return (
        labor.groupby(["Order Merge Key", "Plant", "Material Number"], dropna=False, as_index=False)
        .agg({
            "Order": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip()))),
            "Labor HR.Act": "sum",
            "FOH-Others.Act": "sum",
            "Selected Hours": "sum",
            "Labor Source file": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip())))
        })
        .rename(columns={"Labor Source file": "Labor Source files"})
    )


def attach_labor_hours(out: pd.DataFrame, labor: pd.DataFrame) -> pd.DataFrame:
    """Attach labor hours to production output.

    Matching priority:
    1. Order / Order Number after normalization
    2. Plant + Material Number fallback
    """
    out = out.copy()

    if "Order Merge Key" not in out.columns:
        out["Order Merge Key"] = out["Order"].apply(normalize_order_key)

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"]:
        out[col] = 0 if col != "Labor Source files" else ""

    if labor is None or labor.empty:
        return out

    labor = labor.copy()
    if "Order Merge Key" not in labor.columns:
        labor["Order Merge Key"] = labor["Order"].apply(normalize_order_key)

    order_labor = (
        labor[labor["Order Merge Key"].astype(str).str.strip() != ""]
        .groupby(["Order Merge Key"], dropna=False, as_index=False)
        .agg({
            "Labor HR.Act": "sum",
            "FOH-Others.Act": "sum",
            "Selected Hours": "sum",
            "Labor Source files": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip())))
        })
    )

    if not order_labor.empty:
        out = out.merge(order_labor, on="Order Merge Key", how="left", suffixes=("", "_labor"))
        for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
            labor_col = f"{col}_labor"
            if labor_col in out.columns:
                out[col] = pd.to_numeric(out[labor_col], errors="coerce").fillna(0)
                out = out.drop(columns=[labor_col])
        if "Labor Source files_labor" in out.columns:
            out["Labor Source files"] = out["Labor Source files_labor"].fillna("").astype(str)
            out = out.drop(columns=["Labor Source files_labor"])

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    fallback_mask = out["Selected Hours"].eq(0)
    pm_labor = (
        labor[
            (labor["Plant"].astype(str).str.strip() != "")
            & (labor["Material Number"].astype(str).str.strip() != "")
        ]
        .groupby(["Plant", "Material Number"], dropna=False, as_index=False)
        .agg({
            "Labor HR.Act": "sum",
            "FOH-Others.Act": "sum",
            "Selected Hours": "sum",
            "Labor Source files": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip())))
        })
    )

    if fallback_mask.any() and not pm_labor.empty:
        fallback = out.loc[fallback_mask, ["Plant", "Material Number"]].merge(
            pm_labor,
            on=["Plant", "Material Number"],
            how="left",
        )
        for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
            out.loc[fallback_mask, col] = pd.to_numeric(fallback[col], errors="coerce").fillna(0).to_numpy()
        out.loc[fallback_mask, "Labor Source files"] = fallback["Labor Source files"].fillna("").astype(str).to_numpy()

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    out["Labor Source files"] = out["Labor Source files"].fillna("").astype(str)
    return out


def process_files(
    paths: list[Path],
    year: Optional[int],
    labor_paths: Optional[list[Path]] = None,
    labor_mode: str = "both",
) -> tuple[Path, dict]:
    masters = build_masters()
    out = load_production_dataframe(paths)
    labor = load_labor_dataframe(labor_paths or [], labor_mode)
    out = attach_labor_hours(out, labor)

    if year:
        out = out[out["Year"] == int(year)].copy()

    parsed = out["Material description"].apply(lambda desc: parse_product_series(desc, masters))
    out["Product series"] = parsed.apply(lambda x: x[0])
    out["解析說明"] = parsed.apply(lambda x: x[1])

    classified = out.apply(
        lambda r: classify(
            r["Material Number"],
            r["Material description"],
            r["Product series"],
            r["Plant"],
            masters,
        ),
        axis=1,
    )

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
        out.groupby(group_cols, dropna=False, as_index=False)
        .agg({
            "Delivered quantity": "sum",
            "Selected Hours": "sum",
        })
        .rename(columns={
            "Delivered quantity": "年度生產量",
            "Selected Hours": "年度工時",
        })
        .sort_values(["Plant", "Material Number"])
    )

    plant_qty_total = annual.groupby(["Year", "Plant"], dropna=False)["年度生產量"].transform("sum")
    plant_hour_total = annual.groupby(["Year", "Plant"], dropna=False)["年度工時"].transform("sum")
    annual["生產數量占比(%)"] = 0.0
    annual["生產工時占比(%)"] = 0.0

    qty_mask = plant_qty_total.ne(0)
    hour_mask = plant_hour_total.ne(0)

    annual.loc[qty_mask, "生產數量占比(%)"] = (
        annual.loc[qty_mask, "年度生產量"] / plant_qty_total.loc[qty_mask] * 100
    )
    annual.loc[hour_mask, "生產工時占比(%)"] = (
        annual.loc[hour_mask, "年度工時"] / plant_hour_total.loc[hour_mask] * 100
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
        out.groupby(["判斷來源", "規則判定結果", "命中規則"], dropna=False, as_index=False)["Delivered quantity"]
        .agg(筆數="count", 生產量="sum")
    )

    file_summary = (
        out.groupby(["Source file"], dropna=False, as_index=False)["Delivered quantity"]
        .agg(筆數="count", 生產量="sum")
        .sort_values(["Source file"])
    )

    wip = out[out["Is_WIP"] == "Y"].copy()

    file_id = uuid.uuid4().hex[:10]
    output_path = OUTPUT_DIR / f"年度產品產量與分類結果_v6_{year or 'ALL'}_{file_id}.xlsx"

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out.to_excel(writer, index=False, sheet_name="工單明細_已分類")
        annual.to_excel(writer, index=False, sheet_name="Plant_Material年度產量")
        type_summary.to_excel(writer, index=False, sheet_name="Plant_產品類型年度產量")
        customer_summary.to_excel(writer, index=False, sheet_name="Plant_客戶年度產量")
        source_summary.to_excel(writer, index=False, sheet_name="判斷來源摘要")
        file_summary.to_excel(writer, index=False, sheet_name="來源檔案摘要")
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
        "files": int(len(paths)),
        "labor_files": int(len(labor_paths or [])),
        "rows": int(len(out)),
        "annual_rows": int(len(annual)),
        "total_qty": float(out["Delivered quantity"].sum()),
        "total_hours": float(out["Selected Hours"].sum()) if "Selected Hours" in out.columns else 0.0,
        "wip_rows": int(len(wip)),
        "output_filename": output_path.name,
        "year": year or "ALL",
    }

    return output_path, summary


def process_file(path: Path, year: Optional[int]) -> tuple[Path, dict]:
    return process_files([path], year, None, "both")

def save_uploaded_rule(file_path: Path) -> int:
    if file_path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        df = pd.read_excel(file_path, dtype=str).fillna("")
    else:
        df = pd.read_csv(file_path, dtype=str, encoding="utf-8-sig").fillna("")
    df = normalize_rule_upload(df)
    df.to_csv(DATA_DIR / "rule_master.csv", index=False, encoding="utf-8-sig")
    return len(df)


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/process")
async def process(request: Request):
    """Step 1 processing endpoint.

    Manual multipart reader:
    - avoids FastAPI 422 validation
    - accepts file objects from ANY form key
    - classifies labor files by key name containing labor/hour/worktime
    """
    try:
        form = await request.form()
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(
            {"ok": False, "message": f"無法讀取上傳表單：{exc}"},
            status_code=400,
        )

    available_keys = list(form.keys())

    def is_upload_file_like(item) -> bool:
        return bool(getattr(item, "filename", None)) and hasattr(item, "read")

    upload_files = []
    labor_uploads = []

    for item in form.getlist("files") + form.getlist("file"):
        if is_upload_file_like(item):
            upload_files.append(item)

    for item in form.getlist("labor_files") + form.getlist("labor_file"):
        if is_upload_file_like(item):
            labor_uploads.append(item)

    seen_ids = {id(x) for x in upload_files + labor_uploads}
    for key in available_keys:
        for item in form.getlist(key):
            if not is_upload_file_like(item) or id(item) in seen_ids:
                continue

            key_lower = str(key).lower()
            if "labor" in key_lower or "hour" in key_lower or "worktime" in key_lower:
                labor_uploads.append(item)
            else:
                upload_files.append(item)
            seen_ids.add(id(item))

    if not upload_files:
        detail = []
        for key in available_keys:
            for item in form.getlist(key):
                detail.append({
                    "key": str(key),
                    "type": type(item).__name__,
                    "filename": str(getattr(item, "filename", "")),
                })
        return JSONResponse(
            {
                "ok": False,
                "message": (
                    "請至少上傳一個 Excel 生產數量工單檔案。"
                    f" 後端收到的表單欄位：{available_keys}；欄位內容摘要：{detail}"
                ),
            },
            status_code=400,
        )

    labor_mode = str(form.get("labor_mode") or "both").strip()
    year = form.get("year")

    saved_paths: list[Path] = []
    for upload in upload_files:
        filename = str(getattr(upload, "filename", "") or "")
        if not filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse({"ok": False, "message": f"{filename} 不是 Excel 檔案"}, status_code=400)

        saved = UPLOAD_DIR / f"{uuid.uuid4().hex}_{Path(filename).name}"
        saved.write_bytes(await upload.read())
        saved_paths.append(saved)

    saved_labor_paths: list[Path] = []
    for upload in labor_uploads:
        filename = str(getattr(upload, "filename", "") or "")
        if not filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse({"ok": False, "message": f"{filename} 不是 Excel 工時檔案"}, status_code=400)

        saved = UPLOAD_DIR / f"labor_{uuid.uuid4().hex}_{Path(filename).name}"
        saved.write_bytes(await upload.read())
        saved_labor_paths.append(saved)

    try:
        year_value: Optional[int] = None
        if year is not None and str(year).strip() != "":
            year_value = int(str(year).strip())

        output_path, summary = process_files(saved_paths, year_value, saved_labor_paths, labor_mode)
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)

    return {"ok": True, "summary": summary, "download_url": f"/download/{output_path.name}"}



# =========================================================
# Module 1 · Step 2 Batch Data Formatting
# Step 1 Output + Product Activity Bulk Template -> Formatted Bulk File
# =========================================================
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
    output_path = OUTPUT_DIR / f"formatted_product_activity_data_bulk_{token}.xlsx"

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



@app.post("/upload-rule-master")
async def upload_rule_master(file: UploadFile = File(...)):
    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
        return JSONResponse({"ok": False, "message": "請上傳 Excel 或 CSV 檔案"}, status_code=400)
    saved = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    saved.write_bytes(await file.read())
    try:
        count = save_uploaded_rule(saved)
    except Exception as exc:
        traceback.print_exc()
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
    path = DATA_DIR / "rule_master.csv"
    if not path.exists():
        ensure_master_files()
    return FileResponse(path, filename="rule_master.csv", media_type="text/csv")


@app.get("/download-product-series-master")
def download_product_series_master():
    path = DATA_DIR / "product_series_master.csv"
    if not path.exists():
        ensure_master_files()
    return FileResponse(path, filename="product_series_master.csv", media_type="text/csv")

# =========================================================
# Module 2 · BOM Expansion
# Standard BOM + Raw Material Bulk Template -> Raw Material Bulk
# =========================================================
from bom_formatter import generate_raw_material_bulk_file


@app.post("/process-bom-expansion")
async def process_bom_expansion(
    bom_file: UploadFile = File(...),
    template_file: UploadFile = File(...),
):
    if not bom_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "Standard BOM 請上傳 Excel 檔案"}, status_code=400)

    if not template_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "Raw Material Bulk Template 請上傳 Excel 檔案"}, status_code=400)

    token = uuid.uuid4().hex[:10]

    bom_path = UPLOAD_DIR / f"standard_bom_{token}_{bom_file.filename}"
    template_path = UPLOAD_DIR / f"raw_material_template_{token}_{template_file.filename}"
    output_path = OUTPUT_DIR / f"raw_material_activity_data_bulk_{token}.xlsx"

    bom_path.write_bytes(await bom_file.read())
    template_path.write_bytes(await template_file.read())

    try:
        summary = generate_raw_material_bulk_file(
            bom_path=bom_path,
            raw_material_template_path=template_path,
            output_path=output_path,
        )
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)

    return {
        "ok": True,
        "message": "BOM Expansion completed successfully.",
        "summary": summary,
        "download_url": f"/download/{output_path.name}",
    }
