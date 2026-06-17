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
from bulk_formatter import generate_product_activity_bulk_file, generate_product_activity_bulk_files_by_site, generate_product_activity_bulk_files_by_site_zip
from bom_formatter import generate_raw_material_bulk_file

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Annual Output Platform v6", version="6.0.0")
print("===== CMP MAIN VERSION: GOLDEN_V1_ORDER_NUMBER_DROPDOWN_HOURS_FIX =====")
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

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
    "MITS", "US", "UK", "JP", "JIS", "105K", "106K", "110K",
]

# Product series extraction noise words:
# these words describe process/material/assembly context and must not become Product series.
SERIES_NOISE_WORDS = [
    "SMTFOR", "SMT FOR", "SMT", "FOR", "ASSY", "ASSEMBLY", "MODULE",
    "PCB", "ELECTRON", "ELECTRONIC", "PCBA", "BL", "PC", "PET", "PCPET",
    "CH",
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

FINISHED_PRODUCT_PREFIXES = ["SG-"]

def normalize_order_key(value: object) -> str:
    """Normalize SAP order number for matching production orders with working-hour orders.

    Handles SAP/Excel differences:
    - 000307544123 vs 307544123
    - 307544123.0 vs 307544123
    - spaces, hyphens, slashes, and other separators
    """
    if pd.isna(value):
        return ""

    text = str(value).strip().upper()
    text = re.sub(r"\.0+$", "", text)
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[^A-Z0-9]", "", text)

    if text.isdigit():
        text = text.lstrip("0") or "0"

    return text


def parse_labor_number(series: pd.Series) -> pd.Series:
    """Parse labor-hour values safely from Excel text."""
    text = series.astype(str).str.strip()
    text = text.str.replace(",", "", regex=False)
    text = text.str.replace("，", "", regex=False)
    text = text.str.replace("－", "-", regex=False)
    text = text.replace({"": "0", "nan": "0", "None": "0", "-": "0"})
    return pd.to_numeric(text, errors="coerce").fillna(0)

def normalize_labor_mode(value: object) -> str:
    """Normalize working-hour source selection.

    Canonical values:
    - both: Labor HR.Act + FOH-Others.Act
    - labor_hr: Labor HR.Act only
    - foh: FOH-Others.Act only
    """
    text = str(value or "both").strip().lower()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)

    if text in {"labor_hr", "laborhr", "labor-hr", "labor", "hr"}:
        return "labor_hr"
    if text in {"foh", "foh_others", "fohothers", "foh-others"}:
        return "foh"
    if text in {"both", "all", "labor+foh", "laborhr.act+foh-others.act"}:
        return "both"

    if "人員+設備" in text or "人員設備" in text:
        return "both"
    if "人員工時" in text and "設備工時" in text:
        return "both"
    if "人員工時" in text:
        return "labor_hr"
    if "設備工時" in text:
        return "foh"

    if "only" in text:
        if "foh" in text:
            return "foh"
        if "labor" in text:
            return "labor_hr"

    if "foh" in text and "labor" not in text:
        return "foh"
    if "labor" in text and "foh" not in text:
        return "labor_hr"

    return "both"



RULE_COLUMNS = [
    "Priority", "Rule Type", "Key", "Product Type", "Customer",
    "Customer Code Logic", "Is_WIP", "Enabled", "Product Line", "Production Site",
]

DEFAULT_RULE_MASTER = (
    "Priority,Rule Type,Key,Product Type,Customer,Customer Code Logic,Is_WIP,Enabled,Product Line,Production Site\n"
    "1,Material Number Prefix,851-,WIP,,,Y,Y,,\n"
    "2,Material Number Prefix,852-,WIP,,,Y,Y,,\n"
    "10,Series Prefix,SN,NB,,NB_3RD_CHAR,N,Y,,\n"
    "11,Series Prefix,FU,NB,,NB_3RD_CHAR,N,Y,,\n"
    "12,Series Prefix,SP,TP,,NB_3RD_CHAR,N,Y,,\n"
    "13,Series Prefix,SM,DT Mouse,,DT_3_4_CHAR,N,Y,,\n"
    "14,Series Prefix,SA,DT Accessory,,DT_3_4_CHAR,N,Y,,\n"
    "15,Description Contains,RECEIVER,DT Dongle,,,N,Y,,\n"
    "16,Series Prefix,SK,DT Keyboard,,DT_3_4_CHAR,N,Y,,\n"
    "17,Series Prefix,SB,DT Keyboard+Mouse,,DT_3_4_CHAR,N,Y,,\n"
    "18,Series Prefix,ST,DT Tablet Keyboard,,DT_3_4_CHAR,N,Y,,\n"
    "19,Description Contains,TOUCH PAD MODULE,TP,,,N,Y,,\n"
    "20,Description Contains,TOUCHPAD MODULE,TP,,,N,Y,,\n"
    "21,Series Prefix,SCMC,WIP,,,Y,Y,,\n"
    "90,Description Contains,ASSY,WIP,,,Y,Y,,\n"
    "999,Default,*,WIP,,,Y,Y,,\n"
)


def ensure_master_files() -> None:
    rule_path = DATA_DIR / "rule_master.csv"
    series_path = DATA_DIR / "product_series_master.csv"
    if not rule_path.exists():
        rule_path.write_text(DEFAULT_RULE_MASTER, encoding="utf-8-sig")
    if not series_path.exists():
        series_path.write_text("Plant,Product series,產品類型,客戶代碼,客戶名稱\n", encoding="utf-8-sig")


ensure_master_files()



def production_site_from_line(product_line: object) -> str:
    """Map Product Line to Production Site. Used only as fallback when Rule Master does not provide a site."""
    line = str(product_line or "").strip().upper()
    if line == "NB":
        return "常州廠(A2)-IPS"
    if line == "TP":
        return "常州廠(A9)-IPS"
    return ""


def resolve_production_site(product_line: object, production_site: object = "") -> str:
    """Resolve Production Site from Rule Master output.

    Priority:
    1. Production Site from rule_master.csv
    2. Product Line fallback mapping
    3. Legacy fallback for unclassified / non-NB / non-TP records
    """
    site = str(production_site or "").strip()
    if site:
        return site

    fallback_site = production_site_from_line(product_line)
    if fallback_site:
        return fallback_site

    # Do not assign a default site when Product Line cannot be determined.
    # This avoids unresolved Product Series being incorrectly assigned to 石碣廠-IPS.
    return ""


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


def _strip_leading_noise(text: str) -> str:
    """Remove process/material words before Product series extraction."""
    value = str(text or "").upper()
    value = value.replace("&", " ").replace("+", " ")
    value = re.sub(r"SMT\s*FOR", " ", value)
    for word in SERIES_NOISE_WORDS:
        value = re.sub(rf"\b{re.escape(word)}\b", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _extract_series_core(candidate: str) -> str:
    """Extract the core product series and remove trailing region/layout descriptors.

    Examples:
    - SN3103B02US105KBLANKNOKEYCANEW -> SN3103B02
    - SP7D01B02MITS -> SP7D01B02
    - SL6D00B00US -> SL6D00B00
    - SCMC50B11XXX -> SCMC50B11
    """
    value = str(candidate or "").upper()
    if not value:
        return ""

    # First cut at known description words.
    cut_positions = [pos for word in SERIES_STOP_WORDS if (pos := value.find(word)) > 0]

    # Country/layout descriptors often follow the real series immediately,
    # e.g. SN3103B02US105K -> SN3103B02, SL6D00B00US -> SL6D00B00.
    for token in ["US", "UK", "JP", "JIS"]:
        pos = value.find(token)
        if pos > 0:
            cut_positions.append(pos)

    if cut_positions:
        value = value[:min(cut_positions)]

    # Prefer the longest core ending with Letter + 2 digits.
    # This covers common Lite-On series such as SN3103B02, SP7D01B02, SL6D00B00, SCMC50B11.
    core_match = re.match(r"^([A-Z]{2,5}[A-Z0-9]*\d[A-Z]\d{2})", value)
    if core_match:
        return core_match.group(1)

    return value


def trim_series_candidate(candidate: str) -> str:
    """Clean a candidate product series.

    Avoid treating process/material words such as SMTFOR, ASSY, MODULE, PC+PET as part of Product series.
    """
    candidate = str(candidate or "").upper()
    candidate = re.sub(r"[^A-Z0-9].*$", "", candidate)
    candidate = _extract_series_core(candidate)
    return candidate


def _series_candidate_from_text(text: str, pattern: re.Pattern) -> Optional[str]:
    """在一段文字中尋找第一個符合 Series Prefix 的候選值，並套用描述詞截斷。"""
    if not text:
        return None

    # Remove process/material words first, so "SMTfor SP7D01B02" becomes "SP7D01B02".
    cleaned_text = _strip_leading_noise(str(text).upper())

    # 移除空白與非英數符號，讓 CHSN4396BL1 / AssyCH SN5372BL 都能被抓到。
    compact_text = re.sub(r"[^A-Z0-9]+", "", cleaned_text)

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
    - Assy,PCB&Electron, SMTfor SP7D01B02,Mits -> SP7D01B02
    - Assy,Module,BL,PC+PET,SL6D00B00,US -> SL6D00B00
    - Assy,CH,SN3103B02US105KBlanknokeycaNEW -> SN3103B02
    - Assy,CHSN4396BL1,UK,106KBlank,nokeycaNEW -> SN4396BL1
    - SP2B20XF0AssyHapticForce Touchpad module -> SP2B20XF0
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

        product_line = str(row.get("Product Line", "") or "").strip()
        if not product_line and product_type.upper() != "WIP":
            product_line = product_type

        production_site = str(row.get("Production Site", "") or "").strip()

        return {
            "產品類型": product_type,
            "Product Line": product_line,
            "Production Site": production_site,
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
    product_line = product_type if product_type.upper() != "WIP" else ""
    return {
        "產品類型": product_type,
        "Product Line": product_line,
        "Production Site": production_site_from_line(product_line),
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
        "Product Line": "",
        "Production Site": "",
        "客戶代碼": "",
        "客戶名稱": "",
        "判斷來源": "Default WIP",
        "規則判定結果": "WIP",
        "命中規則": "No rule matched → WIP",
        "Is_WIP": "Y",
    }





def is_finished_product_whitelist(material_number: object) -> bool:
    """Finished product whitelist.

    Prefixes such as SG- mean:
    - This is a finished product number.
    - Do not let ASSY / SFG / MODULE / PCBA override it to WIP.
    - Product Type / Product Line / Production Site must still be decided by Product Series rules.
    """
    material_number_u = str(material_number or "").upper().strip()
    return any(material_number_u.startswith(prefix) for prefix in FINISHED_PRODUCT_PREFIXES)


def is_wip_by_rule_master(material_number: object, description: object, series: str, masters: dict) -> bool:
    """Detect WIP independently from Product Line / Production Site attribution.

    Product Line may be inferred from SN/FU/SP/SCMC rules, but Product Type must remain WIP
    when any WIP rule matches, e.g. 850-/851-/852-/H50-, SFG, ASSY, SCMC.
    Default WIP is excluded here because it is only a fallback when no rule matches.
    """
    material_number_u = str(material_number or "").upper().strip()
    if is_finished_product_whitelist(material_number_u):
        return False
    description_u = str(description or "").upper().strip()
    series_u = str(series or "").upper().strip()

    rules: pd.DataFrame = masters["rules"]
    for _, row in rules.iterrows():
        rule_type = str(row.get("Rule Type", "") or "").strip()
        key = str(row.get("Key", "") or "").strip()

        if normalize_rule_type(rule_type) in ["default", "預設"]:
            continue

        product_type = str(row.get("Product Type", "") or "").strip().upper()
        is_wip = str(row.get("Is_WIP", "") or "").strip().upper()

        if product_type != "WIP" and is_wip not in ["Y", "YES", "TRUE", "1"]:
            continue

        if rule_matches(rule_type, key, material_number_u, description_u, series_u):
            return True

    return False



def infer_product_type_line_site_from_series_rules(description: object, series: str, masters: dict) -> tuple[str, str, str]:
    """Infer Product Type / Product Line / Production Site from series or description rules.

    Used for finished-product whitelist material numbers such as SG-:
    SG- itself only means non-WIP; SN/SP/FU/etc. from Product Series decides NB/TP and site.
    """
    description_u = str(description or "").upper().strip()
    series_u = str(series or "").upper().strip()
    rules: pd.DataFrame = masters["rules"]

    for _, row in rules.iterrows():
        rule_type = str(row.get("Rule Type", "") or "").strip()
        key = str(row.get("Key", "") or "").strip()
        rt = normalize_rule_type(rule_type)

        if rt not in ["series prefix", "product series prefix", "產品系列前綴", "系列前綴",
                      "series exact", "product series exact", "產品系列", "產品系列完全符合",
                      "description contains", "material description contains", "描述包含", "品名包含"]:
            continue

        if not rule_matches(rule_type, key, "", description_u, series_u):
            continue

        product_type = str(row.get("Product Type", "") or "").strip()
        if product_type.upper() == "WIP":
            # A finished-product whitelist item should not become WIP because of SCMC/ASSY/SFG style rules.
            continue

        product_line = str(row.get("Product Line", "") or "").strip()
        if not product_line:
            product_line = product_type

        production_site = str(row.get("Production Site", "") or "").strip()
        if not production_site:
            production_site = production_site_from_line(product_line)

        if product_type or product_line or production_site:
            return product_type, product_line, production_site

    return "", "", ""


def infer_product_line_site_from_rules(description: object, series: str, masters: dict) -> tuple[str, str]:
    """Infer Product Line / Production Site for WIP without changing Product Type.

    WIP prefix rules identify semi-finished goods only. For Production Site,
    infer the original product line from Series Prefix or Description Contains rules.
    """
    description_u = str(description or "").upper().strip()
    series_u = str(series or "").upper().strip()
    rules: pd.DataFrame = masters["rules"]

    for _, row in rules.iterrows():
        rule_type = str(row.get("Rule Type", "") or "").strip()
        key = str(row.get("Key", "") or "").strip()
        rt = normalize_rule_type(rule_type)

        if rt not in [
            "series prefix", "product series prefix", "產品系列前綴", "系列前綴",
            "series exact", "product series exact", "產品系列", "產品系列完全符合",
            "description contains", "material description contains", "描述包含", "品名包含",
        ]:
            continue

        if not rule_matches(rule_type, key, "", description_u, series_u):
            continue

        product_line = str(row.get("Product Line", "") or "").strip()
        product_type = str(row.get("Product Type", "") or "").strip()
        if not product_line and product_type.upper() != "WIP":
            product_line = product_type

        production_site = str(row.get("Production Site", "") or "").strip()
        if not production_site:
            production_site = production_site_from_line(product_line)

        if product_line or production_site:
            return product_line, production_site

    return "", ""

def load_production_dataframe(paths: list[Path]) -> pd.DataFrame:
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
        raise ValueError("沒有可處理的工單資料")

    return pd.concat(frames, ignore_index=True)



def load_labor_dataframe(paths: list[Path], labor_mode: str = "both") -> pd.DataFrame:
    """Load production working-hour order files.

    Matching key:
    - Production quantity file: Order
    - Working-hour file: Order Number / Order
    - Match by normalized Order Merge Key only.
    - No Plant + Material Number fallback.
    """
    columns = [
        "Order", "Order Merge Key", "Plant", "Material Number",
        "Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"
    ]
    if not paths:
        return pd.DataFrame(columns=columns)

    mode = normalize_labor_mode(labor_mode)

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

        part["Labor HR.Act"] = parse_labor_number(df[cols["labor_hr"]])
        part["FOH-Others.Act"] = parse_labor_number(df[cols["foh_others"]])

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
    labor = labor[labor["Order Merge Key"].astype(str).str.strip() != ""].copy()

    if labor.empty:
        return pd.DataFrame(columns=columns)

    return (
        labor.groupby(["Order Merge Key"], dropna=False, as_index=False)
        .agg({
            "Order": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip()))),
            "Plant": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip()))),
            "Material Number": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip()))),
            "Labor HR.Act": "sum",
            "FOH-Others.Act": "sum",
            "Selected Hours": "sum",
            "Labor Source file": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip())))
        })
        .rename(columns={"Labor Source file": "Labor Source files"})
    )


def attach_labor_hours(out: pd.DataFrame, labor: pd.DataFrame) -> pd.DataFrame:
    """Attach labor hours to production output by Order Number only.

    No Plant + Material Number fallback is used.
    """
    out = out.copy()
    if "Order Merge Key" not in out.columns:
        out["Order Merge Key"] = out["Order"].apply(normalize_order_key)

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"]:
        if col not in out.columns:
            out[col] = 0 if col != "Labor Source files" else ""

    if labor is None or labor.empty:
        for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)
        out["Labor Source files"] = out["Labor Source files"].fillna("").astype(str)
        out["Matched Labor Order Number"] = ""
        out["Labor Match Status"] = "No labor file"
        return out

    labor = labor.copy()
    if "Order Merge Key" not in labor.columns:
        labor["Order Merge Key"] = labor["Order"].apply(normalize_order_key)

    order_labor = (
        labor[labor["Order Merge Key"].astype(str).str.strip() != ""]
        .groupby(["Order Merge Key"], dropna=False, as_index=False)
        .agg({
            "Order": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip()))),
            "Labor HR.Act": "sum",
            "FOH-Others.Act": "sum",
            "Selected Hours": "sum",
            "Labor Source files": lambda s: "; ".join(sorted(set(str(x) for x in s if str(x).strip())))
        })
        .rename(columns={"Order": "Matched Labor Order Number"})
    )

    out = out.merge(order_labor, on="Order Merge Key", how="left", suffixes=("", "_labor"))

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        labor_col = f"{col}_labor"
        if labor_col in out.columns:
            out[col] = pd.to_numeric(out[labor_col], errors="coerce").fillna(0)
            out = out.drop(columns=[labor_col])
        else:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    if "Labor Source files_labor" in out.columns:
        out["Labor Source files"] = out["Labor Source files_labor"].fillna("").astype(str)
        out = out.drop(columns=["Labor Source files_labor"])

    if "Matched Labor Order Number" not in out.columns:
        out["Matched Labor Order Number"] = ""

    out["Matched Labor Order Number"] = out["Matched Labor Order Number"].fillna("").astype(str)
    out["Labor Match Status"] = out["Matched Labor Order Number"].apply(
        lambda x: "Matched" if str(x).strip() else "Not matched"
    )
    out["Labor Source files"] = out["Labor Source files"].fillna("").astype(str)

    return out


def process_files(
    paths: list[Path],
    year: Optional[int],
    labor_paths: Optional[list[Path]] = None,
    labor_mode: str = "both",
) -> tuple[Path, dict]:
    masters = build_masters()
    labor_mode = normalize_labor_mode(labor_mode)

    out = load_production_dataframe(paths)
    labor = load_labor_dataframe(labor_paths or [], labor_mode)
    out = attach_labor_hours(out, labor)

    # Final safety guard:
    # Total working hours follows the platform dropdown selection.
    # - 人員+設備工時: Labor HR.Act + FOH-Others.Act
    # - 人員工時: Labor HR.Act
    # - 設備工時: FOH-Others.Act
    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        if col in out.columns:
            out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    if labor_mode == "labor_hr":
        out["Selected Hours"] = out["Labor HR.Act"]
    elif labor_mode == "foh":
        out["Selected Hours"] = out["FOH-Others.Act"]
    else:
        out["Selected Hours"] = out["Labor HR.Act"] + out["FOH-Others.Act"]

    out["Working Hour Source"] = labor_mode

    if year:
        out = out[out["Year"] == int(year)].copy()

    parsed = out["Material description"].apply(lambda desc: parse_product_series(desc, masters))
    out["Product series"] = parsed.apply(lambda x: x[0])
    out["解析說明"] = parsed.apply(lambda x: x[1])

    classified = out.apply(
        lambda r: classify(r["Material Number"], r["Material description"], r["Product series"], r["Plant"], masters),
        axis=1,
    )
    out["產品類型"] = classified.apply(lambda x: x.get("產品類型", ""))
    out["Product Line"] = classified.apply(lambda x: x.get("Product Line", ""))
    out["Production Site"] = classified.apply(lambda x: x.get("Production Site", ""))
    out["客戶代碼"] = classified.apply(lambda x: x.get("客戶代碼", ""))
    out["客戶名稱"] = classified.apply(lambda x: x.get("客戶名稱", ""))
    out["判斷來源"] = classified.apply(lambda x: x.get("判斷來源", ""))
    out["規則判定結果"] = classified.apply(lambda x: x.get("規則判定結果", ""))
    out["命中規則"] = classified.apply(lambda x: x.get("命中規則", ""))
    out["Is_WIP"] = classified.apply(lambda x: x.get("Is_WIP", "N"))

    # Minimal change: do not alter Product Type. Only infer missing Product Line / Production Site.
    missing_line_mask = out["Product Line"].astype(str).str.strip().eq("")
    if missing_line_mask.any():
        inferred = out.loc[missing_line_mask].apply(
            lambda r: infer_product_line_site_from_rules(r["Material description"], r["Product series"], masters),
            axis=1,
        )
        out.loc[missing_line_mask, "Product Line"] = inferred.apply(lambda x: x[0]).to_numpy()
        out.loc[missing_line_mask, "Production Site"] = inferred.apply(lambda x: x[1]).to_numpy()

    out["Production Site"] = out.apply(
        lambda r: resolve_production_site(r["Product Line"], r["Production Site"]),
        axis=1,
    )

    # Final safety guard:
    # Product Line / Production Site can be inferred from series rules, but Product Type must remain WIP
    # if WIP rules such as 850-/851-/852-/H50-/SFG/ASSY/SCMC match.
    wip_rule_mask = out.apply(
        lambda r: is_wip_by_rule_master(r["Material Number"], r["Material description"], r["Product series"], masters),
        axis=1,
    )
    if wip_rule_mask.any():
        out.loc[wip_rule_mask, "產品類型"] = "WIP"
        out.loc[wip_rule_mask, "Is_WIP"] = "Y"
        out.loc[wip_rule_mask, "規則判定結果"] = "WIP"

    # Finished-product whitelist:
    # SG- means non-WIP only. Product Type / Product Line / Production Site still comes from Product Series.
    finished_product_mask = out["Material Number"].apply(is_finished_product_whitelist)
    if finished_product_mask.any():
        inferred_finished = out.loc[finished_product_mask].apply(
            lambda r: infer_product_type_line_site_from_series_rules(
                r["Material description"], r["Product series"], masters
            ),
            axis=1,
        )
        inferred_product_type = inferred_finished.apply(lambda x: x[0]).to_numpy()
        inferred_product_line = inferred_finished.apply(lambda x: x[1]).to_numpy()
        inferred_production_site = inferred_finished.apply(lambda x: x[2]).to_numpy()

        idx = out.index[finished_product_mask]
        for pos, row_idx in enumerate(idx):
            if inferred_product_type[pos]:
                out.at[row_idx, "產品類型"] = inferred_product_type[pos]
            if inferred_product_line[pos]:
                out.at[row_idx, "Product Line"] = inferred_product_line[pos]
            if inferred_production_site[pos]:
                out.at[row_idx, "Production Site"] = inferred_production_site[pos]
            out.at[row_idx, "Is_WIP"] = "N"

        out.loc[finished_product_mask, "Production Site"] = out.loc[finished_product_mask].apply(
            lambda r: resolve_production_site(r["Product Line"], r["Production Site"]),
            axis=1,
        )

    group_cols = [
        "Year", "Plant", "Production Site", "Product Line", "Material Number", "Material description", "Product series",
        "產品類型", "客戶代碼", "客戶名稱", "判斷來源", "Is_WIP"
    ]
    annual = (
        out.groupby(group_cols, dropna=False, as_index=False)
        .agg({
            "Delivered quantity": "sum",
            "Labor HR.Act": "sum",
            "FOH-Others.Act": "sum",
            "Selected Hours": "sum",
        })
        .rename(columns={
            "Delivered quantity": "年度生產量",
            "Labor HR.Act": "年度人員工時",
            "FOH-Others.Act": "年度設備工時",
            "Selected Hours": "年度總工時",
        })
        .sort_values(["Plant", "Material Number"])
    )

    plant_qty_total = annual.groupby(["Year", "Plant"], dropna=False)["年度生產量"].transform("sum")
    plant_hour_total = annual.groupby(["Year", "Plant"], dropna=False)["年度總工時"].transform("sum")
    annual["生產數量占比(%)"] = 0.0
    annual["生產工時占比(%)"] = 0.0

    qty_mask = plant_qty_total.ne(0)
    hour_mask = plant_hour_total.ne(0)

    annual.loc[qty_mask, "生產數量占比(%)"] = (
        annual.loc[qty_mask, "年度生產量"] / plant_qty_total.loc[qty_mask] * 100
    )
    annual.loc[hour_mask, "生產工時占比(%)"] = (
        annual.loc[hour_mask, "年度總工時"] / plant_hour_total.loc[hour_mask] * 100
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

    # Export view: only show the selected working-hour source in Excel.
    out_export = out.copy()
    annual_export = annual.copy()

    # Rename Excel output header only; keep internal calculation column as Selected Hours.
    out_export = out_export.rename(columns={"Selected Hours": "Total working hours"})

    # Keep Labor HR.Act, FOH-Others.Act, and Total working hours for traceability.

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_export.to_excel(writer, index=False, sheet_name="工單明細_已分類")
        annual_export.to_excel(writer, index=False, sheet_name="Plant_Material年度產量")
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
        "labor_mode": labor_mode,
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
    """Backward-compatible wrapper for single-file processing."""
    return process_files([path], year, None, "both")


def normalize_rule_upload(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {}
    for c in df.columns:
        key = str(c).strip().lower()
        if key in ["priority", "優先順序", "排序"]:
            rename_map[c] = "Priority"
        elif key in ["rule type", "規則類型", "判斷類型"]:
            rename_map[c] = "Rule Type"
        elif key in ["key", "規則值", "關鍵字", "prefix", "前綴"]:
            rename_map[c] = "Key"
        elif key in ["product type", "產品分類", "產品類型"]:
            rename_map[c] = "Product Type"
        elif key in ["customer", "客戶", "客戶名稱"]:
            rename_map[c] = "Customer"
        elif key in ["customer code logic", "客戶代碼邏輯"]:
            rename_map[c] = "Customer Code Logic"
        elif key in ["is_wip", "is wip", "wip", "半品"]:
            rename_map[c] = "Is_WIP"
        elif key in ["product line", "產品線", "歸屬類型", "production site type"]:
            rename_map[c] = "Product Line"
        elif key in ["production site", "生產廠區", "廠區", "廠別"]:
            rename_map[c] = "Production Site"
        elif key in ["enabled", "啟用"]:
            rename_map[c] = "Enabled"
    df = df.rename(columns=rename_map).fillna("")
    for col in RULE_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[RULE_COLUMNS].copy()
    for c in RULE_COLUMNS:
        df[c] = df[c].astype(str).str.strip()
    df["Enabled"] = df["Enabled"].replace("", "Y")
    df["Is_WIP"] = df["Is_WIP"].replace("", "N")
    df = df[(df["Rule Type"] != "") & (df["Key"] != "") & (df["Product Type"] != "")]
    return df


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




@app.get("/debug-version")
def debug_version():
    return {
        "ok": True,
        "app": "Carbon Management Platform",
        "version": "PROCESS_MANUAL_FORM_V6",
        "process_endpoint": "manual form compatible",
        "supports": ["files multi-upload", "file single-upload", "blank year"],
    }

@app.post("/process")
async def process(request: Request):
    """Step 1 processing endpoint.

    Manual multipart reader:
    - accepts production quantity files from files/file
    - accepts working-hour files from labor_files/labor_file
    - avoids FastAPI 422 validation when optional fields are omitted
    """
    try:
        form = await request.form()
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": f"無法讀取上傳表單：{exc}"}, status_code=400)

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
        return JSONResponse({"ok": False, "message": "請至少上傳一個 Excel 生產數量工單檔案"}, status_code=400)

    labor_mode = normalize_labor_mode(form.get("labor_mode") or "both")
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
# Step 2 · Batch Data Formatting
# Step1 Output + Bulk Template -> Formatted Product Activity Bulk
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

    step1_path = UPLOAD_DIR / f"step1_output_{token}_{Path(step1_file.filename).name}"
    template_path = UPLOAD_DIR / f"bulk_template_{token}_{Path(template_file.filename).name}"

    step1_path.write_bytes(await step1_file.read())
    template_path.write_bytes(await template_file.read())

    try:
        summary = generate_product_activity_bulk_files_by_site_zip(
            step1_output_path=step1_path,
            bulk_template_path=template_path,
            output_dir=OUTPUT_DIR,
            token=token,
        )
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)

    return {
        "ok": True,
        "message": "Bulk file generated successfully.",
        "summary": summary,
        "download_url": summary.get("download_url", ""),
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

    bom_path = UPLOAD_DIR / f"standard_bom_{token}_{Path(bom_file.filename).name}"
    template_path = UPLOAD_DIR / f"raw_material_template_{token}_{Path(template_file.filename).name}"
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

