from __future__ import annotations

import re
import traceback
import uuid
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

import pandas as pd
from fastapi import FastAPI, File, Form, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from bulk_formatter import generate_product_activity_bulk_file, generate_product_activity_bulk_files_by_site, generate_product_activity_bulk_files_by_site_zip
from bom_formatter import BOM_FORMATTER_VERSION, generate_raw_material_bulk_file, generate_raw_material_bulk_files_by_site_zip, export_bom_structure_file, generate_working_hour_rollup_file
from factor_selector import FACTOR_SELECTOR_VERSION, apply_ccl_factors_to_raw_material_bulk, collect_factor_library_geographies, preload_factor_libraries, search_factor_library

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
DATA_DIR = BASE_DIR / "data"
RULE_LIBRARY_DIR = BASE_DIR / "rule_library"
FACTOR_LIBRARY_DIR = DATA_DIR / "factor_library"
LATEST_BOM_STRUCTURE_PATH = OUTPUT_DIR / "bom_structure_latest.xlsx"
LATEST_WORKING_HOUR_ROLLUP_PATH = OUTPUT_DIR / "working_hour_rollup_latest.xlsx"
LATEST_RAW_MATERIAL_BULK_PATH = OUTPUT_DIR / "raw_material_activity_data_bulk_latest.xlsx"

RULE_SET_MAP = {
    "IPS": "LiteOn_IPS",
    "AE": "LiteOn_AE",
    "PC_CE": "LiteOn_PC_CE",
}

DEFAULT_RULE_SET = "IPS"

UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DATA_DIR.mkdir(exist_ok=True)
RULE_LIBRARY_DIR.mkdir(exist_ok=True)
FACTOR_LIBRARY_DIR.mkdir(exist_ok=True)

app = FastAPI(title="Annual Output Platform v6", version="6.0.0")
print("===== CMP MAIN VERSION: CMP_V15_0_RULE_MASTER_ENGINE_SITE_PREFIX_WHITELIST =====")
print(f"===== BOM FORMATTER VERSION: {BOM_FORMATTER_VERSION} =====")

MODULE3_CCL_EXECUTOR = ThreadPoolExecutor(max_workers=2)
MODULE3_CCL_JOBS: Dict[str, Dict[str, Any]] = {}

def _set_module3_ccl_job(job_id: str, **updates: Any) -> None:
    job = MODULE3_CCL_JOBS.setdefault(job_id, {})
    job.update(updates)
    # 後台統一估算剩餘秒數；前台只顯示 remaining_seconds。
    progress = int(job.get("progress") or 0)
    started_at = job.get("started_at")
    if started_at and 5 < progress < 100:
        try:
            start_dt = datetime.fromisoformat(started_at)
            elapsed = max(1.0, (datetime.now() - start_dt).total_seconds())
            job["remaining_seconds"] = max(1, int(round(elapsed * (100 - progress) / progress)))
        except Exception:
            job["remaining_seconds"] = job.get("remaining_seconds", 30)
    elif progress >= 100:
        job["remaining_seconds"] = 0
    else:
        job.setdefault("remaining_seconds", 30)
    job["updated_at"] = datetime.now().isoformat(timespec="seconds")

def _run_module3_ccl_job(job_id: str, raw_path: Path, ccl_path: Path, output_path: Path) -> None:
    def report(progress: int, step: str) -> None:
        _set_module3_ccl_job(job_id, status="running", progress=max(0, min(100, int(progress))), step=step)

    try:
        report(1, "建立 CCL 係數對應工作")
        summary = apply_ccl_factors_to_raw_material_bulk(raw_path, ccl_path, output_path, progress_callback=report)
        summary["app_version"] = "CMP_MODULE3_CCL_JOB_V1"
        _set_module3_ccl_job(
            job_id,
            status="success",
            progress=100,
            step="CCL 係數對應完成",
            message="CCL 係數對應完成。",
            summary=summary,
            download_url=summary.get("download_url", f"/download/{output_path.name}"),
        )
    except Exception as exc:
        traceback.print_exc()
        _set_module3_ccl_job(
            job_id,
            status="error",
            progress=100,
            step="CCL 係數對應失敗",
            message=str(exc),
        )

app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))

# =========================================================
# v6 分類邏輯：
# 1. rule_master.csv 依 Priority 由小到大判斷
#    - Material Number Exact
#    - Material Number Prefix
#    - Description Contains
#    - Series Prefix
#    - Plant Exact / Plant Prefix
#    - Default
# 2. 若 rule_master 未命中，再查 product_series_master.csv
# 3. Strict Site 未命中則排除；非 Strict Site 未命中才寫 WIP
# =========================================================

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

FINISHED_PRODUCT_PREFIXES: list[str] = []

def normalize_order_key(value: object) -> str:
    """Normalize SAP order number for matching quantity orders with working-hour orders.

    Rules:
    - Remove trailing .0 generated by Excel numeric conversion.
    - Remove spaces.
    - Keep only letters and numbers.
    - If the order number is numeric, remove leading zeros.
      Example: 000307544123 -> 307544123
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
    "Priority", "Rule Type", "Key", "Product Type", "Product Line",
    "Production Site", "Is_WIP", "Enabled", "Prefix Length",
]

DEFAULT_RULE_MASTER = (
    "Priority,Rule Type,Key,Product Type,Product Line,Production Site,Is_WIP,Enabled,Prefix Length\n"
    "1,Material Number Prefix,851-,WIP,,,Y,Y,\n"
    "2,Material Number Prefix,852-,WIP,,,Y,Y,\n"
    "10,Series Prefix,SN,NB,,,,Y,\n"
    "11,Series Prefix,FU,NB,,,,Y,\n"
    "12,Series Prefix,SP,TP,,,,Y,\n"
    "13,Series Prefix,SM,DT Mouse,DT,,,Y,\n"
    "14,Series Prefix,SA,DT Accessory,DT,,,Y,\n"
    "15,Description Contains,RECEIVER,DT Dongle,DT,,,Y,\n"
    "16,Series Prefix,SK,DT Keyboard,DT,,,Y,\n"
    "17,Series Prefix,SB,DT Keyboard+Mouse,DT,,,Y,\n"
    "18,Series Prefix,ST,DT Tablet Keyboard,DT,,,Y,\n"
    "19,Description Contains,TOUCH PAD MODULE,TP,,,,Y,\n"
    "20,Description Contains,TOUCHPAD MODULE,TP,,,,Y,\n"
    "21,Series Prefix,SCMC,WIP,,,Y,Y,\n"
    "90,Description Contains,ASSY,WIP,,,Y,Y,\n"
    "999,Default,*,WIP,,,Y,Y,\n"
)

def normalize_rule_set(value: object) -> str:
    text = str(value or DEFAULT_RULE_SET).strip().upper()
    text = text.replace("&", "_").replace("-", "_").replace(" ", "_")
    if text in {"PC&CE", "PC_CE", "PCCE", "PC CE"}:
        text = "PC_CE"
    if text not in RULE_SET_MAP:
        text = DEFAULT_RULE_SET
    return text


def get_rule_set_dir(rule_set: object = DEFAULT_RULE_SET) -> Path:
    normalized = normalize_rule_set(rule_set)
    return RULE_LIBRARY_DIR / RULE_SET_MAP[normalized]


def ensure_master_files() -> None:
    rule_path = DATA_DIR / "rule_master.csv"
    series_path = DATA_DIR / "product_series_master.csv"
    if not rule_path.exists():
        rule_path.write_text(DEFAULT_RULE_MASTER, encoding="utf-8-sig")
    if not series_path.exists():
        series_path.write_text("Plant,Product series,產品類型\n", encoding="utf-8-sig")

    for rule_set_key in RULE_SET_MAP:
        rule_dir = get_rule_set_dir(rule_set_key)
        rule_dir.mkdir(parents=True, exist_ok=True)
        bu_rule_path = rule_dir / "rule_master.csv"
        bu_series_path = rule_dir / "product_series_master.csv"

        if not bu_rule_path.exists():
            if rule_path.exists():
                bu_rule_path.write_bytes(rule_path.read_bytes())
            else:
                bu_rule_path.write_text(DEFAULT_RULE_MASTER, encoding="utf-8-sig")

        if not bu_series_path.exists():
            if series_path.exists():
                bu_series_path.write_bytes(series_path.read_bytes())
            else:
                bu_series_path.write_text("Plant,Product series,產品類型\n", encoding="utf-8-sig")


ensure_master_files()



def production_site_from_line(product_line: object) -> str:
    """Rule Master only mode.

    Production Site must come from rule_master.csv.
    This function is kept only for backward compatibility and never infers a site.
    """
    return ""

def resolve_production_site(product_line: object, production_site: object = "") -> str:
    """Rule Master only mode.

    Production Site is controlled by rule_master.csv only.
    If rule_master.csv leaves Production Site blank, keep it blank.
    """
    return str(production_site or "").strip()

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


def load_rule_master(rule_set: object = DEFAULT_RULE_SET) -> pd.DataFrame:
    df = read_csv_flexible(get_rule_set_dir(rule_set) / "rule_master.csv", RULE_COLUMNS)
    for c in RULE_COLUMNS:
        df[c] = df[c].astype(str).str.strip()
    df["Enabled"] = df["Enabled"].str.upper().replace("", "Y")
    df["Rule Type"] = df["Rule Type"].str.strip()
    df["Key"] = df["Key"].str.upper().str.strip()
    df["Priority_num"] = pd.to_numeric(df["Priority"], errors="coerce").fillna(9999)
    df = df[df["Enabled"].isin(["Y", "YES", "TRUE", "1"])]
    return df.sort_values(["Priority_num", "Rule Type", "Key"], kind="stable").reset_index(drop=True)


def load_product_series_master(rule_set: object = DEFAULT_RULE_SET) -> pd.DataFrame:
    path = get_rule_set_dir(rule_set) / "product_series_master.csv"
    df = read_csv_flexible(path)
    for col in ["Plant", "Product series", "產品類型"]:
        if col not in df.columns:
            df[col] = ""
        df[col] = df[col].astype(str).str.strip()
    df["Plant"] = df["Plant"].str.replace(r"\.0$", "", regex=True)
    df["Product series"] = df["Product series"].str.upper().str.replace(r"\s+", "", regex=True)
    return df[["Plant", "Product series", "產品類型"]].copy()


def build_masters(rule_set: object = DEFAULT_RULE_SET) -> dict:
    rule_set = normalize_rule_set(rule_set)
    rule_master = load_rule_master(rule_set)
    series_master = load_product_series_master(rule_set)

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
        "rule_set": rule_set,
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



def normalize_rule_type(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def rule_matches(rule_type: str, key: str, material_number: str, description: str, series: str, plant: str = "") -> bool:
    rt = normalize_rule_type(rule_type)
    key_u = str(key or "").upper().strip()
    plant_u = str(plant or "").upper().strip().replace(".0", "")
    if not key_u:
        return False

    # 支援英文與常見中文寫法
    if rt in ["plant exact", "plant", "plant code exact", "廠別", "廠別完全符合", "廠區", "廠區完全符合"]:
        return plant_u == key_u
    if rt in ["plant prefix", "plant code prefix", "廠別前綴", "廠區前綴"]:
        return plant_u.startswith(key_u)
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


def resolve_plant_production_site_from_rule_master(plant: object, masters: dict) -> tuple[str, str]:
    """Resolve Production Site by Plant rules in rule_master.csv.

    Supported Rule Type:
    - Plant Exact
    - Plant
    - Plant Prefix

    This function only resolves Production Site. It does not change Product Type or Product Line.
    """
    plant_u = str(plant or "").upper().strip().replace(".0", "")
    if not plant_u:
        return "", ""

    rules: pd.DataFrame = masters["rules"]
    for _, row in rules.iterrows():
        rule_type = str(row.get("Rule Type", "") or "").strip()
        key = str(row.get("Key", "") or "").strip()
        rt = normalize_rule_type(rule_type)

        if rt not in [
            "plant exact", "plant", "plant code exact", "廠別", "廠別完全符合", "廠區", "廠區完全符合",
            "plant prefix", "plant code prefix", "廠別前綴", "廠區前綴",
        ]:
            continue

        if not rule_matches(rule_type, key, "", "", "", plant_u):
            continue

        production_site = str(row.get("Production Site", "") or "").strip()
        if production_site:
            return production_site, f"{rule_type}={key}"

    return "", ""




def _normalize_production_site_key(value: object) -> str:
    """Normalize Production Site text for strict-site comparison."""
    return re.sub(r"\s+", "", str(value or "").strip().upper())


def get_strict_production_sites(masters: dict) -> set[str]:
    """Read strict Production Site rules from rule_master.csv.

    Rule purpose:
    - Shared sites such as 越南海防廠-IPS may contain products from other BUs.
    - If no normal classification rule/Product Series Master rule matches, rows from
      these sites must be excluded instead of falling back to Default WIP.

    Supported Rule Type values:
    - Production Site Strict Rule Match
    - Production Site Strict
    - Strict Production Site
    - Default Exclude
    - Exclude If No Match

    Site value can be maintained in either Key or Production Site.
    """
    strict_sites: set[str] = set()
    rules: pd.DataFrame = masters.get("rules", pd.DataFrame())
    for _, row in rules.iterrows():
        rt = normalize_rule_type(row.get("Rule Type", ""))
        if rt not in [
            "production site strict rule match",
            "production site strict",
            "strict production site",
            "default exclude",
            "exclude if no match",
            "廠區嚴格規則",
            "生產廠區嚴格規則",
            "未命中排除",
        ]:
            continue

        for col in ["Production Site", "Key"]:
            site = str(row.get(col, "") or "").strip()
            if site and site not in ["*", "DEFAULT"]:
                strict_sites.add(_normalize_production_site_key(site))
    return strict_sites


def should_exclude_unmatched_by_strict_site(plant: object, masters: dict, current_production_site: object = "") -> tuple[bool, str, str]:
    """Return whether an unmatched row should be excluded by strict-site rule."""
    strict_sites = get_strict_production_sites(masters)
    if not strict_sites:
        return False, "", ""

    production_site = str(current_production_site or "").strip()
    plant_rule_hit = ""
    if not production_site:
        production_site, plant_rule_hit = resolve_plant_production_site_from_rule_master(plant, masters)

    site_key = _normalize_production_site_key(production_site)
    if site_key and site_key in strict_sites:
        return True, production_site, plant_rule_hit
    return False, production_site, plant_rule_hit


def _rule_site_is_compatible(rule_site: object, current_production_site: object) -> bool:
    """Return whether a classification rule may apply to the current plant's Production Site.

    Blank Production Site on a classification rule means the rule is generic and may be used
    by multiple sites. When a rule has a Production Site value, it may only apply to rows
    whose Plant has been resolved to the same Production Site. This prevents a site-specific
    DT/NB/TP rule from assigning the wrong site to another plant while still allowing shared
    generic product-line rules such as DT Mouse/DT Keyboard.
    """
    rule_site_key = _normalize_production_site_key(rule_site)
    if not rule_site_key:
        return True
    current_site_key = _normalize_production_site_key(current_production_site)
    if not current_site_key:
        return True
    return rule_site_key == current_site_key



def _material_number_prefix(value: object, length: int) -> str:
    """Return the first N non-space characters of a material number."""
    text = str(value or "").upper().strip()
    text = re.sub(r"\s+", "", text)
    if length <= 0:
        return text
    return text[:length]


def _as_int_or_default(value: object, default: int = 2) -> int:
    try:
        return int(float(str(value or "").strip()))
    except Exception:
        return default


def get_site_material_prefix_whitelist(masters: dict, current_production_site: object = "") -> list[dict]:
    """Read optional Production Site material-prefix whitelist rules from rule_master.csv.

    Supported Rule Type values:
    - Production Site Material Prefix Whitelist
    - Site Material Prefix Whitelist
    - Material Number Prefix Whitelist
    - Site Prefix Whitelist

    If any whitelist rows exist for the current Production Site, a row must match one
    of the maintained prefixes before normal classification is allowed to continue.
    This is generic and BU-neutral: each BU/site controls its own whitelist in Rule Master.

    Optional Prefix Length values:
    - PREFIX2, FIRST2, 2 => compare first 2 characters.
    - PREFIX3, FIRST3, 3 => compare first 3 characters.
    - FULL, EXACT => compare the full Key length.
    Blank defaults to first 2 characters for backward compatibility with current IPS rules.
    """
    site_key = _normalize_production_site_key(current_production_site)
    if not site_key:
        return []

    rows: list[dict] = []
    rules: pd.DataFrame = masters.get("rules", pd.DataFrame())
    for _, row in rules.iterrows():
        rt = normalize_rule_type(row.get("Rule Type", ""))
        if rt not in [
            "production site material prefix whitelist",
            "site material prefix whitelist",
            "material number prefix whitelist",
            "site prefix whitelist",
            "廠區料號前綴白名單",
            "料號前綴白名單",
        ]:
            continue

        rule_site = str(row.get("Production Site", "") or "").strip()
        if not _rule_site_is_compatible(rule_site, current_production_site):
            continue

        key = str(row.get("Key", "") or "").upper().strip()
        if not key or key in ["*", "DEFAULT"]:
            continue

        logic = str(row.get("Prefix Length", "") or "").strip().upper()
        if logic in ["PREFIX3", "FIRST3", "3"]:
            prefix_len = 3
        elif logic in ["PREFIX4", "FIRST4", "4"]:
            prefix_len = 4
        elif logic in ["FULL", "EXACT"]:
            prefix_len = len(key)
        else:
            prefix_len = _as_int_or_default(logic, 2)

        rows.append({
            "key": key,
            "prefix_len": prefix_len,
            "production_site": rule_site,
            "rule_type": str(row.get("Rule Type", "") or "").strip(),
        })

    return rows


def enforce_site_material_prefix_whitelist(material_number: object, masters: dict, current_production_site: object = "") -> dict:
    """Apply Rule Master-driven site material-prefix whitelist.

    When a Production Site has whitelist rows, only allowed material-number prefixes
    may continue into the normal classifier. Non-whitelisted rows are excluded before
    generic Series Prefix / Description Contains rules can accidentally classify products
    from other BUs at shared factories.
    """
    whitelist = get_site_material_prefix_whitelist(masters, current_production_site)
    if not whitelist:
        return {}

    material_u = str(material_number or "").upper().strip()
    material_compact = re.sub(r"\s+", "", material_u)
    for item in whitelist:
        prefix = _material_number_prefix(material_compact, item["prefix_len"])
        if prefix == item["key"]:
            return {
                "_site_prefix_whitelist_allowed": True,
                "_site_prefix_whitelist_rule": f"{item['rule_type']}={item['key']}",
            }

    prefixes = sorted({item["key"] for item in whitelist})
    return {
        "產品類型": "",
        "Product Line": "",
        "Production Site": str(current_production_site or "").strip(),
        "判斷來源": "Rule Master",
        "規則判定結果": "Excluded",
        "命中規則": "Material Number prefix not in Production Site whitelist: "
                  + (str(material_number or "").strip() or "(blank)")
                  + " not in "
                  + "/".join(prefixes),
        "Is_WIP": "N",
        "_exclude": True,
        "_site_prefix_whitelist_excluded": True,
    }


def classify_by_rule_master(material_number: object, description: object, series: str, masters: dict, current_production_site: object = "") -> dict:
    material_number_u = str(material_number or "").upper().strip()
    description_u = str(description or "").upper().strip()
    series_u = str(series or "").upper().strip()

    rules: pd.DataFrame = masters["rules"]
    for _, row in rules.iterrows():
        rule_type = str(row.get("Rule Type", "") or "").strip()
        key = str(row.get("Key", "") or "").strip()

        # Default 規則只作為文件與下載範本保留；實際 Default WIP 放在 Product Series Master 之後。
        if normalize_rule_type(rule_type) in ["default", "預設", "default exclude", "exclude if no match", "未命中排除"]:
            continue

        if not rule_matches(rule_type, key, material_number_u, description_u, series_u):
            continue

        product_type = str(row.get("Product Type", "") or "").strip()
        is_wip = str(row.get("Is_WIP", "") or "").upper().strip()
        if not is_wip:
            is_wip = "Y" if product_type.upper() == "WIP" else "N"

        # Rules with blank Product Type are markers/metadata, not final classification.
        # Example: 越南海防廠-IPS prefix 90 is only a finished-product whitelist;
        # it must continue through the normal Rule Master / Product Series Master flow.
        if not product_type:
            continue

        # Rule Master only mode:
        # Product Line and Production Site must come from rule_master.csv.
        # Do not auto-fill Product Line from Product Type.
        product_line = str(row.get("Product Line", "") or "").strip()
        production_site = str(row.get("Production Site", "") or "").strip()
        if not _rule_site_is_compatible(production_site, current_production_site):
            continue

        return {
            "產品類型": product_type,
            "Product Line": product_line,
            "Production Site": production_site,
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
    # Rule Master only mode:
    # Product Series Master may classify Product Type only.
    # Product Line / Production Site are not inferred here.
    product_line = ""
    return {
        "產品類型": product_type,
        "Product Line": product_line,
        "Production Site": "",
        "判斷來源": "Product Series Master",
        "規則判定結果": "符合" if product_type else "待補產品分類",
        "命中規則": f"Product series={series_u}",
        "Is_WIP": "Y" if product_type.upper() == "WIP" else "N",
    }


def classify(material_number: object, description: object, series: str, plant: object, masters: dict, current_production_site: object = "") -> dict:
    whitelist_result = enforce_site_material_prefix_whitelist(material_number, masters, current_production_site)
    if whitelist_result.get("_exclude"):
        return whitelist_result

    result = classify_by_rule_master(material_number, description, series, masters, current_production_site)
    if result.get("產品類型"):
        return result

    result = classify_by_series_master(plant, series, masters)
    if result.get("產品類型"):
        return result

    exclude, strict_site, plant_rule_hit = should_exclude_unmatched_by_strict_site(plant, masters, current_production_site)
    if exclude:
        return {
            "產品類型": "",
            "Product Line": "",
            "Production Site": strict_site,
            "判斷來源": "Rule Master",
            "規則判定結果": "Excluded",
            "命中規則": f"No rule matched → Excluded by strict Production Site ({strict_site})" + (f" | {plant_rule_hit}" if plant_rule_hit else ""),
            "Is_WIP": "N",
            "_exclude": True,
        }

    return {
        "產品類型": "WIP",
        "Product Line": "",
        "Production Site": "",
        "判斷來源": "Default WIP",
        "規則判定結果": "WIP",
        "命中規則": "No rule matched → WIP",
        "Is_WIP": "Y",
        "_exclude": False,
    }





def is_finished_product_whitelist(material_number: object) -> bool:
    """Deprecated no-op. Finished product handling is controlled by rule_master.csv."""
    return False


def is_wip_by_rule_master(material_number: object, description: object, series: str, masters: dict) -> bool:
    """Deprecated no-op. Rule Master priority is now the single source of truth."""
    return False


def infer_product_type_line_site_from_series_rules(description: object, series: str, masters: dict) -> tuple[str, str, str]:
    """Infer finished-product classification from rule_master.csv only.

    SG- only means non-WIP. Product Type / Product Line / Production Site
    must still come from matched rule_master.csv fields.
    No fallback from Product Type to Product Line, and no fallback from Product Line to Production Site.
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
            continue

        product_line = str(row.get("Product Line", "") or "").strip()
        production_site = str(row.get("Production Site", "") or "").strip()

        if product_type or product_line or production_site:
            return product_type, product_line, production_site

    return "", "", ""

def infer_product_line_site_from_rules(description: object, series: str, masters: dict) -> tuple[str, str]:
    """Infer Product Line / Production Site for WIP from rule_master.csv only.

    WIP Product Type remains WIP. Product Line and Production Site may be filled
    only when the matched rule_master.csv row explicitly provides them.
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
        production_site = str(row.get("Production Site", "") or "").strip()

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
    """Load production labor work order files and aggregate labor hours by Order first,
    with Plant + Material Number retained as fallback merge keys.
    """
    if not paths:
        return pd.DataFrame(
            columns=[
                "Order", "Order Merge Key", "Plant", "Material Number",
                "Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"
            ]
        )

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
        return pd.DataFrame(
            columns=[
                "Order", "Order Merge Key", "Plant", "Material Number",
                "Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"
            ]
        )

    labor = pd.concat(frames, ignore_index=True)
    labor = labor[labor["Order Merge Key"].astype(str).str.strip() != ""].copy()

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
    """Attach labor hours to production output.

    Primary key only:
    - Production quantity file Order
    - Working-hour file Order Number / Order

    Both sides are matched by normalized Order Merge Key.
    No Plant + Material Number fallback is used.
    """
    out = out.copy()
    if "Order Merge Key" not in out.columns:
        out["Order Merge Key"] = out["Order"].apply(normalize_order_key)

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"]:
        if col not in out.columns:
            out[col] = 0 if col != "Labor Source files" else ""

    if labor is None or labor.empty:
        return out

    if "Order Merge Key" not in labor.columns:
        labor = labor.copy()
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

    if not order_labor.empty:
        out = out.merge(order_labor, on="Order Merge Key", how="left", suffixes=("", "_labor"))
        for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"]:
            labor_col = f"{col}_labor"
            if labor_col in out.columns:
                if col == "Labor Source files":
                    out[col] = out[labor_col].fillna("").astype(str)
                else:
                    out[col] = pd.to_numeric(out[labor_col], errors="coerce")
                out = out.drop(columns=[labor_col])

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)

    for col in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0)
    out["Labor Source files"] = out["Labor Source files"].fillna("").astype(str)

    if "Matched Labor Order Number" not in out.columns:
        out["Matched Labor Order Number"] = ""
    out["Matched Labor Order Number"] = out["Matched Labor Order Number"].fillna("").astype(str)
    out["Labor Match Status"] = out["Matched Labor Order Number"].apply(
        lambda x: "Matched" if str(x).strip() else "Not matched"
    )

    return out

def process_files(
    paths: list[Path],
    year: Optional[int],
    labor_paths: Optional[list[Path]] = None,
    labor_mode: str = "both",
    rule_set: str = DEFAULT_RULE_SET,
) -> tuple[Path, dict]:
    rule_set = normalize_rule_set(rule_set)
    masters = build_masters(rule_set)
    labor_mode = normalize_labor_mode(labor_mode)

    out = load_production_dataframe(paths)
    labor = load_labor_dataframe(labor_paths or [], labor_mode)
    out = attach_labor_hours(out, labor)

    # Final safety guard: total working hours must always follow selected working-hour source.
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

    # Resolve Plant -> Production Site before classification.
    # This makes Rule Master classification plant-aware:
    # - site-specific rules only apply to the same Production Site
    # - blank-site rules remain generic and can be used by both 越南海防廠-IPS and 廣州石碣廠-IPS
    initial_plant_site_rules = out["Plant"].apply(lambda p: resolve_plant_production_site_from_rule_master(p, masters))
    out["_Plant Production Site"] = initial_plant_site_rules.apply(lambda x: x[0])
    out["_Plant Production Site Rule"] = initial_plant_site_rules.apply(lambda x: x[1])

    classified = out.apply(
        lambda r: classify(
            r["Material Number"],
            r["Material description"],
            r["Product series"],
            r["Plant"],
            masters,
            r.get("_Plant Production Site", ""),
        ),
        axis=1,
    )
    out["產品類型"] = classified.apply(lambda x: x.get("產品類型", ""))
    out["Product Line"] = classified.apply(lambda x: x.get("Product Line", ""))
    out["Production Site"] = classified.apply(lambda x: x.get("Production Site", ""))
    out["判斷來源"] = classified.apply(lambda x: x.get("判斷來源", ""))
    out["規則判定結果"] = classified.apply(lambda x: x.get("規則判定結果", ""))
    out["命中規則"] = classified.apply(lambda x: x.get("命中規則", ""))
    out["Is_WIP"] = classified.apply(lambda x: x.get("Is_WIP", "N"))
    # Generic rule-engine behavior:
    # If Rule Master classifies a row as WIP and does not explicitly provide Product Line,
    # keep Product Line and Product Series blank. This prevents downstream inference from
    # re-attaching product-family attributes to WIP/SFG rows.
    wip_without_line_mask = (
        out["Is_WIP"].astype(str).str.upper().isin(["Y", "YES", "TRUE", "1"])
        & out["Product Line"].astype(str).str.strip().eq("")
    )
    if wip_without_line_mask.any():
        out.loc[wip_without_line_mask, "Product series"] = ""
        out.loc[wip_without_line_mask, "Product Line"] = ""
        out.loc[wip_without_line_mask, "解析說明"] = "Rule Master：WIP且未指定Product Line，不使用Product Series"

    # If a generic product classification rule is used, keep its Product Type/Product Line
    # and fill Production Site from Plant Exact/Prefix rules. This prevents shared DT rules
    # from forcing all outputs to 越南海防廠-IPS when 石碣廠 uses the same product types.
    blank_site_mask = out["Production Site"].astype(str).str.strip().eq("")
    plant_site_available_mask = out["_Plant Production Site"].astype(str).str.strip().ne("")
    fill_site_mask = blank_site_mask & plant_site_available_mask
    if fill_site_mask.any():
        out.loc[fill_site_mask, "Production Site"] = out.loc[fill_site_mask, "_Plant Production Site"].to_numpy()

    strict_site_excluded_rows = int(classified.apply(lambda x: bool(x.get("_exclude", False))).sum())
    if strict_site_excluded_rows:
        keep_mask = ~classified.apply(lambda x: bool(x.get("_exclude", False)))
        out = out.loc[keep_mask].copy().reset_index(drop=True)

    # Rule Master only mode:
    # Missing Product Line / Production Site may be filled only from explicit rule_master.csv rows.
    missing_line_mask = out["Product Line"].astype(str).str.strip().eq("")
    missing_line_mask = missing_line_mask & ~out["Is_WIP"].astype(str).str.upper().isin(["Y", "YES", "TRUE", "1"])
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

    # Plant Rule Master override:
    # Production Site may be controlled directly by Plant Exact / Plant Prefix rules in rule_master.csv.
    # This only updates Production Site and does not change Product Type / Product Line / WIP status.
    plant_site_rules = out["Plant"].apply(lambda p: resolve_plant_production_site_from_rule_master(p, masters))
    plant_sites = plant_site_rules.apply(lambda x: x[0])
    plant_rule_hits = plant_site_rules.apply(lambda x: x[1])

    plant_site_mask = plant_sites.astype(str).str.strip().ne("")
    if plant_site_mask.any():
        out.loc[plant_site_mask, "Production Site"] = plant_sites.loc[plant_site_mask].to_numpy()
        out.loc[plant_site_mask, "命中規則"] = out.loc[plant_site_mask, "命中規則"].astype(str) + " | " + plant_rule_hits.loc[plant_site_mask].astype(str)
        out.loc[plant_site_mask, "判斷來源"] = out.loc[plant_site_mask, "判斷來源"].astype(str).where(
            out.loc[plant_site_mask, "判斷來源"].astype(str).str.strip().ne(""),
            "Rule Master"
        )

    out = out.drop(columns=["_Plant Production Site", "_Plant Production Site Rule"], errors="ignore")

    group_cols = [
        "Year", "Plant", "Production Site", "Product Line", "Material Number", "Material description", "Product series",
        "產品類型", "判斷來源", "Is_WIP"
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


    source_summary = (
        out.groupby(["判斷來源", "規則判定結果", "命中規則"], dropna=False, as_index=False)["Delivered quantity"]
        .agg(筆數="count", 生產量="sum")
    )

    file_summary = (
        out.groupby(["Source file"], dropna=False, as_index=False)["Delivered quantity"]
        .agg(筆數="count", 生產量="sum")
        .sort_values(["Source file"])
    )

    # Labor source file summary: helps verify whether each labor file was loaded and how many hours were read.
    if labor is not None and not labor.empty:
        labor_source_summary = (
            labor.copy()
            .assign(**{
                "Labor HR.Act": pd.to_numeric(labor["Labor HR.Act"], errors="coerce").fillna(0),
                "FOH-Others.Act": pd.to_numeric(labor["FOH-Others.Act"], errors="coerce").fillna(0),
                "Selected Hours": pd.to_numeric(labor["Selected Hours"], errors="coerce").fillna(0),
            })
            .groupby(["Labor Source files"], dropna=False, as_index=False)
            .agg({
                "Order Merge Key": "count",
                "Labor HR.Act": "sum",
                "FOH-Others.Act": "sum",
                "Selected Hours": "sum",
            })
            .rename(columns={
                "Labor Source files": "Labor Source file",
                "Order Merge Key": "工時Order數",
                "Labor HR.Act": "Labor HR.Act合計",
                "FOH-Others.Act": "FOH-Others.Act合計",
                "Selected Hours": "Selected Hours合計",
            })
        )
    else:
        labor_source_summary = pd.DataFrame(columns=[
            "Labor Source file", "工時Order數", "Labor HR.Act合計", "FOH-Others.Act合計", "Selected Hours合計"
        ])

    # Labor matching diagnostics: shows how production Orders matched labor Order Numbers.
    diagnostic_cols = [
        "Source file", "Order", "Order Merge Key", "Matched Labor Order Number",
        "Labor Match Status", "Material Number", "Material description",
        "Labor HR.Act", "FOH-Others.Act", "Selected Hours", "Labor Source files"
    ]
    for c in diagnostic_cols:
        if c not in out.columns:
            out[c] = ""
    labor_match_diagnostics = out[diagnostic_cols].copy()
    for c in ["Labor HR.Act", "FOH-Others.Act", "Selected Hours"]:
        labor_match_diagnostics[c] = pd.to_numeric(labor_match_diagnostics[c], errors="coerce").fillna(0)
    labor_match_diagnostics = labor_match_diagnostics.rename(columns={
        "Selected Hours": "Total working hours"
    })

    labor_match_summary = (
        labor_match_diagnostics.groupby(["Labor Match Status"], dropna=False, as_index=False)["Order"]
        .count()
        .rename(columns={"Order": "筆數"})
    )

    wip = out[out["Is_WIP"] == "Y"].copy()

    file_id = uuid.uuid4().hex[:10]
    output_path = OUTPUT_DIR / f"年度產品產量與分類結果_v6_{year or 'ALL'}_{file_id}.xlsx"

    # Export view: only show the selected working-hour source in Excel.
    out_export = out.copy()
    annual_export = annual.copy()

    # Rename Excel output header only; keep internal calculation column as Selected Hours.
    out_export = out_export.rename(columns={"Selected Hours": "Total working hours"})

    if labor_mode == "labor_hr":
        out_export = out_export.drop(columns=["FOH-Others.Act"], errors="ignore")
        annual_export = annual_export.drop(columns=["年度設備工時"], errors="ignore")
    elif labor_mode == "foh":
        out_export = out_export.drop(columns=["Labor HR.Act"], errors="ignore")
        annual_export = annual_export.drop(columns=["年度人員工時"], errors="ignore")

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        out_export.to_excel(writer, index=False, sheet_name="工單明細_已分類")
        annual_export.to_excel(writer, index=False, sheet_name="Plant_Material年度產量")
        type_summary.to_excel(writer, index=False, sheet_name="Plant_產品類型年度產量")
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
        "rule_set": rule_set,
        "rows": int(len(out)),
        "annual_rows": int(len(annual)),
        "total_qty": float(out["Delivered quantity"].sum()),
        "total_hours": float(out["Selected Hours"].sum()) if "Selected Hours" in out.columns else 0.0,
        "wip_rows": int(len(wip)),
        "strict_site_excluded_rows": int(strict_site_excluded_rows),
        "output_filename": output_path.name,
        "year": year or "ALL",
    }
    return output_path, summary

def process_file(path: Path, year: Optional[int]) -> tuple[Path, dict]:
    """Backward-compatible wrapper for single-file processing."""
    return process_files([path], year, None, "both", DEFAULT_RULE_SET)


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
        elif key in ["prefix length", "prefix_len", "match length", "match_length", "前綴長度", "比對長度"]:
            rename_map[c] = "Prefix Length"
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
    # Product Type may be blank for rules that only maintain Production Site,
    # e.g. Plant Exact, 3760 -> Production Site.
    df = df[(df["Rule Type"] != "") & (df["Key"] != "")]
    return df


def save_uploaded_rule(file_path: Path, rule_set: object = DEFAULT_RULE_SET) -> int:
    if file_path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        df = pd.read_excel(file_path, dtype=str).fillna("")
    else:
        df = pd.read_csv(file_path, dtype=str, encoding="utf-8-sig").fillna("")
    df = normalize_rule_upload(df)
    target_dir = get_rule_set_dir(rule_set)
    target_dir.mkdir(parents=True, exist_ok=True)
    df.to_csv(target_dir / "rule_master.csv", index=False, encoding="utf-8-sig")
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
        "supports": ["files multi-upload", "file single-upload", "Module 2 multi-BOM upload", "blank year", "BU rule library"],
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
    rule_set = normalize_rule_set(form.get("rule_set") or DEFAULT_RULE_SET)
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

        output_path, summary = process_files(saved_paths, year_value, saved_labor_paths, labor_mode, rule_set)
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
    working_hour_source: str = Form("direct"),
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

    working_hour_source = str(working_hour_source or "direct").strip()
    bom_structure_path = None
    working_hour_rollup_path = None
    if working_hour_source in ["include_semi", "semi", "semi_finished", "rollup", "rolled_up", "total"]:
        if not LATEST_WORKING_HOUR_ROLLUP_PATH.exists():
            return JSONResponse({
                "ok": False,
                "message": "No Working Hour Roll-up result found. Please complete Module 2 → BOM Expansion with Step 1 Output first, then return to Step 2 to generate Product Activity Data Bulk."
            }, status_code=400)
        working_hour_rollup_path = LATEST_WORKING_HOUR_ROLLUP_PATH
        bom_structure_path = LATEST_BOM_STRUCTURE_PATH if LATEST_BOM_STRUCTURE_PATH.exists() else None

    try:
        summary = generate_product_activity_bulk_files_by_site_zip(
            step1_output_path=step1_path,
            bulk_template_path=template_path,
            output_dir=OUTPUT_DIR,
            token=token,
            working_hour_source=working_hour_source,
            bom_structure_path=bom_structure_path,
            working_hour_rollup_path=working_hour_rollup_path,
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
async def upload_rule_master(request: Request):
    try:
        form = await request.form()
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": f"無法讀取上傳表單：{exc}"}, status_code=400)

    file = form.get("file")
    rule_set = normalize_rule_set(form.get("rule_set") or DEFAULT_RULE_SET)

    if not file or not getattr(file, "filename", None):
        return JSONResponse({"ok": False, "message": "請上傳 Rule Master 檔案"}, status_code=400)

    if not file.filename.lower().endswith((".xlsx", ".xlsm", ".xls", ".csv")):
        return JSONResponse({"ok": False, "message": "請上傳 Excel 或 CSV 檔案"}, status_code=400)

    saved = UPLOAD_DIR / f"{uuid.uuid4().hex}_{file.filename}"
    saved.write_bytes(await file.read())
    try:
        count = save_uploaded_rule(saved, rule_set)
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)
    return {"ok": True, "count": count, "rule_set": rule_set}


@app.get("/download/{filename}")
def download(filename: str):
    path = OUTPUT_DIR / filename
    if not path.exists():
        return JSONResponse({"ok": False, "message": "檔案不存在"}, status_code=404)
    media_type = "application/zip" if path.suffix.lower() == ".zip" else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    return FileResponse(path, filename=filename, media_type=media_type)


@app.get("/download-rule-master")
def download_rule_master(rule_set: str = DEFAULT_RULE_SET):
    ensure_master_files()
    rule_set = normalize_rule_set(rule_set)
    path = get_rule_set_dir(rule_set) / "rule_master.csv"
    return FileResponse(path, filename=f"rule_master_{rule_set}.csv", media_type="text/csv")


@app.get("/download-product-series-master")
def download_product_series_master(rule_set: str = DEFAULT_RULE_SET):
    ensure_master_files()
    rule_set = normalize_rule_set(rule_set)
    path = get_rule_set_dir(rule_set) / "product_series_master.csv"
    return FileResponse(path, filename=f"product_series_master_{rule_set}.csv", media_type="text/csv")



# =========================================================
# Module 3 · Carbon Emission Factor Selection
# CCL Mapping + Factor Library Search
# =========================================================

@app.get("/module3/latest-raw-material-bulk")
def module3_latest_raw_material_bulk():
    if not LATEST_RAW_MATERIAL_BULK_PATH.exists():
        return {"ok": True, "available": False, "message": "尚未找到 Module 2 最新產出的 raw_material_activity_data_bulk 檔案。"}
    stat = LATEST_RAW_MATERIAL_BULK_PATH.stat()
    return {
        "ok": True,
        "available": True,
        "filename": LATEST_RAW_MATERIAL_BULK_PATH.name,
        "size": stat.st_size,
        "updated_at": datetime.fromtimestamp(stat.st_mtime).isoformat(timespec="seconds"),
        "download_url": f"/download/{LATEST_RAW_MATERIAL_BULK_PATH.name}",
    }

@app.post("/module3/apply-ccl-factors-job")
async def module3_apply_ccl_factors_job(
    ccl_mapping_file: UploadFile = File(...),
    raw_material_file: UploadFile | None = File(None),
):
    token = uuid.uuid4().hex[:10]
    job_id = token

    ccl_filename = str(getattr(ccl_mapping_file, "filename", "") or "")
    if not ccl_filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "CCL 係數組配表 請上傳 Excel 檔案"}, status_code=400)

    if raw_material_file is not None and getattr(raw_material_file, "filename", None):
        raw_filename = str(getattr(raw_material_file, "filename", "") or "")
        if not raw_filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse({"ok": False, "message": "Raw Material Bulk 請上傳 Excel 檔案"}, status_code=400)
        raw_path = UPLOAD_DIR / f"module3_raw_material_{token}_{Path(raw_filename).name}"
        raw_path.write_bytes(await raw_material_file.read())
        raw_source = "uploaded"
    else:
        raw_path = LATEST_RAW_MATERIAL_BULK_PATH
        raw_source = "module2_latest"
        if not raw_path.exists():
            return JSONResponse(
                {"ok": False, "message": "找不到 Module 2 最新產出的 raw_material_activity_data_bulk 檔案，請先完成 Module 2 BOM Expansion。"},
                status_code=400,
            )

    ccl_path = UPLOAD_DIR / f"module3_ccl_mapping_{token}_{Path(ccl_filename).name}"
    output_path = OUTPUT_DIR / f"module3_ccl_factor_filled_{token}.xlsx"
    ccl_path.write_bytes(await ccl_mapping_file.read())

    _set_module3_ccl_job(
        job_id,
        status="queued",
        progress=0,
        step="工作已建立，等待背景處理",
        message="CCL 係數對應已開始。",
        created_at=datetime.now().isoformat(timespec="seconds"),
        started_at=datetime.now().isoformat(timespec="seconds"),
        remaining_seconds=30,
        raw_source=raw_source,
        raw_material_filename=raw_path.name,
    )
    MODULE3_CCL_EXECUTOR.submit(_run_module3_ccl_job, job_id, raw_path, ccl_path, output_path)
    return {"ok": True, "job_id": job_id, "message": "CCL 係數對應已開始。", "raw_source": raw_source, "raw_material_filename": raw_path.name}


@app.get("/module3/ccl-job/{job_id}")
def module3_get_ccl_job(job_id: str):
    job = MODULE3_CCL_JOBS.get(job_id)
    if not job:
        return JSONResponse({"ok": False, "message": "找不到 CCL 對應工作，請重新執行。"}, status_code=404)
    return {"ok": True, "job": job}


@app.post("/module3/apply-ccl-factors")
async def module3_apply_ccl_factors(
    ccl_mapping_file: UploadFile = File(...),
    raw_material_file: UploadFile | None = File(None),
):
    token = uuid.uuid4().hex[:10]

    ccl_filename = str(getattr(ccl_mapping_file, "filename", "") or "")
    if not ccl_filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse({"ok": False, "message": "CCL 係數組配表 請上傳 Excel 檔案"}, status_code=400)

    if raw_material_file is not None and getattr(raw_material_file, "filename", None):
        raw_filename = str(getattr(raw_material_file, "filename", "") or "")
        if not raw_filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse({"ok": False, "message": "Raw Material Bulk 請上傳 Excel 檔案"}, status_code=400)
        raw_path = UPLOAD_DIR / f"module3_raw_material_{token}_{Path(raw_filename).name}"
        raw_path.write_bytes(await raw_material_file.read())
        raw_source = "uploaded"
    else:
        raw_path = LATEST_RAW_MATERIAL_BULK_PATH
        raw_source = "module2_latest"
        if not raw_path.exists():
            return JSONResponse({"ok": False, "message": "找不到 Module 2 最新產出的 raw_material_activity_data_bulk 檔案，請先完成 Module 2 BOM Expansion。"}, status_code=400)

    ccl_path = UPLOAD_DIR / f"module3_ccl_mapping_{token}_{Path(ccl_filename).name}"
    output_path = OUTPUT_DIR / f"module3_ccl_factor_filled_{token}.xlsx"
    ccl_path.write_bytes(await ccl_mapping_file.read())

    try:
        summary = apply_ccl_factors_to_raw_material_bulk(raw_path, ccl_path, output_path)
        summary["app_version"] = "CMP_MODULE3_CCL_SYNC_V2_1"
        summary["raw_source"] = raw_source
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)

    return {
        "ok": True,
        "message": "CCL 係數對應完成。",
        "summary": summary,
        "download_url": summary.get("download_url", f"/download/{output_path.name}"),
    }


def _module3_factor_library_paths():
    return (
        FACTOR_LIBRARY_DIR / "APOS Cumulative LCIA v3.12(顧問).xlsx",
        FACTOR_LIBRARY_DIR / "Cut-off Cumulative LCIA v3.12(顧問).xlsx",
    )


@app.on_event("startup")
def module3_preload_factor_libraries_on_startup():
    """Preload APOS / Cut-off factor libraries into memory for faster Module 3 searches."""
    try:
        apos_path, cutoff_path = _module3_factor_library_paths()
        summary = preload_factor_libraries(apos_path, cutoff_path)
        print(f"===== MODULE3 FACTOR LIBRARY PRELOAD: {summary} =====")
    except Exception as exc:
        print(f"===== MODULE3 FACTOR LIBRARY PRELOAD FAILED: {exc} =====")


@app.get("/module3/factor-library-filters")
def module3_factor_library_filters():
    apos_path, cutoff_path = _module3_factor_library_paths()
    try:
        geographies = ["GLO", "RoW", "RER"]
        return {
            "ok": True,
            "geographies": geographies,
            "sources": ["APOS", "Cut-off"],
            "process_types": ["production", "market_for"],
            "app_version": "CMP_MODULE3_STAGE2_V21",
        }
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)


@app.get("/module3/search-factor-library")
def module3_search_factor_library(
    keyword: str = "",
    activity_name_keyword: str = "",
    reference_product_keyword: str = "",
    limit: int = 10,
    source: str = "all",
    geography: str = "all",
    process_type: str = "all",
    page: int = 1,
    page_size: int = 10,
):
    apos_path, cutoff_path = _module3_factor_library_paths()
    try:
        result = search_factor_library(
            keyword,
            apos_path,
            cutoff_path,
            limit=limit,
            source=source,
            geography=geography,
            process_type=process_type,
            page=page,
            page_size=page_size,
            activity_name_keyword=activity_name_keyword,
            reference_product_keyword=reference_product_keyword,
        )
        result["ok"] = True
        result["app_version"] = "CMP_MODULE3_STAGE2_V20"
        return result
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": str(exc)}, status_code=400)



# =========================================================
# Module 2 · BOM Expansion
# Standard BOM + Raw Material Bulk Template -> Raw Material Bulk
# =========================================================
@app.post("/process-bom-expansion")
async def process_bom_expansion(request: Request):
    """Module 2 BOM Expansion.

    V13 supports multiple Standard BOM Excel files.
    Accepted file field names:
    - bom_files: new multi-upload field
    - bom_file: backward-compatible single-upload field
    """
    try:
        form = await request.form()
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse({"ok": False, "message": f"無法讀取上傳表單：{exc}"}, status_code=400)

    def is_upload_file_like(item) -> bool:
        return bool(getattr(item, "filename", None)) and hasattr(item, "read")

    bom_uploads = []
    for item in form.getlist("bom_files") + form.getlist("bom_file"):
        if is_upload_file_like(item):
            bom_uploads.append(item)

    template_file = form.get("template_file")
    step1_file = form.get("step1_file")

    supplier_uploads = []
    for item in form.getlist("supplier_files") + form.getlist("supplier_file"):
        if is_upload_file_like(item):
            supplier_uploads.append(item)

    parent_col = str(form.get("parent_col") or "")
    component_col = str(form.get("component_col") or "")
    qty_col = str(form.get("qty_col") or "")
    unit_col = str(form.get("unit_col") or "")
    description_col = str(form.get("description_col") or "")
    material_group_col = str(form.get("material_group_col") or "")
    valid_from_col = str(form.get("valid_from_col") or "")

    if not bom_uploads:
        return JSONResponse(
            {"ok": False, "message": "請至少上傳一個 Standard BOM Excel 檔案"},
            status_code=400,
        )

    for bom_file in bom_uploads:
        filename = str(getattr(bom_file, "filename", "") or "")
        if not filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse(
                {"ok": False, "message": f"{filename} 不是 Standard BOM Excel 檔案"},
                status_code=400,
            )

    if not template_file or not getattr(template_file, "filename", None):
        return JSONResponse(
            {"ok": False, "message": "請上傳 Raw Material Bulk Template"},
            status_code=400,
        )

    if not template_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return JSONResponse(
            {"ok": False, "message": "Raw Material Bulk Template 請上傳 Excel 檔案"},
            status_code=400,
        )

    if step1_file is not None and getattr(step1_file, "filename", None):
        if not step1_file.filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse(
                {"ok": False, "message": "Step 1 Output 請上傳 Excel 檔案"},
                status_code=400,
            )

    for supplier_file in supplier_uploads:
        filename = str(getattr(supplier_file, "filename", "") or "")
        if filename and not filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
            return JSONResponse(
                {"ok": False, "message": f"{filename} 不是 Supplier Excel 檔案"},
                status_code=400,
            )

    token = uuid.uuid4().hex[:10]

    bom_paths: list[Path] = []
    for idx, bom_file in enumerate(bom_uploads, start=1):
        filename = str(getattr(bom_file, "filename", "") or f"bom_{idx}.xlsx")
        saved_bom = UPLOAD_DIR / f"standard_bom_{token}_{idx}_{Path(filename).name}"
        saved_bom.write_bytes(await bom_file.read())
        bom_paths.append(saved_bom)

    template_path = UPLOAD_DIR / f"raw_material_template_{token}_{Path(template_file.filename).name}"
    output_path = OUTPUT_DIR / f"raw_material_activity_data_bulk_{token}.xlsx"
    working_hour_rollup_output_path = OUTPUT_DIR / f"working_hour_rollup_{token}.xlsx"
    supplier_bulk_template_path = BASE_DIR / "templates" / "supplier_bulk_create_template_v1.xlsx"
    supplier_bulk_output_path = OUTPUT_DIR / f"supplier_bulk_create_{token}.xlsx"
    step1_path = None
    supplier_paths: list[Path] = []

    template_path.write_bytes(await template_file.read())
    if step1_file is not None and getattr(step1_file, "filename", None):
        step1_path = UPLOAD_DIR / f"step1_for_rollup_{token}_{Path(step1_file.filename).name}"
        step1_path.write_bytes(await step1_file.read())

    for idx, supplier_file in enumerate(supplier_uploads, start=1):
        filename = str(getattr(supplier_file, "filename", "") or f"supplier_{idx}.xlsx")
        supplier_path = UPLOAD_DIR / f"supplier_{token}_{idx}_{Path(filename).name}"
        supplier_path.write_bytes(await supplier_file.read())
        supplier_paths.append(supplier_path)

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
        if step1_path is not None:
            summary = generate_raw_material_bulk_files_by_site_zip(
                bom_path=bom_paths,
                raw_material_template_path=template_path,
                output_dir=OUTPUT_DIR,
                token=token,
                step1_output_path=step1_path,
                mapping=mapping,
                supplier_paths=supplier_paths,
                supplier_bulk_template_path=supplier_bulk_template_path,
                supplier_bulk_output_path=supplier_bulk_output_path,
            )
            output_path = OUTPUT_DIR / str(summary.get("output_filename", f"raw_material_activity_data_bulk_by_site_{token}.zip"))
        else:
            summary = generate_raw_material_bulk_file(
                bom_path=bom_paths,
                raw_material_template_path=template_path,
                output_path=output_path,
                mapping=mapping,
                supplier_paths=supplier_paths,
                supplier_bulk_template_path=supplier_bulk_template_path,
                supplier_bulk_output_path=supplier_bulk_output_path,
            )
        # 供 Module 3 直接串接使用：保留 Module 2 最新產出的 Raw Material Bulk。
        try:
            if output_path.exists() and output_path.suffix.lower() in {".xlsx", ".xlsm"}:
                import shutil as _shutil
                _shutil.copy2(output_path, LATEST_RAW_MATERIAL_BULK_PATH)
                summary["raw_material_bulk_latest"] = LATEST_RAW_MATERIAL_BULK_PATH.name
                summary["raw_material_bulk_latest_download_url"] = f"/download/{LATEST_RAW_MATERIAL_BULK_PATH.name}"
            else:
                summary["raw_material_bulk_latest"] = ""
        except Exception as _latest_exc:
            summary["raw_material_bulk_latest"] = ""
            summary["raw_material_bulk_latest_error"] = str(_latest_exc)

        bom_structure_summary = export_bom_structure_file(
            bom_path=bom_paths,
            output_path=LATEST_BOM_STRUCTURE_PATH,
            mapping=mapping,
        )
        summary["bom_structure_latest"] = LATEST_BOM_STRUCTURE_PATH.name
        summary["bom_structure_rows"] = int(bom_structure_summary.get("structure_rows", 0))
        summary["bom_structure_download_url"] = f"/download/{LATEST_BOM_STRUCTURE_PATH.name}"
        summary["bom_files"] = int(bom_structure_summary.get("bom_files", summary.get("bom_files", len(bom_paths))))
        summary["bom_rows_before_dedup"] = int(bom_structure_summary.get("bom_rows_before_dedup", summary.get("bom_rows_before_dedup", 0)))
        summary["bom_rows_after_dedup"] = int(bom_structure_summary.get("bom_rows_after_dedup", summary.get("bom_rows_after_dedup", 0)))
        summary["bom_duplicate_rows_removed"] = int(bom_structure_summary.get("bom_duplicate_rows_removed", summary.get("bom_duplicate_rows_removed", 0)))

        if step1_path is not None:
            rollup_summary = generate_working_hour_rollup_file(
                step1_output_path=step1_path,
                bom_structure_path=LATEST_BOM_STRUCTURE_PATH,
                output_path=working_hour_rollup_output_path,
            )
            LATEST_WORKING_HOUR_ROLLUP_PATH.write_bytes(working_hour_rollup_output_path.read_bytes())
            summary["working_hour_rollup_filename"] = working_hour_rollup_output_path.name
            summary["working_hour_rollup_download_url"] = f"/download/{working_hour_rollup_output_path.name}"
            summary["working_hour_rollup_latest"] = LATEST_WORKING_HOUR_ROLLUP_PATH.name
            summary["working_hour_rollup_latest_download_url"] = f"/download/{LATEST_WORKING_HOUR_ROLLUP_PATH.name}"
            summary["working_hour_rollup_rows"] = int(rollup_summary.get("summary_rows", 0))
            summary["working_hour_rollup_detail_rows"] = int(rollup_summary.get("detail_rows", 0))
            summary["working_hour_rollup_total_direct_hours"] = float(rollup_summary.get("total_direct_hours", 0))
            summary["working_hour_rollup_total_semi_hours"] = float(rollup_summary.get("total_semi_hours", 0))
            summary["working_hour_rollup_total_hours"] = float(rollup_summary.get("total_hours", 0))
        else:
            summary["working_hour_rollup_filename"] = ""
            summary["working_hour_rollup_download_url"] = ""
            summary["working_hour_rollup_rows"] = 0

        summary["supplier_upload_files"] = len(supplier_paths)
        summary["app_version"] = "CMP_V16_0_SUPPLIER_MASTER_CLEAN_PATCH"
        summary["bom_formatter_version"] = BOM_FORMATTER_VERSION
    except Exception as exc:
        traceback.print_exc()
        return JSONResponse(
            {"ok": False, "message": str(exc)},
            status_code=400,
        )

    return {
        "ok": True,
        "message": "BOM Expansion completed successfully.",
        "app_version": "CMP_V16_0_SUPPLIER_MASTER_CLEAN_PATCH",
        "bom_formatter_version": BOM_FORMATTER_VERSION,
        "summary": summary,
        "download_url": summary.get("download_url", f"/download/{output_path.name}"),
    }
