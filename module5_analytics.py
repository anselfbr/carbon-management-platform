from __future__ import annotations
import io, zipfile, math
from pathlib import Path
from typing import Any, Iterable
from openpyxl import load_workbook

SHEET_ALIASES = {"input sheet activity data", "raw material activity data", "activity data"}
ALIASES = {
 "product": ["Allocated Target Product/Service","Product Name","Target Product","target_product","產品代碼","產品名稱"],
 "material": ["Raw Material Code","Raw Material Number","Material","raw_material_code","料號","原物料代碼"],
 "material_name": ["Raw Material Name","raw_material_name","原物料名稱"],
 "usage": ["Usage","Activity Data","activity_data","使用量"],
 "unit": ["Activity Data Unit","Unit","activity_data_unit","單位"],
 "net_weight": ["Net Weight (optional)","Net weight","Net Weight","net_weight","淨重"],
 "weight_unit": ["Weight Unit (optional)","Weight Unit","weight_unit","重量單位"],
 "factor": ["Emission Factor","Carbon Factor","emission_factor","碳係數"],
 "supplier": ["Supplier Name (optional)","Supplier Name","supplier_name","供應商名稱"],
 "plant": ["Transportation Destination","Production Site","Unit Name","transportation_destination","廠區"],
 "country": ["Country/Area","Country Area","Country","country_area","國家地區"],
}

def _norm(v: Any) -> str:
    return " ".join(str(v or "").replace("\n"," ").strip().lower().split())

def _num(v: Any) -> float:
    if v is None or v == "": return 0.0
    try:
        x=float(str(v).replace(",","").strip())
        return x if math.isfinite(x) else 0.0
    except Exception: return 0.0

def _header_map(rows: list[tuple[Any,...]]) -> dict[str,int]:
    candidates=[]
    for i in range(max(len(r) for r in rows)):
        vals=[]
        for r in rows:
            if i < len(r) and str(r[i] or "").strip(): vals.append(str(r[i]).strip())
        candidates.append(vals)
    out={}
    for key,names in ALIASES.items():
        wanted={_norm(x) for x in names}
        for idx,vals in enumerate(candidates):
            if any(_norm(v) in wanted for v in vals): out[key]=idx; break
    return out

def _iter_workbooks(path: Path) -> Iterable[tuple[str, bytes]]:
    if path.suffix.lower()==".zip":
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.lower().endswith((".xlsx",".xlsm")) and not Path(n).name.startswith("~$"):
                    yield Path(n).name, z.read(n)
    else:
        yield path.name, path.read_bytes()

def _weight_to_kg(value: float, unit: str) -> float:
    u=_norm(unit)
    if u in {"g","gram","grams","公克"}: return value/1000
    if u in {"mg","milligram","milligrams","毫克"}: return value/1_000_000
    if u in {"t","ton","tonne","tonnes","metric ton","公噸"}: return value*1000
    return value

def analyze_bulk(path: Path) -> dict[str,Any]:
    records=[]; skipped=0; files=0
    for name,data in _iter_workbooks(path):
        files+=1
        wb=load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        ws=None
        for s in wb.sheetnames:
            if _norm(s) in SHEET_ALIASES: ws=wb[s]; break
        if ws is None: skipped+=1; wb.close(); continue
        rows=ws.iter_rows(values_only=True)
        h1=next(rows,()); h2=next(rows,())
        hm=_header_map([h1,h2])
        required={"product","material","usage","factor"}
        if not required.issubset(hm): skipped+=1; wb.close(); continue
        for row in rows:
            product=str(row[hm["product"]] or "").strip() if hm["product"] < len(row) else ""
            material=str(row[hm["material"]] or "").strip() if hm["material"] < len(row) else ""
            usage=_num(row[hm["usage"]]) if hm["usage"] < len(row) else 0
            factor=_num(row[hm["factor"]]) if hm["factor"] < len(row) else 0
            if not product or not material or usage==0 or factor==0: continue
            unit=str(row[hm.get("unit",-1)] or "").strip() if hm.get("unit",-1)>=0 and hm["unit"] < len(row) else ""
            nw=_num(row[hm.get("net_weight",-1)]) if hm.get("net_weight",-1)>=0 and hm["net_weight"] < len(row) else 0
            wu=str(row[hm.get("weight_unit",-1)] or "").strip() if hm.get("weight_unit",-1)>=0 and hm["weight_unit"] < len(row) else ""
            weight_activity={"kg","kilogram","kilograms","公斤","g","gram","grams","公克","mg","milligram","milligrams","毫克","t","ton","tonne","tonnes","公噸"}
            if _norm(unit) in weight_activity:
                activity_kg=_weight_to_kg(usage,unit); method="Usage(weight) × EF"
            elif nw>0:
                activity_kg=usage*_weight_to_kg(nw,wu); method="Usage × Net Weight × EF"
            else:
                activity_kg=usage; method="Usage × EF"
            emission=activity_kg*factor
            def val(k,default=""):
                i=hm.get(k,-1); return str(row[i] or "").strip() if i>=0 and i<len(row) else default
            records.append({"product":product,"material":material,"material_name":val("material_name",material),"supplier":val("supplier","Unassigned") or "Unassigned","plant":val("plant","Unassigned") or "Unassigned","country":val("country",""),"usage":usage,"activity_kg":activity_kg,"factor":factor,"emission":emission,"method":method,"source_file":name})
        wb.close()
    if not records: raise ValueError("找不到可分析的 Activity Data；請確認檔案含 Product、Raw Material、Usage 與 Emission Factor 欄位。")
    def agg(key):
        d={}
        for r in records: d[r[key]]=d.get(r[key],0)+r["emission"]
        return [{"name":k,"emission":v} for k,v in sorted(d.items(),key=lambda x:x[1],reverse=True)]
    products=agg("product"); materials=agg("material"); suppliers=agg("supplier"); plants=agg("plant")
    total=sum(r["emission"] for r in records)
    by_product={}
    for r in records:
        p=by_product.setdefault(r["product"],{"materials":{},"suppliers":{}})
        p["materials"][r["material"]]=p["materials"].get(r["material"],0)+r["emission"]
        p["suppliers"][r["supplier"]]=p["suppliers"].get(r["supplier"],0)+r["emission"]
    drill={p:{"materials":[{"name":k,"emission":v} for k,v in sorted(vv["materials"].items(),key=lambda x:x[1],reverse=True)],"suppliers":[{"name":k,"emission":v} for k,v in sorted(vv["suppliers"].items(),key=lambda x:x[1],reverse=True)]} for p,vv in by_product.items()}
    return {"ok":True,"summary":{"total_emission":total,"product_count":len(products),"material_count":len(materials),"supplier_count":len(suppliers),"plant_count":len(plants),"record_count":len(records),"file_count":files,"skipped_files":skipped},"products":products,"materials":materials,"suppliers":suppliers,"plants":plants,"drilldown":drill,"calculation_methods":sorted({r["method"] for r in records})}
