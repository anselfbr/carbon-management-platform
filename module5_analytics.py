from __future__ import annotations
import io, zipfile, math
from pathlib import Path
from typing import Any, Iterable
from openpyxl import load_workbook

SHEET_ALIASES = {"input sheet activity data", "raw material activity data", "activity data"}
ALIASES = {
 "product": ["Allocated Target Product/Service","Product Name","Target Product","target_product","product_name","產品代碼","產品名稱"],
 "material": ["Raw Material Code","Raw Material Number","Material","raw_material_code","料號","原物料代碼"],
 "material_name": ["Raw Material Name","raw_material_name","原物料名稱"],
 "usage": ["Usage","Activity Data","activity_data","使用量","用量","生產數量","年度生產量","Annual Quantity","Delivered quantity"],
 "unit": ["Activity Data Unit","Unit","activity_data_unit","單位"],
 "net_weight": ["Net Weight (optional)","Net weight","Net Weight","net_weight","淨重"],
 "weight_unit": ["Weight Unit (optional)","Weight Unit","weight_unit","重量單位"],
 "factor": ["Emission Factor","Carbon Factor","emission_factor","碳係數","排放係數"],
 "supplier": ["Supplier Name (optional)","Supplier Name","supplier_name","供應商名稱"],
 "plant": ["Transportation Destination","Production Site","Unit Name","transportation_destination","production_site","廠區","製造場所"],
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
    width=max((len(r) for r in rows), default=0)
    candidates=[]
    for i in range(width):
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

def _cell(row: tuple[Any,...], hm: dict[str,int], key: str, default: str="") -> str:
    i=hm.get(key,-1)
    return str(row[i] or "").strip() if i>=0 and i<len(row) else default

def analyze_bulk_many(paths: Iterable[Path]) -> dict[str,Any]:
    raw_rows=[]
    quantity_rows=[]
    skipped=0
    files=0
    product_bulk_files=0
    raw_material_files=0

    for path in paths:
      for name,data in _iter_workbooks(path):
          files+=1
          wb=load_workbook(io.BytesIO(data), read_only=True, data_only=True)
          ws=None
          for s in wb.sheetnames:
              if _norm(s) in SHEET_ALIASES: ws=wb[s]; break
          if ws is None:
              skipped+=1; wb.close(); continue
          rows=ws.iter_rows(values_only=True)
          h1=next(rows,()); h2=next(rows,())
          hm=_header_map([h1,h2])

          # 原物料 Bulk：具備成品、原物料、使用量與係數。
          if {"product","material","usage","factor"}.issubset(hm):
              raw_material_files+=1
              for row in rows:
                  product=_cell(row,hm,"product")
                  material=_cell(row,hm,"material")
                  usage=_num(row[hm["usage"]]) if hm["usage"] < len(row) else 0
                  factor=_num(row[hm["factor"]]) if hm["factor"] < len(row) else 0
                  if not product or not material or usage==0 or factor==0: continue
                  unit=_cell(row,hm,"unit")
                  nw=_num(row[hm["net_weight"]]) if hm.get("net_weight",-1)>=0 and hm["net_weight"] < len(row) else 0
                  wu=_cell(row,hm,"weight_unit")
                  weight_activity={"kg","kilogram","kilograms","公斤","g","gram","grams","公克","mg","milligram","milligrams","毫克","t","ton","tonne","tonnes","公噸"}
                  if _norm(unit) in weight_activity:
                      activity_kg=_weight_to_kg(usage,unit); method="Usage(weight) × EF ÷ Production Quantity"
                  elif nw>0:
                      activity_kg=usage*_weight_to_kg(nw,wu); method="Usage × Net Weight × EF ÷ Production Quantity"
                  else:
                      activity_kg=usage; method="Usage × EF ÷ Production Quantity"
                  raw_rows.append({
                      "product":product,"material":material,
                      "material_name":_cell(row,hm,"material_name",material),
                      "supplier":_cell(row,hm,"supplier","Unassigned") or "Unassigned",
                      "plant":_cell(row,hm,"plant","Unassigned") or "Unassigned",
                      "country":_cell(row,hm,"country"),"usage":usage,
                      "activity_kg_total":activity_kg,"factor":factor,
                      "total_emission":activity_kg*factor,"method":method,"source_file":name,
                  })
          # 成品 Bulk：具備成品與 Activity Data，但不具備原物料及係數。
          elif {"product","usage"}.issubset(hm) and "material" not in hm and "factor" not in hm:
              product_bulk_files+=1
              for row in rows:
                  product=_cell(row,hm,"product")
                  qty=_num(row[hm["usage"]]) if hm["usage"] < len(row) else 0
                  if not product or qty<=0: continue
                  quantity_rows.append({"product":product,"plant":_cell(row,hm,"plant"),"quantity":qty,"source_file":name})
          else:
              skipped+=1
          wb.close()

    if not raw_rows:
        raise ValueError("找不到可分析的原物料 Activity Data；請確認檔案含 Product、Raw Material、Usage 與 Emission Factor 欄位。")
    if not quantity_rows:
        raise ValueError("缺少成品生產數量。請將 M1B 產出的成品 Bulk 與 M3A 原物料 Bulk 一起上傳，系統才能換算每 1 PC 成品的碳排放與原物料用量。")

    qty_by_product={}
    qty_by_product_plant={}
    plants_by_product={}
    for r in quantity_rows:
        p=_norm(r["product"]); pl=_norm(r["plant"])
        qty_by_product[p]=qty_by_product.get(p,0.0)+r["quantity"]
        if pl:
            qty_by_product_plant[(p,pl)]=qty_by_product_plant.get((p,pl),0.0)+r["quantity"]
            plants_by_product.setdefault(p,set()).add(pl)

    missing=[]
    ambiguous=[]
    records=[]
    for r in raw_rows:
        p=_norm(r["product"]); pl=_norm(r["plant"])
        qty=qty_by_product_plant.get((p,pl),0.0) if pl else 0.0
        if qty<=0:
            known=plants_by_product.get(p,set())
            if len(known)>1 and pl and (p,pl) not in qty_by_product_plant:
                ambiguous.append(f'{r["product"]}（原物料廠區：{r["plant"]}）')
                continue
            qty=qty_by_product.get(p,0.0)
        if qty<=0:
            missing.append(r["product"]); continue
        unit_activity=r["activity_kg_total"]/qty
        unit_emission=r["total_emission"]/qty
        records.append({**r,"production_quantity":qty,"activity_kg":unit_activity,"emission":unit_emission})

    if ambiguous:
        sample="、".join(sorted(set(ambiguous))[:8])
        raise ValueError(f"部分成品有多個生產廠區，但原物料廠區無法與成品 Bulk 對應：{sample}。請確認 M1B Production Site 與 M3A Transportation Destination 名稱一致。")
    if missing:
        sample="、".join(sorted(set(missing))[:12])
        raise ValueError(f"下列成品找不到生產數量，無法換算每 PC：{sample}。請一併上傳包含這些成品的 M1B 成品 Bulk。")
    if not records:
        raise ValueError("沒有可完成每 PC 換算的資料。")

    def agg(key):
        d={}
        for r in records:
            item=d.setdefault(r[key],{"emission":0.0,"total_emission":0.0})
            item["emission"]+=r["emission"]
            item["total_emission"]+=r["total_emission"]
        return [{"name":k,**v} for k,v in sorted(d.items(),key=lambda x:x[1]["emission"],reverse=True)]

    products=agg("product"); materials=agg("material"); suppliers=agg("supplier"); plants=agg("plant")
    total_unit=sum(x["emission"] for x in products)
    total_absolute=sum(r["total_emission"] for r in records)
    by_product={}
    material_detail={}
    product_material_detail={}
    product_qty={}

    for r in records:
        product_qty[r["product"]]=max(product_qty.get(r["product"],0.0),r["production_quantity"])
        p=by_product.setdefault(r["product"],{"materials":{},"suppliers":{},"production_quantity":r["production_quantity"],"total_emission":0.0})
        p["total_emission"]+=r["total_emission"]
        for group,key in (("materials","material"),("suppliers","supplier")):
            item=p[group].setdefault(r[key],{"emission":0.0,"total_emission":0.0})
            item["emission"]+=r["emission"]; item["total_emission"]+=r["total_emission"]

        md=material_detail.setdefault(r["material"],{"name":r["material_name"],"emission":0.0,"total_emission":0.0,"activity_kg":0.0,"total_activity_kg":0.0,"suppliers":{},"products":{},"plants":{},"records":0})
        md["emission"]+=r["emission"]; md["total_emission"]+=r["total_emission"]
        md["activity_kg"]+=r["activity_kg"]; md["total_activity_kg"]+=r["activity_kg_total"]; md["records"]+=1
        for group,key in (("suppliers","supplier"),("products","product"),("plants","plant")):
            item=md[group].setdefault(r[key],{"emission":0.0,"total_emission":0.0})
            item["emission"]+=r["emission"]; item["total_emission"]+=r["total_emission"]

        key=f'{r["product"]}|||{r["material"]}'
        pmd=product_material_detail.setdefault(key,{"product":r["product"],"material":r["material"],"name":r["material_name"],"emission":0.0,"total_emission":0.0,"activity_kg":0.0,"total_activity_kg":0.0,"production_quantity":r["production_quantity"],"suppliers":{},"plants":{},"records":0})
        pmd["emission"]+=r["emission"]; pmd["total_emission"]+=r["total_emission"]
        pmd["activity_kg"]+=r["activity_kg"]; pmd["total_activity_kg"]+=r["activity_kg_total"]; pmd["records"]+=1
        for group,key_name in (("suppliers","supplier"),("plants","plant")):
            item=pmd[group].setdefault(r[key_name],{"emission":0.0,"total_emission":0.0})
            item["emission"]+=r["emission"]; item["total_emission"]+=r["total_emission"]

    def ranked(d):
        return [{"name":k,**v} for k,v in sorted(d.items(),key=lambda x:x[1]["emission"],reverse=True)]

    drill={p:{"materials":ranked(v["materials"]),"suppliers":ranked(v["suppliers"]),"production_quantity":v["production_quantity"],"total_emission":v["total_emission"]} for p,v in by_product.items()}
    material_details={k:{**v,"suppliers":ranked(v["suppliers"]),"products":ranked(v["products"]),"plants":ranked(v["plants"])} for k,v in material_detail.items()}
    product_material_details={k:{**v,"suppliers":ranked(v["suppliers"]),"plants":ranked(v["plants"])} for k,v in product_material_detail.items()}

    return {
        "ok":True,
        "summary":{
            "total_emission":total_unit,
            "absolute_total_emission":total_absolute,
            "product_count":len(products),"material_count":len(materials),"supplier_count":len(suppliers),"plant_count":len(plants),
            "record_count":len(records),"file_count":files,"skipped_files":skipped,
            "product_bulk_files":product_bulk_files,"raw_material_files":raw_material_files,
            "production_quantity_rows":len(quantity_rows),
        },
        "products":products,"materials":materials,"suppliers":suppliers,"plants":plants,
        "drilldown":drill,"material_details":material_details,"product_material_details":product_material_details,
        "product_quantities":product_qty,
        "calculation_methods":sorted({r["method"] for r in records}),
        "analysis_basis":"per_pc",
    }

def analyze_bulk(path: Path) -> dict[str,Any]:
    return analyze_bulk_many([path])
