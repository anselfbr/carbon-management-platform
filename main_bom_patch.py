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
