from fastapi import FastAPI, UploadFile, Form
from fastapi.responses import JSONResponse
import tempfile
import json
import os
import traceback
import importlib

# import your helper (the same file you just ran successfully)
fieldmap = importlib.import_module("field_map_generate_Occudo")

app = FastAPI(title="PDF Field Template → JSON API")


@app.post("/generate_pdf_json/")
async def generate_pdf_json(
    file: UploadFile,
    rebuild_template: bool = Form(default=True)
):
    """
    Upload a PDF and return its parsed field template JSON.
    (Uses your existing field_map_generate.py logic.)
    """
    try:
        # Save uploaded file temporarily
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(await file.read())
            tmp_pdf_path = tmp_pdf.name

        # Temporary file paths
        tmp_json = os.path.splitext(tmp_pdf_path)[0] + "_template.json"
        tmp_xlsx = os.path.splitext(tmp_pdf_path)[0] + "_lookup.xlsx"

        # Call your working CLI logic
        fieldmap.export_lookup_template_from_pdf(
            input_pdf=tmp_pdf_path,
            template_json=tmp_json,
            out_path=tmp_xlsx,
            rebuild_template=rebuild_template
        )

        # Verify template JSON created
        if not os.path.exists(tmp_json):
            return JSONResponse(
                {"error": "Template JSON not created."},
                status_code=500
            )

        # Load JSON data
        with open(tmp_json, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        # Cleanup temp files
        for path in [tmp_pdf_path, tmp_json, tmp_xlsx]:
            try:
                os.remove(path)
            except Exception:
                pass

        return {
            "status": "success",
            "total_fields": len(json_data.get("fields", [])),
            "data": json_data
        }

    except Exception as e:
        traceback.print_exc()
        return JSONResponse(
            {"error": str(e), "trace": traceback.format_exc()},
            status_code=500
        )
