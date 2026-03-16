"""FastAPI server for ECL Automation."""

import os
import uuid
import shutil
import tempfile

from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import FileResponse, JSONResponse

from ecl_engine import ECLEngine

app = FastAPI(title="ECL Automation")

BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
OUTPUTS_DIR = os.path.join(BASE_DIR, "outputs")
os.makedirs(OUTPUTS_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")
templates = Jinja2Templates(directory=os.path.join(BASE_DIR, "templates"))


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/api/compute")
async def compute(
    dpd_file: UploadFile = File(...),
    weo_file: UploadFile = File(...),
    shock: float = Form(0.10),
    tm_start_year: int = Form(2020),
    hist_cutoff: int = Form(2024),
):
    job_id  = str(uuid.uuid4())[:8]
    tmp_dir = os.path.join(tempfile.gettempdir(), f"ecl_{job_id}")
    os.makedirs(tmp_dir, exist_ok=True)

    dpd_path    = os.path.join(tmp_dir, "dpd.xlsx")
    weo_path    = os.path.join(tmp_dir, "weo.xlsx")
    output_path = os.path.join(OUTPUTS_DIR, f"ECL_Output_{job_id}.xlsx")

    with open(dpd_path, "wb") as f:
        shutil.copyfileobj(dpd_file.file, f)
    with open(weo_path, "wb") as f:
        shutil.copyfileobj(weo_file.file, f)

    try:
        engine  = ECLEngine(dpd_path, weo_path, output_path, {
            "shock":         shock,
            "tm_start_year": tm_start_year,
            "hist_cutoff":   hist_cutoff,
        })
        results = engine.run()
        results["download_url"] = f"/api/download/{job_id}"
        return JSONResponse(results)

    except Exception as e:
        import traceback
        traceback.print_exc()
        return JSONResponse({"error": str(e)}, status_code=400)

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


@app.get("/api/download/{job_id}")
async def download(job_id: str):
    path = os.path.join(OUTPUTS_DIR, f"ECL_Output_{job_id}.xlsx")
    if not os.path.exists(path):
        return JSONResponse({"error": "File not found"}, status_code=404)
    return FileResponse(
        path,
        filename="ECL_Output.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
