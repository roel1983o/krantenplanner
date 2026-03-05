import os
import uuid
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from krantenplanner.pipeline import run_pipeline

BASE_DIR = Path(__file__).resolve().parent.parent
STATIC_DIR = BASE_DIR / "app" / "static"
RUNS_DIR = BASE_DIR / "runs"
RUNS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Krantenplanner V1.1")

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")

@app.get("/", response_class=HTMLResponse)
def index():
    return (STATIC_DIR / "index.html").read_text(encoding="utf-8")

def _save_upload(upload: UploadFile, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    with dest.open("wb") as f:
        f.write(upload.file.read())

@app.post("/upload/kordiam")
async def upload_kordiam(run_id: str, file: UploadFile = File(...)):
    run_dir = RUNS_DIR / run_id
    _save_upload(file, run_dir / "kordiam_report.xlsx")
    return {"run_id": run_id}

@app.post("/upload/posities")
async def upload_posities(file: UploadFile = File(...)):
    run_id = (await _get_or_create_run_id())
    run_dir = RUNS_DIR / run_id
    _save_upload(file, run_dir / "posities_en_kenmerken.xlsx")
    return {"run_id": run_id}

async def _get_or_create_run_id():
    # simple: client passes run_id header; else create new and return
    # (the frontend stores it)
    return str(uuid.uuid4())

@app.post("/generate")
async def generate(payload: dict):
    run_id = payload.get("run_id")
    if not run_id:
        raise HTTPException(status_code=400, detail="Missing run_id")

    run_dir = RUNS_DIR / run_id
    kordiam = run_dir / "kordiam_report.xlsx"
    posities = run_dir / "posities_en_kenmerken.xlsx"
    if not kordiam.exists() or not posities.exists():
        raise HTTPException(status_code=400, detail="Upload both files first")

    # Run pipeline
    out = run_pipeline(
        kordiam_report_xlsx=str(kordiam),
        posities_xlsx=str(posities),
        workdir=str(run_dir),
    )

    return {
        "krantenplanning_xlsx": f"/download/{run_id}/krantenplanning",
        "handout_pdf": f"/download/{run_id}/handout",
    }

@app.get("/download/{run_id}/krantenplanning")
async def download_krantenplanning(run_id: str):
    p = RUNS_DIR / run_id / "Krantenplanning.xlsx"
    if not p.exists():
        raise HTTPException(status_code=404, detail="Krantenplanning not found")
    return FileResponse(str(p), filename="Krantenplanning.xlsx")

@app.get("/download/{run_id}/handout")
async def download_handout(run_id: str):
    p = RUNS_DIR / run_id / "handout_modern_v3.pdf"
    if not p.exists():
        raise HTTPException(status_code=404, detail="PDF not found")
    return FileResponse(str(p), filename="handout_modern_v3.pdf")
