from __future__ import annotations

import json
import shutil
import uuid
from pathlib import Path
from typing import List, Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel

from pipeline import RUNS_DIR, run_pipeline, validate_reference_output, LLM_PROVIDER, OPENAI_MODEL, OPENAI_API_KEY, AZURE_OPENAI_DEPLOYMENT, AZURE_OPENAI_API_KEY

app = FastAPI(title="Insurance KPI Ontology Mapping Tool")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

LAST_UPLOAD_ID: Optional[str] = None


class RunMappingRequest(BaseModel):
    upload_id: Optional[str] = None
    run_llm: bool = True  # ignored in this full-LLM version; backend always requires configured LLM/API


@app.get("/api/health")
def health():
    configured = bool(OPENAI_API_KEY or AZURE_OPENAI_API_KEY)
    model = AZURE_OPENAI_DEPLOYMENT if LLM_PROVIDER in {"azure", "azure_openai"} else OPENAI_MODEL
    return {"status": "healthy", "llm_provider": LLM_PROVIDER, "llm_model": model, "llm_configured": configured}


@app.post("/api/upload")
async def upload_files(reports: List[UploadFile] = File(...), ontology: UploadFile = File(...)):
    global LAST_UPLOAD_ID
    if len(reports) < 2:
        raise HTTPException(status_code=400, detail="Upload at least two report JSON files.")
    upload_id = str(uuid.uuid4())
    upload_dir = RUNS_DIR / upload_id / "uploads"
    report_dir = upload_dir / "reports"
    ontology_dir = upload_dir / "ontology"
    report_dir.mkdir(parents=True, exist_ok=True)
    ontology_dir.mkdir(parents=True, exist_ok=True)

    saved_reports = []
    for f in reports:
        if not f.filename.lower().endswith(".json"):
            raise HTTPException(status_code=400, detail=f"Report must be JSON: {f.filename}")
        target = report_dir / Path(f.filename).name
        with target.open("wb") as out:
            shutil.copyfileobj(f.file, out)
        saved_reports.append(target.name)

    if not ontology.filename.lower().endswith((".xlsx", ".xls", ".csv", ".json")):
        raise HTTPException(status_code=400, detail="Ontology must be Excel, CSV, or JSON.")
    ontology_target = ontology_dir / Path(ontology.filename).name
    with ontology_target.open("wb") as out:
        shutil.copyfileobj(ontology.file, out)

    LAST_UPLOAD_ID = upload_id
    return {"upload_id": upload_id, "reports": saved_reports, "ontology": ontology_target.name}


@app.post("/api/run-mapping")
def run_mapping(req: RunMappingRequest):
    upload_id = req.upload_id or LAST_UPLOAD_ID
    if not upload_id:
        raise HTTPException(status_code=400, detail="Upload files first, then run mapping.")

    upload_dir = RUNS_DIR / upload_id / "uploads"
    report_dir = upload_dir / "reports"
    ontology_dir = upload_dir / "ontology"
    if not report_dir.exists() or not ontology_dir.exists():
        raise HTTPException(status_code=404, detail=f"Upload ID not found: {upload_id}")

    report_paths = sorted(report_dir.glob("*.json"))
    ontology_paths = [p for p in ontology_dir.iterdir() if p.is_file()]
    if len(report_paths) < 2 or not ontology_paths:
        raise HTTPException(status_code=400, detail="Missing uploaded reports or ontology file.")

    run_id = str(uuid.uuid4())
    output_dir = RUNS_DIR / run_id / "outputs"
    try:
        excel_path, results, metadata = run_pipeline(report_paths, ontology_paths[0], output_dir, run_llm=True)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Pipeline failed: {e}") from e
    return {
        "run_id": run_id,
        "message": "Full LLM mapping completed.",
        "mode": metadata.get("mode"),
        "download_url": f"/api/download/{run_id}",
        "sheet_count": len(results),
    }


@app.get("/api/results/{run_id}")
def get_results(run_id: str):
    path = RUNS_DIR / run_id / "outputs" / "results.json"
    if not path.exists():
        raise HTTPException(status_code=404, detail="Results not found.")
    return json.loads(path.read_text(encoding="utf-8"))


@app.get("/api/download/{run_id}")
def download_excel(run_id: str):
    path = RUNS_DIR / run_id / "outputs" / "final_output_with_all_steps.xlsx"
    if not path.exists():
        raise HTTPException(status_code=404, detail="Excel output not found.")
    return FileResponse(
        path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename="final_output_with_all_steps.xlsx",
    )


@app.get("/api/validate/{run_id}")
def validate_run(run_id: str):
    path = RUNS_DIR / run_id / "outputs" / "final_output_with_all_steps.xlsx"
    if not path.exists():
        raise HTTPException(status_code=404, detail="Excel output not found.")
    return validate_reference_output(path)
