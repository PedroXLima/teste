"""
FastAPI application — serves the frontend and handles DOCX → PDF conversion.
"""

from __future__ import annotations

import os
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from .docx_parser import parse_docx
from .pdf_renderer import generate_pdf

APP_DIR = Path(__file__).parent
STATIC_DIR = APP_DIR / "static"
UPLOAD_DIR = Path(tempfile.gettempdir()) / "docx2pdf_uploads"
OUTPUT_DIR = Path(tempfile.gettempdir()) / "docx2pdf_output"

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="DOCX → PDF Designer", version="1.0.0")

app.mount("/static", StaticFiles(directory=str(STATIC_DIR)), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    index_path = STATIC_DIR / "index.html"
    return HTMLResponse(content=index_path.read_text(encoding="utf-8"))


@app.post("/api/convert")
async def convert_docx_to_pdf(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Please upload a valid .docx file.",
        )

    job_id = uuid.uuid4().hex[:12]
    docx_path = UPLOAD_DIR / f"{job_id}.docx"
    pdf_path = OUTPUT_DIR / f"{job_id}.pdf"

    try:
        contents = await file.read()
        docx_path.write_bytes(contents)

        doc = parse_docx(str(docx_path))
        generate_pdf(doc, str(pdf_path))

        safe_name = Path(file.filename).stem + "_designed.pdf"

        return FileResponse(
            path=str(pdf_path),
            media_type="application/pdf",
            filename=safe_name,
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Conversion failed: {str(e)}")

    finally:
        if docx_path.exists():
            docx_path.unlink(missing_ok=True)


@app.post("/api/preview")
async def preview_html(file: UploadFile = File(...)):
    """Return the generated HTML for preview (debug endpoint)."""
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Please upload a valid .docx file.")

    job_id = uuid.uuid4().hex[:12]
    docx_path = UPLOAD_DIR / f"{job_id}.docx"

    try:
        contents = await file.read()
        docx_path.write_bytes(contents)

        doc = parse_docx(str(docx_path))
        from .pdf_renderer import render_html
        html_content = render_html(doc)
        return HTMLResponse(content=html_content)

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

    finally:
        if docx_path.exists():
            docx_path.unlink(missing_ok=True)
