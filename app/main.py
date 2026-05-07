"""HTTP API and UI for DOCX → institutional PDF conversion."""

from __future__ import annotations

import io
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse

from app.docx_parser import parse_docx_bytes
from app.pdf_render import render_pdf

app = FastAPI(
    title="Relatório institucional PDF",
    description="Converte DOCX em PDF no padrão Observatório da Indústria / FIEA.",
)

_INDEX = Path(__file__).resolve().parent / "templates" / "index.html"


@app.get("/", response_class=HTMLResponse)
def index() -> str:
    return _INDEX.read_text(encoding="utf-8")


@app.post("/api/convert")
async def convert_docx(file: UploadFile = File(...)) -> StreamingResponse:
    name = (file.filename or "").lower()
    if not name.endswith(".docx"):
        raise HTTPException(
            status_code=400, detail="Envie um arquivo .docx (.doc não é suportado)."
        )
    data = await file.read()
    if len(data) > 25 * 1024 * 1024:
        raise HTTPException(
            status_code=400, detail="Arquivo muito grande (máximo 25 MB)."
        )
    if len(data) < 100:
        raise HTTPException(status_code=400, detail="Arquivo vazio ou inválido.")

    try:
        ctx = parse_docx_bytes(data)
        pdf_bytes = render_pdf(ctx)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"Não foi possível ler ou montar o PDF: {exc}",
        ) from exc

    return StreamingResponse(
        io.BytesIO(pdf_bytes),
        media_type="application/pdf",
        headers={
            "Content-Disposition": 'attachment; filename="relatorio.pdf"',
        },
    )
