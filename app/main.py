from __future__ import annotations

import os
import tempfile
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from starlette.background import BackgroundTask

from app.converter import convert_docx_to_pdf


app = FastAPI(
    title="Institutional DOCX to PDF Designer",
    description="Upload DOCX files and receive professionally designed institutional PDFs.",
)


INDEX_HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Institutional DOCX to PDF Designer</title>
  <style>
    :root {
      --primary-blue: #005A9C;
      --dark-navy: #003B6F;
      --medium-blue: #0B63B6;
      --light-blue: #D9EAF7;
      --very-light-blue: #EEF6FC;
      --text: #2B2B2B;
      --secondary: #6B7280;
      --border: #D1D5DB;
      font-family: Inter, Roboto, "Source Sans Pro", Arial, sans-serif;
    }

    * { box-sizing: border-box; }

    body {
      margin: 0;
      min-height: 100vh;
      color: var(--text);
      background:
        radial-gradient(circle at 15% 20%, rgba(217, 234, 247, 0.9), transparent 34rem),
        linear-gradient(135deg, #ffffff 0%, #eef6fc 48%, #d9eaf7 100%);
    }

    .hero {
      min-height: 100vh;
      display: grid;
      grid-template-columns: minmax(18rem, 1fr) minmax(20rem, 30rem);
      gap: 3rem;
      align-items: center;
      width: min(72rem, calc(100% - 3rem));
      margin: 0 auto;
      padding: 4rem 0;
    }

    .eyebrow {
      color: var(--medium-blue);
      font-weight: 800;
      letter-spacing: 0.14em;
      text-transform: uppercase;
      font-size: 0.78rem;
      margin-bottom: 1rem;
    }

    h1 {
      margin: 0;
      color: var(--dark-navy);
      font-size: clamp(2.4rem, 6vw, 5.4rem);
      line-height: 0.95;
      letter-spacing: -0.055em;
    }

    .lead {
      margin: 1.5rem 0 0;
      max-width: 44rem;
      color: var(--secondary);
      font-size: 1.16rem;
      line-height: 1.7;
    }

    .features {
      display: grid;
      grid-template-columns: repeat(2, minmax(0, 1fr));
      gap: 0.9rem;
      margin-top: 2rem;
      max-width: 43rem;
    }

    .feature {
      border-left: 4px solid var(--primary-blue);
      background: rgba(255, 255, 255, 0.75);
      padding: 0.9rem 1rem;
      border-radius: 0 1rem 1rem 0;
      box-shadow: 0 10px 30px rgba(0, 59, 111, 0.08);
      font-weight: 700;
      color: var(--dark-navy);
    }

    .card {
      position: relative;
      overflow: hidden;
      background: rgba(255, 255, 255, 0.9);
      border: 1px solid rgba(0, 90, 156, 0.14);
      border-radius: 1.6rem;
      box-shadow: 0 30px 80px rgba(0, 59, 111, 0.18);
      padding: 2rem;
    }

    .card::before {
      content: "";
      position: absolute;
      inset: 0 0 auto;
      height: 0.45rem;
      background: linear-gradient(90deg, var(--dark-navy), var(--primary-blue), var(--medium-blue));
    }

    .upload-zone {
      border: 2px dashed rgba(0, 90, 156, 0.35);
      background: var(--very-light-blue);
      border-radius: 1.2rem;
      padding: 2rem;
      text-align: center;
      transition: 0.2s ease;
    }

    .upload-zone:hover {
      border-color: var(--primary-blue);
      background: #fff;
    }

    input[type="file"] {
      width: 100%;
      margin: 1.2rem 0;
      padding: 0.85rem;
      background: #fff;
      border: 1px solid var(--border);
      border-radius: 0.85rem;
    }

    button {
      width: 100%;
      border: 0;
      border-radius: 0.95rem;
      background: linear-gradient(135deg, var(--dark-navy), var(--primary-blue));
      color: #fff;
      font-size: 1rem;
      font-weight: 800;
      padding: 1rem 1.2rem;
      cursor: pointer;
      box-shadow: 0 14px 30px rgba(0, 90, 156, 0.26);
    }

    .small {
      color: var(--secondary);
      font-size: 0.88rem;
      line-height: 1.5;
      margin-top: 1rem;
    }

    @media (max-width: 860px) {
      .hero { grid-template-columns: 1fr; }
      .features { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <main class="hero">
    <section>
      <div class="eyebrow">Document automation / strategic intelligence</div>
      <h1>DOCX to institutional PDF</h1>
      <p class="lead">
        Upload a Word document and generate a print-ready A4 PDF with a corporate Brazilian
        institutional report design: blue identity system, cover, page headers, styled tables,
        figures, captions, section hierarchy and back cover.
      </p>
      <div class="features">
        <div class="feature">Preserves text, headings and tables</div>
        <div class="feature">Applies analytical report styling</div>
        <div class="feature">Includes cover and back cover</div>
        <div class="feature">Returns a downloadable PDF</div>
      </div>
    </section>

    <section class="card" aria-label="Upload DOCX form">
      <form class="upload-zone" method="post" action="/convert" enctype="multipart/form-data">
        <h2>Upload DOCX file</h2>
        <p class="small">Accepted format: .docx. The original document content is preserved while the PDF layout is redesigned.</p>
        <input type="file" name="file" accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" required>
        <button type="submit">Create designed PDF</button>
      </form>
      <p class="small">
        Processing happens on demand. Large files with many images may take longer while figures are embedded into the final PDF.
      </p>
    </section>
  </main>
</body>
</html>
"""


@app.get("/", response_class=HTMLResponse)
async def index() -> str:
    return INDEX_HTML


@app.post("/convert")
async def convert(file: UploadFile = File(...)) -> FileResponse:
    original_name = Path(file.filename or "document.docx").name
    if not original_name.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Please upload a .docx file.")

    with tempfile.TemporaryDirectory(prefix="docx-pdf-") as temp_dir:
        temp_path = Path(temp_dir)
        input_path = temp_path / original_name
        output_path = temp_path / f"{input_path.stem}-institutional-report.pdf"

        with input_path.open("wb") as buffer:
            while chunk := await file.read(1024 * 1024):
                buffer.write(chunk)

        try:
            convert_docx_to_pdf(input_path, output_path)
        except Exception as exc:  # noqa: BLE001 - surface a clean API error for malformed DOCX files.
            raise HTTPException(status_code=422, detail=f"Could not convert DOCX: {exc}") from exc

        download_name = f"{input_path.stem}-institutional-report.pdf"
        stable_output = Path(tempfile.gettempdir()) / f"{os.urandom(8).hex()}-{download_name}"
        output_path.replace(stable_output)

    return FileResponse(
        stable_output,
        filename=download_name,
        media_type="application/pdf",
        background=BackgroundTask(lambda path: Path(path).unlink(missing_ok=True), stable_output),
    )
