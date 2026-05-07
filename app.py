from __future__ import annotations

from pathlib import Path
from uuid import uuid4

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from pdf_generator import generate_institutional_pdf

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_EXTENSIONS = {"docx"}

UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.secret_key = "institutional-pdf-generator"


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/generate")
def generate_pdf():
    uploaded = request.files.get("docx_file")
    if uploaded is None or uploaded.filename == "":
        flash("Selecione um arquivo .docx para continuar.")
        return redirect(url_for("index"))

    if not allowed_file(uploaded.filename):
        flash("Formato inválido. Envie um arquivo com extensão .docx.")
        return redirect(url_for("index"))

    safe_name = secure_filename(uploaded.filename)
    token = uuid4().hex
    docx_path = UPLOAD_DIR / f"{token}_{safe_name}"
    pdf_path = OUTPUT_DIR / f"{Path(safe_name).stem}_{token}.pdf"

    uploaded.save(docx_path)

    try:
        generate_institutional_pdf(docx_path=docx_path, pdf_path=pdf_path)
    except Exception as exc:  # noqa: BLE001
        flash(f"Falha ao gerar PDF: {exc}")
        return redirect(url_for("index"))

    return send_file(
        pdf_path,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"{Path(safe_name).stem}_institucional.pdf",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)
