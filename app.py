from __future__ import annotations

import os
import tempfile
from io import BytesIO
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from docx_parser import parse_docx
from pdf_builder import build_pdf


ALLOWED_EXTENSIONS = {"docx"}


app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-key-change-me")
app.config["MAX_CONTENT_LENGTH"] = 30 * 1024 * 1024  # 30 MB


def is_allowed(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/generate")
def generate():
    uploaded_file = request.files.get("docx_file")
    if not uploaded_file or uploaded_file.filename == "":
        flash("Selecione um arquivo DOCX antes de continuar.", "error")
        return redirect(url_for("index"))

    if not is_allowed(uploaded_file.filename):
        flash("Formato inválido. Envie um arquivo .docx.", "error")
        return redirect(url_for("index"))

    safe_name = secure_filename(uploaded_file.filename)
    output_name = f"{Path(safe_name).stem}_institucional.pdf"

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=True) as temp_docx:
        uploaded_file.save(temp_docx.name)
        model = parse_docx(temp_docx.name)
        pdf_bytes = build_pdf(model)

    return send_file(
        BytesIO(pdf_bytes),
        as_attachment=True,
        download_name=output_name,
        mimetype="application/pdf",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
