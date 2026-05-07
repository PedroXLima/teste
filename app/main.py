from __future__ import annotations

import os
import uuid
from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from app.pdf_generator import build_pdf_from_docx


BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "generated" / "uploads"
OUTPUT_DIR = BASE_DIR / "generated" / "output"
ALLOWED_EXTENSIONS = {"docx"}


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder=str(BASE_DIR / "app" / "templates"),
        static_folder=str(BASE_DIR / "app" / "static"),
    )
    app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")
    app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    @app.get("/")
    def index() -> str:
        return render_template("index.html")

    @app.post("/generate")
    def generate():
        uploaded_file = request.files.get("docx_file")
        if uploaded_file is None or uploaded_file.filename == "":
            flash("Selecione um arquivo DOCX para continuar.", "error")
            return redirect(url_for("index"))

        filename = secure_filename(uploaded_file.filename)
        extension = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
        if extension not in ALLOWED_EXTENSIONS:
            flash("Apenas arquivos .docx sao suportados.", "error")
            return redirect(url_for("index"))

        token = uuid.uuid4().hex
        input_path = UPLOAD_DIR / f"{token}-{filename}"
        uploaded_file.save(input_path)

        try:
            pdf_path = build_pdf_from_docx(input_path, OUTPUT_DIR)
        except Exception as exc:  # pragma: no cover - surfaced through UI
            flash(f"Falha ao gerar o PDF: {exc}", "error")
            return redirect(url_for("index"))

        download_name = pdf_path.name
        return send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=download_name,
            max_age=0,
        )

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
