"""Flask web app for uploading a DOCX and downloading the styled PDF."""

from __future__ import annotations

import io
import os
import re
import time
from pathlib import Path

from flask import Flask, Response, abort, jsonify, render_template, request, send_file

from .converter import docx_to_html, docx_to_pdf_bytes


_BASE_DIR = Path(__file__).resolve().parent
MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder=str(_BASE_DIR / "templates"),
        static_folder=str(_BASE_DIR / "static"),
    )
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

    @app.get("/")
    def index() -> str:
        return render_template("index.html")

    @app.get("/healthz")
    def healthz():
        return jsonify({"ok": True})

    @app.post("/convert")
    def convert():
        upload = request.files.get("docx")
        if upload is None or upload.filename == "":
            abort(400, description="Nenhum arquivo enviado.")

        filename = upload.filename or "documento.docx"
        if not filename.lower().endswith(".docx"):
            abort(400, description="Envie um arquivo .docx válido.")

        data = upload.read()
        if not data:
            abort(400, description="Arquivo vazio.")

        try:
            pdf_bytes = docx_to_pdf_bytes(data)
        except Exception as exc:  # pragma: no cover - surfaces parser errors to UI
            app.logger.exception("Falha ao converter DOCX")
            return Response(
                f"Erro ao converter o arquivo: {exc}",
                status=500,
                mimetype="text/plain; charset=utf-8",
            )

        out_name = re.sub(r"\.docx$", "", filename, flags=re.IGNORECASE) + ".pdf"
        return send_file(
            io.BytesIO(pdf_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=out_name,
        )

    @app.post("/preview")
    def preview():
        """Return rendered HTML for debugging in the browser."""
        upload = request.files.get("docx")
        if upload is None or upload.filename == "":
            abort(400, description="Nenhum arquivo enviado.")
        data = upload.read()
        if not data:
            abort(400, description="Arquivo vazio.")
        html_str = docx_to_html(data)
        return Response(html_str, mimetype="text/html; charset=utf-8")

    @app.errorhandler(413)
    def too_large(_):
        return Response(
            "Arquivo muito grande. Limite: 50 MB.",
            status=413,
            mimetype="text/plain; charset=utf-8",
        )

    return app


app = create_app()


if __name__ == "__main__":  # pragma: no cover
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
