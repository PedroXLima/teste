"""
Flask application – serves the UI and the /api/convert endpoint.
"""

import os
import json
import uuid
import shutil
import tempfile

from flask import Flask, request, jsonify, send_file, render_template

from .docx_parser import parse_docx
from .pdf_generator import generate_pdf

MAX_UPLOAD_SIZE = 50 * 1024 * 1024  # 50 MB

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE

UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "docx2pdf_uploads")
OUTPUT_DIR = os.path.join(tempfile.gettempdir(), "docx2pdf_output")
IMAGES_DIR = os.path.join(tempfile.gettempdir(), "docx2pdf_images")

for d in (UPLOAD_DIR, OUTPUT_DIR, IMAGES_DIR):
    os.makedirs(d, exist_ok=True)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/convert", methods=["POST"])
def convert():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    file = request.files["file"]
    if not file.filename:
        return jsonify({"error": "Empty filename."}), 400

    if not file.filename.lower().endswith(".docx"):
        return jsonify({"error": "Only .docx files are supported."}), 400

    job_id = uuid.uuid4().hex[:12]
    job_upload = os.path.join(UPLOAD_DIR, job_id)
    job_images = os.path.join(IMAGES_DIR, job_id)
    os.makedirs(job_upload, exist_ok=True)
    os.makedirs(job_images, exist_ok=True)

    docx_path = os.path.join(job_upload, file.filename)
    file.save(docx_path)

    try:
        parsed = parse_docx(docx_path, job_images)

        meta_override = request.form.get("metadata")
        if meta_override:
            try:
                overrides = json.loads(meta_override)
                meta = parsed["metadata"]
                if overrides.get("title"):
                    meta["title"] = overrides["title"]
                if overrides.get("subtitle"):
                    meta["subtitle"] = overrides["subtitle"]
                if overrides.get("authors"):
                    meta["authors"] = [a.strip() for a in overrides["authors"].split(",")]
                if overrides.get("institution"):
                    meta["institution"] = overrides["institution"]
                if overrides.get("version"):
                    meta["version"] = overrides["version"]
                if overrides.get("date"):
                    meta["date"] = overrides["date"]
            except (json.JSONDecodeError, AttributeError):
                pass

        base_name = os.path.splitext(file.filename)[0]
        pdf_name = f"{base_name}_designed.pdf"
        pdf_path = os.path.join(OUTPUT_DIR, f"{job_id}_{pdf_name}")

        generate_pdf(parsed, pdf_path)

        return send_file(
            pdf_path,
            mimetype="application/pdf",
            as_attachment=True,
            download_name=pdf_name,
        )

    except Exception as e:
        return jsonify({"error": f"PDF generation failed: {str(e)}"}), 500

    finally:
        try:
            shutil.rmtree(job_upload, ignore_errors=True)
            shutil.rmtree(job_images, ignore_errors=True)
        except Exception:
            pass


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
