"""Top-level DOCX → PDF conversion pipeline."""

from __future__ import annotations

import io
import os
from pathlib import Path
from typing import IO, Union

from weasyprint import HTML

from .docx_parser import parse_docx
from .renderer import render_html
from .structurer import structure


_BASE_DIR = Path(__file__).resolve().parent
_CSS_PATH = _BASE_DIR / "static" / "report.css"


def _load_css() -> str:
    return _CSS_PATH.read_text(encoding="utf-8")


def docx_to_pdf_bytes(docx_input: Union[bytes, IO[bytes], str, os.PathLike]) -> bytes:
    """Convert a DOCX file (path, bytes, or file-like) into PDF bytes."""
    if isinstance(docx_input, (bytes, bytearray)):
        file_obj: IO[bytes] = io.BytesIO(bytes(docx_input))
    elif isinstance(docx_input, (str, os.PathLike)):
        file_obj = open(docx_input, "rb")
    elif hasattr(docx_input, "read"):
        file_obj = docx_input
    else:
        raise TypeError("Unsupported docx_input type")

    try:
        parsed = parse_docx(file_obj)
    finally:
        if isinstance(docx_input, (str, os.PathLike)):
            file_obj.close()

    structured = structure(parsed)
    css = _load_css()
    html_str = render_html(structured, css)

    pdf_bytes = HTML(string=html_str, base_url=str(_BASE_DIR)).write_pdf()
    return pdf_bytes


def docx_to_html(docx_input: Union[bytes, IO[bytes], str, os.PathLike]) -> str:
    """Render to styled HTML (mostly useful for debugging)."""
    if isinstance(docx_input, (bytes, bytearray)):
        file_obj: IO[bytes] = io.BytesIO(bytes(docx_input))
    elif isinstance(docx_input, (str, os.PathLike)):
        file_obj = open(docx_input, "rb")
    elif hasattr(docx_input, "read"):
        file_obj = docx_input
    else:
        raise TypeError("Unsupported docx_input type")

    try:
        parsed = parse_docx(file_obj)
    finally:
        if isinstance(docx_input, (str, os.PathLike)):
            file_obj.close()

    structured = structure(parsed)
    css = _load_css()
    return render_html(structured, css)
