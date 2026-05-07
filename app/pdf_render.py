"""Render institutional PDF from parsed DOCX context using WeasyPrint."""

from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader, select_autoescape
from weasyprint import HTML


def render_pdf(context: dict) -> bytes:
    base = Path(__file__).resolve().parent
    env = Environment(
        loader=FileSystemLoader(base / "templates"),
        autoescape=select_autoescape(["html", "xml"]),
    )
    tpl = env.get_template("report.html")
    html_out = tpl.render(**context)
    return HTML(string=html_out, base_url=str(base)).write_pdf()
