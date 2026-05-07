"""
PDF Renderer — transforms parsed DOCX data into a styled HTML document
then converts it to PDF via WeasyPrint.

Implements the full institutional design system.
"""

from __future__ import annotations

import html
import os
import re
import textwrap
from pathlib import Path

from weasyprint import HTML

from .docx_parser import (
    BlockType,
    ContentBlock,
    DocumentMetadata,
    ParsedDocument,
    RunFragment,
)

ASSETS_DIR = Path(__file__).parent / "assets"

# ─────────────────────── colour palette ───────────────────────
PRIMARY_BLUE = "#005A9C"
DARK_NAVY = "#003B6F"
MEDIUM_BLUE = "#0B63B6"
LIGHT_BLUE = "#D9EAF7"
VERY_LIGHT_BLUE = "#EEF6FC"
WHITE = "#FFFFFF"
MAIN_TEXT = "#2B2B2B"
SECONDARY_TEXT = "#6B7280"
TABLE_BORDER = "#D1D5DB"

# ─────────────────────── CSS ──────────────────────────────────

CSS = f"""
@page {{
    size: A4 portrait;
    margin: 20mm 22mm 22mm 22mm;

    @top-left {{
        content: string(section-title);
        font-family: 'Inter', 'Roboto', sans-serif;
        font-size: 8pt;
        color: {SECONDARY_TEXT};
        border-bottom: 0.5pt solid {PRIMARY_BLUE};
        padding-bottom: 4pt;
    }}
    @top-right {{
        content: counter(page);
        font-family: 'Inter', 'Roboto', sans-serif;
        font-size: 8pt;
        color: {SECONDARY_TEXT};
        border-bottom: 0.5pt solid {PRIMARY_BLUE};
        padding-bottom: 4pt;
    }}
    @bottom-center {{
        content: "";
        border-top: 0.3pt solid {TABLE_BORDER};
    }}
}}

@page cover {{
    margin: 0;
    @top-left {{ content: none; border: none; }}
    @top-right {{ content: none; border: none; }}
    @bottom-center {{ content: none; border: none; }}
}}

@page back-cover {{
    margin: 0;
    @top-left {{ content: none; border: none; }}
    @top-right {{ content: none; border: none; }}
    @bottom-center {{ content: none; border: none; }}
}}

@page divider {{
    margin: 0;
    @top-left {{ content: none; border: none; }}
    @top-right {{ content: none; border: none; }}
    @bottom-center {{ content: none; border: none; }}
}}

@page credits {{
    @top-left {{ content: none; border: none; }}
    @top-right {{ content: none; border: none; }}
}}

@page toc {{
    @top-left {{ content: none; border: none; }}
    @top-right {{ content: none; border: none; }}
}}

/* ── reset ── */
* {{ margin: 0; padding: 0; box-sizing: border-box; }}

body {{
    font-family: 'Inter', 'Roboto', 'Source Sans Pro', sans-serif;
    font-size: 10.5pt;
    line-height: 1.65;
    color: {MAIN_TEXT};
    background: {WHITE};
}}

/* ── cover ── */
.cover-page {{
    page: cover;
    page-break-after: always;
    width: 210mm;
    height: 297mm;
    position: relative;
    overflow: hidden;
    background: linear-gradient(135deg, {DARK_NAVY} 0%, {PRIMARY_BLUE} 60%, {MEDIUM_BLUE} 100%);
}}

.cover-network {{
    position: absolute;
    top: 0; left: 0; right: 0; bottom: 0;
    background-image:
        radial-gradient(circle at 20% 80%, rgba(255,255,255,0.04) 0%, transparent 50%),
        radial-gradient(circle at 80% 20%, rgba(255,255,255,0.03) 0%, transparent 40%),
        radial-gradient(circle at 50% 50%, rgba(255,255,255,0.02) 0%, transparent 60%);
    background-size: 100% 100%;
}}

.cover-geometric {{
    position: absolute;
    bottom: 0; left: 0; right: 0;
    height: 60%;
    background-image:
        linear-gradient(30deg, rgba(255,255,255,0.015) 12%, transparent 12.5%,
            transparent 87%, rgba(255,255,255,0.015) 87.5%, rgba(255,255,255,0.015)),
        linear-gradient(150deg, rgba(255,255,255,0.015) 12%, transparent 12.5%,
            transparent 87%, rgba(255,255,255,0.015) 87.5%, rgba(255,255,255,0.015)),
        linear-gradient(60deg, rgba(255,255,255,0.015) 25%, transparent 25.5%,
            transparent 75%, rgba(255,255,255,0.015) 75%, rgba(255,255,255,0.015));
    background-size: 80px 140px;
}}

.cover-content {{
    position: absolute;
    bottom: 40mm;
    left: 25mm;
    right: 25mm;
}}

.cover-accent-bar {{
    width: 60px;
    height: 4px;
    background: {LIGHT_BLUE};
    margin-bottom: 18px;
    border-radius: 2px;
}}

.cover-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 28pt;
    font-weight: 800;
    color: {WHITE};
    text-transform: uppercase;
    letter-spacing: 0.5pt;
    line-height: 1.15;
    margin-bottom: 12px;
}}

.cover-subtitle {{
    font-family: 'Inter', 'Roboto', sans-serif;
    font-size: 13pt;
    font-weight: 400;
    color: rgba(255,255,255,0.85);
    line-height: 1.4;
    margin-bottom: 20px;
}}

.cover-meta {{
    font-family: 'Inter', 'Roboto', sans-serif;
    font-size: 8.5pt;
    color: rgba(255,255,255,0.6);
    line-height: 1.6;
}}

.cover-badge {{
    position: absolute;
    bottom: 15mm;
    right: 15mm;
    background: rgba(255,255,255,0.12);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 6px;
    padding: 6px 14px;
    font-family: 'Inter', sans-serif;
    font-size: 7.5pt;
    color: rgba(255,255,255,0.8);
}}

.cover-top-bar {{
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 6px;
    background: linear-gradient(90deg, {LIGHT_BLUE}, rgba(255,255,255,0.3));
}}

/* ── credits page ── */
.credits-page {{
    page: credits;
    page-break-after: always;
    padding-top: 15mm;
}}

.credits-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 20pt;
    font-weight: 700;
    color: {PRIMARY_BLUE};
    margin-bottom: 8mm;
    text-transform: uppercase;
    letter-spacing: 0.5pt;
}}

.credits-group {{
    margin-bottom: 6mm;
}}

.credits-label {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 9.5pt;
    font-weight: 700;
    color: {DARK_NAVY};
    text-transform: uppercase;
    letter-spacing: 0.3pt;
    margin-bottom: 2mm;
}}

.credits-names {{
    font-size: 10pt;
    color: {MAIN_TEXT};
    line-height: 1.6;
}}

/* ── section divider ── */
.section-divider {{
    page: divider;
    page-break-after: always;
    width: 210mm;
    height: 297mm;
    position: relative;
    overflow: hidden;
    background: linear-gradient(160deg, {DARK_NAVY} 0%, {PRIMARY_BLUE} 100%);
}}

.section-divider .cover-network {{
    opacity: 0.6;
}}

.divider-content {{
    position: absolute;
    bottom: 50mm;
    left: 25mm;
    right: 25mm;
}}

.divider-number {{
    font-family: 'Montserrat', sans-serif;
    font-size: 60pt;
    font-weight: 800;
    color: rgba(255,255,255,0.1);
    line-height: 1;
    margin-bottom: 5mm;
}}

.divider-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 26pt;
    font-weight: 700;
    color: {WHITE};
    text-transform: uppercase;
    letter-spacing: 0.5pt;
    line-height: 1.2;
}}

.divider-bar {{
    width: 60px;
    height: 4px;
    background: {LIGHT_BLUE};
    margin-top: 12px;
    border-radius: 2px;
}}

/* ── TOC ── */
.toc-page {{
    page: toc;
    page-break-after: always;
    padding-top: 10mm;
}}

.toc-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 20pt;
    font-weight: 700;
    color: {PRIMARY_BLUE};
    text-transform: uppercase;
    margin-bottom: 8mm;
    letter-spacing: 0.5pt;
}}

.toc-divider {{
    width: 100%;
    height: 1px;
    background: {TABLE_BORDER};
    margin-bottom: 6mm;
}}

/* ── headings ── */
h1 {{
    string-set: section-title content();
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 18pt;
    font-weight: 700;
    color: {PRIMARY_BLUE};
    margin-top: 10mm;
    margin-bottom: 3mm;
    line-height: 1.25;
    letter-spacing: 0.2pt;
    page-break-after: avoid;
}}

h1::after {{
    content: "";
    display: block;
    width: 100%;
    height: 1.5px;
    background: linear-gradient(90deg, {PRIMARY_BLUE}, {LIGHT_BLUE}, transparent);
    margin-top: 3mm;
}}

h2 {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 14pt;
    font-weight: 700;
    color: {DARK_NAVY};
    margin-top: 8mm;
    margin-bottom: 2.5mm;
    line-height: 1.3;
    page-break-after: avoid;
}}

h3 {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 12pt;
    font-weight: 600;
    color: {MEDIUM_BLUE};
    margin-top: 6mm;
    margin-bottom: 2mm;
    line-height: 1.35;
    page-break-after: avoid;
}}

h4 {{
    font-family: 'Inter', 'Roboto', sans-serif;
    font-size: 11pt;
    font-weight: 600;
    color: {DARK_NAVY};
    margin-top: 5mm;
    margin-bottom: 1.5mm;
    line-height: 1.4;
    page-break-after: avoid;
}}

/* ── body ── */
p {{
    margin-bottom: 3mm;
    text-align: justify;
    hyphens: auto;
    orphans: 3;
    widows: 3;
}}

p.left {{ text-align: left; }}
p.center {{ text-align: center; }}
p.right {{ text-align: right; }}

/* ── lists ── */
ul, ol {{
    margin-left: 6mm;
    margin-bottom: 3mm;
}}

li {{
    margin-bottom: 1.5mm;
    line-height: 1.6;
}}

/* ── tables ── */
.table-wrapper {{
    margin: 5mm 0;
    page-break-inside: avoid;
}}

.table-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 9.5pt;
    font-weight: 700;
    color: {DARK_NAVY};
    margin-bottom: 2mm;
}}

table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 9pt;
    line-height: 1.45;
}}

thead tr {{
    background: {DARK_NAVY};
}}

thead th {{
    color: {WHITE};
    font-weight: 700;
    padding: 3mm 2.5mm;
    text-align: left;
    font-size: 8.5pt;
    text-transform: uppercase;
    letter-spacing: 0.3pt;
}}

tbody tr:nth-child(even) {{
    background: {VERY_LIGHT_BLUE};
}}

tbody tr:nth-child(odd) {{
    background: {WHITE};
}}

tbody td {{
    padding: 2.5mm 2.5mm;
    border-bottom: 0.5pt solid {TABLE_BORDER};
    vertical-align: top;
}}

.table-source {{
    font-size: 7.5pt;
    color: {SECONDARY_TEXT};
    text-align: center;
    margin-top: 1.5mm;
    font-style: italic;
}}

/* ── images ── */
.figure-wrapper {{
    margin: 5mm 0;
    text-align: center;
    page-break-inside: avoid;
}}

.figure-wrapper img {{
    max-width: 100%;
    height: auto;
    border-radius: 4px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
}}

.figure-caption {{
    font-size: 8.5pt;
    color: {SECONDARY_TEXT};
    margin-top: 2mm;
    text-align: center;
    font-style: italic;
}}

/* ── captions / sources ── */
.caption-text {{
    font-size: 8.5pt;
    color: {SECONDARY_TEXT};
    font-style: italic;
    margin-bottom: 2mm;
    text-align: center;
}}

/* ── callout boxes ── */
.callout {{
    background: {VERY_LIGHT_BLUE};
    border-left: 4px solid {PRIMARY_BLUE};
    padding: 4mm 5mm;
    margin: 4mm 0;
    border-radius: 0 4px 4px 0;
    page-break-inside: avoid;
}}

.callout-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 9.5pt;
    font-weight: 700;
    color: {PRIMARY_BLUE};
    margin-bottom: 1.5mm;
}}

/* ── back cover ── */
.back-cover {{
    page: back-cover;
    page-break-before: always;
    width: 210mm;
    height: 297mm;
    position: relative;
    overflow: hidden;
    background: linear-gradient(160deg, {DARK_NAVY} 0%, {PRIMARY_BLUE} 100%);
}}

.back-cover .cover-network {{
    opacity: 0.5;
}}

.back-footer {{
    position: absolute;
    bottom: 20mm;
    left: 25mm;
    right: 25mm;
}}

.back-divider {{
    width: 100%;
    height: 1px;
    background: rgba(255,255,255,0.25);
    margin-bottom: 8mm;
}}

.back-title {{
    font-family: 'Montserrat', 'Poppins', sans-serif;
    font-size: 14pt;
    font-weight: 700;
    color: {WHITE};
    margin-bottom: 3mm;
}}

.back-subtitle {{
    font-size: 10pt;
    color: rgba(255,255,255,0.7);
    margin-bottom: 6mm;
}}

.back-contact {{
    font-size: 8.5pt;
    color: rgba(255,255,255,0.6);
    text-align: right;
}}

/* ── utilities ── */
.page-break {{
    page-break-after: always;
}}

strong {{ font-weight: 700; }}
em {{ font-style: italic; }}

.run-super {{ vertical-align: super; font-size: 0.7em; }}
.run-sub {{ vertical-align: sub; font-size: 0.7em; }}

/* ── references ── */
.references-block {{
    margin-bottom: 2mm;
    padding-left: 8mm;
    text-indent: -8mm;
    font-size: 9.5pt;
    line-height: 1.55;
}}
"""

# ─────────────────────── helpers ──────────────────────────────


def _esc(text: str) -> str:
    return html.escape(text) if text else ""


def _render_runs(runs: list[RunFragment], fallback_text: str = "") -> str:
    if not runs:
        return _esc(fallback_text)
    parts: list[str] = []
    for r in runs:
        t = _esc(r.text)
        if r.bold:
            t = f"<strong>{t}</strong>"
        if r.italic:
            t = f"<em>{t}</em>"
        if r.underline:
            t = f"<u>{t}</u>"
        if r.superscript:
            t = f'<span class="run-super">{t}</span>'
        if r.subscript:
            t = f'<span class="run-sub">{t}</span>'
        parts.append(t)
    return "".join(parts)


def _is_section_divider_candidate(text: str) -> bool:
    """Check if a heading should get a full-page section divider."""
    lower = text.strip().lower()
    stripped = re.sub(r"^\d+[\.\)]\s*", "", lower)
    keywords = {
        "introdução", "introducao", "introduction",
        "metodologia", "methodology",
        "desenvolvimento", "development",
        "resultados", "results",
        "conclusão", "conclusao", "conclusion", "conclusões", "conclusoes",
        "referências", "referencias", "references",
        "anexos", "anexo", "annex", "annexes",
        "considerações finais", "consideracoes finais",
        "recomendações", "recomendacoes",
        "apresentação", "apresentacao",
    }
    return stripped in keywords


_section_counter = 0


def _render_section_divider(text: str) -> str:
    global _section_counter
    _section_counter += 1
    return f"""
    <div class="section-divider">
        <div class="cover-network"></div>
        <div class="cover-geometric"></div>
        <div class="divider-content">
            <div class="divider-number">{_section_counter:02d}</div>
            <div class="divider-title">{_esc(text)}</div>
            <div class="divider-bar"></div>
        </div>
    </div>
    """


def _render_cover(meta: DocumentMetadata) -> str:
    meta_lines: list[str] = []
    if meta.authors:
        meta_lines.append(f"Autores: {', '.join(meta.authors)}")
    if meta.reviewers:
        meta_lines.append(f"Revisão: {', '.join(meta.reviewers)}")
    if meta.description:
        meta_lines.append(meta.description)

    badge = ""
    badge_parts = []
    if meta.version:
        badge_parts.append(meta.version)
    if meta.date:
        badge_parts.append(meta.date)
    if badge_parts:
        badge = f'<div class="cover-badge">{_esc(" | ".join(badge_parts))}</div>'

    return f"""
    <div class="cover-page">
        <div class="cover-top-bar"></div>
        <div class="cover-network"></div>
        <div class="cover-geometric"></div>
        <div class="cover-content">
            <div class="cover-accent-bar"></div>
            <div class="cover-title">{_esc(meta.title or "Relatório")}</div>
            {"<div class='cover-subtitle'>" + _esc(meta.subtitle) + "</div>" if meta.subtitle else ""}
            <div class="cover-meta">{"<br>".join(_esc(l) for l in meta_lines)}</div>
        </div>
        {badge}
    </div>
    """


def _render_credits(meta: DocumentMetadata) -> str:
    if not meta.credits and not meta.authors and not meta.reviewers:
        return ""

    groups = ""
    if meta.authors:
        groups += f"""
        <div class="credits-group">
            <div class="credits-label">Autores</div>
            <div class="credits-names">{", ".join(_esc(a) for a in meta.authors)}</div>
        </div>
        """
    if meta.reviewers:
        groups += f"""
        <div class="credits-group">
            <div class="credits-label">Revisão</div>
            <div class="credits-names">{", ".join(_esc(r) for r in meta.reviewers)}</div>
        </div>
        """
    for label, names in meta.credits.items():
        groups += f"""
        <div class="credits-group">
            <div class="credits-label">{_esc(label)}</div>
            <div class="credits-names">{", ".join(_esc(n) for n in names)}</div>
        </div>
        """
    if not groups:
        return ""

    return f"""
    <div class="credits-page">
        <div class="credits-title">Ficha Técnica</div>
        {groups}
    </div>
    """


def _render_back_cover(meta: DocumentMetadata) -> str:
    return f"""
    <div class="back-cover">
        <div class="cover-network"></div>
        <div class="cover-geometric"></div>
        <div class="back-footer">
            <div class="back-divider"></div>
            <div class="back-title">{_esc(meta.title or "Relatório")}</div>
            <div class="back-subtitle">{_esc(meta.subtitle or "")}</div>
            {"<div class='back-contact'>" + _esc(meta.contact) + "</div>" if meta.contact else ""}
        </div>
    </div>
    """


def _render_table(block: ContentBlock) -> str:
    if not block.table_data:
        return ""

    rows = block.table_data
    header = rows[0] if rows else []
    body = rows[1:] if len(rows) > 1 else []

    thead = "<thead><tr>" + "".join(f"<th>{_esc(c)}</th>" for c in header) + "</tr></thead>"
    tbody_rows = ""
    for row in body:
        tbody_rows += "<tr>" + "".join(f"<td>{_esc(c)}</td>" for c in row) + "</tr>"
    tbody = f"<tbody>{tbody_rows}</tbody>" if tbody_rows else ""

    return f"""
    <div class="table-wrapper">
        <table>{thead}{tbody}</table>
    </div>
    """


def _render_image(block: ContentBlock) -> str:
    caption = ""
    if block.caption:
        caption = f'<div class="figure-caption">{_esc(block.caption)}</div>'
    return f"""
    <div class="figure-wrapper">
        <img src="{block.image_data}" alt="{_esc(block.caption)}" />
        {caption}
    </div>
    """


# ─────────────────────── main renderer ────────────────────────


def render_html(doc: ParsedDocument) -> str:
    """Convert a ParsedDocument into a complete HTML string."""
    global _section_counter
    _section_counter = 0

    meta = doc.metadata
    blocks = doc.blocks
    parts: list[str] = []

    parts.append(_render_cover(meta))

    credits_html = _render_credits(meta)
    if credits_html:
        parts.append(credits_html)

    in_list = False
    list_type = "ul"
    skip_early_meta = True
    content_started = False

    i = 0
    while i < len(blocks):
        block = blocks[i]

        if skip_early_meta and block.block_type in (
            BlockType.TITLE, BlockType.SUBTITLE
        ):
            i += 1
            continue

        if block.block_type == BlockType.PAGE_BREAK:
            if in_list:
                parts.append(f"</{list_type}>")
                in_list = False
            parts.append('<div class="page-break"></div>')
            i += 1
            continue

        if block.block_type != BlockType.LIST_ITEM and in_list:
            parts.append(f"</{list_type}>")
            in_list = False

        skip_early_meta = False
        content_started = True

        if block.block_type == BlockType.HEADING1:
            text = block.text.strip()
            if _is_section_divider_candidate(text):
                parts.append(_render_section_divider(text))
            else:
                rendered = _render_runs(block.runs, text)
                parts.append(f"<h1>{rendered}</h1>")

        elif block.block_type == BlockType.HEADING2:
            rendered = _render_runs(block.runs, block.text)
            parts.append(f"<h2>{rendered}</h2>")

        elif block.block_type == BlockType.HEADING3:
            rendered = _render_runs(block.runs, block.text)
            parts.append(f"<h3>{rendered}</h3>")

        elif block.block_type == BlockType.HEADING4:
            rendered = _render_runs(block.runs, block.text)
            parts.append(f"<h4>{rendered}</h4>")

        elif block.block_type == BlockType.PARAGRAPH:
            rendered = _render_runs(block.runs, block.text)
            align_cls = block.alignment if block.alignment != "justify" else ""
            parts.append(f'<p class="{align_cls}">{rendered}</p>')

        elif block.block_type == BlockType.CAPTION:
            rendered = _render_runs(block.runs, block.text)
            parts.append(f'<p class="caption-text">{rendered}</p>')

        elif block.block_type == BlockType.LIST_ITEM:
            if not in_list:
                list_type = "ol" if block.list_style == "number" else "ul"
                parts.append(f"<{list_type}>")
                in_list = True
            rendered = _render_runs(block.runs, block.text)
            parts.append(f"<li>{rendered}</li>")

        elif block.block_type == BlockType.TABLE:
            parts.append(_render_table(block))

        elif block.block_type == BlockType.IMAGE:
            parts.append(_render_image(block))

        i += 1

    if in_list:
        parts.append(f"</{list_type}>")

    parts.append(_render_back_cover(meta))

    body = "\n".join(parts)

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<style>
{CSS}
</style>
</head>
<body>
{body}
</body>
</html>"""


def generate_pdf(doc: ParsedDocument, output_path: str) -> str:
    """Generate the final PDF file and return its path."""
    html_content = render_html(doc)
    html_obj = HTML(string=html_content)
    html_obj.write_pdf(output_path)
    return output_path
