"""Render a StructuredDocument into HTML for WeasyPrint.

The HTML is intentionally designed to be styled by a CSS Paged Media
stylesheet (see ``static/report.css``). The renderer focuses on emitting
semantic, class-rich markup rather than inline styling so the design
system can evolve in the stylesheet alone.
"""

from __future__ import annotations

import html
import re
from typing import Iterable, List, Optional

from .docx_parser import Block, ImageBlock, ParsedDocument, Run, TableBlock
from .structurer import Section, StructuredDocument


def _esc(value: str) -> str:
    return html.escape(value or "", quote=True)


def _runs_to_html(runs: Iterable[Run]) -> str:
    parts: List[str] = []
    for run in runs:
        text = _esc(run.text).replace("\n", "<br/>")
        if not text:
            continue
        if run.bold:
            text = f"<strong>{text}</strong>"
        if run.italic:
            text = f"<em>{text}</em>"
        if run.underline:
            text = f"<u>{text}</u>"
        parts.append(text)
    return "".join(parts)


# ---------------------------------------------------------------------------
# Block-level rendering
# ---------------------------------------------------------------------------


def _render_heading(block: Block) -> str:
    level = max(1, min(block.level or 1, 4))
    text = _runs_to_html(block.runs) or _esc(block.text)
    cls = f"h h-{level}"
    return f'<h{level + 1} class="{cls}">{text}</h{level + 1}>'


def _render_paragraph(block: Block) -> str:
    text = _runs_to_html(block.runs) or _esc(block.text)
    if not text.strip():
        return '<p class="p empty">&nbsp;</p>'
    return f'<p class="p">{text}</p>'


def _render_caption(block: Block) -> str:
    text = _runs_to_html(block.runs) or _esc(block.text)
    return f'<p class="caption">{text}</p>'


def _render_quote(block: Block) -> str:
    text = _runs_to_html(block.runs) or _esc(block.text)
    return f'<blockquote class="callout"><p>{text}</p></blockquote>'


def _render_list(items: List[Block]) -> str:
    tag = "ul" if items[0].list_kind == "bullet" else "ol"
    li = "".join(
        f"<li>{_runs_to_html(b.runs) or _esc(b.text)}</li>" for b in items
    )
    return f'<{tag} class="list">{li}</{tag}>'


def _emu_to_px(emu: int) -> Optional[float]:
    if not emu:
        return None
    return emu / 9525.0


def _render_image(block: ImageBlock) -> str:
    width = _emu_to_px(block.width_emu)
    style = ""
    if width:
        # Cap to body content width (~16cm = 605px) for safety
        width_capped = min(width, 605)
        style = f' style="max-width:{width_capped:.0f}px"'
    return (
        '<figure class="figure">'
        f'<img class="figure-img" src="{_esc(block.data_uri)}" alt=""{style}/>'
        "</figure>"
    )


def _render_table(block: TableBlock) -> str:
    if not block.rows:
        return ""
    head_cells = block.rows[0]
    body_rows = block.rows[1:]
    head_html = "".join(
        f"<th>{_runs_to_html(cell)}</th>" for cell in head_cells
    )
    body_html_parts: List[str] = []
    for idx, row in enumerate(body_rows):
        cls = "alt" if idx % 2 == 1 else ""
        cells = "".join(f"<td>{_runs_to_html(cell)}</td>" for cell in row)
        body_html_parts.append(f'<tr class="{cls}">{cells}</tr>')
    return (
        '<div class="table-wrap">'
        '<table class="data-table">'
        f"<thead><tr>{head_html}</tr></thead>"
        f"<tbody>{''.join(body_html_parts)}</tbody>"
        "</table>"
        "</div>"
    )


def _render_blocks(blocks: List[Block]) -> str:
    out: List[str] = []
    list_buffer: List[Block] = []

    def flush_list():
        if list_buffer:
            out.append(_render_list(list_buffer))
            list_buffer.clear()

    for block in blocks:
        if block.kind == "list_item":
            if list_buffer and list_buffer[-1].list_kind != block.list_kind:
                flush_list()
            list_buffer.append(block)
            continue
        flush_list()

        if block.kind == "page_break":
            out.append('<div class="page-break"></div>')
        elif block.kind == "heading":
            out.append(_render_heading(block))
        elif block.kind == "caption":
            out.append(_render_caption(block))
        elif block.kind == "quote":
            out.append(_render_quote(block))
        elif block.kind == "image":
            out.append(_render_image(block))  # type: ignore[arg-type]
        elif block.kind == "table":
            out.append(_render_table(block))  # type: ignore[arg-type]
        elif block.kind == "paragraph":
            out.append(_render_paragraph(block))
        # ignore unknown kinds silently (defensive)

    flush_list()
    return "\n".join(out)


# ---------------------------------------------------------------------------
# Section rendering
# ---------------------------------------------------------------------------


def _render_cover(section: Section, metadata: dict) -> str:
    title = (metadata.get("title") or section.title or "Relatório Institucional").strip()
    subtitle = (metadata.get("subtitle") or "").strip()
    author = (metadata.get("author") or "").strip()
    reviewer = (metadata.get("last_modified_by") or "").strip()
    version = (metadata.get("version") or "").strip()
    modified = (metadata.get("modified") or "").strip()
    keywords = (metadata.get("keywords") or "").strip()
    category = (metadata.get("category") or "").strip()

    badge_text = ""
    if version and modified:
        badge_text = f"v{_esc(version)} · {_esc(modified[:10])}"
    elif version:
        badge_text = f"v{_esc(version)}"
    elif modified:
        badge_text = _esc(modified[:10])

    meta_lines: List[str] = []
    if category:
        meta_lines.append(
            f'<div class="cover-meta-line"><span class="cover-meta-label">Categoria</span><span>{_esc(category)}</span></div>'
        )
    if author:
        meta_lines.append(
            f'<div class="cover-meta-line"><span class="cover-meta-label">Autor</span><span>{_esc(author)}</span></div>'
        )
    if reviewer and reviewer != author:
        meta_lines.append(
            f'<div class="cover-meta-line"><span class="cover-meta-label">Revisão</span><span>{_esc(reviewer)}</span></div>'
        )
    if version:
        meta_lines.append(
            f'<div class="cover-meta-line"><span class="cover-meta-label">Versão</span><span>{_esc(version)}</span></div>'
        )
    if modified:
        meta_lines.append(
            f'<div class="cover-meta-line"><span class="cover-meta-label">Data</span><span>{_esc(modified[:10])}</span></div>'
        )

    badge_html = (
        f'<div class="cover-badge">{badge_text}</div>' if badge_text else ""
    )

    return f"""
<section class="page cover-page">
  <div class="cover-bg"></div>
  <div class="cover-network"></div>
  <div class="cover-overlay"></div>
  <div class="cover-top">
    <div class="cover-eyebrow">Relatório Institucional</div>
  </div>
  <div class="cover-panel">
    <h1 class="cover-title">{_esc(title.upper())}</h1>
    {f'<p class="cover-subtitle">{_esc(subtitle)}</p>' if subtitle else ''}
    <div class="cover-divider"></div>
    <div class="cover-meta">{''.join(meta_lines)}</div>
  </div>
  {badge_html}
</section>
""".strip()


def _render_credits(section: Section) -> str:
    body = _render_blocks(section.blocks)
    return f"""
<section class="page content-page credits-page">
  <header class="page-header">
    <span class="page-header-title">Créditos Institucionais</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    <h2 class="section-title">{_esc(section.title or 'Ficha Técnica')}</h2>
    <div class="section-rule"></div>
    <div class="credits-body">{body}</div>
  </main>
</section>
""".strip()


def _render_copyright(section: Section) -> str:
    body = _render_blocks(section.blocks)
    return f"""
<section class="page content-page copyright-page">
  <main class="page-main">
    <h2 class="section-title small">{_esc(section.title or 'Ficha Catalográfica')}</h2>
    <div class="section-rule"></div>
    <div class="copyright-body">{body}</div>
  </main>
</section>
""".strip()


def _render_presentation(section: Section) -> str:
    body = _render_blocks(section.blocks)
    return f"""
<section class="page content-page presentation-page">
  <header class="page-header">
    <span class="page-header-title">{_esc(section.title or 'Apresentação')}</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    <h2 class="section-title">{_esc(section.title or 'Apresentação')}</h2>
    <div class="section-rule"></div>
    <div class="presentation-body">{body}</div>
  </main>
</section>
""".strip()


def _render_toc(section: Section, structured: StructuredDocument) -> str:
    """Render a generated table of contents based on the structured doc."""
    rows: List[str] = []
    for s in structured.sections:
        if s.kind in {"cover", "back_cover", "toc", "copyright"}:
            continue
        title = s.title or _section_default_title(s.kind)
        if not title:
            continue
        rows.append(
            '<li class="toc-item">'
            f'<span class="toc-title">{_esc(title)}</span>'
            '<span class="toc-leader"></span>'
            '<span class="toc-page" data-target="'
            f'{_section_anchor(s)}"></span>'
            "</li>"
        )
        # Sub-headings inside body sections
        if s.kind == "section":
            for block in s.blocks:
                if block.kind == "heading" and block.level in (2, 3):
                    indent = "indent-1" if block.level == 2 else "indent-2"
                    rows.append(
                        f'<li class="toc-item sub {indent}">'
                        f'<span class="toc-title">{_esc(block.text)}</span>'
                        '<span class="toc-leader"></span>'
                        f'<span class="toc-page"></span>'
                        "</li>"
                    )
    body_blocks = _render_blocks(section.blocks) if section.blocks else ""
    return f"""
<section class="page content-page toc-page">
  <header class="page-header">
    <span class="page-header-title">Sumário</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    <h2 class="section-title">Sumário</h2>
    <div class="section-rule"></div>
    {body_blocks}
    <ol class="toc-list">{''.join(rows)}</ol>
  </main>
</section>
""".strip()


def _render_simple_list_section(section: Section, label: str) -> str:
    body = _render_blocks(section.blocks)
    return f"""
<section class="page content-page list-page">
  <header class="page-header">
    <span class="page-header-title">{_esc(label)}</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    <h2 class="section-title">{_esc(section.title or label)}</h2>
    <div class="section-rule"></div>
    {body}
  </main>
</section>
""".strip()


def _render_section_divider(title: str) -> str:
    return f"""
<section class="page section-divider">
  <div class="divider-bg"></div>
  <div class="divider-overlay"></div>
  <div class="divider-network"></div>
  <div class="divider-content">
    <div class="divider-eyebrow">Seção</div>
    <h2 class="divider-title">{_esc(title.upper())}</h2>
    <div class="divider-accent"></div>
  </div>
</section>
""".strip()


def _slugify(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "-", value or "").strip("-").lower()
    return value or "section"


def _section_anchor(section: Section) -> str:
    return f"sec-{_slugify(section.title or section.kind)}"


def _section_default_title(kind: str) -> str:
    return {
        "credits": "Ficha Técnica",
        "copyright": "Ficha Catalográfica",
        "presentation": "Apresentação",
        "toc": "Sumário",
        "list_figures": "Lista de Figuras",
        "list_tables": "Lista de Tabelas",
        "abbreviations": "Lista de Abreviações",
        "references": "Referências",
        "annexes": "Anexos",
        "back_cover": "Contato",
        "section": "",
    }.get(kind, "")


def _render_body_section(section: Section) -> str:
    title = section.title.strip()
    body_html = _render_blocks(section.blocks)
    anchor = _section_anchor(section)
    divider = _render_section_divider(title) if title else ""
    return f"""
{divider}
<section id="{anchor}" class="page content-page body-section">
  <header class="page-header">
    <span class="page-header-title">{_esc(title or 'Conteúdo')}</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    {body_html}
  </main>
</section>
""".strip()


def _render_references(section: Section) -> str:
    items: List[str] = []
    for block in section.blocks:
        if block.kind == "paragraph" and (block.text or "").strip():
            items.append(
                f'<li class="ref-item">{_runs_to_html(block.runs) or _esc(block.text)}</li>'
            )
        elif block.kind == "list_item":
            items.append(
                f'<li class="ref-item">{_runs_to_html(block.runs) or _esc(block.text)}</li>'
            )
    body = f'<ul class="ref-list">{"".join(items)}</ul>' if items else _render_blocks(section.blocks)
    return f"""
<section class="page content-page references-page">
  <header class="page-header">
    <span class="page-header-title">Referências</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    <h2 class="section-title">{_esc(section.title or 'Referências')}</h2>
    <div class="section-rule"></div>
    {body}
  </main>
</section>
""".strip()


def _render_annexes(section: Section) -> str:
    body = _render_blocks(section.blocks)
    return f"""
<section class="page content-page annexes-page">
  <header class="page-header">
    <span class="page-header-title">Anexos</span>
    <span class="page-header-divider"></span>
  </header>
  <main class="page-main">
    <h2 class="section-title big">{_esc((section.title or 'Anexos').upper())}</h2>
    <div class="section-rule"></div>
    {body}
  </main>
</section>
""".strip()


def _render_back_cover(section: Section, metadata: dict) -> str:
    title = (metadata.get("title") or "").strip()
    subtitle = (metadata.get("subtitle") or "").strip()
    contact_blocks = _render_blocks(section.blocks) if section.blocks else ""
    return f"""
<section class="page back-cover-page">
  <div class="back-bg"></div>
  <div class="back-network"></div>
  <div class="back-content">
    <div class="back-bottom">
      <div class="back-bottom-left">
        {f'<div class="back-title">{_esc(title)}</div>' if title else ''}
        {f'<div class="back-subtitle">{_esc(subtitle)}</div>' if subtitle else ''}
      </div>
      <div class="back-bottom-right">{contact_blocks}</div>
    </div>
    <div class="back-divider"></div>
    <div class="back-footer">
      Observatório da Indústria · Sistema FIEA
    </div>
  </div>
</section>
""".strip()


# ---------------------------------------------------------------------------
# Top-level render
# ---------------------------------------------------------------------------


def render_html(structured: StructuredDocument, css: str) -> str:
    metadata = structured.metadata
    body_parts: List[str] = []

    for section in structured.sections:
        if section.kind == "cover":
            body_parts.append(_render_cover(section, metadata))
        elif section.kind == "credits":
            body_parts.append(_render_credits(section))
        elif section.kind == "copyright":
            body_parts.append(_render_copyright(section))
        elif section.kind == "presentation":
            body_parts.append(_render_presentation(section))
        elif section.kind == "toc":
            body_parts.append(_render_toc(section, structured))
        elif section.kind == "list_figures":
            body_parts.append(_render_simple_list_section(section, "Lista de Figuras"))
        elif section.kind == "list_tables":
            body_parts.append(_render_simple_list_section(section, "Lista de Tabelas"))
        elif section.kind == "abbreviations":
            body_parts.append(_render_simple_list_section(section, "Lista de Abreviações"))
        elif section.kind == "references":
            body_parts.append(_render_references(section))
        elif section.kind == "annexes":
            body_parts.append(_render_annexes(section))
        elif section.kind == "back_cover":
            body_parts.append(_render_back_cover(section, metadata))
        else:
            body_parts.append(_render_body_section(section))

    title_text = _esc(metadata.get("title") or "Relatório Institucional")

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="utf-8"/>
<title>{title_text}</title>
<style>{css}</style>
</head>
<body>
{''.join(body_parts)}
</body>
</html>
"""
