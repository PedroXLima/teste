"""Parse DOCX into a structured model for institutional PDF rendering."""

from __future__ import annotations

import html
import io
import re
from typing import Any, Iterator, List, Optional

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph


def iter_block_items(document: Document) -> Iterator[DocxParagraph | DocxTable]:
    """Yield paragraphs and tables in document order."""
    body = document.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield DocxParagraph(child, document)
        elif child.tag == qn("w:tbl"):
            yield DocxTable(child, document)


def _paragraph_style_name(paragraph: DocxParagraph) -> str:
    try:
        return paragraph.style.name if paragraph.style else "Normal"
    except (AttributeError, ValueError):
        return "Normal"


def _runs_to_html(paragraph: DocxParagraph) -> str:
    parts: List[str] = []
    for run in paragraph.runs:
        t = html.escape(run.text)
        if not t:
            continue
        if run.bold:
            t = f"<strong>{t}</strong>"
        if run.italic:
            t = f"<em>{t}</em>"
        if run.underline:
            t = f'<span class="run-underline">{t}</span>'
        parts.append(t)
    if parts:
        return "".join(parts)
    return html.escape(paragraph.text or "")


_RGX_META = re.compile(
    r"^\s*(vers[aã]o|version|revis[aã]o|data|date|autor(?:es)?|revisor(?:es)?)\s*[:]\s*(.+)$",
    re.I,
)


def _guess_heading_level(style_name: str) -> Optional[int]:
    s = style_name.strip().lower()
    for i in range(1, 10):
        if s == f"heading {i}":
            return i
        if s.startswith(f"heading {i}"):
            return i
        if s == f"título {i}" or s.startswith(f"título {i}"):
            return i
    return None


def _is_toc_style(style_name: str) -> bool:
    s = style_name.upper()
    return "TOC" in s or s.startswith("ÍNDICE") or "TOC " in style_name


def _is_callout_style(style_name: str) -> bool:
    s = style_name.lower()
    return any(
        x in s
        for x in (
            "quote",
            "citação",
            "citacao",
            "intense quote",
            "destaque",
            "callout",
        )
    )


def _blob_to_data_uri(content_type: str, blob: bytes) -> str:
    ct = content_type or "image/png"
    if ct == "image/x-emf":
        ct = "image/png"
    import base64

    b64 = base64.b64encode(blob).decode("ascii")
    return f"data:{ct};base64,{b64}"


def _image_parts_from_paragraph(
    paragraph: DocxParagraph, document: Document
) -> List[str]:
    """Extract inline images as data URIs."""
    part = document.part
    uris: List[str] = []
    root = paragraph._element
    embed_tag = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
    for blip in root.findall(".//" + qn("a:blip")):
        rid = blip.get(embed_tag)
        if not rid:
            continue
        try:
            rel = part.rels[rid]
            if getattr(rel, "is_external", False):
                continue
            ipart = rel.target_part
            blob = ipart.blob
            ctype = getattr(ipart, "content_type", None) or "image/png"
            uris.append(_blob_to_data_uri(str(ctype), blob))
        except Exception:
            continue
    return uris


def _parse_table(table: DocxTable, table_index: int) -> dict[str, Any]:
    rows_raw = list(table.rows)
    if not rows_raw:
        return {
            "kind": "table",
            "table_index": table_index,
            "title": "",
            "headers": [],
            "rows": [],
            "source": "",
            "has_header": False,
        }

    if len(rows_raw) == 1:
        cells = [html.escape(c.text.strip()) for c in rows_raw[0].cells]
        return {
            "kind": "table",
            "table_index": table_index,
            "title": "",
            "headers": [],
            "rows": [cells],
            "source": "",
            "has_header": False,
        }

    headers = [html.escape(c.text.strip()) for c in rows_raw[0].cells]
    body_rows: List[List[str]] = []
    for r in rows_raw[1:]:
        body_rows.append([html.escape(c.text.strip()) for c in r.cells])

    return {
        "kind": "table",
        "table_index": table_index,
        "title": "",
        "headers": headers,
        "rows": body_rows,
        "source": "",
        "has_header": True,
    }


def _normalize_ws(text: str) -> str:
    return " ".join(text.split())


def parse_docx_bytes(data: bytes) -> dict[str, Any]:
    doc = Document(io.BytesIO(data))
    cp = doc.core_properties

    blocks: List[dict[str, Any]] = []
    cover_logo_uris: List[str] = []
    table_idx = 0
    figure_idx = 0

    for block in iter_block_items(doc):
        if isinstance(block, DocxParagraph):
            style = _paragraph_style_name(block)
            text_plain = (block.text or "").strip()
            imgs = _image_parts_from_paragraph(block, doc)

            if imgs:
                st_low = style.lower()
                is_caption_style = "caption" in st_low or "legenda" in st_low
                for uri in imgs:
                    if len(cover_logo_uris) < 6 and len(blocks) < 12:
                        cover_logo_uris.append(uri)
                    cap = ""
                    if is_caption_style and text_plain:
                        cap = html.escape(_normalize_ws(text_plain))
                    blocks.append(
                        {
                            "kind": "figure",
                            "figure_index": figure_idx,
                            "src": uri,
                            "caption_html": cap,
                            "caption_style": style,
                            "body_html": _runs_to_html(block)
                            if text_plain and not is_caption_style
                            else "",
                        }
                    )
                    figure_idx += 1
                if text_plain and not is_caption_style:
                    blocks.append(
                        {
                            "kind": "paragraph",
                            "style": style,
                            "html": _runs_to_html(block),
                            "plain": text_plain,
                            "is_toc": _is_toc_style(style),
                            "heading_level": _guess_heading_level(style),
                            "is_callout": _is_callout_style(style),
                        }
                    )
                continue

            if not text_plain:
                continue

            hl = _guess_heading_level(style)
            blocks.append(
                {
                    "kind": "paragraph",
                    "style": style,
                    "html": _runs_to_html(block),
                    "plain": text_plain,
                    "is_toc": _is_toc_style(style),
                    "heading_level": hl,
                    "is_callout": _is_callout_style(style),
                }
            )
        else:
            tbl = _parse_table(block, table_idx)
            blocks.append(tbl)
            table_idx += 1

    first_body_idx = 0
    for i, b in enumerate(blocks):
        if b["kind"] == "paragraph" and b.get("heading_level") == 1:
            first_body_idx = i
            break
        if b["kind"] == "paragraph" and re.match(
            r"^(sum[aá]rio|índice|conteúdo|conteudo|table of contents)\b",
            b["plain"],
            re.I,
        ):
            first_body_idx = i
            break

    front_blocks = blocks[:first_body_idx]
    body_blocks = blocks[first_body_idx:]

    title = (cp.title or "").strip()
    subtitle = (cp.subject or "").strip()
    authors = (cp.author or "").strip()
    version = ""
    doc_date = ""
    keywords = (cp.keywords or "").strip()

    if keywords:
        for part in keywords.replace(";", ",").split(","):
            part = part.strip()
            m = re.match(r"version\s*[:]\s*(.+)", part, re.I)
            if m:
                version = m.group(1).strip()
            m = re.match(r"data\s*[:]\s*(.+)", part, re.I)
            if m:
                doc_date = m.group(1).strip()

    front_lines = []
    for b in front_blocks:
        if b["kind"] != "paragraph":
            continue
        front_lines.append(b["plain"])

    if not title:
        for line in front_lines:
            if len(line) > 3:
                title = line
                break
    if not title:
        for b in body_blocks:
            if b["kind"] == "paragraph":
                title = b["plain"][:500]
                break
        title = title or "Relatório"

    if not subtitle and len(front_lines) >= 2:
        cand = front_lines[1]
        if cand != title and len(cand) < 400:
            subtitle = cand

    credits_groups: List[dict[str, str]] = []
    for b in front_blocks:
        if b["kind"] != "paragraph":
            continue
        m = _RGX_META.match(b["plain"])
        if m:
            label = m.group(1).strip().title()
            value = m.group(2).strip()
            credits_groups.append({"label": label, "value": value})
            if label.lower().startswith("vers"):
                version = version or value
            if label.lower().startswith("data"):
                doc_date = doc_date or value

    toc_entries: List[dict[str, Any]] = []
    anchor_counter = [0]

    def next_anchor() -> str:
        anchor_counter[0] += 1
        return f"h-{anchor_counter[0]}"

    prepared_body: List[dict[str, Any]] = []
    first_h1_done = False

    for b in body_blocks:
        if b["kind"] == "paragraph":
            hl = b.get("heading_level")
            plain = b.get("plain") or ""
            is_toc = b.get("is_toc")
            if hl == 1:
                aid = next_anchor()
                b = {**b, "anchor": aid}
                if not is_toc:
                    toc_entries.append(
                        {"level": 1, "text": plain, "anchor": aid}
                    )
                if first_h1_done:
                    prepared_body.append(
                        {
                            "kind": "section_divider",
                            "title_html": html.escape(plain),
                            "slug": plain[:80],
                            "banner": "metodologia" in plain.lower(),
                        }
                    )
                first_h1_done = True
            elif hl == 2:
                aid = next_anchor()
                b = {**b, "anchor": aid}
                if not is_toc:
                    toc_entries.append(
                        {"level": 2, "text": plain, "anchor": aid}
                    )
            elif hl == 3:
                aid = next_anchor()
                b = {**b, "anchor": aid}
                if not is_toc:
                    toc_entries.append(
                        {"level": 3, "text": plain, "anchor": aid}
                    )

            prepared_body.append(b)
        else:
            prepared_body.append(b)

    ref_mode = False
    stamped: List[dict[str, Any]] = []
    for b in prepared_body:
        if b["kind"] == "paragraph":
            hl = b.get("heading_level")
            plain = (b.get("plain") or "").strip().lower()
            if hl in (1, 2):
                if re.match(r"^(referências|referencias|bibliografia)\b", plain):
                    ref_mode = True
                elif hl == 1:
                    ref_mode = False
            elif ref_mode and not hl:
                b = {**b, "hang_reference": True}
        stamped.append(b)
    prepared_body = stamped

    cover_badges: List[str] = []
    if version:
        cover_badges.append(version)
    if doc_date:
        cover_badges.append(doc_date)

    return {
        "meta": {
            "title": title,
            "subtitle": subtitle,
            "authors": authors,
            "version": version,
            "date": doc_date,
            "keywords": keywords,
        },
        "cover_logos": cover_logo_uris[:4],
        "credits_groups": credits_groups,
        "front_has_content": bool(front_blocks),
        "blocks": prepared_body,
        "toc": toc_entries,
        "report_title_short": title[:120],
        "cover_badges": cover_badges,
    }
