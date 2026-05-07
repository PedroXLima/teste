"""DOCX → structured document model.

This module reads a .docx file with python-docx and produces a list of
"blocks" that the renderer can consume. The goal is to preserve all
original content (text, headings, tables, images, captions, lists,
references, annexes) while inferring enough structural metadata to drive
the institutional design system used by the renderer.
"""

from __future__ import annotations

import base64
import io
import os
import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from docx import Document
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------


@dataclass
class Run:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False


@dataclass
class Block:
    kind: str  # paragraph | heading | caption | quote | list_item | table | image | toc | page_break | section_break
    level: int = 0  # heading level
    runs: List[Run] = field(default_factory=list)
    text: str = ""
    style: str = ""
    list_kind: str = ""  # bullet | number
    extra: Dict[str, Any] = field(default_factory=dict)


@dataclass
class TableBlock(Block):
    rows: List[List[List[Run]]] = field(default_factory=list)


@dataclass
class ImageBlock(Block):
    data_uri: str = ""
    width_emu: int = 0
    height_emu: int = 0


@dataclass
class ParsedDocument:
    blocks: List[Block]
    metadata: Dict[str, Any]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _runs_from_paragraph(paragraph: Paragraph) -> List[Run]:
    runs: List[Run] = []
    for r in paragraph.runs:
        text = r.text or ""
        if not text:
            continue
        runs.append(
            Run(
                text=text,
                bold=bool(r.bold),
                italic=bool(r.italic),
                underline=bool(r.underline),
            )
        )
    if not runs and paragraph.text:
        runs.append(Run(text=paragraph.text))
    return runs


def _strip_accents(value: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFKD", value) if not unicodedata.combining(c)
    )


def _normalize_style(name: str) -> str:
    return _strip_accents((name or "").lower()).strip()


_HEADING_RE = re.compile(r"^heading\s*(\d+)|^titulo\s*(\d+)", re.IGNORECASE)


def _detect_heading_level(style_name: str) -> Optional[int]:
    if not style_name:
        return None
    norm = _normalize_style(style_name)
    m = _HEADING_RE.match(norm)
    if m:
        for g in m.groups():
            if g:
                try:
                    return max(1, min(6, int(g)))
                except ValueError:
                    pass
    if norm in {"title", "titulo"}:
        return 1
    if norm in {"subtitle", "subtitulo"}:
        return 2
    return None


def _detect_caption(style_name: str, text: str) -> bool:
    norm = _normalize_style(style_name)
    if "caption" in norm or "legenda" in norm:
        return True
    # heuristic for Brazilian institutional docs
    if re.match(r"^(figura|tabela|quadro|gr[áa]fico)\s+\d+\s*[-–:]", text.strip(), re.I):
        return True
    return False


def _detect_quote(style_name: str) -> bool:
    norm = _normalize_style(style_name)
    return any(token in norm for token in ("quote", "citacao", "citation"))


def _detect_list_kind(paragraph: Paragraph) -> Optional[str]:
    style_name = _normalize_style(paragraph.style.name if paragraph.style else "")
    if "list bullet" in style_name or "lista com marcadores" in style_name:
        return "bullet"
    if "list number" in style_name or "lista numerada" in style_name:
        return "number"
    # check numbering xml
    p = paragraph._p
    numpr = p.find(
        ".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numPr"
    )
    if numpr is not None:
        return "bullet"
    return None


def _iter_block_items(parent):
    """Yield paragraphs and tables in document order."""
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("unsupported parent")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# ---------------------------------------------------------------------------
# Image extraction
# ---------------------------------------------------------------------------


_DRAWING_NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
}


def _paragraph_images(paragraph: Paragraph, doc) -> List[ImageBlock]:
    blocks: List[ImageBlock] = []
    blips = paragraph._p.findall(
        ".//a:blip", _DRAWING_NS
    )
    if not blips:
        return blocks
    extents = paragraph._p.findall(".//wp:extent", _DRAWING_NS)
    for idx, blip in enumerate(blips):
        rId = blip.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
        )
        if not rId:
            continue
        try:
            part = doc.part.related_parts[rId]
        except KeyError:
            continue
        content_type = getattr(part, "content_type", "image/png") or "image/png"
        data = part.blob
        b64 = base64.b64encode(data).decode("ascii")
        data_uri = f"data:{content_type};base64,{b64}"
        cx = cy = 0
        if idx < len(extents):
            try:
                cx = int(extents[idx].get("cx", 0))
                cy = int(extents[idx].get("cy", 0))
            except ValueError:
                cx = cy = 0
        blocks.append(
            ImageBlock(
                kind="image",
                data_uri=data_uri,
                width_emu=cx,
                height_emu=cy,
            )
        )
    return blocks


# ---------------------------------------------------------------------------
# Metadata inference
# ---------------------------------------------------------------------------


def _infer_metadata(doc, blocks: List[Block]) -> Dict[str, Any]:
    cp = doc.core_properties
    metadata: Dict[str, Any] = {
        "title": (cp.title or "").strip(),
        "subtitle": (cp.subject or "").strip(),
        "author": (cp.author or "").strip(),
        "last_modified_by": (cp.last_modified_by or "").strip(),
        "created": cp.created.isoformat() if cp.created else "",
        "modified": cp.modified.isoformat() if cp.modified else "",
        "version": (cp.revision and str(cp.revision)) or "",
        "keywords": (cp.keywords or "").strip(),
        "category": (cp.category or "").strip(),
        "comments": (cp.comments or "").strip(),
    }

    # If title missing, take the first level-1 heading or the first non-empty paragraph
    if not metadata["title"]:
        for b in blocks:
            if b.kind == "heading" and b.level == 1 and b.text.strip():
                metadata["title"] = b.text.strip()
                break
        if not metadata["title"]:
            for b in blocks:
                if b.kind == "paragraph" and b.text.strip():
                    metadata["title"] = b.text.strip()[:120]
                    break

    if not metadata["subtitle"]:
        # try a level-2 heading right after the title
        seen_title = False
        for b in blocks:
            if b.kind == "heading" and b.level == 1:
                seen_title = True
                continue
            if seen_title and b.kind == "heading" and b.level in (2, 3) and b.text.strip():
                metadata["subtitle"] = b.text.strip()
                break
            if seen_title and b.kind == "paragraph" and b.text.strip():
                metadata["subtitle"] = b.text.strip()[:160]
                break
    return metadata


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def parse_docx(file_obj) -> ParsedDocument:
    """Parse a DOCX file into structured blocks + metadata."""
    if hasattr(file_obj, "read"):
        data = file_obj.read()
        doc = Document(io.BytesIO(data))
    else:
        doc = Document(file_obj)

    blocks: List[Block] = []

    for item in _iter_block_items(doc):
        if isinstance(item, Paragraph):
            blocks.extend(_paragraph_to_blocks(item, doc))
        elif isinstance(item, Table):
            blocks.append(_table_to_block(item))

    metadata = _infer_metadata(doc, blocks)
    return ParsedDocument(blocks=blocks, metadata=metadata)


# ---------------------------------------------------------------------------
# Paragraph & table conversion
# ---------------------------------------------------------------------------


_TOC_HINT_RE = re.compile(
    r"^(sum[áa]rio|table of contents|conte[úu]do|[ií]ndice)\b", re.IGNORECASE
)


def _paragraph_to_blocks(paragraph: Paragraph, doc) -> List[Block]:
    text = paragraph.text or ""
    style_name = paragraph.style.name if paragraph.style else ""
    runs = _runs_from_paragraph(paragraph)
    images = _paragraph_images(paragraph, doc)

    # Page break detection
    breaks = paragraph._p.findall(
        ".//w:br[@w:type='page']",
        {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
    )

    out: List[Block] = []

    if breaks:
        out.append(Block(kind="page_break"))

    for img in images:
        out.append(img)

    if not text.strip() and not images and not breaks:
        # blank paragraph - keep some breathing room but don't be excessive
        out.append(Block(kind="paragraph", text="", runs=[], style=style_name))
        return out

    if not text.strip():
        return out

    heading_level = _detect_heading_level(style_name)
    if heading_level is not None:
        out.append(
            Block(
                kind="heading",
                level=heading_level,
                runs=runs,
                text=text,
                style=style_name,
            )
        )
        return out

    if _detect_caption(style_name, text):
        out.append(
            Block(
                kind="caption",
                runs=runs,
                text=text,
                style=style_name,
            )
        )
        return out

    if _detect_quote(style_name):
        out.append(
            Block(
                kind="quote",
                runs=runs,
                text=text,
                style=style_name,
            )
        )
        return out

    list_kind = _detect_list_kind(paragraph)
    if list_kind:
        out.append(
            Block(
                kind="list_item",
                runs=runs,
                text=text,
                style=style_name,
                list_kind=list_kind,
            )
        )
        return out

    out.append(
        Block(
            kind="paragraph",
            runs=runs,
            text=text,
            style=style_name,
        )
    )
    return out


def _cell_runs(cell: _Cell) -> List[Run]:
    runs: List[Run] = []
    first = True
    for para in cell.paragraphs:
        if not first and para.text:
            runs.append(Run(text="\n"))
        for r in _runs_from_paragraph(para):
            runs.append(r)
        first = False
    return runs


def _table_to_block(table: Table) -> TableBlock:
    rows: List[List[List[Run]]] = []
    for row in table.rows:
        row_cells: List[List[Run]] = []
        for cell in row.cells:
            row_cells.append(_cell_runs(cell))
        rows.append(row_cells)
    return TableBlock(kind="table", rows=rows)
