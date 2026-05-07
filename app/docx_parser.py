"""
DOCX Parser — extracts structured content from .docx files.

Produces a list of content blocks with semantic types so the
PDF renderer can apply the design system without losing any
original content.
"""

from __future__ import annotations

import base64
import io
import os
import re
from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Any

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from PIL import Image


class BlockType(Enum):
    TITLE = auto()
    SUBTITLE = auto()
    HEADING1 = auto()
    HEADING2 = auto()
    HEADING3 = auto()
    HEADING4 = auto()
    PARAGRAPH = auto()
    TABLE = auto()
    IMAGE = auto()
    LIST_ITEM = auto()
    PAGE_BREAK = auto()
    FOOTNOTE = auto()
    CAPTION = auto()
    TOC = auto()


@dataclass
class RunFragment:
    text: str
    bold: bool = False
    italic: bool = False
    underline: bool = False
    superscript: bool = False
    subscript: bool = False
    font_size: float | None = None
    color: str | None = None


@dataclass
class ContentBlock:
    block_type: BlockType
    text: str = ""
    runs: list[RunFragment] = field(default_factory=list)
    level: int = 0
    list_style: str = ""
    table_data: list[list[str]] = field(default_factory=list)
    table_header_rows: int = 1
    image_data: str = ""
    image_format: str = "png"
    image_width: int = 0
    image_height: int = 0
    caption: str = ""
    alignment: str = "left"
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class DocumentMetadata:
    title: str = ""
    subtitle: str = ""
    authors: list[str] = field(default_factory=list)
    reviewers: list[str] = field(default_factory=list)
    version: str = ""
    date: str = ""
    description: str = ""
    institution: str = ""
    credits: dict[str, list[str]] = field(default_factory=dict)
    contact: str = ""


@dataclass
class ParsedDocument:
    metadata: DocumentMetadata
    blocks: list[ContentBlock]
    images: dict[str, str] = field(default_factory=dict)


# ────────────────────────────────────────────────────────────────
# Heuristic patterns used to identify metadata in early paragraphs
# ────────────────────────────────────────────────────────────────

_RE_VERSION = re.compile(
    r"(?:vers[ãa]o|version|v\.?\s*)\s*[:.]?\s*(.+)", re.IGNORECASE
)
_RE_DATE = re.compile(
    r"(?:data|date)\s*[:.]?\s*(.+)", re.IGNORECASE
)
_RE_AUTHOR = re.compile(
    r"(?:autor(?:es|a)?|author[s]?)\s*[:.]?\s*(.+)", re.IGNORECASE
)
_RE_REVIEWER = re.compile(
    r"(?:revis[ãa]o|reviewer?[s]?|revisor(?:es)?)\s*[:.]?\s*(.+)",
    re.IGNORECASE,
)
_RE_SUBTITLE = re.compile(
    r"(?:subt[ií]tulo|subtitle)\s*[:.]?\s*(.+)", re.IGNORECASE
)

_CREDIT_LABELS = [
    "realização", "realizacao", "execução", "execucao",
    "coordenação", "coordenacao", "autores", "autoras",
    "projeto gráfico", "projeto grafico", "revisão", "revisao",
    "diagramação", "diagramacao", "colaboração", "colaboracao",
    "equipe técnica", "equipe tecnica", "supervisão", "supervisao",
]

_SECTION_KEYWORDS = {
    "introdução", "introducao", "introduction",
    "metodologia", "methodology",
    "desenvolvimento", "development",
    "resultados", "results",
    "conclusão", "conclusao", "conclusion", "conclusões", "conclusoes",
    "referências", "referencias", "references",
    "anexos", "anexo", "annex", "annexes",
    "apêndice", "apendice", "appendix",
    "sumário", "sumario", "summary",
    "apresentação", "apresentacao", "presentation",
    "prefácio", "prefacio", "preface",
    "lista de figuras", "lista de tabelas",
    "lista de abreviaturas", "abreviaturas",
}

_TOC_PATTERNS = re.compile(
    r"(sumário|sumario|table of contents|índice|indice|conteúdo|conteudo)",
    re.IGNORECASE,
)


def _extract_runs(paragraph) -> list[RunFragment]:
    fragments: list[RunFragment] = []
    for run in paragraph.runs:
        if not run.text:
            continue
        rp = run.font
        color = None
        if rp.color and rp.color.rgb:
            color = str(rp.color.rgb)
        fragments.append(
            RunFragment(
                text=run.text,
                bold=bool(rp.bold),
                italic=bool(rp.italic),
                underline=bool(rp.underline),
                superscript=bool(rp.superscript),
                subscript=bool(rp.subscript),
                font_size=rp.size.pt if rp.size else None,
                color=color,
            )
        )
    return fragments


def _classify_heading_level(paragraph) -> BlockType | None:
    style_name = (paragraph.style.name or "").lower()
    if "title" in style_name and "subtitle" not in style_name:
        return BlockType.TITLE
    if "subtitle" in style_name:
        return BlockType.SUBTITLE
    if style_name.startswith("heading"):
        parts = style_name.split()
        for p in parts:
            if p.isdigit():
                level = int(p)
                mapping = {
                    1: BlockType.HEADING1,
                    2: BlockType.HEADING2,
                    3: BlockType.HEADING3,
                }
                return mapping.get(level, BlockType.HEADING4)
        return BlockType.HEADING1
    return None


def _is_list_paragraph(paragraph) -> tuple[bool, str]:
    style_name = (paragraph.style.name or "").lower()
    if "list" in style_name:
        if "bullet" in style_name or "unordered" in style_name:
            return True, "bullet"
        if "number" in style_name or "ordered" in style_name:
            return True, "number"
        return True, "bullet"

    numPr = paragraph._element.find(qn("w:pPr"))
    if numPr is not None:
        numId = numPr.find(qn("w:numPr"))
        if numId is not None:
            return True, "bullet"
    return False, ""


def _get_alignment(paragraph) -> str:
    alignment = paragraph.alignment
    if alignment is None:
        return "left"
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    mapping = {
        WD_ALIGN_PARAGRAPH.CENTER: "center",
        WD_ALIGN_PARAGRAPH.RIGHT: "right",
        WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
        WD_ALIGN_PARAGRAPH.LEFT: "left",
    }
    return mapping.get(alignment, "left")


def _extract_table(table: Table) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.rows:
        cells = []
        for cell in row.cells:
            cells.append(cell.text.strip())
        rows.append(cells)
    return rows


def _extract_images(doc: Document) -> dict[str, str]:
    """Extract all images from the document as base64 strings keyed by rId."""
    images: dict[str, str] = {}
    for rel_id, rel in doc.part.rels.items():
        if "image" in rel.reltype:
            try:
                blob = rel.target_part.blob
                img = Image.open(io.BytesIO(blob))
                fmt = img.format or "PNG"
                buf = io.BytesIO()
                if img.mode in ("RGBA", "P"):
                    img = img.convert("RGBA")
                    fmt = "PNG"
                else:
                    img = img.convert("RGB")
                    if fmt.upper() not in ("JPEG", "JPG", "PNG"):
                        fmt = "PNG"
                img.save(buf, format=fmt)
                b64 = base64.b64encode(buf.getvalue()).decode("ascii")
                mime = "image/png" if fmt.upper() == "PNG" else "image/jpeg"
                images[rel_id] = f"data:{mime};base64,{b64}"
            except Exception:
                pass
    return images


def _paragraph_has_image(paragraph) -> str | None:
    """Return the rId of an embedded image if the paragraph contains one."""
    for run in paragraph.runs:
        drawing_elements = run._element.findall(
            f".//{qn('wp:inline')}"
        ) + run._element.findall(f".//{qn('wp:anchor')}")
        for drawing in drawing_elements:
            blip = drawing.find(f".//{qn('a:blip')}")
            if blip is not None:
                embed = blip.get(qn("r:embed"))
                if embed:
                    return embed
    return None


def _guess_metadata(blocks: list[ContentBlock]) -> DocumentMetadata:
    """Try to extract metadata heuristically from the first blocks."""
    meta = DocumentMetadata()
    scan_limit = min(len(blocks), 20)

    for i in range(scan_limit):
        b = blocks[i]
        text = b.text.strip()
        if not text:
            continue

        if b.block_type == BlockType.TITLE and not meta.title:
            meta.title = text
            continue
        if b.block_type == BlockType.SUBTITLE and not meta.subtitle:
            meta.subtitle = text
            continue

        m = _RE_VERSION.match(text)
        if m:
            meta.version = m.group(1).strip()
            continue
        m = _RE_DATE.match(text)
        if m:
            meta.date = m.group(1).strip()
            continue
        m = _RE_AUTHOR.match(text)
        if m:
            meta.authors = [
                a.strip() for a in re.split(r"[,;/]", m.group(1)) if a.strip()
            ]
            continue
        m = _RE_REVIEWER.match(text)
        if m:
            meta.reviewers = [
                r.strip() for r in re.split(r"[,;/]", m.group(1)) if r.strip()
            ]
            continue

        lower = text.lower()
        for label in _CREDIT_LABELS:
            if lower.startswith(label):
                rest = text[len(label):].strip().lstrip(":").strip()
                if rest:
                    meta.credits[label.title()] = [
                        n.strip() for n in re.split(r"[,;/\n]", rest) if n.strip()
                    ]
                break

    if not meta.title:
        for b in blocks[:10]:
            if b.block_type in (BlockType.HEADING1,) and b.text.strip():
                meta.title = b.text.strip()
                break
    if not meta.title:
        for b in blocks[:10]:
            if b.text.strip():
                meta.title = b.text.strip()
                break

    return meta


def parse_docx(file_path: str) -> ParsedDocument:
    """Parse a DOCX file and return structured content blocks."""
    doc = Document(file_path)
    images = _extract_images(doc)
    blocks: list[ContentBlock] = []

    for element in doc.element.body:
        tag = element.tag.split("}")[-1] if "}" in element.tag else element.tag

        if tag == "p":
            from docx.text.paragraph import Paragraph
            paragraph = Paragraph(element, doc)
            text = paragraph.text.strip()

            img_rid = _paragraph_has_image(paragraph)
            if img_rid and img_rid in images:
                blocks.append(
                    ContentBlock(
                        block_type=BlockType.IMAGE,
                        image_data=images[img_rid],
                        caption=text if text else "",
                    )
                )
                continue

            heading_type = _classify_heading_level(paragraph)
            if heading_type:
                blocks.append(
                    ContentBlock(
                        block_type=heading_type,
                        text=text,
                        runs=_extract_runs(paragraph),
                        alignment=_get_alignment(paragraph),
                    )
                )
                continue

            is_list, list_style = _is_list_paragraph(paragraph)
            if is_list:
                blocks.append(
                    ContentBlock(
                        block_type=BlockType.LIST_ITEM,
                        text=text,
                        runs=_extract_runs(paragraph),
                        list_style=list_style,
                        alignment=_get_alignment(paragraph),
                    )
                )
                continue

            if text:
                lower = text.lower().strip()
                is_caption = lower.startswith(("figura", "figure", "tabela", "table", "gráfico", "grafico", "chart", "quadro", "fonte:", "source:"))
                blocks.append(
                    ContentBlock(
                        block_type=BlockType.CAPTION if is_caption else BlockType.PARAGRAPH,
                        text=text,
                        runs=_extract_runs(paragraph),
                        alignment=_get_alignment(paragraph),
                    )
                )

            if element.findall(f".//{qn('w:br')}[@{qn('w:type')}='page']"):
                blocks.append(ContentBlock(block_type=BlockType.PAGE_BREAK))

        elif tag == "tbl":
            table = Table(element, doc)
            table_data = _extract_table(table)
            if table_data:
                blocks.append(
                    ContentBlock(
                        block_type=BlockType.TABLE,
                        table_data=table_data,
                    )
                )

    metadata = _guess_metadata(blocks)
    return ParsedDocument(metadata=metadata, blocks=blocks, images=images)
