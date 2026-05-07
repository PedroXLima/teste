from __future__ import annotations

from dataclasses import dataclass, field
from io import BytesIO
from typing import Iterator

from docx import Document
from docx.document import Document as DocumentClass
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from PIL import Image as PILImage


@dataclass
class ContentBlock:
    kind: str
    text: str = ""
    level: int = 0
    style_name: str = ""
    rows: list[list[str]] = field(default_factory=list)
    image_bytes: bytes | None = None
    image_width: int | None = None
    image_height: int | None = None


@dataclass
class DocModel:
    title: str
    subtitle: str
    metadata: dict[str, str]
    credits: dict[str, list[str]]
    blocks: list[ContentBlock]


ROLE_KEYS = (
    "realização",
    "realizacao",
    "execução",
    "execucao",
    "coordenação",
    "coordenacao",
    "autores",
    "autor",
    "revisão",
    "revisao",
    "projeto gráfico",
    "projeto grafico",
    "equipe técnica",
    "equipe tecnica",
)


def _iter_blocks(parent: DocumentClass) -> Iterator[DocxParagraph | DocxTable]:
    body = parent.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield DocxParagraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield DocxTable(child, parent)


def _extract_image_bytes(paragraph: DocxParagraph, document: DocumentClass) -> bytes | None:
    for run in paragraph.runs:
        blips = run._element.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
        for blip in blips:
            rel_id = blip.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
            if rel_id and rel_id in document.part.related_parts:
                return document.part.related_parts[rel_id].blob
    return None


def _infer_heading_level(style_name: str, text: str) -> int:
    lower_style = style_name.lower()
    if "heading 1" in lower_style or "título 1" in lower_style or "titulo 1" in lower_style:
        return 1
    if "heading 2" in lower_style or "título 2" in lower_style or "titulo 2" in lower_style:
        return 2
    if "heading 3" in lower_style or "título 3" in lower_style or "titulo 3" in lower_style:
        return 3

    # Fallback heuristic for uppercase short lines.
    if text and len(text) < 80 and text == text.upper():
        return 1
    return 0


def _parse_key_value(text: str) -> tuple[str, str] | None:
    if ":" not in text:
        return None
    key, value = text.split(":", 1)
    key = key.strip()
    value = value.strip()
    if not key or not value:
        return None
    return key, value


def _normalize_key(key: str) -> str:
    return " ".join(key.strip().lower().split())


def parse_docx(path: str) -> DocModel:
    doc = Document(path)

    non_empty_paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    core_title = (doc.core_properties.title or "").strip()

    title = core_title or (non_empty_paragraphs[0] if non_empty_paragraphs else "Relatório Técnico")
    subtitle = non_empty_paragraphs[1] if len(non_empty_paragraphs) > 1 else ""

    metadata: dict[str, str] = {}
    credits: dict[str, list[str]] = {}
    blocks: list[ContentBlock] = []

    for block in _iter_blocks(doc):
        if isinstance(block, DocxParagraph):
            text = block.text.strip()
            style_name = block.style.name if block.style else "Normal"
            heading_level = _infer_heading_level(style_name, text)

            if text:
                parsed = _parse_key_value(text)
                if parsed:
                    key, value = parsed
                    normalized = _normalize_key(key)
                    metadata.setdefault(normalized, value)
                    if normalized in ROLE_KEYS:
                        credits.setdefault(key.strip(), []).append(value)

                blocks.append(
                    ContentBlock(
                        kind="heading" if heading_level else "paragraph",
                        text=text,
                        level=heading_level,
                        style_name=style_name,
                    )
                )
            else:
                image_blob = _extract_image_bytes(block, doc)
                if image_blob:
                    width = None
                    height = None
                    try:
                        image = PILImage.open(BytesIO(image_blob))
                        width, height = image.size
                    except Exception:
                        pass
                    blocks.append(
                        ContentBlock(
                            kind="image",
                            image_bytes=image_blob,
                            image_width=width,
                            image_height=height,
                        )
                    )
        else:
            rows: list[list[str]] = []
            for row in block.rows:
                rows.append([cell.text.strip() for cell in row.cells])
            if rows:
                blocks.append(ContentBlock(kind="table", rows=rows))

    return DocModel(
        title=title,
        subtitle=subtitle,
        metadata=metadata,
        credits=credits,
        blocks=blocks,
    )
