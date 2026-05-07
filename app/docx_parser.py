"""
DOCX Parser – extracts structured content from a .docx file for PDF generation.

Produces a list of "content blocks" that the PDF renderer consumes.
Each block is a dict with at least a "type" key.
"""

import os
import re
import io
import hashlib
from typing import List, Dict, Any, Optional

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_ALIGN_PARAGRAPH


HEADING_STYLES = {
    "heading 1", "heading 2", "heading 3", "heading 4",
    "título 1", "título 2", "título 3", "título 4",
    "title", "titulo",
}

COVER_KEYWORDS = [
    "observatório", "relatório", "estudo", "boletim", "nota técnica",
    "análise", "panorama", "diagnóstico", "pesquisa", "caderno",
]

SECTION_DIVIDER_KEYWORDS = [
    "introdução", "metodologia", "desenvolvimento", "resultados",
    "conclusão", "conclusões", "considerações finais",
    "referências", "referências bibliográficas", "anexos", "anexo",
    "apêndice", "apêndices",
]

CREDITS_KEYWORDS = [
    "realização", "execução", "coordenação", "autores", "autor",
    "revisão", "projeto gráfico", "equipe técnica", "equipe",
    "elaboração", "colaboração", "supervisão",
]

ABBREV_KEYWORDS = [
    "lista de abreviaturas", "abreviaturas", "siglas",
    "lista de siglas", "acrônimos",
]


def _paragraph_text(para: Paragraph) -> str:
    return para.text.strip()


def _is_bold(para: Paragraph) -> bool:
    for run in para.runs:
        if run.bold:
            return True
    return False


def _detect_style(para: Paragraph) -> str:
    style_name = (para.style.name or "").lower()
    if style_name in HEADING_STYLES or style_name.startswith("heading"):
        return style_name
    return "body"


def _heading_level(style_name: str) -> int:
    m = re.search(r"(\d)", style_name)
    if m:
        return int(m.group(1))
    if "title" in style_name or "titulo" in style_name or "título" in style_name:
        return 1
    return 1


def _extract_images(doc: Document, output_dir: str) -> Dict[str, str]:
    """Extract all images from the DOCX and return rId -> filepath mapping."""
    images = {}
    os.makedirs(output_dir, exist_ok=True)
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            img_data = rel.target_part.blob
            ext = os.path.splitext(rel.target_part.partname)[-1] or ".png"
            fname = hashlib.md5(img_data).hexdigest() + ext
            fpath = os.path.join(output_dir, fname)
            with open(fpath, "wb") as f:
                f.write(img_data)
            images[rel.rId] = fpath
    return images


def _get_inline_images(para: Paragraph, images: Dict[str, str]) -> List[str]:
    """Return list of image paths embedded in a paragraph."""
    found = []
    for run in para.runs:
        xml = run._element.xml
        blip_matches = re.findall(r'r:embed="([^"]+)"', xml)
        for rid in blip_matches:
            if rid in images:
                found.append(images[rid])
    return found


def _parse_table(table: Table) -> Dict[str, Any]:
    """Parse a DOCX table into a structured dict."""
    rows = []
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        rows.append(cells)
    return {
        "type": "table",
        "rows": rows,
        "num_cols": len(rows[0]) if rows else 0,
        "num_rows": len(rows),
    }


def _detect_block_role(text: str, level: int) -> Optional[str]:
    """Detect if a heading signals a special section."""
    lower = text.lower().strip()
    if lower in ("sumário", "índice", "table of contents", "conteúdo"):
        return "toc"
    if lower in ("lista de figuras",):
        return "list_of_figures"
    if lower in ("lista de tabelas",):
        return "list_of_tables"
    for kw in ABBREV_KEYWORDS:
        if kw in lower:
            return "abbreviations"
    if lower in ("apresentação", "prefácio", "foreword", "presentation"):
        return "presentation"
    for kw in SECTION_DIVIDER_KEYWORDS:
        if lower == kw or lower.startswith(kw):
            if level <= 1:
                return "section_divider"
    if lower in ("referências", "referências bibliográficas", "references"):
        return "references"
    if lower.startswith("anexo"):
        return "annexes"
    return None


def parse_docx(filepath: str, images_dir: str) -> Dict[str, Any]:
    """
    Parse a DOCX file and return structured content.

    Returns a dict with:
      - "metadata": extracted cover/metadata info
      - "blocks": list of content blocks
    """
    doc = Document(filepath)
    images = _extract_images(doc, images_dir)
    blocks: List[Dict[str, Any]] = []

    metadata: Dict[str, Any] = {
        "title": "",
        "subtitle": "",
        "authors": [],
        "reviewers": [],
        "version": "",
        "date": "",
        "institution": "",
        "description": "",
    }

    core = doc.core_properties
    if core.title:
        metadata["title"] = core.title
    if core.author:
        metadata["authors"] = [a.strip() for a in core.author.split(",")]
    if core.subject:
        metadata["subtitle"] = core.subject
    if core.modified:
        metadata["date"] = str(core.modified.strftime("%d/%m/%Y")) if core.modified else ""

    cover_detected = False
    credits_section = False

    body_elements = doc.element.body
    element_index = 0

    for child in body_elements:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            para = Paragraph(child, doc)
            text = _paragraph_text(para)
            style = _detect_style(para)
            inline_imgs = _get_inline_images(para, images)

            if not text and not inline_imgs:
                element_index += 1
                continue

            if inline_imgs:
                for img_path in inline_imgs:
                    blocks.append({
                        "type": "image",
                        "path": img_path,
                        "caption": "",
                    })

            if not text:
                element_index += 1
                continue

            if style != "body":
                level = _heading_level(style)

                if not cover_detected and level == 1:
                    if not metadata["title"]:
                        metadata["title"] = text
                    cover_detected = True

                role = _detect_block_role(text, level)

                if role == "section_divider":
                    blocks.append({
                        "type": "section_divider",
                        "title": text,
                        "level": level,
                    })
                    credits_section = False
                elif role:
                    blocks.append({
                        "type": role,
                        "title": text,
                        "level": level,
                    })
                    credits_section = False
                else:
                    blocks.append({
                        "type": "heading",
                        "text": text,
                        "level": level,
                        "bold": _is_bold(para),
                    })

                    is_credit_heading = False
                    lower = text.lower()
                    for kw in CREDITS_KEYWORDS:
                        if kw in lower:
                            is_credit_heading = True
                            credits_section = True
                            blocks[-1]["type"] = "credits_heading"
                            break

                    if not is_credit_heading and level <= 1:
                        credits_section = False
            else:
                is_caption = False
                lower = text.lower()

                if re.match(r"^(figura|figure|gráfico|quadro)\s+\d", lower):
                    if blocks and blocks[-1]["type"] == "image":
                        blocks[-1]["caption"] = text
                        element_index += 1
                        continue
                    is_caption = True

                if re.match(r"^(tabela|table)\s+\d", lower):
                    is_caption = True

                if re.match(r"^fonte:", lower) or re.match(r"^source:", lower):
                    if blocks and blocks[-1]["type"] in ("image", "table"):
                        blocks[-1].setdefault("source", "")
                        blocks[-1]["source"] = text
                        element_index += 1
                        continue

                if is_caption:
                    blocks.append({
                        "type": "caption",
                        "text": text,
                    })
                elif credits_section:
                    blocks.append({
                        "type": "credits_body",
                        "text": text,
                        "bold": _is_bold(para),
                    })
                else:
                    alignment = None
                    if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                        alignment = "center"
                    elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                        alignment = "right"
                    elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                        alignment = "justify"

                    blocks.append({
                        "type": "paragraph",
                        "text": text,
                        "bold": _is_bold(para),
                        "alignment": alignment,
                    })

        elif tag == "tbl":
            table = Table(child, doc)
            tbl_data = _parse_table(table)
            blocks.append(tbl_data)

        element_index += 1

    if not metadata["title"] and blocks:
        for b in blocks:
            if b["type"] == "heading" and b.get("level", 99) <= 1:
                metadata["title"] = b["text"]
                break
        if not metadata["title"]:
            for b in blocks:
                if b["type"] in ("paragraph", "heading"):
                    metadata["title"] = b.get("text", "")[:100]
                    break

    if not metadata["subtitle"]:
        for i, b in enumerate(blocks):
            if b["type"] == "heading" and b.get("level", 99) == 2:
                metadata["subtitle"] = b["text"]
                break
            if b["type"] == "paragraph" and i < 5:
                metadata["subtitle"] = b["text"][:200]
                break

    return {
        "metadata": metadata,
        "blocks": blocks,
    }
