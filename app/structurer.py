"""Group parsed DOCX blocks into the institutional report sections.

The renderer expects a structured representation with named logical
sections (cover, credits, copyright, presentation, toc, body sections,
references, annexes, back cover). This module performs that
classification using heading-name heuristics tuned for Brazilian
institutional reports.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .docx_parser import Block, ImageBlock, ParsedDocument, TableBlock


def _norm(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(c for c in value if not unicodedata.combining(c))
    return value.lower().strip()


# Section label rules ---------------------------------------------------------

SECTION_RULES = [
    (
        "credits",
        re.compile(
            r"^(ficha\s+t[eé]cnica|cr[eé]ditos|equipe|expediente|realiza[cç][aã]o|execu[cç][aã]o|coordena[cç][aã]o|autores?|revis[aã]o|projeto\s+gr[aá]fico)\b"
        ),
    ),
    (
        "copyright",
        re.compile(
            r"^(cataloga[cç][aã]o|cdu|cdd|copyright|direitos\s+reservados|ficha\s+catalogr[aá]fica|isbn)\b"
        ),
    ),
    (
        "presentation",
        re.compile(
            r"^(apresenta[cç][aã]o|pref[aá]cio|introdu[cç][aã]o\s+institucional|carta\s+do\s+presidente|mensagem)\b"
        ),
    ),
    (
        "toc",
        re.compile(r"^(sum[aá]rio|[ií]ndice|conte[uú]do|table\s+of\s+contents)\b"),
    ),
    (
        "list_figures",
        re.compile(r"^(lista\s+de\s+figuras|figuras|list\s+of\s+figures)\b"),
    ),
    (
        "list_tables",
        re.compile(r"^(lista\s+de\s+tabelas|tabelas|list\s+of\s+tables)\b"),
    ),
    (
        "abbreviations",
        re.compile(
            r"^(lista\s+de\s+(abrevia[cç][oõ]es|siglas)|abrevia[cç][oõ]es|siglas|gloss[aá]rio)\b"
        ),
    ),
    (
        "references",
        re.compile(r"^(refer[eê]ncias|bibliografia|references|bibliography)\b"),
    ),
    (
        "annexes",
        re.compile(r"^(anexos?|ap[eê]ndices?|annex(es)?|appendix)\b"),
    ),
    (
        "back_cover",
        re.compile(r"^(contra\s*capa|back\s*cover|contato|contatos|fale\s+conosco)\b"),
    ),
]


@dataclass
class Section:
    kind: str  # cover | credits | copyright | presentation | toc | list_figures | list_tables | abbreviations | section | references | annexes | back_cover
    title: str = ""
    blocks: List[Block] = field(default_factory=list)
    extra: Dict = field(default_factory=dict)


@dataclass
class StructuredDocument:
    metadata: Dict
    sections: List[Section]


def _classify_heading(text: str) -> Optional[str]:
    n = _norm(text)
    for kind, pattern in SECTION_RULES:
        if pattern.match(n):
            return kind
    return None


def _block_is_empty(block: Block) -> bool:
    if block.kind != "paragraph":
        return False
    return not (block.text or "").strip()


def structure(parsed: ParsedDocument) -> StructuredDocument:
    blocks = parsed.blocks
    sections: List[Section] = []

    current = Section(kind="section", title="")
    current_logical_kind = "section"

    def flush():
        nonlocal current
        # drop pure-empty sections except the cover seed
        non_empty = any(not _block_is_empty(b) for b in current.blocks)
        if current.blocks and (non_empty or current.kind == "cover"):
            sections.append(current)

    for block in blocks:
        if block.kind == "page_break":
            # Page breaks signal a new logical section only if the next
            # heading reclassifies it. Keep it as a marker so the renderer
            # can preserve it inside body sections.
            current.blocks.append(block)
            continue

        if block.kind == "heading":
            level = block.level or 1
            classified = _classify_heading(block.text)
            if classified or level == 1:
                # New top-level section
                flush()
                kind = classified or "section"
                current = Section(kind=kind, title=block.text.strip())
                current_logical_kind = kind
                # For named meta-sections, do not include the heading itself
                # as a body block since the renderer prints its own title.
                if kind in {
                    "credits",
                    "copyright",
                    "presentation",
                    "toc",
                    "list_figures",
                    "list_tables",
                    "abbreviations",
                    "references",
                    "annexes",
                    "back_cover",
                }:
                    continue
                current.blocks.append(block)
                continue

        current.blocks.append(block)

    flush()

    # Insert synthetic cover section at the very front using metadata
    cover = Section(kind="cover", title=parsed.metadata.get("title", ""))
    cover.extra = {"metadata": parsed.metadata}
    sections.insert(0, cover)

    # Ensure a back cover at the end (synthetic if missing)
    if not any(s.kind == "back_cover" for s in sections):
        sections.append(Section(kind="back_cover", title="", extra={"metadata": parsed.metadata}))

    return StructuredDocument(metadata=parsed.metadata, sections=sections)
