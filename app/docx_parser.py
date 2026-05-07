from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph


ROLE_PATTERN = re.compile(
    r"^(Realizacao|Realização|Execucao|Execução|Coordenacao|Coordenação|Autores?|Projeto Grafico|Projeto Gráfico|Revisao|Revisão|Equipe Tecnica|Equipe Técnica|Colaboracao|Colaboração|Apoio|Supervisao|Supervisão)\s*:?\s*(.+)?$",
    re.IGNORECASE,
)
VERSION_PATTERN = re.compile(r"\b(?:vers[aã]o|v\.?)\s*[:\-]?\s*([A-Za-z0-9\.\-_/]+)", re.IGNORECASE)
DATE_PATTERN = re.compile(
    r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|[A-Za-zçÇ]+(?:\s+de)?\s+\d{4}|\d{4})\b",
    re.IGNORECASE,
)
EMAIL_PATTERN = re.compile(r"[\w\.-]+@[\w\.-]+\.\w+")
WEBSITE_PATTERN = re.compile(r"(?:https?://|www\.)\S+", re.IGNORECASE)
CAPTION_PATTERN = re.compile(r"^(Figura|Imagem|Gr[aá]fico|Quadro|Tabela)\s+\w+", re.IGNORECASE)
SOURCE_PATTERN = re.compile(r"^(Fonte|Source)\s*[:\-]", re.IGNORECASE)


@dataclass
class Block:
    kind: str


@dataclass
class ParagraphBlock(Block):
    text: str
    style_name: str = ""
    level: int | None = None


@dataclass
class TableBlock(Block):
    rows: list[list[str]]


@dataclass
class ImageBlock(Block):
    path: Path
    alt_text: str = ""


@dataclass
class ParsedDocument:
    source_name: str
    title: str
    subtitle: str = ""
    authors: list[str] = field(default_factory=list)
    reviewers: list[str] = field(default_factory=list)
    version: str = ""
    date: str = ""
    credits: dict[str, list[str]] = field(default_factory=dict)
    catalog_lines: list[str] = field(default_factory=list)
    contacts: list[str] = field(default_factory=list)
    logos: list[Path] = field(default_factory=list)
    blocks: list[Block] = field(default_factory=list)


def iter_block_items(parent: DocxDocument | _Cell) -> Iterator[Paragraph | Table]:
    parent_element = parent.element.body if isinstance(parent, DocxDocument) else parent._tc
    for child in parent_element.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def parse_docx(docx_path: Path, asset_dir: Path) -> ParsedDocument:
    asset_dir.mkdir(parents=True, exist_ok=True)
    document = Document(str(docx_path))
    blocks: list[Block] = []
    leading_images: list[Path] = []
    text_entries: list[tuple[int, str, str]] = []
    front_matter_open = True

    for index, block in enumerate(iter_block_items(document)):
        if isinstance(block, Paragraph):
            style_name = block.style.name if block.style is not None else ""
            text = normalize_text(block.text)
            images = extract_images_from_paragraph(block, asset_dir, f"paragraph-{index}")

            if images and front_matter_open and not text:
                leading_images.extend(images[:2])

            if text:
                text_entries.append((len(blocks), text, style_name))
                blocks.append(
                    ParagraphBlock(
                        kind=classify_paragraph_kind(text, style_name),
                        text=text,
                        style_name=style_name,
                        level=extract_heading_level(style_name),
                    )
                )
                front_matter_open = False

            for image_path in images:
                blocks.append(ImageBlock(kind="image", path=image_path, alt_text=text))
                front_matter_open = False

        else:
            rows = []
            for row in block.rows:
                rows.append([normalize_text(cell.text) for cell in row.cells])
            if any(any(cell for cell in row) for row in rows):
                blocks.append(TableBlock(kind="table", rows=rows))
                front_matter_open = False

    metadata = infer_metadata(docx_path, document, text_entries)
    metadata.blocks = blocks
    metadata.logos = leading_images[:2]
    return metadata


def infer_metadata(
    docx_path: Path,
    document: DocxDocument,
    text_entries: list[tuple[int, str, str]],
) -> ParsedDocument:
    non_empty = [entry for entry in text_entries if entry[1]]
    core = document.core_properties
    title = core.title.strip() if core.title else ""
    subtitle = ""

    if not title and non_empty:
        title = choose_title(non_empty)

    if non_empty:
        title_index = next((idx for idx, (_, text, _) in enumerate(non_empty) if text == title), 0)
        for _, candidate, _ in non_empty[title_index + 1 : title_index + 4]:
            if candidate and candidate != title and len(candidate) <= 180 and not looks_like_role_line(candidate):
                subtitle = candidate
                break

    credits: dict[str, list[str]] = {}
    catalog_lines: list[str] = []
    contacts: list[str] = []
    authors: list[str] = []
    reviewers: list[str] = []
    version = ""
    date = ""

    for _, text, _ in non_empty[:80]:
        role_match = ROLE_PATTERN.match(text)
        if role_match:
            role = role_match.group(1).strip().title()
            people = split_people(role_match.group(2) or "")
            if people:
                credits.setdefault(role, []).extend(people)
                if role.startswith("Autor"):
                    authors.extend(people)
                if role.startswith("Revis"):
                    reviewers.extend(people)
            continue

        if not version:
            version_match = VERSION_PATTERN.search(text)
            if version_match:
                version = version_match.group(1).strip()

        if not date:
            date_match = DATE_PATTERN.search(text)
            if date_match:
                date = date_match.group(0).strip()

        if EMAIL_PATTERN.search(text) or WEBSITE_PATTERN.search(text) or "telefone" in text.lower():
            contacts.append(text)
            catalog_lines.append(text)
        elif any(token in text.lower() for token in ("copyright", "direitos reservados", "endereco", "endereço", "cep")):
            catalog_lines.append(text)

    if not authors and core.author:
        authors.append(core.author.strip())

    return ParsedDocument(
        source_name=docx_path.stem,
        title=title or docx_path.stem.replace("_", " "),
        subtitle=subtitle,
        authors=dedupe(authors),
        reviewers=dedupe(reviewers),
        version=version,
        date=date,
        credits={key: dedupe(value) for key, value in credits.items()},
        catalog_lines=dedupe(catalog_lines),
        contacts=dedupe(contacts),
        blocks=[],
    )


def extract_images_from_paragraph(paragraph: Paragraph, asset_dir: Path, prefix: str) -> list[Path]:
    image_paths: list[Path] = []
    blips = paragraph._element.xpath(".//a:blip")
    for image_index, blip in enumerate(blips):
        rel_id = blip.get(qn("r:embed"))
        if not rel_id:
            continue
        related_part = paragraph.part.related_parts.get(rel_id)
        if related_part is None:
            continue
        extension = Path(related_part.filename).suffix or ".png"
        output_path = asset_dir / f"{prefix}-{image_index}{extension.lower()}"
        output_path.write_bytes(related_part.blob)
        image_paths.append(output_path)
    return image_paths


def choose_title(non_empty: list[tuple[int, str, str]]) -> str:
    for _, text, style_name in non_empty[:6]:
        if extract_heading_level(style_name) == 1:
            return text
    return non_empty[0][1]


def classify_paragraph_kind(text: str, style_name: str) -> str:
    if extract_heading_level(style_name) is not None:
        return "heading"
    if SOURCE_PATTERN.match(text):
        return "source"
    if CAPTION_PATTERN.match(text):
        return "caption"
    if looks_like_callout(text):
        return "callout"
    if looks_like_list(style_name, text):
        return "list_item"
    return "paragraph"


def extract_heading_level(style_name: str) -> int | None:
    match = re.search(r"(?:heading|titulo|t[íi]tulo)\s*([1-6])", style_name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    return None


def looks_like_callout(text: str) -> bool:
    lowered = text.lower()
    return any(
        lowered.startswith(prefix)
        for prefix in (
            "insight:",
            "insight ",
            "recomendacao:",
            "recomendação:",
            "nota:",
            "observacao:",
            "observação:",
            "atencao:",
            "atenção:",
        )
    )


def looks_like_list(style_name: str, text: str) -> bool:
    lowered = style_name.lower()
    return "list" in lowered or "lista" in lowered or text.startswith("- ") or text.startswith("• ")


def looks_like_role_line(text: str) -> bool:
    return ROLE_PATTERN.match(text) is not None


def split_people(value: str) -> list[str]:
    value = normalize_text(value)
    if not value:
        return []
    parts = re.split(r";|,|\n| e ", value)
    return [part.strip() for part in parts if part.strip()]


def normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", value).strip()


def dedupe(values: list[str]) -> list[str]:
    seen: set[str] = set()
    ordered: list[str] = []
    for value in values:
        if value and value not in seen:
            seen.add(value)
            ordered.append(value)
    return ordered
