from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Literal
from xml.sax.saxutils import escape

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table as DocxTable
from docx.table import _Cell
from docx.text.paragraph import Paragraph as DocxParagraph
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, StyleSheet1
from reportlab.lib.units import cm, mm
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    Image,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)


PRIMARY_BLUE = colors.HexColor("#005A9C")
DARK_NAVY = colors.HexColor("#003B6F")
MEDIUM_BLUE = colors.HexColor("#0B63B6")
LIGHT_BLUE = colors.HexColor("#D9EAF7")
VERY_LIGHT_BLUE = colors.HexColor("#EEF6FC")
WHITE = colors.HexColor("#FFFFFF")
MAIN_TEXT = colors.HexColor("#2B2B2B")
SECONDARY_TEXT = colors.HexColor("#6B7280")
TABLE_BORDER = colors.HexColor("#D1D5DB")

PAGE_WIDTH, PAGE_HEIGHT = A4
BODY_LEFT = 21 * mm
BODY_RIGHT = 21 * mm
BODY_TOP = 23 * mm
BODY_BOTTOM = 21 * mm


@dataclass
class DocumentMetadata:
    title: str
    subtitle: str = ""
    author: str = ""
    subject: str = ""
    date: str = ""
    filename: str = ""
    keywords: list[str] = field(default_factory=list)


@dataclass
class ParagraphBlock:
    kind: Literal["paragraph"] = "paragraph"
    text: str = ""
    style_name: str = ""
    level: int | None = None
    is_list: bool = False
    is_numbered: bool = False


@dataclass
class TableBlock:
    kind: Literal["table"] = "table"
    rows: list[list[str]] = field(default_factory=list)


@dataclass
class ImageBlock:
    kind: Literal["image"] = "image"
    data: bytes = b""
    content_type: str = "image/png"
    alt_text: str = ""


DocumentBlock = ParagraphBlock | TableBlock | ImageBlock


def convert_docx_to_pdf(input_path: Path | str, output_path: Path | str) -> None:
    """Convert a DOCX file into a designed institutional PDF."""
    input_path = Path(input_path)
    output_path = Path(output_path)
    document = Document(input_path)
    blocks = parse_docx(document)
    metadata = extract_metadata(document, blocks, input_path.name)
    build_pdf(blocks, metadata, output_path)


def parse_docx(document: DocumentObject) -> list[DocumentBlock]:
    blocks: list[DocumentBlock] = []
    seen_image_hashes: set[int] = set()

    for item in iter_block_items(document):
        if isinstance(item, DocxParagraph):
            text = normalize_text(item.text)
            images = extract_images_from_paragraph(document, item)

            if text:
                style_name = item.style.name if item.style is not None else ""
                blocks.append(
                    ParagraphBlock(
                        text=text,
                        style_name=style_name,
                        level=heading_level(style_name, text),
                        is_list=is_list_style(style_name),
                        is_numbered=is_numbered_style(style_name),
                    )
                )

            for image in images:
                marker = hash(image.data)
                if marker not in seen_image_hashes:
                    seen_image_hashes.add(marker)
                    blocks.append(image)

        elif isinstance(item, DocxTable):
            rows = [[normalize_text(cell.text) for cell in row.cells] for row in item.rows]
            if any(any(cell for cell in row) for row in rows):
                blocks.append(TableBlock(rows=rows))

    return blocks


def iter_block_items(parent: DocumentObject | _Cell) -> Iterable[DocxParagraph | DocxTable]:
    if isinstance(parent, DocumentObject):
        parent_elm = parent.element.body
        parent_obj = parent
    else:
        parent_elm = parent._tc
        parent_obj = parent

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield DocxParagraph(child, parent_obj)
        elif isinstance(child, CT_Tbl):
            yield DocxTable(child, parent_obj)


def extract_images_from_paragraph(document: DocumentObject, paragraph: DocxParagraph) -> list[ImageBlock]:
    images: list[ImageBlock] = []
    for run in paragraph.runs:
        for blip in run._element.xpath(".//a:blip"):
            relationship_id = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
            if not relationship_id:
                continue
            image_part = document.part.related_parts.get(relationship_id)
            if image_part is None:
                continue
            images.append(
                ImageBlock(
                    data=image_part.blob,
                    content_type=getattr(image_part, "content_type", "image/png"),
                    alt_text="",
                )
            )
    return images


def normalize_text(text: str) -> str:
    return re.sub(r"[ \t]+\n", "\n", text.replace("\xa0", " ")).strip()


def heading_level(style_name: str, text: str) -> int | None:
    style = style_name.lower()
    match = re.search(r"(?:heading|t[íi]tulo)\s*([1-6])", style)
    if match:
        return int(match.group(1))
    if style in {"title", "título"}:
        return 0
    if style in {"subtitle", "subtítulo"}:
        return 7

    # Many DOCX files exported from editors lose semantic styles. A conservative
    # all-caps detector helps preserve major hierarchy without rewriting content.
    clean = re.sub(r"[^A-Za-zÀ-ÿ0-9 ]", "", text)
    if 4 <= len(clean) <= 90 and clean.upper() == clean and re.search(r"[A-ZÀ-Ý]", clean):
        return 1
    return None


def is_list_style(style_name: str) -> bool:
    return "list" in style_name.lower() or "lista" in style_name.lower()


def is_numbered_style(style_name: str) -> bool:
    style = style_name.lower()
    return "number" in style or "número" in style or "numerada" in style


def extract_metadata(document: DocumentObject, blocks: list[DocumentBlock], filename: str) -> DocumentMetadata:
    props = document.core_properties
    paragraphs = [block for block in blocks if isinstance(block, ParagraphBlock)]
    first_heading = next((block.text for block in paragraphs if block.level in {0, 1}), "")
    first_text = next((block.text for block in paragraphs if block.text), "")

    title = normalize_text(props.title or first_heading or first_text or Path(filename).stem)
    if len(title) > 150 and first_heading:
        title = first_heading
    if len(title) > 180:
        title = f"{title[:177].rstrip()}..."

    subtitle = normalize_text(props.subject or "")
    if not subtitle:
        subtitle = infer_subtitle(paragraphs, title)

    date = ""
    if props.modified:
        date = props.modified.strftime("%d/%m/%Y")
    elif props.created:
        date = props.created.strftime("%d/%m/%Y")

    explicit_date = infer_date(paragraphs[:25])
    if explicit_date:
        date = explicit_date

    return DocumentMetadata(
        title=title,
        subtitle=subtitle,
        author=normalize_text(props.author or ""),
        subject=normalize_text(props.subject or ""),
        date=date,
        filename=filename,
        keywords=infer_keywords(paragraphs[:40]),
    )


def infer_subtitle(paragraphs: list[ParagraphBlock], title: str) -> str:
    for paragraph in paragraphs[:12]:
        if paragraph.text == title:
            continue
        if paragraph.level in {0, 1}:
            continue
        if 12 <= len(paragraph.text) <= 180:
            return paragraph.text
    return ""


def infer_date(paragraphs: list[ParagraphBlock]) -> str:
    date_patterns = [
        r"\b\d{1,2}/\d{1,2}/\d{2,4}\b",
        r"\b\d{4}-\d{2}-\d{2}\b",
        r"\b(?:janeiro|fevereiro|mar[çc]o|abril|maio|junho|julho|agosto|setembro|outubro|novembro|dezembro)\s+de\s+\d{4}\b",
    ]
    for paragraph in paragraphs:
        for pattern in date_patterns:
            match = re.search(pattern, paragraph.text, flags=re.IGNORECASE)
            if match:
                return match.group(0)
    return ""


def infer_keywords(paragraphs: list[ParagraphBlock]) -> list[str]:
    keyword_pool = [
        "indústria",
        "economia",
        "inteligência",
        "competitividade",
        "inovação",
        "desenvolvimento",
        "mercado",
        "estratégia",
        "produtividade",
    ]
    joined = " ".join(paragraph.text.lower() for paragraph in paragraphs)
    return [keyword for keyword in keyword_pool if keyword in joined][:4]


def build_pdf(blocks: list[DocumentBlock], metadata: DocumentMetadata, output_path: Path) -> None:
    styles = build_styles()
    doc = InstitutionalDocTemplate(str(output_path), metadata=metadata)
    story: list[Flowable] = [CoverPage(metadata), NextPageTemplate("Body"), PageBreak()]

    if not blocks:
        story.append(Paragraph("Documento sem conteúdo textual extraível.", styles["Body"]))

    h1_count = 0
    for index, block in enumerate(blocks):
        if isinstance(block, ParagraphBlock):
            flowables = paragraph_to_flowables(block, styles, h1_count, index)
            if block.level == 1 and should_create_section_divider(block, index):
                h1_count += 1
            story.extend(flowables)
        elif isinstance(block, TableBlock):
            story.extend(table_to_flowables(block, styles, doc.body_width))
        elif isinstance(block, ImageBlock):
            story.extend(image_to_flowables(block, styles, doc.body_width))

    story.extend([NextPageTemplate("BackCover"), PageBreak(), BackCoverPage(metadata)])
    doc.build(story)


class InstitutionalDocTemplate(BaseDocTemplate):
    def __init__(self, filename: str, metadata: DocumentMetadata) -> None:
        self.metadata = metadata
        self.body_width = PAGE_WIDTH - BODY_LEFT - BODY_RIGHT
        self.body_height = PAGE_HEIGHT - BODY_TOP - BODY_BOTTOM
        super().__init__(
            filename,
            pagesize=A4,
            leftMargin=BODY_LEFT,
            rightMargin=BODY_RIGHT,
            topMargin=BODY_TOP,
            bottomMargin=BODY_BOTTOM,
            title=metadata.title,
            author=metadata.author,
        )

        cover_frame = Frame(0, 0, PAGE_WIDTH, PAGE_HEIGHT, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        body_frame = Frame(
            BODY_LEFT,
            BODY_BOTTOM,
            self.body_width,
            self.body_height,
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )
        self.addPageTemplates(
            [
                PageTemplate(id="Cover", frames=[cover_frame]),
                PageTemplate(id="Body", frames=[body_frame], onPage=self.draw_body_chrome),
                PageTemplate(id="BackCover", frames=[cover_frame]),
            ]
        )

    def draw_body_chrome(self, canvas, doc) -> None:  # noqa: ANN001 - ReportLab callback signature.
        canvas.saveState()
        canvas.setFillColor(WHITE)
        canvas.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, stroke=0, fill=1)

        header_y = PAGE_HEIGHT - 15 * mm
        canvas.setStrokeColor(PRIMARY_BLUE)
        canvas.setLineWidth(1.2)
        canvas.line(BODY_LEFT, header_y - 5 * mm, PAGE_WIDTH - BODY_RIGHT, header_y - 5 * mm)

        canvas.setFont("Helvetica-Bold", 8.5)
        canvas.setFillColor(DARK_NAVY)
        header_title = truncate(self.metadata.title, 78)
        canvas.drawString(BODY_LEFT, header_y, header_title)

        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(SECONDARY_TEXT)
        body_page = max(1, canvas.getPageNumber() - 1)
        canvas.drawRightString(PAGE_WIDTH - BODY_RIGHT, header_y, f"p. {body_page}")

        footer_y = 12 * mm
        canvas.setStrokeColor(TABLE_BORDER)
        canvas.setLineWidth(0.5)
        canvas.line(BODY_LEFT, footer_y + 5 * mm, PAGE_WIDTH - BODY_RIGHT, footer_y + 5 * mm)
        canvas.setFont("Helvetica", 7.5)
        canvas.setFillColor(SECONDARY_TEXT)
        canvas.drawString(BODY_LEFT, footer_y, "Relatório institucional gerado a partir do documento original")
        if self.metadata.date:
            canvas.drawRightString(PAGE_WIDTH - BODY_RIGHT, footer_y, self.metadata.date)
        canvas.restoreState()


def build_styles() -> StyleSheet1:
    styles = StyleSheet1()
    base_font = "Helvetica"
    bold_font = "Helvetica-Bold"

    styles.add(
        ParagraphStyle(
            name="Body",
            fontName=base_font,
            fontSize=10.2,
            leading=15.2,
            textColor=MAIN_TEXT,
            alignment=TA_JUSTIFY,
            spaceAfter=7,
        )
    )
    styles.add(
        ParagraphStyle(
            name="BodyLeft",
            parent=styles["Body"],
            alignment=TA_LEFT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Bullet",
            parent=styles["BodyLeft"],
            leftIndent=13,
            firstLineIndent=-8,
            bulletIndent=0,
            spaceAfter=5,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Heading1",
            fontName=bold_font,
            fontSize=20,
            leading=24,
            textColor=PRIMARY_BLUE,
            spaceBefore=18,
            spaceAfter=8,
            keepWithNext=True,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Heading2",
            fontName=bold_font,
            fontSize=15,
            leading=19,
            textColor=DARK_NAVY,
            spaceBefore=13,
            spaceAfter=7,
            keepWithNext=True,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Heading3",
            fontName=bold_font,
            fontSize=12.5,
            leading=16,
            textColor=MEDIUM_BLUE,
            spaceBefore=10,
            spaceAfter=5,
            keepWithNext=True,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Caption",
            fontName=base_font,
            fontSize=8.2,
            leading=11,
            textColor=SECONDARY_TEXT,
            alignment=TA_CENTER,
            spaceBefore=2,
            spaceAfter=9,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Source",
            parent=styles["Caption"],
            fontSize=7.8,
            italic=True,
        )
    )
    styles.add(
        ParagraphStyle(
            name="TableCell",
            fontName=base_font,
            fontSize=8.2,
            leading=10.5,
            textColor=MAIN_TEXT,
            alignment=TA_LEFT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="TableHeader",
            parent=styles["TableCell"],
            fontName=bold_font,
            textColor=WHITE,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CalloutTitle",
            fontName=bold_font,
            fontSize=10.2,
            leading=13,
            textColor=DARK_NAVY,
            spaceAfter=3,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CalloutBody",
            parent=styles["BodyLeft"],
            fontSize=9.4,
            leading=13,
            spaceAfter=0,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CoverTitle",
            fontName=bold_font,
            fontSize=28,
            leading=32,
            textColor=WHITE,
            alignment=TA_LEFT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="CoverSubtitle",
            fontName=base_font,
            fontSize=12,
            leading=16,
            textColor=WHITE,
            alignment=TA_LEFT,
        )
    )
    return styles


def paragraph_to_flowables(block: ParagraphBlock, styles: StyleSheet1, h1_count: int, index: int) -> list[Flowable]:
    text = clean_for_reportlab(block.text)
    if not text:
        return []

    if is_callout(block.text):
        return [build_callout(block.text, styles), Spacer(1, 5)]

    if is_caption(block.text):
        style = styles["Source"] if block.text.lower().startswith(("fonte:", "source:")) else styles["Caption"]
        return [Paragraph(text, style)]

    if block.level == 1:
        flowables: list[Flowable] = []
        if should_create_section_divider(block, index):
            flowables.extend([PageBreak(), SectionDivider(block.text), PageBreak()])
        flowables.extend([Paragraph(text, styles["Heading1"]), divider_line(PRIMARY_BLUE, width=1.0), Spacer(1, 5)])
        return flowables
    if block.level == 2:
        return [Paragraph(text, styles["Heading2"])]
    if block.level and block.level >= 3:
        return [Paragraph(text, styles["Heading3"])]
    if block.is_list:
        bullet = "•" if not block.is_numbered else None
        return [Paragraph(text, styles["Bullet"], bulletText=bullet)]
    return [Paragraph(text, styles["Body"])]


def should_create_section_divider(block: ParagraphBlock, index: int) -> bool:
    lowered = strip_accents(block.text.lower())
    front_matter_terms = (
        "sumario",
        "sumário",
        "lista de figuras",
        "lista de tabelas",
        "abreviaturas",
        "creditos",
        "créditos",
        "catalog",
        "copyright",
        "apresentacao",
        "apresentação",
    )
    if index < 4:
        return False
    if any(term in lowered for term in front_matter_terms):
        return False
    return 3 <= len(block.text) <= 110


def table_to_flowables(block: TableBlock, styles: StyleSheet1, available_width: float) -> list[Flowable]:
    col_count = max((len(row) for row in block.rows), default=1)
    normalized_rows = [row + [""] * (col_count - len(row)) for row in block.rows]

    table_data: list[list[Paragraph]] = []
    for row_index, row in enumerate(normalized_rows):
        row_style = styles["TableHeader"] if row_index == 0 else styles["TableCell"]
        table_data.append([Paragraph(clean_for_reportlab(cell), row_style) for cell in row])

    col_widths = calculate_col_widths(normalized_rows, available_width)
    table = Table(table_data, colWidths=col_widths, repeatRows=1, hAlign="CENTER", splitByRow=1)
    style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK_NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.35, TABLE_BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]

    for row_index in range(1, len(normalized_rows)):
        if row_index % 2 == 1:
            style_commands.append(("BACKGROUND", (0, row_index), (-1, row_index), VERY_LIGHT_BLUE))
        for col_index, value in enumerate(normalized_rows[row_index]):
            if looks_numeric(value):
                style_commands.append(("ALIGN", (col_index, row_index), (col_index, row_index), "RIGHT"))

    table.setStyle(TableStyle(style_commands))
    return [Spacer(1, 6), table, Spacer(1, 10)]


def calculate_col_widths(rows: list[list[str]], available_width: float) -> list[float]:
    col_count = max((len(row) for row in rows), default=1)
    weights = [1.0] * col_count
    for col_index in range(col_count):
        max_len = max((len(row[col_index]) for row in rows if col_index < len(row)), default=1)
        weights[col_index] = min(max(max_len / 16, 0.8), 3.2)

    total = sum(weights)
    return [available_width * weight / total for weight in weights]


def image_to_flowables(block: ImageBlock, styles: StyleSheet1, available_width: float) -> list[Flowable]:
    try:
        image_buffer = io.BytesIO(block.data)
        with PILImage.open(io.BytesIO(block.data)) as pil_image:
            original_width, original_height = pil_image.size
        max_width = available_width * 0.94
        max_height = PAGE_HEIGHT * 0.42
        ratio = min(max_width / original_width, max_height / original_height, 1.0)
        image = Image(image_buffer, width=original_width * ratio, height=original_height * ratio, hAlign="CENTER")
        return [Spacer(1, 8), image, Spacer(1, 8)]
    except Exception:  # noqa: BLE001 - skip malformed embedded media while preserving surrounding content.
        return [Paragraph("Figura incorporada não pôde ser renderizada no PDF.", styles["Caption"])]


def build_callout(text: str, styles: StyleSheet1) -> Table:
    title, body = split_callout(text)
    content = [
        [
            Paragraph(clean_for_reportlab(title), styles["CalloutTitle"]),
            Paragraph(clean_for_reportlab(body), styles["CalloutBody"]),
        ]
    ]
    table = Table(content, colWidths=[4.5 * cm, PAGE_WIDTH - BODY_LEFT - BODY_RIGHT - 4.5 * cm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), VERY_LIGHT_BLUE),
                ("BACKGROUND", (0, 0), (0, 0), LIGHT_BLUE),
                ("LINEBEFORE", (0, 0), (0, 0), 4, DARK_NAVY),
                ("BOX", (0, 0), (-1, -1), 0.4, LIGHT_BLUE),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 9),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )
    return table


class CoverPage(Flowable):
    def __init__(self, metadata: DocumentMetadata) -> None:
        super().__init__()
        self.metadata = metadata
        self.width = PAGE_WIDTH
        self.height = PAGE_HEIGHT

    def wrap(self, _available_width: float, _available_height: float) -> tuple[float, float]:
        return PAGE_WIDTH, PAGE_HEIGHT

    def draw(self) -> None:
        canvas = self.canv
        draw_blue_background(canvas)
        draw_network_pattern(canvas, intensity=0.38)

        canvas.saveState()
        canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.12))
        canvas.roundRect(20 * mm, PAGE_HEIGHT - 39 * mm, 64 * mm, 11 * mm, 5, stroke=0, fill=1)
        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 7.8)
        canvas.drawString(25 * mm, PAGE_HEIGHT - 35 * mm, "INTELIGÊNCIA ESTRATÉGICA")
        canvas.restoreState()

        panel_x = 0
        panel_y = 48 * mm
        panel_h = 96 * mm
        canvas.setFillColor(colors.Color(0, 59 / 255, 111 / 255, alpha=0.72))
        canvas.rect(panel_x, panel_y, PAGE_WIDTH, panel_h, stroke=0, fill=1)
        canvas.setFillColor(colors.Color(0, 90 / 255, 156 / 255, alpha=0.86))
        canvas.rect(0, panel_y + panel_h - 7 * mm, PAGE_WIDTH, 7 * mm, stroke=0, fill=1)

        styles = build_styles()
        title = clean_for_reportlab(self.metadata.title.upper())
        title_paragraph = Paragraph(title, styles["CoverTitle"])
        title_width = PAGE_WIDTH - 40 * mm
        title_height = title_paragraph.wrap(title_width, panel_h)[1]
        title_paragraph.drawOn(canvas, 20 * mm, panel_y + panel_h - 20 * mm - title_height)

        if self.metadata.subtitle:
            subtitle = Paragraph(clean_for_reportlab(self.metadata.subtitle), styles["CoverSubtitle"])
            subtitle_height = subtitle.wrap(title_width, 35 * mm)[1]
            subtitle.drawOn(canvas, 20 * mm, panel_y + 22 * mm - subtitle_height)

        meta_parts = [part for part in [self.metadata.author, self.metadata.date] if part]
        if meta_parts:
            canvas.setFont("Helvetica", 8.5)
            canvas.setFillColor(WHITE)
            canvas.drawString(20 * mm, panel_y + 12 * mm, " | ".join(meta_parts))

        if self.metadata.date:
            canvas.setFillColor(WHITE)
            canvas.roundRect(PAGE_WIDTH - 55 * mm, 19 * mm, 35 * mm, 11 * mm, 5, stroke=0, fill=1)
            canvas.setFillColor(DARK_NAVY)
            canvas.setFont("Helvetica-Bold", 7.2)
            canvas.drawCentredString(PAGE_WIDTH - 37.5 * mm, 23 * mm, self.metadata.date)


class SectionDivider(Flowable):
    def __init__(self, title: str) -> None:
        super().__init__()
        self.title = title
        self.width = PAGE_WIDTH
        self.height = PAGE_HEIGHT - BODY_TOP - BODY_BOTTOM

    def wrap(self, _available_width: float, _available_height: float) -> tuple[float, float]:
        return PAGE_WIDTH - BODY_LEFT - BODY_RIGHT, PAGE_HEIGHT - BODY_TOP - BODY_BOTTOM

    def draw(self) -> None:
        canvas = self.canv
        x = -BODY_LEFT
        y = -BODY_BOTTOM
        canvas.saveState()
        canvas.setFillColor(DARK_NAVY)
        canvas.rect(x, y, PAGE_WIDTH, PAGE_HEIGHT, stroke=0, fill=1)
        canvas.setFillColor(PRIMARY_BLUE)
        canvas.rect(x, y, PAGE_WIDTH, PAGE_HEIGHT * 0.42, stroke=0, fill=1)
        draw_network_pattern(canvas, origin_x=x, origin_y=y, intensity=0.24)
        canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.10))
        for offset in range(0, 8):
            canvas.circle(x + PAGE_WIDTH - (18 + offset * 13) * mm, y + (28 + offset * 10) * mm, 26 * mm, stroke=0, fill=1)

        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 27)
        lines = split_title_lines(self.title.upper(), max_chars=24)
        start_y = y + 76 * mm + (len(lines) - 1) * 15
        for line_index, line in enumerate(lines):
            canvas.drawString(x + 22 * mm, start_y - line_index * 32, line)
        canvas.setStrokeColor(WHITE)
        canvas.setLineWidth(1.6)
        canvas.line(x + 22 * mm, y + 55 * mm, x + 83 * mm, y + 55 * mm)
        canvas.setStrokeColor(LIGHT_BLUE)
        canvas.setLineWidth(5)
        canvas.line(x + 22 * mm, y + 49 * mm, x + 60 * mm, y + 49 * mm)
        canvas.restoreState()


class BackCoverPage(Flowable):
    def __init__(self, metadata: DocumentMetadata) -> None:
        super().__init__()
        self.metadata = metadata
        self.width = PAGE_WIDTH
        self.height = PAGE_HEIGHT

    def wrap(self, _available_width: float, _available_height: float) -> tuple[float, float]:
        return PAGE_WIDTH, PAGE_HEIGHT

    def draw(self) -> None:
        canvas = self.canv
        draw_blue_background(canvas, dark=True)
        draw_network_pattern(canvas, intensity=0.32)

        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 11)
        canvas.drawCentredString(PAGE_WIDTH / 2, PAGE_HEIGHT - 47 * mm, "RELATÓRIO INSTITUCIONAL")

        canvas.setStrokeColor(WHITE)
        canvas.setLineWidth(0.8)
        canvas.line(20 * mm, 34 * mm, PAGE_WIDTH - 20 * mm, 34 * mm)

        canvas.setFont("Helvetica-Bold", 12)
        canvas.drawString(20 * mm, 24 * mm, truncate(self.metadata.title, 54))
        if self.metadata.subtitle:
            canvas.setFont("Helvetica", 8.5)
            canvas.drawString(20 * mm, 18 * mm, truncate(self.metadata.subtitle, 72))

        canvas.setFont("Helvetica", 8.5)
        footer = self.metadata.author or self.metadata.date or "Documento gerado a partir do arquivo DOCX original"
        canvas.drawRightString(PAGE_WIDTH - 20 * mm, 24 * mm, truncate(footer, 48))


def draw_blue_background(canvas, dark: bool = False) -> None:  # noqa: ANN001 - ReportLab canvas type.
    canvas.saveState()
    canvas.setFillColor(DARK_NAVY if dark else PRIMARY_BLUE)
    canvas.rect(0, 0, PAGE_WIDTH, PAGE_HEIGHT, stroke=0, fill=1)
    canvas.setFillColor(MEDIUM_BLUE if dark else DARK_NAVY)
    canvas.rect(0, PAGE_HEIGHT * 0.56, PAGE_WIDTH, PAGE_HEIGHT * 0.44, stroke=0, fill=1)
    canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.07))
    canvas.circle(PAGE_WIDTH * 0.86, PAGE_HEIGHT * 0.86, 70 * mm, stroke=0, fill=1)
    canvas.circle(PAGE_WIDTH * 0.12, PAGE_HEIGHT * 0.18, 58 * mm, stroke=0, fill=1)
    canvas.setFillColor(colors.Color(0.85, 0.93, 0.98, alpha=0.10))
    canvas.rect(0, 0, PAGE_WIDTH, 42 * mm, stroke=0, fill=1)
    canvas.restoreState()


def draw_network_pattern(canvas, origin_x: float = 0, origin_y: float = 0, intensity: float = 0.26) -> None:  # noqa: ANN001
    canvas.saveState()
    canvas.setStrokeColor(colors.Color(1, 1, 1, alpha=intensity))
    canvas.setFillColor(colors.Color(1, 1, 1, alpha=min(intensity + 0.12, 0.5)))
    canvas.setLineWidth(0.35)
    nodes = [
        (35, 66),
        (58, 92),
        (92, 77),
        (126, 101),
        (160, 80),
        (181, 114),
        (44, 134),
        (82, 151),
        (122, 139),
        (168, 157),
    ]
    scaled = [(origin_x + x * mm, origin_y + y * mm) for x, y in nodes]
    for start, end in zip(scaled, scaled[1:]):
        canvas.line(start[0], start[1], end[0], end[1])
    for index, point in enumerate(scaled):
        if index + 2 < len(scaled):
            other = scaled[index + 2]
            canvas.line(point[0], point[1], other[0], other[1])
        canvas.circle(point[0], point[1], 1.35 * mm, stroke=0, fill=1)
    canvas.restoreState()


def divider_line(color: colors.Color, width: float = 0.7) -> Table:
    table = Table([[""]], colWidths=[PAGE_WIDTH - BODY_LEFT - BODY_RIGHT], rowHeights=[1])
    table.setStyle(TableStyle([("LINEABOVE", (0, 0), (-1, -1), width, color)]))
    return table


def is_callout(text: str) -> bool:
    return bool(
        re.match(
            r"^\s*(nota|atenção|atencao|recomendação|recomendacao|insight|oportunidade|risco|importante)\s*[:\-]",
            text,
            flags=re.IGNORECASE,
        )
    )


def split_callout(text: str) -> tuple[str, str]:
    match = re.match(r"^\s*([^:\-]{2,40})\s*[:\-]\s*(.*)$", text, flags=re.DOTALL)
    if not match:
        return "Destaque", text
    return match.group(1).strip().title(), match.group(2).strip()


def is_caption(text: str) -> bool:
    return bool(
        re.match(
            r"^\s*(figura|gr[áa]fico|quadro|tabela|fonte|source|nota)\s+\d*[\s:–-]",
            text,
            flags=re.IGNORECASE,
        )
        or re.match(r"^\s*(fonte|source|nota)\s*:", text, flags=re.IGNORECASE)
    )


def looks_numeric(value: str) -> bool:
    stripped = value.strip()
    return bool(re.fullmatch(r"[-+]?R?\$?\s*[\d.,]+%?", stripped))


def clean_for_reportlab(text: str) -> str:
    return escape(text).replace("\n", "<br/>")


def truncate(text: str, max_length: int) -> str:
    if len(text) <= max_length:
        return text
    return f"{text[: max_length - 3].rstrip()}..."


def split_title_lines(title: str, max_chars: int) -> list[str]:
    words = title.split()
    lines: list[str] = []
    current: list[str] = []
    for word in words:
        if current and len(" ".join([*current, word])) > max_chars:
            lines.append(" ".join(current))
            current = [word]
        else:
            current.append(word)
    if current:
        lines.append(" ".join(current))
    return lines[:4] or [title]


def strip_accents(text: str) -> str:
    replacements = str.maketrans(
        {
            "á": "a",
            "à": "a",
            "â": "a",
            "ã": "a",
            "é": "e",
            "ê": "e",
            "í": "i",
            "ó": "o",
            "ô": "o",
            "õ": "o",
            "ú": "u",
            "ü": "u",
            "ç": "c",
        }
    )
    return text.translate(replacements)
