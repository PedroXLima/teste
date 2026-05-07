from __future__ import annotations

from io import BytesIO
from typing import Iterable

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Flowable,
    Frame,
    HRFlowable,
    Image,
    NextPageTemplate,
    PageBreak,
    Paragraph,
    PageTemplate,
    Spacer,
    Table,
    TableStyle,
)
from reportlab.platypus.tableofcontents import TableOfContents

from docx_parser import ContentBlock, DocModel


PRIMARY_BLUE = colors.HexColor("#005A9C")
DARK_NAVY = colors.HexColor("#003B6F")
MEDIUM_BLUE = colors.HexColor("#0B63B6")
LIGHT_BLUE = colors.HexColor("#D9EAF7")
VERY_LIGHT_BLUE = colors.HexColor("#EEF6FC")
WHITE = colors.HexColor("#FFFFFF")
MAIN_TEXT = colors.HexColor("#2B2B2B")
SECONDARY_TEXT = colors.HexColor("#6B7280")
DIVIDER = colors.HexColor("#D1D5DB")


MAJOR_SECTIONS = {
    "introdução",
    "introducao",
    "metodologia",
    "desenvolvimento",
    "conclusões",
    "conclusoes",
    "referências",
    "referencias",
    "anexos",
}


def _normalize(text: str) -> str:
    return " ".join(text.strip().lower().split())


class FullPageFlowable(Flowable):
    def __init__(self, draw_callback):
        super().__init__()
        self.draw_callback = draw_callback

    def wrap(self, availWidth, availHeight):
        return A4[0], A4[1]

    def draw(self):
        self.draw_callback(self.canv, A4[0], A4[1])


class StyledDocTemplate(BaseDocTemplate):
    def __init__(self, filename, report_title: str, **kwargs):
        super().__init__(filename, **kwargs)
        self.report_title = report_title

        normal_frame = Frame(
            self.leftMargin,
            self.bottomMargin,
            self.width,
            self.height,
            id="normal-frame",
        )
        full_frame = Frame(
            0,
            0,
            A4[0],
            A4[1],
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
            id="full-frame",
        )

        self.addPageTemplates(
            [
                PageTemplate(id="full", frames=[full_frame]),
                PageTemplate(id="normal", frames=[normal_frame], onPage=self._draw_page_chrome),
            ]
        )

    def _draw_page_chrome(self, canvas: Canvas, doc):
        canvas.saveState()
        top_y = A4[1] - 14 * mm
        canvas.setFillColor(PRIMARY_BLUE)
        canvas.rect(0, A4[1] - 16 * mm, A4[0], 5 * mm, stroke=0, fill=1)

        canvas.setFont("Helvetica-Bold", 9)
        canvas.setFillColor(WHITE)
        header = (self.report_title or "Relatório Institucional").strip()[:85]
        canvas.drawString(self.leftMargin, top_y, header)

        canvas.setFont("Helvetica-Bold", 9)
        canvas.drawRightString(A4[0] - self.rightMargin, top_y, f"{canvas.getPageNumber()}")

        canvas.setStrokeColor(MEDIUM_BLUE)
        canvas.setLineWidth(0.7)
        canvas.line(self.leftMargin, A4[1] - 17 * mm, A4[0] - self.rightMargin, A4[1] - 17 * mm)
        canvas.restoreState()

    def afterFlowable(self, flowable):
        if isinstance(flowable, Paragraph) and hasattr(flowable, "_toc_level"):
            text = getattr(flowable, "_toc_text", flowable.getPlainText())
            level = getattr(flowable, "_toc_level", 0)
            key = f"toc-{self.seq.nextf('heading')}"
            self.canv.bookmarkPage(key)
            self.notify("TOCEntry", (level, text, self.page, key))


def _draw_network_pattern(canvas: Canvas, width: float, height: float, color: colors.Color):
    canvas.saveState()
    canvas.setStrokeColor(color)
    canvas.setLineWidth(0.45)
    step = 24
    for idx in range(0, int(width), step):
        y = (idx % 80) + 22
        canvas.line(idx, 0, min(width, idx + 60), y + 38)
    for x in range(16, int(width), 90):
        for y in range(12, int(height * 0.35), 62):
            canvas.circle(x, y, 1.5, stroke=1, fill=0)
    canvas.restoreState()


def _cover_drawer(model: DocModel):
    metadata_lines = []
    for key in ("descrição", "descricao", "revisão", "revisao", "versão", "versao", "data"):
        if key in model.metadata:
            metadata_lines.append(f"{key.title()}: {model.metadata[key]}")

    def _draw(canvas: Canvas, width: float, height: float):
        canvas.saveState()
        canvas.setFillColor(DARK_NAVY)
        canvas.rect(0, 0, width, height, stroke=0, fill=1)

        for i in range(7):
            opacity = 0.12 + (i * 0.03)
            canvas.setFillColor(colors.Color(0.03, 0.37, 0.69, alpha=opacity))
            canvas.rect(0, i * (height / 7), width, (height / 7), stroke=0, fill=1)

        _draw_network_pattern(canvas, width, height * 0.42, colors.Color(1, 1, 1, alpha=0.25))

        panel_height = height * 0.36
        canvas.setFillColor(colors.Color(0.0, 0.25, 0.46, alpha=0.78))
        canvas.rect(0, 0, width, panel_height, stroke=0, fill=1)

        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 26)
        canvas.drawString(28 * mm, 80 * mm, model.title.upper()[:70])

        if model.subtitle:
            canvas.setFont("Helvetica", 13)
            canvas.drawString(28 * mm, 70 * mm, model.subtitle[:100])

        canvas.setFont("Helvetica", 9)
        y = 54 * mm
        for line in metadata_lines[:5]:
            canvas.drawString(28 * mm, y, line)
            y -= 5.3 * mm

        if model.metadata.get("versão") or model.metadata.get("versao") or model.metadata.get("data"):
            badge = f"Versão {model.metadata.get('versão') or model.metadata.get('versao', '-')}"
            if model.metadata.get("data"):
                badge += f" | {model.metadata['data']}"
            canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.18))
            canvas.roundRect(width - 74 * mm, 16 * mm, 60 * mm, 11 * mm, 3 * mm, stroke=0, fill=1)
            canvas.setFillColor(WHITE)
            canvas.setFont("Helvetica-Bold", 8.4)
            canvas.drawCentredString(width - 44 * mm, 20.5 * mm, badge[:48])
        canvas.restoreState()

    return _draw


def _section_drawer(title: str):
    def _draw(canvas: Canvas, width: float, height: float):
        canvas.saveState()
        canvas.setFillColor(colors.HexColor("#1f4f7a"))
        canvas.rect(0, 0, width, height, stroke=0, fill=1)
        canvas.setFillColor(colors.Color(0, 0, 0, alpha=0.24))
        canvas.rect(0, 0, width, height, stroke=0, fill=1)
        _draw_network_pattern(canvas, width, height * 0.5, colors.Color(1, 1, 1, alpha=0.2))

        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 36)
        canvas.drawString(25 * mm, 56 * mm, title[:35])
        canvas.setFillColor(PRIMARY_BLUE)
        canvas.rect(25 * mm, 50 * mm, 58 * mm, 2.7 * mm, stroke=0, fill=1)
        canvas.restoreState()

    return _draw


def _back_cover_drawer(model: DocModel):
    contact = model.metadata.get("contato") or model.metadata.get("site") or "www.fiea.org.br"

    def _draw(canvas: Canvas, width: float, height: float):
        canvas.saveState()
        canvas.setFillColor(PRIMARY_BLUE)
        canvas.rect(0, 0, width, height, stroke=0, fill=1)
        canvas.setFillColor(colors.Color(0, 0.15, 0.32, alpha=0.38))
        canvas.rect(0, 0, width, height, stroke=0, fill=1)
        _draw_network_pattern(canvas, width, height * 0.45, colors.Color(1, 1, 1, alpha=0.24))

        canvas.setStrokeColor(WHITE)
        canvas.setLineWidth(0.7)
        canvas.line(24 * mm, 30 * mm, width - 24 * mm, 30 * mm)

        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 14)
        canvas.drawString(24 * mm, 21 * mm, model.title[:52])
        if model.subtitle:
            canvas.setFont("Helvetica", 10)
            canvas.drawString(24 * mm, 16 * mm, model.subtitle[:62])

        canvas.setFont("Helvetica", 9)
        canvas.drawRightString(width - 24 * mm, 21 * mm, contact[:42])
        canvas.restoreState()

    return _draw


def _build_styles():
    base = getSampleStyleSheet()
    body = ParagraphStyle(
        "BodyInstitutional",
        parent=base["BodyText"],
        fontName="Helvetica",
        fontSize=10.7,
        leading=16,
        textColor=MAIN_TEXT,
        alignment=TA_JUSTIFY,
        spaceAfter=6.5,
    )
    heading_1 = ParagraphStyle(
        "HeadingInstitutional1",
        parent=base["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=21,
        textColor=PRIMARY_BLUE,
        spaceBefore=10,
        spaceAfter=5,
    )
    heading_2 = ParagraphStyle(
        "HeadingInstitutional2",
        parent=base["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=14.5,
        leading=18,
        textColor=MEDIUM_BLUE,
        spaceBefore=7,
        spaceAfter=4,
    )
    heading_3 = ParagraphStyle(
        "HeadingInstitutional3",
        parent=base["Heading3"],
        fontName="Helvetica-Bold",
        fontSize=12.2,
        leading=15,
        textColor=DARK_NAVY,
        spaceBefore=6,
        spaceAfter=3,
    )
    caption = ParagraphStyle(
        "Caption",
        parent=base["BodyText"],
        fontName="Helvetica",
        fontSize=8.8,
        leading=11,
        textColor=SECONDARY_TEXT,
        spaceBefore=2.5,
        spaceAfter=7.5,
    )
    small_note = ParagraphStyle(
        "SmallNote",
        parent=base["BodyText"],
        fontName="Helvetica",
        fontSize=8.4,
        leading=11,
        textColor=SECONDARY_TEXT,
    )
    return body, heading_1, heading_2, heading_3, caption, small_note


def _table_from_rows(rows: list[list[str]]) -> Table:
    max_columns = max(len(row) for row in rows)
    normalized = [row + [""] * (max_columns - len(row)) for row in rows]
    col_width = (A4[0] - 2 * 20 * mm) / max_columns
    table = Table(normalized, colWidths=[col_width] * max_columns, repeatRows=1)

    style_commands: list[tuple] = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK_NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, DIVIDER),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
    ]

    for row_idx in range(1, len(normalized)):
        bg = WHITE if row_idx % 2 else VERY_LIGHT_BLUE
        style_commands.append(("BACKGROUND", (0, row_idx), (-1, row_idx), bg))

    table.setStyle(TableStyle(style_commands))
    return table


def _add_toc(story: list, heading_style: ParagraphStyle):
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle(
            name="TOCLevel1",
            fontName="Helvetica",
            fontSize=10,
            textColor=MAIN_TEXT,
            leftIndent=6,
            firstLineIndent=-6,
            spaceBefore=2,
        ),
        ParagraphStyle(
            name="TOCLevel2",
            fontName="Helvetica",
            fontSize=9,
            textColor=SECONDARY_TEXT,
            leftIndent=20,
            firstLineIndent=-6,
            spaceBefore=1,
        ),
        ParagraphStyle(
            name="TOCLevel3",
            fontName="Helvetica",
            fontSize=8.5,
            textColor=SECONDARY_TEXT,
            leftIndent=32,
            firstLineIndent=-6,
            spaceBefore=1,
        ),
    ]
    story.append(Paragraph("SUMÁRIO", heading_style))
    story.append(HRFlowable(width="100%", color=DIVIDER, thickness=0.6))
    story.append(Spacer(1, 6))
    story.append(toc)
    story.append(PageBreak())


def _is_caption(text: str) -> bool:
    lowered = text.lower()
    return lowered.startswith("figura ") or lowered.startswith("fonte:") or lowered.startswith("tabela ")


def _append_credits_page(story: list, model: DocModel, heading_style: ParagraphStyle, body_style: ParagraphStyle):
    if not model.credits:
        return
    story.append(Paragraph("CRÉDITOS INSTITUCIONAIS", heading_style))
    story.append(HRFlowable(width="100%", color=DIVIDER, thickness=0.6))
    story.append(Spacer(1, 6))
    for role, names in model.credits.items():
        story.append(Paragraph(f"<b>{role}</b>", body_style))
        for name in names:
            story.append(Paragraph(name, body_style))
        story.append(Spacer(1, 2))
    story.append(PageBreak())


def _append_metadata_catalog_page(story: list, model: DocModel, heading_style: ParagraphStyle, small_style: ParagraphStyle):
    if not model.metadata:
        return
    story.append(Paragraph("INFORMAÇÕES INSTITUCIONAIS", heading_style))
    story.append(HRFlowable(width="100%", color=DIVIDER, thickness=0.6))
    story.append(Spacer(1, 8))
    for key, value in model.metadata.items():
        story.append(Paragraph(f"<b>{key.title()}</b>: {value}", small_style))
        story.append(Spacer(1, 1.2))
    story.append(PageBreak())


def _block_to_story(
    block: ContentBlock,
    body_style: ParagraphStyle,
    h1: ParagraphStyle,
    h2: ParagraphStyle,
    h3: ParagraphStyle,
    caption_style: ParagraphStyle,
) -> Iterable:
    if block.kind == "heading":
        style = h1 if block.level <= 1 else h2 if block.level == 2 else h3
        p = Paragraph(block.text, style)
        p._toc_level = max(0, min(block.level - 1, 2))
        p._toc_text = block.text
        yield p
        if block.level <= 2:
            yield HRFlowable(width="100%", color=DIVIDER, thickness=0.5)
            yield Spacer(1, 3)
        return

    if block.kind == "paragraph":
        style = caption_style if _is_caption(block.text) else body_style
        yield Paragraph(block.text, style)
        return

    if block.kind == "table":
        yield Spacer(1, 4)
        yield _table_from_rows(block.rows)
        yield Spacer(1, 8)
        return

    if block.kind == "image" and block.image_bytes:
        stream = BytesIO(block.image_bytes)
        image = Image(stream)
        max_width = A4[0] - 2 * 24 * mm
        max_height = 110 * mm
        ratio = min(max_width / image.imageWidth, max_height / image.imageHeight)
        ratio = min(ratio, 1.0)
        image.drawWidth = image.imageWidth * ratio
        image.drawHeight = image.imageHeight * ratio
        yield Spacer(1, 4)
        yield image
        yield Spacer(1, 5)


def build_pdf(model: DocModel) -> bytes:
    buffer = BytesIO()
    doc = StyledDocTemplate(
        buffer,
        report_title=model.title,
        pagesize=A4,
        leftMargin=20 * mm,
        rightMargin=20 * mm,
        topMargin=24 * mm,
        bottomMargin=22 * mm,
    )

    body_style, h1, h2, h3, caption_style, small_style = _build_styles()

    story: list = []
    story.append(NextPageTemplate("full"))
    story.append(FullPageFlowable(_cover_drawer(model)))
    story.append(PageBreak())
    story.append(NextPageTemplate("normal"))

    _append_credits_page(story, model, h1, body_style)
    _append_metadata_catalog_page(story, model, h1, small_style)
    _add_toc(story, h1)

    first_heading_seen = False
    for block in model.blocks:
        if block.kind == "heading" and block.level <= 1:
            normalized = _normalize(block.text)
            if first_heading_seen and normalized in MAJOR_SECTIONS:
                story.append(NextPageTemplate("full"))
                story.append(PageBreak())
                story.append(FullPageFlowable(_section_drawer(block.text)))
                story.append(PageBreak())
                story.append(NextPageTemplate("normal"))
            first_heading_seen = True

        story.extend(_block_to_story(block, body_style, h1, h2, h3, caption_style))

    story.append(PageBreak())
    story.append(NextPageTemplate("full"))
    story.append(FullPageFlowable(_back_cover_drawer(model)))

    doc.multiBuild(story)
    return buffer.getvalue()
