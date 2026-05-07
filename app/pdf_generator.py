from __future__ import annotations

import os
import re
import uuid
from pathlib import Path
from xml.sax.saxutils import escape

from PIL import Image as PilImage
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    KeepTogether,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Flowable,
)
from reportlab.platypus.tableofcontents import TableOfContents

from app.docx_parser import ImageBlock, ParagraphBlock, ParsedDocument, TableBlock, parse_docx


PRIMARY_BLUE = colors.HexColor("#005A9C")
DARK_NAVY = colors.HexColor("#003B6F")
MEDIUM_BLUE = colors.HexColor("#0B63B6")
LIGHT_BLUE = colors.HexColor("#D9EAF7")
VERY_LIGHT_BLUE = colors.HexColor("#EEF6FC")
TEXT_COLOR = colors.HexColor("#2B2B2B")
SECONDARY_TEXT = colors.HexColor("#6B7280")
DIVIDER_COLOR = colors.HexColor("#D1D5DB")
WHITE = colors.white


class FontPack:
    def __init__(self) -> None:
        self.body = "Helvetica"
        self.body_bold = "Helvetica-Bold"
        self.heading = "Helvetica-Bold"
        self.heading_bold = "Helvetica-Bold"


def register_fonts() -> FontPack:
    font_pack = FontPack()
    font_candidates = [
        (
            "DejaVuSans",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        ),
    ]
    for family_name, regular_path, bold_path in font_candidates:
        if os.path.exists(regular_path) and os.path.exists(bold_path):
            pdfmetrics.registerFont(TTFont(family_name, regular_path))
            pdfmetrics.registerFont(TTFont(f"{family_name}-Bold", bold_path))
            font_pack.body = family_name
            font_pack.body_bold = f"{family_name}-Bold"
            font_pack.heading = family_name
            font_pack.heading_bold = f"{family_name}-Bold"
            break
    return font_pack


class DividerLine(Flowable):
    def __init__(self, width: float, color: colors.Color = DIVIDER_COLOR, thickness: float = 1) -> None:
        super().__init__()
        self.line_width = width
        self.color = color
        self.thickness = thickness
        self.height = thickness + 2

    def wrap(self, avail_width: float, avail_height: float) -> tuple[float, float]:
        return min(avail_width, self.line_width), self.height

    def draw(self) -> None:
        self.canv.saveState()
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, self.height / 2, self.line_width, self.height / 2)
        self.canv.restoreState()


class FullPageArtwork(Flowable):
    def __init__(
        self,
        title: str,
        subtitle: str = "",
        metadata_lines: list[str] | None = None,
        logos: list[Path] | None = None,
        contacts: list[str] | None = None,
        variant: str = "cover",
    ) -> None:
        super().__init__()
        self.title = title
        self.subtitle = subtitle
        self.metadata_lines = metadata_lines or []
        self.logos = logos or []
        self.contacts = contacts or []
        self.variant = variant

    def wrap(self, avail_width: float, avail_height: float) -> tuple[float, float]:
        return avail_width, avail_height

    def draw(self) -> None:
        canvas = self.canv
        page_width, page_height = canvas._pagesize
        canvas.saveState()

        if self.variant == "back":
            paint_gradient(canvas, page_width, page_height, DARK_NAVY, PRIMARY_BLUE)
            draw_network_pattern(canvas, page_width, page_height, bottom_only=False, stroke=colors.Color(1, 1, 1, alpha=0.22))
            draw_cover_text(
                canvas,
                page_width,
                page_height,
                self.title,
                self.subtitle,
                self.contacts,
                self.logos,
                back_cover=True,
            )
        else:
            paint_gradient(canvas, page_width, page_height, PRIMARY_BLUE, DARK_NAVY)
            draw_network_pattern(canvas, page_width, page_height, bottom_only=True, stroke=colors.Color(1, 1, 1, alpha=0.18))
            draw_cover_text(
                canvas,
                page_width,
                page_height,
                self.title,
                self.subtitle,
                self.metadata_lines,
                self.logos,
                back_cover=False,
                section_divider=self.variant == "section",
            )

        canvas.restoreState()


class IntelligenceReportTemplate(BaseDocTemplate):
    def __init__(self, filename: str, report_title: str, fonts: FontPack) -> None:
        self.report_title = report_title
        self.fonts = fonts
        self._heading_counter = 0

        left_margin = 1.8 * cm
        right_margin = 1.8 * cm
        top_margin = 2.0 * cm
        bottom_margin = 1.8 * cm

        super().__init__(
            filename,
            pagesize=A4,
            leftMargin=left_margin,
            rightMargin=right_margin,
            topMargin=top_margin,
            bottomMargin=bottom_margin,
            title=report_title,
        )

        body_frame = Frame(
            left_margin,
            bottom_margin,
            A4[0] - left_margin - right_margin,
            A4[1] - top_margin - bottom_margin,
            id="body",
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )
        full_frame = Frame(0, 0, A4[0], A4[1], id="full", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)

        self.addPageTemplates(
            [
                PageTemplate(id="FullPage", frames=[full_frame]),
                PageTemplate(id="Body", frames=[body_frame], onPage=self.draw_body_page),
            ]
        )

    def beforeDocument(self) -> None:
        self._heading_counter = 0

    def draw_body_page(self, canvas, doc) -> None:
        canvas.saveState()
        page_width, page_height = A4

        canvas.setFillColor(PRIMARY_BLUE)
        canvas.rect(doc.leftMargin, page_height - 1.55 * cm, 4.1 * cm, 0.22 * cm, stroke=0, fill=1)
        canvas.setFillColor(TEXT_COLOR)
        canvas.setFont(self.fonts.body_bold, 9)
        canvas.drawString(doc.leftMargin, page_height - 1.1 * cm, truncate_text(self.report_title.upper(), 68))
        canvas.setFont(self.fonts.body, 9)
        canvas.setFillColor(SECONDARY_TEXT)
        canvas.drawRightString(page_width - doc.rightMargin, page_height - 1.1 * cm, f"{canvas.getPageNumber():02d}")

        canvas.setStrokeColor(DIVIDER_COLOR)
        canvas.setLineWidth(0.7)
        canvas.line(doc.leftMargin, page_height - 1.7 * cm, page_width - doc.rightMargin, page_height - 1.7 * cm)

        canvas.setStrokeColor(DIVIDER_COLOR)
        canvas.setLineWidth(0.5)
        canvas.line(doc.leftMargin, doc.bottomMargin - 0.35 * cm, page_width - doc.rightMargin, doc.bottomMargin - 0.35 * cm)
        canvas.restoreState()

    def afterFlowable(self, flowable: Flowable) -> None:
        if isinstance(flowable, Paragraph):
            style_name = getattr(flowable.style, "name", "")
            if style_name in {"ReportHeading1", "ReportHeading2"}:
                level = 0 if style_name == "ReportHeading1" else 1
                text = flowable.getPlainText()
                slug = re.sub(r"[^a-z0-9]+", "-", text.lower()).strip("-") or "section"
                bookmark = f"heading-{self._heading_counter}-{slug}"
                self._heading_counter += 1
                self.canv.bookmarkPage(bookmark)
                self.notify("TOCEntry", (level, text, self.page, bookmark))


def build_pdf_from_docx(docx_path: Path, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    asset_dir = output_dir / f"{docx_path.stem}-assets"
    parsed = parse_docx(docx_path, asset_dir)
    fonts = register_fonts()
    styles = build_styles(fonts)

    output_path = output_dir / f"{docx_path.stem}-{uuid.uuid4().hex[:8]}.pdf"
    document = IntelligenceReportTemplate(str(output_path), parsed.title, fonts)
    story = build_story(parsed, styles)
    document.multiBuild(story)
    return output_path


def build_styles(fonts: FontPack) -> dict[str, ParagraphStyle]:
    stylesheet = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "CoverTitle",
            parent=stylesheet["Heading1"],
            fontName=fonts.heading_bold,
            fontSize=25,
            leading=30,
            textColor=WHITE,
            alignment=TA_LEFT,
            spaceAfter=8,
        ),
        "subtitle": ParagraphStyle(
            "CoverSubtitle",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=11,
            leading=15,
            textColor=WHITE,
            alignment=TA_LEFT,
        ),
        "PageHeading": ParagraphStyle(
            "PageHeading",
            parent=stylesheet["Heading1"],
            fontName=fonts.heading_bold,
            fontSize=17,
            leading=22,
            textColor=PRIMARY_BLUE,
            spaceAfter=6,
        ),
        "ReportHeading1": ParagraphStyle(
            "ReportHeading1",
            parent=stylesheet["Heading1"],
            fontName=fonts.heading_bold,
            fontSize=16,
            leading=21,
            textColor=PRIMARY_BLUE,
            spaceBefore=12,
            spaceAfter=5,
        ),
        "ReportHeading2": ParagraphStyle(
            "ReportHeading2",
            parent=stylesheet["Heading2"],
            fontName=fonts.heading_bold,
            fontSize=12,
            leading=16,
            textColor=DARK_NAVY,
            spaceBefore=10,
            spaceAfter=4,
        ),
        "ReportHeading3": ParagraphStyle(
            "ReportHeading3",
            parent=stylesheet["Heading3"],
            fontName=fonts.body_bold,
            fontSize=10.5,
            leading=14,
            textColor=TEXT_COLOR,
            spaceBefore=8,
            spaceAfter=3,
        ),
        "Body": ParagraphStyle(
            "Body",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=10.2,
            leading=15.2,
            alignment=TA_JUSTIFY,
            textColor=TEXT_COLOR,
            spaceAfter=6,
        ),
        "ListItem": ParagraphStyle(
            "ListItem",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=10.2,
            leading=15.2,
            alignment=TA_LEFT,
            leftIndent=14,
            firstLineIndent=-10,
            bulletIndent=0,
            textColor=TEXT_COLOR,
            spaceAfter=4,
        ),
        "Caption": ParagraphStyle(
            "Caption",
            parent=stylesheet["BodyText"],
            fontName=fonts.body_bold,
            fontSize=8.8,
            leading=11.5,
            alignment=TA_CENTER,
            textColor=TEXT_COLOR,
            spaceBefore=3,
            spaceAfter=2,
        ),
        "Source": ParagraphStyle(
            "Source",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=8,
            leading=10,
            alignment=TA_CENTER,
            textColor=SECONDARY_TEXT,
            spaceBefore=2,
            spaceAfter=8,
        ),
        "SmallLabel": ParagraphStyle(
            "SmallLabel",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=8.4,
            leading=10.8,
            alignment=TA_LEFT,
            textColor=SECONDARY_TEXT,
            spaceAfter=3,
        ),
        "CreditsRole": ParagraphStyle(
            "CreditsRole",
            parent=stylesheet["BodyText"],
            fontName=fonts.body_bold,
            fontSize=10,
            leading=13,
            textColor=DARK_NAVY,
            spaceBefore=6,
            spaceAfter=1,
        ),
        "CreditsName": ParagraphStyle(
            "CreditsName",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=10,
            leading=13.5,
            textColor=TEXT_COLOR,
            leftIndent=10,
            spaceAfter=1,
        ),
        "TOCEntry": ParagraphStyle(
            "TOCEntry",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=10,
            leading=13,
            textColor=TEXT_COLOR,
            leftIndent=0,
            spaceAfter=2,
        ),
        "TOCLevel2": ParagraphStyle(
            "TOCLevel2",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=9.4,
            leading=12.5,
            textColor=SECONDARY_TEXT,
            leftIndent=16,
            firstLineIndent=0,
            spaceAfter=1,
        ),
        "Reference": ParagraphStyle(
            "Reference",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=9.3,
            leading=12.6,
            textColor=TEXT_COLOR,
            leftIndent=16,
            firstLineIndent=-14,
            spaceAfter=4,
        ),
        "Contact": ParagraphStyle(
            "Contact",
            parent=stylesheet["BodyText"],
            fontName=fonts.body,
            fontSize=8.8,
            leading=11,
            alignment=TA_RIGHT,
            textColor=WHITE,
        ),
    }


def build_story(parsed: ParsedDocument, styles: dict[str, ParagraphStyle]) -> list[Flowable]:
    story: list[Flowable] = []

    metadata_lines = []
    if parsed.authors:
        metadata_lines.append(f"Autores: {', '.join(parsed.authors)}")
    if parsed.reviewers:
        metadata_lines.append(f"Revisao: {', '.join(parsed.reviewers)}")
    if parsed.version:
        metadata_lines.append(f"Versao: {parsed.version}")
    if parsed.date:
        metadata_lines.append(f"Data: {parsed.date}")

    story.append(
        FullPageArtwork(
            title=parsed.title.upper(),
            subtitle=parsed.subtitle,
            metadata_lines=metadata_lines,
            logos=parsed.logos,
            variant="cover",
        )
    )
    story.append(NextPageTemplate("Body"))
    story.append(PageBreak())

    if parsed.credits:
        story.extend(build_credits_page(parsed, styles))
    if parsed.catalog_lines:
        story.extend(build_catalog_page(parsed, styles))

    story.extend(build_toc_page(styles))

    body_started = False
    reference_mode = False
    annex_mode = False

    for block in parsed.blocks:
        if not body_started and not (
            isinstance(block, ParagraphBlock) and block.kind == "heading" and block.level == 1
        ):
            story.append(NextPageTemplate("Body"))
            story.append(PageBreak())
            body_started = True

        if isinstance(block, ParagraphBlock) and block.kind == "heading" and block.level == 1:
            reference_mode = "refer" in block.text.lower()
            annex_mode = "anex" in block.text.lower()
            story.append(NextPageTemplate("FullPage"))
            story.append(PageBreak())

            story.append(
                FullPageArtwork(
                    title=block.text.upper(),
                    subtitle=parsed.title,
                    logos=parsed.logos,
                    variant="section",
                )
            )
            story.append(NextPageTemplate("Body"))
            story.append(PageBreak())
            body_started = True

        story.extend(render_block(block, styles, reference_mode=reference_mode, annex_mode=annex_mode))

    story.append(NextPageTemplate("FullPage"))
    story.append(PageBreak())
    story.append(
        FullPageArtwork(
            title=parsed.title.upper(),
            subtitle=parsed.subtitle,
            logos=parsed.logos,
            contacts=parsed.contacts[:3],
            variant="back",
        )
    )

    return story


def build_credits_page(parsed: ParsedDocument, styles: dict[str, ParagraphStyle]) -> list[Flowable]:
    flowables: list[Flowable] = [
        Paragraph("CREDITOS INSTITUCIONAIS", styles["PageHeading"]),
        DividerLine(14 * cm, color=PRIMARY_BLUE, thickness=1.1),
        Spacer(1, 0.45 * cm),
    ]
    for role, names in parsed.credits.items():
        flowables.append(Paragraph(role, styles["CreditsRole"]))
        for name in names:
            flowables.append(Paragraph(escape(name), styles["CreditsName"]))
        flowables.append(Spacer(1, 0.12 * cm))
    flowables.append(PageBreak())
    return flowables


def build_catalog_page(parsed: ParsedDocument, styles: dict[str, ParagraphStyle]) -> list[Flowable]:
    flowables: list[Flowable] = [
        Paragraph("INFORMACOES INSTITUCIONAIS", styles["PageHeading"]),
        DividerLine(14 * cm, color=PRIMARY_BLUE, thickness=1.1),
        Spacer(1, 0.4 * cm),
    ]
    for line in parsed.catalog_lines:
        flowables.append(Paragraph(escape(line), styles["SmallLabel"]))
    flowables.append(PageBreak())
    return flowables


def build_toc_page(styles: dict[str, ParagraphStyle]) -> list[Flowable]:
    toc = TableOfContents()
    toc.levelStyles = [styles["TOCEntry"], styles["TOCLevel2"]]
    return [
        Paragraph("SUMARIO", styles["PageHeading"]),
        DividerLine(14 * cm, color=PRIMARY_BLUE, thickness=1.1),
        Spacer(1, 0.4 * cm),
        toc,
    ]


def render_block(
    block,
    styles: dict[str, ParagraphStyle],
    reference_mode: bool = False,
    annex_mode: bool = False,
) -> list[Flowable]:
    if isinstance(block, ParagraphBlock):
        safe_text = format_paragraph_text(block.text)

        if block.kind == "heading":
            if block.level == 1:
                return [
                    Paragraph(safe_text, styles["ReportHeading1"]),
                    DividerLine(14 * cm, color=DIVIDER_COLOR, thickness=0.9),
                    Spacer(1, 0.15 * cm),
                ]
            if block.level == 2:
                return [Paragraph(safe_text, styles["ReportHeading2"])]
            return [Paragraph(safe_text, styles["ReportHeading3"])]

        if block.kind == "caption":
            return [Paragraph(safe_text, styles["Caption"])]
        if block.kind == "source":
            return [Paragraph(safe_text, styles["Source"])]
        if block.kind == "callout":
            return [build_callout(block.text, styles)]
        if block.kind == "list_item":
            text = block.text[2:] if block.text.startswith(("- ", "• ")) else block.text
            return [Paragraph(format_paragraph_text(text), styles["ListItem"], bulletText="-")]
        if reference_mode:
            return [Paragraph(safe_text, styles["Reference"])]
        if annex_mode:
            return [Paragraph(safe_text, styles["Body"])]
        return [Paragraph(safe_text, styles["Body"])]

    if isinstance(block, TableBlock):
        return [build_table(block.rows, styles)]

    if isinstance(block, ImageBlock):
        return [build_image_flowable(block.path)]

    return []


def build_callout(text: str, styles: dict[str, ParagraphStyle]) -> Flowable:
    paragraph = Paragraph(format_paragraph_text(text), styles["Body"])
    table = Table([[paragraph]], colWidths=[16.3 * cm])
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), VERY_LIGHT_BLUE),
                ("LINEBEFORE", (0, 0), (0, -1), 4, PRIMARY_BLUE),
                ("BOX", (0, 0), (-1, -1), 0.4, DIVIDER_COLOR),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return KeepTogether([table, Spacer(1, 0.15 * cm)])


def build_table(rows: list[list[str]], styles: dict[str, ParagraphStyle]) -> Flowable:
    if not rows:
        return Spacer(1, 0)

    normalized = rows if any(rows) else [[""]]
    column_count = max(len(row) for row in normalized)
    padded_rows = [row + [""] * (column_count - len(row)) for row in normalized]
    usable_width = 16.3 * cm
    column_width = usable_width / max(column_count, 1)

    table_data = []
    for row_index, row in enumerate(padded_rows):
        cells = []
        for cell in row:
            style = styles["SmallLabel"] if row_index == 0 else styles["Body"]
            if row_index == 0:
                style = ParagraphStyle(
                    f"TableHeader-{uuid.uuid4().hex}",
                    parent=styles["SmallLabel"],
                    fontName=styles["ReportHeading3"].fontName,
                    textColor=WHITE,
                    alignment=TA_CENTER,
                )
            cells.append(Paragraph(format_paragraph_text(cell or " "), style))
        table_data.append(cells)

    table = Table(table_data, colWidths=[column_width] * column_count, repeatRows=1, hAlign="LEFT")
    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK_NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("GRID", (0, 0), (-1, -1), 0.5, DIVIDER_COLOR),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]

    for row_index in range(1, len(table_data)):
        fill = VERY_LIGHT_BLUE if row_index % 2 == 1 else WHITE
        table_style.append(("BACKGROUND", (0, row_index), (-1, row_index), fill))

    table.setStyle(TableStyle(table_style))
    return KeepTogether([table, Spacer(1, 0.2 * cm)])


def build_image_flowable(image_path: Path) -> Flowable:
    max_width = 15.5 * cm
    max_height = 11.5 * cm
    try:
        width, height = measure_image(image_path)
        scale = min(max_width / width, max_height / height, 1.0)
        image = Image(str(image_path), width=width * scale, height=height * scale, hAlign="CENTER")
        return KeepTogether([image, Spacer(1, 0.15 * cm)])
    except Exception:
        return Spacer(1, 0.01 * cm)


def measure_image(image_path: Path) -> tuple[float, float]:
    with PilImage.open(image_path) as image:
        width_px, height_px = image.size
    dpi = 96.0
    return width_px * 72.0 / dpi, height_px * 72.0 / dpi


def paint_gradient(canvas, page_width: float, page_height: float, start_color, end_color) -> None:
    steps = 28
    for index in range(steps):
        ratio = index / max(steps - 1, 1)
        color = blend_color(start_color, end_color, ratio)
        canvas.setFillColor(color)
        canvas.rect(0, page_height * (index / steps), page_width, page_height / steps + 2, stroke=0, fill=1)

    canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.06))
    canvas.wedge(page_width * 0.62, page_height * 0.58, page_width * 1.28, page_height * 1.2, 90, 250, stroke=0, fill=1)
    canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.04))
    canvas.rect(page_width * 0.56, 0, page_width * 0.44, page_height * 0.32, stroke=0, fill=1)


def draw_cover_text(
    canvas,
    page_width: float,
    page_height: float,
    title: str,
    subtitle: str,
    lines: list[str],
    logos: list[Path],
    *,
    back_cover: bool = False,
    section_divider: bool = False,
) -> None:
    left = 1.7 * cm
    bottom_panel_height = 7.4 * cm if not back_cover else 5.5 * cm
    panel_y = 0 if not back_cover else 0
    panel_color = colors.Color(0.01, 0.22, 0.43, alpha=0.78) if not back_cover else colors.Color(0.0, 0.18, 0.36, alpha=0.7)
    canvas.setFillColor(panel_color)
    canvas.rect(0, panel_y, page_width, bottom_panel_height, stroke=0, fill=1)

    if not back_cover:
        canvas.setFillColor(PRIMARY_BLUE)
        canvas.rect(left, 5.8 * cm, 1.9 * cm, 0.22 * cm, stroke=0, fill=1)
    else:
        canvas.setStrokeColor(WHITE)
        canvas.setLineWidth(0.8)
        canvas.line(left, 3.2 * cm, page_width - left, 3.2 * cm)

    draw_logos(canvas, logos, page_width, page_height, back_cover)

    title_style = ParagraphStyle(
        "FullPageTitle",
        fontName="Helvetica-Bold",
        fontSize=22 if section_divider else 24,
        leading=28 if section_divider else 30,
        textColor=WHITE,
        alignment=TA_LEFT,
    )
    subtitle_style = ParagraphStyle(
        "FullPageSubtitle",
        fontName="Helvetica",
        fontSize=10.5,
        leading=14,
        textColor=WHITE,
        alignment=TA_LEFT if not back_cover else TA_RIGHT,
    )

    title_paragraph = Paragraph(escape(title), title_style)
    title_width, title_height = title_paragraph.wrap(page_width - 3.4 * cm, 3.5 * cm)
    title_y = 4.1 * cm if not section_divider else 5.2 * cm
    title_paragraph.drawOn(canvas, left, title_y)

    if subtitle:
        subtitle_paragraph = Paragraph(escape(subtitle), subtitle_style)
        subtitle_width = page_width - 3.4 * cm
        subtitle_paragraph.wrap(subtitle_width, 1.5 * cm)
        subtitle_y = title_y - 1.25 * cm
        subtitle_x = left if not back_cover else page_width - subtitle_width - left
        subtitle_paragraph.drawOn(canvas, subtitle_x, subtitle_y)

    if lines:
        meta_style = ParagraphStyle(
            "MetaStyle",
            fontName="Helvetica",
            fontSize=8.5,
            leading=11,
            textColor=WHITE,
            alignment=TA_LEFT if not back_cover else TA_RIGHT,
        )
        y = 2.0 * cm if not back_cover else 1.0 * cm
        for line in lines[:4]:
            paragraph = Paragraph(escape(line), meta_style)
            paragraph.wrap(page_width - 3.4 * cm, 0.8 * cm)
            x = left if not back_cover else page_width - (page_width - 3.4 * cm) - left
            paragraph.drawOn(canvas, x, y)
            y -= 0.45 * cm

    if not back_cover and len(lines) >= 2:
        badge_text = lines[-1]
        canvas.setFillColor(colors.Color(1, 1, 1, alpha=0.16))
        badge_width = 3.1 * cm
        badge_height = 0.9 * cm
        badge_x = page_width - badge_width - 1.5 * cm
        badge_y = 1.15 * cm
        canvas.roundRect(badge_x, badge_y, badge_width, badge_height, 6, stroke=0, fill=1)
        canvas.setFillColor(WHITE)
        canvas.setFont("Helvetica-Bold", 8)
        canvas.drawCentredString(badge_x + badge_width / 2, badge_y + 0.33 * cm, truncate_text(badge_text, 26))


def draw_logos(canvas, logos: list[Path], page_width: float, page_height: float, back_cover: bool) -> None:
    if not logos:
        return

    y = page_height - 2.2 * cm if not back_cover else page_height - 3.2 * cm
    x = 1.7 * cm if not back_cover else (page_width / 2) - ((len(logos[:2]) * 2.4 * cm) / 2)

    for logo_path in logos[:2]:
        try:
            width, height = measure_image(logo_path)
            scale = min((2.3 * cm) / max(width, 1), (1.2 * cm) / max(height, 1))
            canvas.drawImage(
                str(logo_path),
                x,
                y,
                width=width * scale,
                height=height * scale,
                mask="auto",
                preserveAspectRatio=True,
            )
            x += 2.5 * cm
        except Exception:
            continue


def draw_network_pattern(canvas, page_width: float, page_height: float, *, bottom_only: bool, stroke) -> None:
    canvas.setStrokeColor(stroke)
    canvas.setFillColor(stroke)
    points = [
        (page_width * 0.07, page_height * 0.12),
        (page_width * 0.18, page_height * 0.22),
        (page_width * 0.31, page_height * 0.17),
        (page_width * 0.46, page_height * 0.24),
        (page_width * 0.62, page_height * 0.15),
        (page_width * 0.79, page_height * 0.22),
        (page_width * 0.91, page_height * 0.14),
    ]
    if not bottom_only:
        points.extend(
            [
                (page_width * 0.14, page_height * 0.78),
                (page_width * 0.34, page_height * 0.68),
                (page_width * 0.55, page_height * 0.74),
                (page_width * 0.81, page_height * 0.7),
            ]
        )

    for index, point in enumerate(points):
        canvas.circle(point[0], point[1], 2, stroke=0, fill=1)
        if index + 1 < len(points):
            next_point = points[index + 1]
            if bottom_only and point[1] > page_height * 0.34:
                continue
            canvas.line(point[0], point[1], next_point[0], next_point[1])

    for start, end in ((0, 2), (1, 3), (2, 4), (3, 5), (4, 6)):
        if end < len(points):
            canvas.line(points[start][0], points[start][1], points[end][0], points[end][1])


def blend_color(color_a, color_b, ratio: float):
    ratio = max(0.0, min(1.0, ratio))
    red = color_a.red + (color_b.red - color_a.red) * ratio
    green = color_a.green + (color_b.green - color_a.green) * ratio
    blue = color_a.blue + (color_b.blue - color_a.blue) * ratio
    return colors.Color(red, green, blue)


def format_paragraph_text(text: str) -> str:
    return escape(text).replace("\n", "<br/>")


def truncate_text(value: str, max_length: int) -> str:
    return value if len(value) <= max_length else f"{value[: max_length - 1]}..."
