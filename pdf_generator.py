from __future__ import annotations

import re
import textwrap
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterator
from xml.sax.saxutils import escape

from docx import Document
from docx.document import Document as DocxDocument
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table as DocxTable
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, StyleSheet1, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate,
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
from reportlab.platypus.flowables import Flowable
from reportlab.platypus.tableofcontents import TableOfContents


PRIMARY_BLUE = colors.HexColor("#005A9C")
DARK_NAVY = colors.HexColor("#003B6F")
MEDIUM_BLUE = colors.HexColor("#0B63B6")
LIGHT_BLUE = colors.HexColor("#D9EAF7")
VERY_LIGHT_BLUE = colors.HexColor("#EEF6FC")
WHITE = colors.HexColor("#FFFFFF")
MAIN_TEXT = colors.HexColor("#2B2B2B")
SECONDARY_TEXT = colors.HexColor("#6B7280")
TABLE_BORDER = colors.HexColor("#D1D5DB")

ROLE_KEYWORDS = {
    "realização",
    "execução",
    "coordenação",
    "autores",
    "autor",
    "projeto gráfico",
    "revisão",
    "equipe",
    "colaboração",
    "organização",
}
MAJOR_SECTIONS = {
    "introdução",
    "metodologia",
    "desenvolvimento",
    "conclusões",
    "conclusão",
    "referências",
    "referencias",
    "anexos",
    "anexo",
}


@dataclass
class ParsedDocument:
    blocks: list[dict]
    title: str
    subtitle: str
    metadata_lines: list[str]
    credits: list[tuple[str, str]]
    catalog_entries: list[str]
    contacts: list[str]


def iter_block_items(parent: DocxDocument) -> Iterator[DocxParagraph | DocxTable]:
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield DocxParagraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield DocxTable(child, parent)


def _heading_level(paragraph: DocxParagraph) -> int | None:
    style_name = paragraph.style.name if paragraph.style is not None else ""
    match = re.search(r"heading\s*(\d+)", style_name, flags=re.IGNORECASE)
    if match:
        return int(match.group(1))
    if style_name.lower().strip() in {"title", "título"}:
        return 1

    text = paragraph.text.strip()
    if not text:
        return None
    if len(text) < 90 and text == text.upper() and text.count(".") <= 1:
        return 1
    return None


def _extract_images(paragraph: DocxParagraph, image_dir: Path) -> list[Path]:
    paths: list[Path] = []
    blips = paragraph._p.xpath(".//a:blip")
    for index, blip in enumerate(blips, start=1):
        rel_id = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
        if rel_id is None:
            continue
        image_part = paragraph.part.related_parts.get(rel_id)
        if image_part is None:
            continue
        extension = image_part.content_type.split("/")[-1].split("+")[0]
        if extension == "jpeg":
            extension = "jpg"
        image_path = image_dir / f"{rel_id}_{index}.{extension}"
        image_path.write_bytes(image_part.blob)
        paths.append(image_path)
    return paths


def _looks_like_caption(text: str) -> bool:
    return bool(re.match(r"^(figura|gráfico|grafico|quadro|imagem)\s*\d*", text, flags=re.IGNORECASE))


def _looks_like_source(text: str) -> bool:
    return bool(re.match(r"^(fonte|source)\s*[:\-]", text, flags=re.IGNORECASE))


def _extract_callout(text: str) -> tuple[str, str] | None:
    match = re.match(
        r"^(insight-chave|insight|recomendação|recomendacao|nota importante|nota|atenção|atencao|risco|oportunidade)\s*[:\-]\s*(.+)$",
        text,
        flags=re.IGNORECASE,
    )
    if not match:
        return None
    title = match.group(1).strip().title()
    body = match.group(2).strip()
    return title, body


def _extract_document_model(doc: DocxDocument, image_dir: Path) -> ParsedDocument:
    blocks: list[dict] = []
    all_text_lines: list[str] = []

    for block in iter_block_items(doc):
        if isinstance(block, DocxParagraph):
            text = block.text.strip()
            for image_path in _extract_images(block, image_dir):
                blocks.append({"type": "image", "path": image_path})

            if not text:
                continue

            all_text_lines.append(text)
            level = _heading_level(block)
            if level is not None:
                blocks.append({"type": "heading", "level": level, "text": text})
                continue

            callout = _extract_callout(text)
            if callout:
                blocks.append({"type": "callout", "title": callout[0], "text": callout[1]})
            elif _looks_like_caption(text):
                blocks.append({"type": "caption", "text": text})
            elif _looks_like_source(text):
                blocks.append({"type": "source", "text": text})
            else:
                blocks.append({"type": "paragraph", "text": text})
        else:
            rows: list[list[str]] = []
            for row in block.rows:
                row_values = [" ".join(cell.text.split()) for cell in row.cells]
                rows.append(row_values)
            if rows:
                all_text_lines.extend([cell for row in rows for cell in row if cell])
                blocks.append({"type": "table", "rows": rows})

    non_empty = [line for line in all_text_lines if line.strip()]
    raw_title = doc.core_properties.title.strip() if doc.core_properties.title else ""
    title = raw_title or (non_empty[0] if non_empty else "RELATÓRIO TÉCNICO")
    subtitle = ""
    if len(non_empty) > 1 and non_empty[1].lower() != title.lower():
        subtitle = non_empty[1]

    metadata_lines: list[str] = []
    credits: list[tuple[str, str]] = []
    catalog_entries: list[str] = []
    contacts: list[str] = []

    for line in non_empty:
        if re.search(r"\b(vers[aã]o|versao|edi[cç][aã]o|data)\b", line, flags=re.IGNORECASE):
            metadata_lines.append(line)
        if ":" in line:
            left, right = line.split(":", 1)
            if left.strip().lower() in ROLE_KEYWORDS and right.strip():
                credits.append((left.strip(), right.strip()))
        if re.search(r"(copyright|todos os direitos|isbn|cataloga[cç][aã]o)", line, flags=re.IGNORECASE):
            catalog_entries.append(line)
        if re.search(r"(@|www\.|http[s]?://|contato|telefone|tel\.)", line, flags=re.IGNORECASE):
            contacts.append(line)

    # Remove duplicates preserving order.
    def dedupe(lines: list[str]) -> list[str]:
        seen: set[str] = set()
        result: list[str] = []
        for value in lines:
            if value not in seen:
                result.append(value)
                seen.add(value)
        return result

    return ParsedDocument(
        blocks=blocks,
        title=title,
        subtitle=subtitle,
        metadata_lines=dedupe(metadata_lines)[:6],
        credits=credits[:20],
        catalog_entries=dedupe(catalog_entries)[:20],
        contacts=dedupe(contacts)[:8],
    )


class CoverFlowable(Flowable):
    def __init__(self, pagesize: tuple[float, float], title: str, subtitle: str, metadata_lines: list[str]):
        super().__init__()
        self.width = pagesize[0]
        self.height = pagesize[1]
        self.title = title.upper()
        self.subtitle = subtitle
        self.metadata_lines = metadata_lines

    def wrap(self, availWidth: float, availHeight: float):
        return self.width, self.height

    def draw(self):
        canv = self.canv
        w, h = self.width, self.height
        steps = 24
        for i in range(steps):
            ratio = i / (steps - 1)
            blend = colors.linearlyInterpolatedColor(PRIMARY_BLUE, DARK_NAVY, 0, 1, ratio)
            canv.setFillColor(blend)
            canv.rect(0, (h / steps) * i, w, (h / steps) + 2, stroke=0, fill=1)

        self._draw_network_pattern(canv, w, h, alpha=0.18)

        panel_h = h * 0.38
        canv.setFillColor(colors.Color(0.02, 0.30, 0.54, alpha=0.78))
        canv.rect(0, 0, w, panel_h, stroke=0, fill=1)

        title_lines = textwrap.wrap(self.title, width=34)[:4] or ["RELATÓRIO"]
        top_y = panel_h - 2.2 * cm
        canv.setFillColor(WHITE)
        canv.setFont("Helvetica-Bold", 28)
        for idx, line in enumerate(title_lines):
            canv.drawString(2.2 * cm, top_y - (idx * 1.0 * cm), line)

        if self.subtitle:
            canv.setFont("Helvetica", 13)
            subtitle_lines = textwrap.wrap(self.subtitle, width=60)[:3]
            for idx, line in enumerate(subtitle_lines):
                canv.drawString(2.2 * cm, top_y - (len(title_lines) * 1.1 * cm) - (idx * 0.7 * cm), line)

        if self.metadata_lines:
            canv.setFont("Helvetica", 9)
            base_y = 1.9 * cm
            for idx, line in enumerate(self.metadata_lines[:4]):
                canv.drawString(2.2 * cm, base_y + (idx * 0.45 * cm), line)

            badge = self.metadata_lines[0]
            canv.setFillColor(colors.Color(1, 1, 1, alpha=0.20))
            badge_w = 6.2 * cm
            badge_h = 0.9 * cm
            badge_x = w - badge_w - 2.2 * cm
            badge_y = 1.6 * cm
            canv.roundRect(badge_x, badge_y, badge_w, badge_h, radius=6, stroke=0, fill=1)
            canv.setFillColor(WHITE)
            canv.setFont("Helvetica-Bold", 8)
            canv.drawRightString(badge_x + badge_w - 0.3 * cm, badge_y + 0.32 * cm, badge[:48])

    @staticmethod
    def _draw_network_pattern(canv, width: float, height: float, alpha: float):
        dot_color = colors.Color(0.85, 0.93, 1.0, alpha=alpha)
        line_color = colors.Color(0.80, 0.90, 1.0, alpha=alpha * 0.7)
        points = [
            (0.72 * width, 0.16 * height),
            (0.80 * width, 0.22 * height),
            (0.87 * width, 0.18 * height),
            (0.92 * width, 0.27 * height),
            (0.78 * width, 0.30 * height),
            (0.66 * width, 0.25 * height),
        ]
        canv.saveState()
        canv.setLineWidth(1)
        canv.setStrokeColor(line_color)
        for i in range(len(points) - 1):
            x1, y1 = points[i]
            x2, y2 = points[i + 1]
            canv.line(x1, y1, x2, y2)
        canv.line(points[0][0], points[0][1], points[3][0], points[3][1])
        canv.setFillColor(dot_color)
        for x, y in points:
            canv.circle(x, y, 3, stroke=0, fill=1)
        canv.restoreState()


class SectionDividerFlowable(Flowable):
    def __init__(self, pagesize: tuple[float, float], title: str):
        super().__init__()
        self.width = pagesize[0]
        self.height = pagesize[1]
        self.title = title

    def wrap(self, availWidth: float, availHeight: float):
        return self.width, self.height

    def draw(self):
        canv = self.canv
        w, h = self.width, self.height
        canv.setFillColor(DARK_NAVY)
        canv.rect(0, 0, w, h, stroke=0, fill=1)

        for i in range(10):
            alpha = 0.03 + i * 0.01
            canv.setFillColor(colors.Color(0.10, 0.42, 0.75, alpha=alpha))
            canv.circle(w * (0.15 + i * 0.08), h * (0.85 - i * 0.05), 55 + i * 8, stroke=0, fill=1)

        CoverFlowable._draw_network_pattern(canv, w, h, alpha=0.23)

        canv.setFillColor(WHITE)
        canv.setFont("Helvetica-Bold", 32)
        wrapped = textwrap.wrap(self.title.upper(), width=28)[:3] or ["SEÇÃO"]
        start_y = 3.0 * cm + (len(wrapped) - 1) * 1.0 * cm
        for idx, line in enumerate(reversed(wrapped)):
            canv.drawString(2.4 * cm, start_y + idx * 1.0 * cm, line)

        canv.setFillColor(PRIMARY_BLUE)
        canv.rect(2.4 * cm, 2.2 * cm, 6.8 * cm, 0.35 * cm, stroke=0, fill=1)


class BackCoverFlowable(Flowable):
    def __init__(self, pagesize: tuple[float, float], title: str, subtitle: str, contacts: list[str]):
        super().__init__()
        self.width = pagesize[0]
        self.height = pagesize[1]
        self.title = title
        self.subtitle = subtitle
        self.contacts = contacts

    def wrap(self, availWidth: float, availHeight: float):
        return self.width, self.height

    def draw(self):
        canv = self.canv
        w, h = self.width, self.height
        steps = 20
        for i in range(steps):
            ratio = i / (steps - 1)
            blend = colors.linearlyInterpolatedColor(PRIMARY_BLUE, DARK_NAVY, 0, 1, ratio)
            canv.setFillColor(blend)
            canv.rect(0, (h / steps) * i, w, (h / steps) + 2, stroke=0, fill=1)

        CoverFlowable._draw_network_pattern(canv, w, h, alpha=0.2)
        canv.setStrokeColor(colors.Color(1, 1, 1, alpha=0.5))
        canv.setLineWidth(0.7)
        canv.line(2.2 * cm, 2.1 * cm, w - 2.2 * cm, 2.1 * cm)

        canv.setFillColor(WHITE)
        canv.setFont("Helvetica-Bold", 13)
        canv.drawString(2.2 * cm, 1.3 * cm, textwrap.shorten(self.title, width=60, placeholder="..."))
        if self.subtitle:
            canv.setFont("Helvetica", 9)
            canv.drawString(2.2 * cm, 0.8 * cm, textwrap.shorten(self.subtitle, width=74, placeholder="..."))

        if self.contacts:
            canv.setFont("Helvetica", 8.5)
            x_pos = w - 2.2 * cm
            y_pos = 1.3 * cm
            for line in self.contacts[:3]:
                canv.drawRightString(x_pos, y_pos, textwrap.shorten(line, width=52, placeholder="..."))
                y_pos -= 0.45 * cm


class InstitutionalDocTemplate(BaseDocTemplate):
    def __init__(self, filename: Path, report_title: str):
        self.pagesize = A4
        self.report_title = report_title
        self.left_margin = 2.1 * cm
        self.right_margin = 2.1 * cm
        self.top_margin = 2.0 * cm
        self.bottom_margin = 2.0 * cm

        super().__init__(
            str(filename),
            pagesize=self.pagesize,
            leftMargin=self.left_margin,
            rightMargin=self.right_margin,
            topMargin=self.top_margin,
            bottomMargin=self.bottom_margin,
        )

        body_frame = Frame(
            self.left_margin,
            self.bottom_margin,
            self.pagesize[0] - self.left_margin - self.right_margin,
            self.pagesize[1] - self.top_margin - self.bottom_margin,
            id="body",
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )
        full_frame = Frame(
            0,
            0,
            self.pagesize[0],
            self.pagesize[1],
            id="full",
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )

        self.addPageTemplates(
            [
                PageTemplate(id="full", frames=[full_frame]),
                PageTemplate(id="body", frames=[body_frame], onPage=self._draw_body_chrome),
            ]
        )

    def _draw_body_chrome(self, canv, doc):
        canv.saveState()
        page_w, page_h = self.pagesize
        top_bar_h = 0.72 * cm
        canv.setFillColor(VERY_LIGHT_BLUE)
        canv.rect(self.left_margin, page_h - self.top_margin + 0.25 * cm, page_w - self.left_margin - self.right_margin, top_bar_h, stroke=0, fill=1)

        canv.setStrokeColor(PRIMARY_BLUE)
        canv.setLineWidth(0.8)
        canv.line(self.left_margin, page_h - self.top_margin + 0.22 * cm, page_w - self.right_margin, page_h - self.top_margin + 0.22 * cm)

        canv.setFillColor(DARK_NAVY)
        canv.setFont("Helvetica-Bold", 9)
        header = textwrap.shorten(self.report_title, width=68, placeholder="...")
        canv.drawString(self.left_margin + 0.25 * cm, page_h - self.top_margin + 0.53 * cm, header)

        canv.setFillColor(SECONDARY_TEXT)
        canv.setFont("Helvetica", 9)
        canv.drawRightString(page_w - self.right_margin, page_h - self.top_margin + 0.53 * cm, str(doc.page))
        canv.restoreState()

    def afterFlowable(self, flowable):
        if getattr(flowable, "_heading_level", None) is None:
            return
        level = max(0, min(2, flowable._heading_level))  # noqa: SLF001
        text = flowable.getPlainText()
        key = f"heading-{self.page}-{abs(hash(text))}"
        self.canv.bookmarkPage(key)
        self.notify("TOCEntry", (level, text, self.page, key))
        try:
            self.canv.addOutlineEntry(text, key, level=level, closed=False)
        except ValueError:
            pass


def _build_styles() -> StyleSheet1:
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="InstitutionalHeading1",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=18,
            leading=22,
            textColor=PRIMARY_BLUE,
            spaceBefore=14,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalHeading2",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=18,
            textColor=MEDIUM_BLUE,
            spaceBefore=12,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalBody",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10.5,
            leading=15,
            textColor=MAIN_TEXT,
            alignment=TA_JUSTIFY,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalCaption",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8.8,
            leading=11,
            textColor=DARK_NAVY,
            alignment=TA_LEFT,
            spaceBefore=2,
            spaceAfter=2,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalSource",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8.2,
            leading=10,
            textColor=SECONDARY_TEXT,
            alignment=TA_CENTER,
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalMetaLabel",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=13,
            textColor=DARK_NAVY,
            alignment=TA_LEFT,
            spaceAfter=1,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalMetaValue",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
            textColor=MAIN_TEXT,
            alignment=TA_LEFT,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalReference",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.6,
            leading=13,
            textColor=MAIN_TEXT,
            leftIndent=0.7 * cm,
            firstLineIndent=-0.6 * cm,
            spaceAfter=4,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalTOC",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
            textColor=MAIN_TEXT,
            alignment=TA_LEFT,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalTOCHeading",
            parent=styles["Heading1"],
            fontName="Helvetica-Bold",
            fontSize=20,
            leading=24,
            textColor=PRIMARY_BLUE,
            alignment=TA_LEFT,
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="InstitutionalSmallMuted",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8.5,
            leading=10.5,
            textColor=SECONDARY_TEXT,
            alignment=TA_LEFT,
            spaceAfter=3,
        )
    )
    return styles


def _make_callout(title: str, text: str, styles: StyleSheet1, width: float) -> Table:
    title_html = f"<b>{escape(title)}</b>"
    body_html = escape(text)
    content = Paragraph(f"{title_html}<br/>{body_html}", styles["InstitutionalBody"])
    callout = Table([[content]], colWidths=[width])
    callout.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), VERY_LIGHT_BLUE),
                ("LINEBEFORE", (0, 0), (0, -1), 3.2, DARK_NAVY),
                ("BOX", (0, 0), (-1, -1), 0.5, TABLE_BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    return callout


def _make_table(rows: list[list[str]], styles: StyleSheet1, width: float) -> Table:
    safe_rows = [[Paragraph(escape(col or " "), styles["InstitutionalBody"]) for col in row] for row in rows]
    col_count = max(len(row) for row in rows)
    col_width = width / max(1, col_count)
    table = Table(safe_rows, colWidths=[col_width] * col_count, repeatRows=1, hAlign="LEFT")
    style = TableStyle(
        [
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
            ("BACKGROUND", (0, 0), (-1, 0), DARK_NAVY),
            ("GRID", (0, 0), (-1, -1), 0.5, TABLE_BORDER),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 5),
            ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]
    )
    for row_idx in range(1, len(rows)):
        bg = WHITE if row_idx % 2 else LIGHT_BLUE
        style.add("BACKGROUND", (0, row_idx), (-1, row_idx), bg)
    table.setStyle(style)
    return table


def _add_credits_page(story: list, parsed: ParsedDocument, styles: StyleSheet1):
    if not parsed.credits:
        return
    story.append(Paragraph("CRÉDITOS INSTITUCIONAIS", styles["InstitutionalTOCHeading"]))
    story.append(Spacer(1, 0.2 * cm))
    for role, names in parsed.credits:
        story.append(Paragraph(escape(role.title()), styles["InstitutionalMetaLabel"]))
        story.append(Paragraph(escape(names), styles["InstitutionalMetaValue"]))
    story.append(PageBreak())


def _add_catalog_page(story: list, parsed: ParsedDocument, styles: StyleSheet1):
    if not parsed.catalog_entries and not parsed.contacts:
        return
    story.append(Paragraph("FICHA CATALOGRÁFICA E CONTATOS", styles["InstitutionalHeading1"]))
    story.append(Spacer(1, 0.2 * cm))
    for line in parsed.catalog_entries:
        story.append(Paragraph(escape(line), styles["InstitutionalSmallMuted"]))
    if parsed.contacts:
        story.append(Spacer(1, 0.4 * cm))
        story.append(Paragraph("CONTATOS INSTITUCIONAIS", styles["InstitutionalHeading2"]))
        for line in parsed.contacts:
            story.append(Paragraph(escape(line), styles["InstitutionalBody"]))
    story.append(PageBreak())


def _add_toc_page(story: list, styles: StyleSheet1):
    toc = TableOfContents()
    toc.levelStyles = [
        ParagraphStyle(
            name="TOCLevel0",
            parent=styles["InstitutionalTOC"],
            fontName="Helvetica-Bold",
            fontSize=10.2,
            textColor=DARK_NAVY,
            leftIndent=0,
            rightIndent=12,
            firstLineIndent=0,
            spaceBefore=2,
            leading=13,
        ),
        ParagraphStyle(
            name="TOCLevel1",
            parent=styles["InstitutionalTOC"],
            fontSize=9.8,
            leftIndent=10,
            rightIndent=12,
            firstLineIndent=0,
            leading=12,
        ),
        ParagraphStyle(
            name="TOCLevel2",
            parent=styles["InstitutionalTOC"],
            fontSize=9.4,
            leftIndent=22,
            rightIndent=12,
            firstLineIndent=0,
            leading=11.5,
        ),
    ]
    story.append(Paragraph("SUMÁRIO", styles["InstitutionalTOCHeading"]))
    story.append(Spacer(1, 0.1 * cm))
    story.append(toc)
    story.append(PageBreak())


def _insert_section_divider(story: list, title: str, pagesize: tuple[float, float]):
    story.append(NextPageTemplate("full"))
    story.append(PageBreak())
    story.append(SectionDividerFlowable(pagesize=pagesize, title=title))
    story.append(NextPageTemplate("body"))
    story.append(PageBreak())


def _scale_image(image_path: Path, max_width: float, max_height: float) -> Image:
    img = Image(str(image_path))
    img.drawWidth = max_width
    img.drawHeight = max_height

    # Keep natural proportions.
    iw, ih = img.imageWidth, img.imageHeight
    ratio = min(max_width / max(iw, 1), max_height / max(ih, 1))
    img.drawWidth = iw * ratio
    img.drawHeight = ih * ratio
    return img


def generate_institutional_pdf(docx_path: Path, pdf_path: Path) -> None:
    source_doc = Document(str(docx_path))
    styles = _build_styles()

    with TemporaryDirectory(prefix="institutional-docx-assets-") as tmp:
        parsed = _extract_document_model(source_doc, Path(tmp))
        doc = InstitutionalDocTemplate(pdf_path, report_title=parsed.title)
        story: list = []

        story.append(CoverFlowable(A4, parsed.title, parsed.subtitle, parsed.metadata_lines))
        story.append(NextPageTemplate("body"))
        story.append(PageBreak())

        _add_credits_page(story, parsed, styles)
        _add_catalog_page(story, parsed, styles)
        _add_toc_page(story, styles)

        in_references = False
        seen_major_sections: set[str] = set()

        for block in parsed.blocks:
            block_type = block["type"]
            if block_type == "heading":
                heading_text = block["text"]
                level = min(block["level"], 3)
                normalized = heading_text.strip().lower()
                if normalized in MAJOR_SECTIONS and normalized not in seen_major_sections:
                    _insert_section_divider(story, heading_text, A4)
                    seen_major_sections.add(normalized)

                if "refer" in normalized:
                    in_references = True
                elif level <= 1 and "refer" not in normalized:
                    in_references = False

                heading_style = styles["InstitutionalHeading1"] if level <= 1 else styles["InstitutionalHeading2"]
                para = Paragraph(escape(heading_text), heading_style)
                para._heading_level = min(level - 1, 2) if level >= 1 else 0  # noqa: SLF001
                story.append(para)
                story.append(
                    Table(
                        [[""]],
                        colWidths=[doc.width],
                        rowHeights=[0.06 * cm],
                        style=[("BACKGROUND", (0, 0), (-1, -1), LIGHT_BLUE), ("LINEBELOW", (0, 0), (-1, 0), 0.4, TABLE_BORDER)],
                    )
                )
                story.append(Spacer(1, 0.08 * cm))
            elif block_type == "paragraph":
                style = styles["InstitutionalReference"] if in_references else styles["InstitutionalBody"]
                story.append(Paragraph(escape(block["text"]), style))
            elif block_type == "caption":
                story.append(Paragraph(escape(block["text"]), styles["InstitutionalCaption"]))
            elif block_type == "source":
                story.append(Paragraph(escape(block["text"]), styles["InstitutionalSource"]))
            elif block_type == "callout":
                story.append(_make_callout(block["title"], block["text"], styles, doc.width))
                story.append(Spacer(1, 0.12 * cm))
            elif block_type == "table":
                table_rows = block["rows"]
                if table_rows:
                    story.append(_make_table(table_rows, styles, doc.width))
                    story.append(Spacer(1, 0.25 * cm))
            elif block_type == "image":
                image_path = block["path"]
                story.append(_scale_image(image_path, max_width=doc.width, max_height=11.5 * cm))
                story.append(Spacer(1, 0.22 * cm))

        story.append(NextPageTemplate("full"))
        story.append(PageBreak())
        story.append(BackCoverFlowable(A4, parsed.title, parsed.subtitle, parsed.contacts))

        doc.multiBuild(story)
