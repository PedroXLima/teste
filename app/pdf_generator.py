"""
PDF Generator – creates a professionally designed PDF using ReportLab
following the institutional design system.
"""

import os
import math
import textwrap
from typing import List, Dict, Any, Optional, Tuple

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm, cm
from reportlab.lib.colors import HexColor, Color, white, black
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    Image, PageBreak, KeepTogether, Frame, PageTemplate,
    BaseDocTemplate, NextPageTemplate, Flowable,
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.shapes import Drawing, Line, Rect, String
from reportlab.graphics import renderPDF
from PIL import Image as PILImage


# ---------------------------------------------------------------------------
# Color Palette
# ---------------------------------------------------------------------------
PRIMARY_BLUE = HexColor("#005A9C")
DARK_NAVY = HexColor("#003B6F")
MEDIUM_BLUE = HexColor("#0B63B6")
LIGHT_BLUE = HexColor("#D9EAF7")
VERY_LIGHT_BLUE = HexColor("#EEF6FC")
WHITE = HexColor("#FFFFFF")
MAIN_TEXT = HexColor("#2B2B2B")
SECONDARY_TEXT = HexColor("#6B7280")
TABLE_BORDER = HexColor("#D1D5DB")

# Page dimensions
PAGE_W, PAGE_H = A4  # 595.27, 841.89 points
MARGIN_TOP = 25 * mm
MARGIN_BOTTOM = 20 * mm
MARGIN_LEFT = 22 * mm
MARGIN_RIGHT = 22 * mm
CONTENT_W = PAGE_W - MARGIN_LEFT - MARGIN_RIGHT


def _register_fonts():
    """Register available fonts; fall back to Helvetica if custom fonts unavailable."""
    font_dirs = [
        "/usr/share/fonts/truetype",
        "/usr/share/fonts",
        "/usr/local/share/fonts",
        os.path.expanduser("~/.fonts"),
    ]

    heading_font = "Helvetica"
    body_font = "Helvetica"

    for d in font_dirs:
        if not os.path.isdir(d):
            continue
        for root, dirs, files in os.walk(d):
            for f in files:
                fl = f.lower()
                full = os.path.join(root, f)
                try:
                    if "montserrat" in fl and fl.endswith(".ttf"):
                        if "bold" in fl and "italic" not in fl:
                            pdfmetrics.registerFont(TTFont("Montserrat-Bold", full))
                        elif "regular" in fl or ("bold" not in fl and "italic" not in fl and "light" not in fl):
                            pdfmetrics.registerFont(TTFont("Montserrat", full))
                            heading_font = "Montserrat"
                    if "inter" in fl and fl.endswith(".ttf"):
                        if "bold" in fl and "italic" not in fl:
                            pdfmetrics.registerFont(TTFont("Inter-Bold", full))
                        elif "regular" in fl or fl == "inter.ttf":
                            pdfmetrics.registerFont(TTFont("Inter", full))
                            body_font = "Inter"
                    if "roboto" in fl and fl.endswith(".ttf"):
                        if "bold" in fl and "italic" not in fl:
                            pdfmetrics.registerFont(TTFont("Roboto-Bold", full))
                        elif "regular" in fl or fl == "roboto.ttf":
                            pdfmetrics.registerFont(TTFont("Roboto", full))
                            if body_font == "Helvetica":
                                body_font = "Roboto"
                    if "dejavusans" in fl.replace("-", "") and fl.endswith(".ttf"):
                        if "bold" in fl:
                            pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", full))
                        elif "oblique" not in fl and "condensed" not in fl:
                            pdfmetrics.registerFont(TTFont("DejaVuSans", full))
                except Exception:
                    pass

    return heading_font, body_font


HEADING_FONT, BODY_FONT = _register_fonts()

HEADING_BOLD = HEADING_FONT + "-Bold" if HEADING_FONT != "Helvetica" else "Helvetica-Bold"
BODY_BOLD = BODY_FONT + "-Bold" if BODY_FONT != "Helvetica" else "Helvetica-Bold"

try:
    pdfmetrics.getFont(HEADING_BOLD)
except KeyError:
    HEADING_BOLD = "Helvetica-Bold"

try:
    pdfmetrics.getFont(BODY_BOLD)
except KeyError:
    BODY_BOLD = "Helvetica-Bold"


def _safe_font(font_name: str) -> str:
    try:
        pdfmetrics.getFont(font_name)
        return font_name
    except KeyError:
        return "Helvetica"


# ---------------------------------------------------------------------------
# Styles
# ---------------------------------------------------------------------------
def _build_styles():
    styles = getSampleStyleSheet()

    hf = _safe_font(HEADING_FONT)
    hfb = _safe_font(HEADING_BOLD)
    bf = _safe_font(BODY_FONT)
    bfb = _safe_font(BODY_BOLD)

    styles.add(ParagraphStyle(
        "DocTitle",
        fontName=hfb,
        fontSize=26,
        leading=32,
        textColor=WHITE,
        alignment=TA_LEFT,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        "DocSubtitle",
        fontName=hf,
        fontSize=14,
        leading=18,
        textColor=WHITE,
        alignment=TA_LEFT,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        "SectionTitle",
        fontName=hfb,
        fontSize=20,
        leading=26,
        textColor=PRIMARY_BLUE,
        spaceBefore=18,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "H2",
        fontName=hfb,
        fontSize=15,
        leading=20,
        textColor=PRIMARY_BLUE,
        spaceBefore=14,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        "H3",
        fontName=hfb,
        fontSize=12,
        leading=16,
        textColor=MEDIUM_BLUE,
        spaceBefore=10,
        spaceAfter=3,
    ))
    styles.add(ParagraphStyle(
        "H4",
        fontName=bfb,
        fontSize=11,
        leading=14,
        textColor=DARK_NAVY,
        spaceBefore=8,
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        "BodyText2",
        fontName=bf,
        fontSize=10,
        leading=15,
        textColor=MAIN_TEXT,
        alignment=TA_JUSTIFY,
        spaceBefore=2,
        spaceAfter=6,
    ))
    styles.add(ParagraphStyle(
        "BodyBold",
        parent=styles["BodyText2"],
        fontName=bfb,
    ))
    styles.add(ParagraphStyle(
        "Caption",
        fontName=bf,
        fontSize=8.5,
        leading=11,
        textColor=SECONDARY_TEXT,
        alignment=TA_CENTER,
        spaceBefore=2,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        "Source",
        fontName=bf,
        fontSize=7.5,
        leading=10,
        textColor=SECONDARY_TEXT,
        alignment=TA_CENTER,
        spaceBefore=1,
        spaceAfter=8,
    ))
    styles.add(ParagraphStyle(
        "SmallMeta",
        fontName=bf,
        fontSize=8,
        leading=10,
        textColor=SECONDARY_TEXT,
        alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        "CreditsLabel",
        fontName=hfb,
        fontSize=10,
        leading=14,
        textColor=DARK_NAVY,
        spaceBefore=10,
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        "CreditsBody",
        fontName=bf,
        fontSize=10,
        leading=14,
        textColor=MAIN_TEXT,
        spaceAfter=2,
    ))
    styles.add(ParagraphStyle(
        "TableHeader",
        fontName=bfb,
        fontSize=9,
        leading=12,
        textColor=WHITE,
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "TableCell",
        fontName=bf,
        fontSize=8.5,
        leading=11,
        textColor=MAIN_TEXT,
        alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        "TableCellCenter",
        fontName=bf,
        fontSize=8.5,
        leading=11,
        textColor=MAIN_TEXT,
        alignment=TA_CENTER,
    ))
    styles.add(ParagraphStyle(
        "ReferenceText",
        fontName=bf,
        fontSize=9,
        leading=13,
        textColor=MAIN_TEXT,
        leftIndent=20,
        firstLineIndent=-20,
        spaceAfter=4,
    ))
    styles.add(ParagraphStyle(
        "CalloutTitle",
        fontName=hfb,
        fontSize=10,
        leading=14,
        textColor=DARK_NAVY,
    ))
    styles.add(ParagraphStyle(
        "CalloutBody",
        fontName=bf,
        fontSize=9.5,
        leading=13,
        textColor=MAIN_TEXT,
    ))
    styles.add(ParagraphStyle(
        "CoverMeta",
        fontName=bf,
        fontSize=9,
        leading=12,
        textColor=HexColor("#D9EAF7"),
        alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        "BackCoverTitle",
        fontName=hfb,
        fontSize=14,
        leading=18,
        textColor=WHITE,
        alignment=TA_LEFT,
    ))
    styles.add(ParagraphStyle(
        "BackCoverBody",
        fontName=bf,
        fontSize=9,
        leading=12,
        textColor=LIGHT_BLUE,
        alignment=TA_LEFT,
    ))

    return styles


# ---------------------------------------------------------------------------
# Custom Flowables
# ---------------------------------------------------------------------------

class BlueDivider(Flowable):
    """Thin horizontal blue divider line."""

    def __init__(self, width=CONTENT_W, thickness=1.2, color=PRIMARY_BLUE):
        super().__init__()
        self.width = width
        self.thickness = thickness
        self.color = color

    def wrap(self, availWidth, availHeight):
        return self.width, self.thickness + 4

    def draw(self):
        self.canv.setStrokeColor(self.color)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 2, self.width, 2)


class GrayDivider(Flowable):
    def __init__(self, width=CONTENT_W, thickness=0.6):
        super().__init__()
        self.width = width
        self.thickness = thickness

    def wrap(self, availWidth, availHeight):
        return self.width, self.thickness + 4

    def draw(self):
        self.canv.setStrokeColor(TABLE_BORDER)
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, 2, self.width, 2)


class NetworkPattern(Flowable):
    """Subtle geometric network pattern background element."""

    def __init__(self, width, height, color=None, opacity=0.08):
        super().__init__()
        self.width = width
        self.height = height
        self.color = color or WHITE
        self.opacity = opacity

    def wrap(self, availWidth, availHeight):
        return 0, 0

    def draw(self):
        c = self.canv
        c.saveState()
        c.setStrokeColor(self.color)
        c.setLineWidth(0.3)

        import random
        random.seed(42)
        points = []
        for _ in range(30):
            x = random.uniform(0, self.width)
            y = random.uniform(0, self.height)
            points.append((x, y))

        for i, (x1, y1) in enumerate(points):
            for x2, y2 in points[i + 1:]:
                dist = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                if dist < self.width * 0.35:
                    alpha = max(0.02, self.opacity * (1 - dist / (self.width * 0.35)))
                    c.setStrokeAlpha(alpha)
                    c.line(x1, y1, x2, y2)

        for x, y in points:
            c.setFillColor(self.color)
            c.setFillAlpha(self.opacity * 1.5)
            c.circle(x, y, 1.5, fill=1, stroke=0)

        c.restoreState()


class FullPageBackground(Flowable):
    """Full-page colored rectangle drawn behind content."""

    def __init__(self, color=DARK_NAVY):
        super().__init__()
        self.bg_color = color

    def wrap(self, availWidth, availHeight):
        return 0, 0

    def draw(self):
        c = self.canv
        c.saveState()
        c.setFillColor(self.bg_color)
        c.rect(0, -PAGE_H, PAGE_W, PAGE_H, fill=1, stroke=0)
        c.restoreState()


class HeaderBar(Flowable):
    """Draws the page header with report title and page number."""

    def __init__(self, title: str, styles):
        super().__init__()
        self.title = title
        self.styles = styles

    def wrap(self, availWidth, availHeight):
        return CONTENT_W, 16

    def draw(self):
        c = self.canv
        c.saveState()
        bf = _safe_font(BODY_FONT)
        c.setFont(bf, 7.5)
        c.setFillColor(SECONDARY_TEXT)

        truncated = self.title[:80]
        c.drawString(0, 4, truncated.upper())

        c.setStrokeColor(PRIMARY_BLUE)
        c.setLineWidth(0.8)
        c.line(0, 0, CONTENT_W, 0)
        c.restoreState()


class CalloutBox(Flowable):
    """Light blue callout/highlight box with left border."""

    def __init__(self, title: str, body: str, styles, width=CONTENT_W):
        super().__init__()
        self.title_text = title
        self.body_text = body
        self.styles = styles
        self.box_width = width
        self._calc_height()

    def _calc_height(self):
        from reportlab.platypus.paragraph import Paragraph as RPara
        title_p = RPara(self.title_text, self.styles["CalloutTitle"])
        body_p = RPara(self.body_text, self.styles["CalloutBody"])
        tw, th = title_p.wrap(self.box_width - 24, 1000)
        bw, bh = body_p.wrap(self.box_width - 24, 1000)
        self.total_h = th + bh + 24
        self._title_para = title_p
        self._body_para = body_p

    def wrap(self, availWidth, availHeight):
        return self.box_width, self.total_h + 8

    def draw(self):
        c = self.canv
        c.saveState()

        c.setFillColor(VERY_LIGHT_BLUE)
        c.roundRect(0, 0, self.box_width, self.total_h, 4, fill=1, stroke=0)

        c.setFillColor(PRIMARY_BLUE)
        c.rect(0, 0, 4, self.total_h, fill=1, stroke=0)

        y = self.total_h - 10
        self._title_para.drawOn(c, 14, y - self._title_para.height)
        y -= self._title_para.height + 4
        self._body_para.drawOn(c, 14, y - self._body_para.height)

        c.restoreState()


# ---------------------------------------------------------------------------
# Page Templates
# ---------------------------------------------------------------------------

def _cover_page(canvas, doc, metadata, styles):
    """Draw the cover page."""
    canvas.saveState()

    canvas.setFillColor(DARK_NAVY)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    gw = PAGE_W
    gh = PAGE_H * 0.4
    import random
    random.seed(42)
    points = []
    for _ in range(40):
        x = random.uniform(0, gw)
        y = random.uniform(0, gh)
        points.append((x, y))

    canvas.setStrokeColor(WHITE)
    canvas.setLineWidth(0.3)
    for i, (x1, y1) in enumerate(points):
        for x2, y2 in points[i + 1:]:
            dist = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
            if dist < gw * 0.3:
                alpha = max(0.02, 0.06 * (1 - dist / (gw * 0.3)))
                canvas.setStrokeAlpha(alpha)
                canvas.line(x1, y1, x2, y2)

    for x, y in points:
        canvas.setFillColor(WHITE)
        canvas.setFillAlpha(0.08)
        canvas.circle(x, y, 2, fill=1, stroke=0)

    canvas.setFillAlpha(1)

    grad_h = PAGE_H * 0.55
    steps = 30
    for i in range(steps):
        frac = i / steps
        r = DARK_NAVY.red * (1 - frac * 0.3) + PRIMARY_BLUE.red * (frac * 0.3)
        g = DARK_NAVY.green * (1 - frac * 0.3) + PRIMARY_BLUE.green * (frac * 0.3)
        b = DARK_NAVY.blue * (1 - frac * 0.3) + PRIMARY_BLUE.blue * (frac * 0.3)
        canvas.setFillColor(Color(r, g, b, 0.85))
        y_pos = PAGE_H * 0.45 - (i * grad_h / steps)
        canvas.rect(0, y_pos, PAGE_W, grad_h / steps + 1, fill=1, stroke=0)

    panel_h = PAGE_H * 0.38
    panel_y = 40
    canvas.setFillColor(Color(PRIMARY_BLUE.red, PRIMARY_BLUE.green, PRIMARY_BLUE.blue, 0.85))
    canvas.rect(0, panel_y, PAGE_W, panel_h, fill=1, stroke=0)

    canvas.setStrokeColor(WHITE)
    canvas.setStrokeAlpha(0.3)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN_LEFT, panel_y + panel_h - 8, PAGE_W - MARGIN_RIGHT, panel_y + panel_h - 8)

    title = metadata.get("title", "Relatório Técnico")
    hfb = _safe_font(HEADING_BOLD)
    canvas.setFont(hfb, 28)
    canvas.setFillColor(WHITE)
    canvas.setFillAlpha(1)

    title_lines = textwrap.wrap(title.upper(), width=35)
    y_title = panel_y + panel_h - 45
    for line in title_lines[:4]:
        canvas.drawString(MARGIN_LEFT + 10, y_title, line)
        y_title -= 36

    subtitle = metadata.get("subtitle", "")
    if subtitle:
        bf = _safe_font(BODY_FONT)
        canvas.setFont(bf, 13)
        canvas.setFillColor(LIGHT_BLUE)
        sub_lines = textwrap.wrap(subtitle, width=60)
        y_sub = y_title - 8
        for line in sub_lines[:3]:
            canvas.drawString(MARGIN_LEFT + 10, y_sub, line)
            y_sub -= 18

    meta_items = []
    if metadata.get("authors"):
        meta_items.append("Autores: " + ", ".join(metadata["authors"]))
    if metadata.get("reviewers"):
        meta_items.append("Revisão: " + ", ".join(metadata["reviewers"]))
    if metadata.get("description"):
        meta_items.append(metadata["description"])

    if meta_items:
        bf = _safe_font(BODY_FONT)
        canvas.setFont(bf, 8.5)
        canvas.setFillColor(LIGHT_BLUE)
        y_meta = panel_y + 60
        for item in meta_items[:4]:
            canvas.drawString(MARGIN_LEFT + 10, y_meta, item[:90])
            y_meta -= 13

    version = metadata.get("version", "")
    date = metadata.get("date", "")
    badge_text = ""
    if version:
        badge_text += f"v{version}"
    if date:
        badge_text += f"  |  {date}" if badge_text else date

    if badge_text:
        canvas.setFillColor(DARK_NAVY)
        badge_w = len(badge_text) * 5.5 + 20
        canvas.roundRect(PAGE_W - MARGIN_RIGHT - badge_w - 10, panel_y + 10, badge_w, 22, 3, fill=1, stroke=0)
        bf = _safe_font(BODY_FONT)
        canvas.setFont(bf, 8)
        canvas.setFillColor(WHITE)
        canvas.drawString(PAGE_W - MARGIN_RIGHT - badge_w, panel_y + 17, badge_text)

    top_bar_h = 60
    canvas.setFillColor(Color(DARK_NAVY.red, DARK_NAVY.green, DARK_NAVY.blue, 0.6))
    canvas.rect(0, PAGE_H - top_bar_h, PAGE_W, top_bar_h, fill=1, stroke=0)

    hfb = _safe_font(HEADING_BOLD)
    canvas.setFont(hfb, 9)
    canvas.setFillColor(WHITE)
    inst = metadata.get("institution", "")
    if inst:
        canvas.drawString(MARGIN_LEFT, PAGE_H - 35, inst.upper())

    canvas.restoreState()


def _back_cover_page(canvas, doc, metadata, styles):
    """Draw the back cover page."""
    canvas.saveState()

    canvas.setFillColor(PRIMARY_BLUE)
    canvas.rect(0, 0, PAGE_W, PAGE_H, fill=1, stroke=0)

    import random
    random.seed(99)
    points = []
    for _ in range(35):
        x = random.uniform(0, PAGE_W)
        y = random.uniform(0, PAGE_H)
        points.append((x, y))

    canvas.setStrokeColor(WHITE)
    canvas.setLineWidth(0.3)
    for i, (x1, y1) in enumerate(points):
        for x2, y2 in points[i + 1:]:
            dist = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
            if dist < PAGE_W * 0.35:
                alpha = 0.05 * (1 - dist / (PAGE_W * 0.35))
                canvas.setStrokeAlpha(max(0.01, alpha))
                canvas.line(x1, y1, x2, y2)

    for x, y in points:
        canvas.setFillColor(WHITE)
        canvas.setFillAlpha(0.06)
        canvas.circle(x, y, 2, fill=1, stroke=0)

    canvas.setFillAlpha(1)

    canvas.setStrokeColor(WHITE)
    canvas.setStrokeAlpha(0.4)
    canvas.setLineWidth(0.5)
    canvas.line(MARGIN_LEFT, 80, PAGE_W - MARGIN_RIGHT, 80)

    title = metadata.get("title", "")
    if title:
        hfb = _safe_font(HEADING_BOLD)
        canvas.setFont(hfb, 14)
        canvas.setFillColor(WHITE)
        canvas.setFillAlpha(1)
        title_lines = textwrap.wrap(title, width=45)
        y = 60
        for line in title_lines[:2]:
            canvas.drawString(MARGIN_LEFT, y, line)
            y -= 18

    subtitle = metadata.get("subtitle", "")
    if subtitle:
        bf = _safe_font(BODY_FONT)
        canvas.setFont(bf, 9)
        canvas.setFillColor(LIGHT_BLUE)
        canvas.drawString(MARGIN_LEFT, 30, subtitle[:80])

    canvas.restoreState()


def _content_header_footer(canvas, doc, metadata, styles):
    """Draw header and footer on standard content pages."""
    canvas.saveState()

    title = metadata.get("title", "Relatório")
    bf = _safe_font(BODY_FONT)
    canvas.setFont(bf, 7)
    canvas.setFillColor(SECONDARY_TEXT)
    canvas.drawString(MARGIN_LEFT, PAGE_H - 18, title[:70].upper())

    canvas.setFont(bf, 7)
    page_num = str(canvas.getPageNumber())
    canvas.drawRightString(PAGE_W - MARGIN_RIGHT, PAGE_H - 18, page_num)

    canvas.setStrokeColor(PRIMARY_BLUE)
    canvas.setLineWidth(0.6)
    canvas.line(MARGIN_LEFT, PAGE_H - 22, PAGE_W - MARGIN_RIGHT, PAGE_H - 22)

    canvas.setStrokeColor(TABLE_BORDER)
    canvas.setLineWidth(0.3)
    canvas.line(MARGIN_LEFT, MARGIN_BOTTOM - 5, PAGE_W - MARGIN_RIGHT, MARGIN_BOTTOM - 5)

    canvas.restoreState()


# ---------------------------------------------------------------------------
# Section Divider Page
# ---------------------------------------------------------------------------

class SectionDividerFlowable(Flowable):
    """Full-page section divider with title."""

    def __init__(self, title: str):
        super().__init__()
        self.title = title

    def wrap(self, availWidth, availHeight):
        return 0, 0

    def draw(self):
        c = self.canv
        c.saveState()

        c.setFillColor(DARK_NAVY)
        c.rect(-MARGIN_LEFT, -PAGE_H + MARGIN_TOP + MARGIN_BOTTOM, PAGE_W, PAGE_H, fill=1, stroke=0)

        import random
        random.seed(hash(self.title) % 10000)
        points = []
        for _ in range(25):
            x = random.uniform(-MARGIN_LEFT, PAGE_W - MARGIN_LEFT)
            y = random.uniform(-PAGE_H + MARGIN_TOP, 0)
            points.append((x, y))

        c.setStrokeColor(WHITE)
        c.setLineWidth(0.3)
        for i, (x1, y1) in enumerate(points):
            for x2, y2 in points[i + 1:]:
                dist = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2)
                if dist < PAGE_W * 0.3:
                    c.setStrokeAlpha(0.04)
                    c.line(x1, y1, x2, y2)

        c.setFillAlpha(1)

        title_y = -PAGE_H + MARGIN_TOP + MARGIN_BOTTOM + 160
        hfb = _safe_font(HEADING_BOLD)
        c.setFont(hfb, 32)
        c.setFillColor(WHITE)
        title_lines = textwrap.wrap(self.title.upper(), width=25)
        for line in title_lines[:3]:
            c.drawString(-MARGIN_LEFT + 40, title_y, line)
            title_y -= 42

        bar_y = title_y - 10
        c.setFillColor(PRIMARY_BLUE)
        c.rect(-MARGIN_LEFT + 40, bar_y, 80, 4, fill=1, stroke=0)

        c.restoreState()


# ---------------------------------------------------------------------------
# Main PDF Builder
# ---------------------------------------------------------------------------

def generate_pdf(parsed_data: Dict[str, Any], output_path: str) -> str:
    """Generate a professionally designed PDF from parsed DOCX data."""

    metadata = parsed_data["metadata"]
    blocks = parsed_data["blocks"]
    styles = _build_styles()

    doc = BaseDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=MARGIN_LEFT,
        rightMargin=MARGIN_RIGHT,
        topMargin=MARGIN_TOP,
        bottomMargin=MARGIN_BOTTOM,
        title=metadata.get("title", ""),
        author=", ".join(metadata.get("authors", [])),
    )

    content_frame = Frame(
        MARGIN_LEFT, MARGIN_BOTTOM,
        CONTENT_W, PAGE_H - MARGIN_TOP - MARGIN_BOTTOM,
        id="content",
    )

    cover_frame = Frame(
        0, 0, PAGE_W, PAGE_H,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        id="cover",
    )

    back_frame = Frame(
        0, 0, PAGE_W, PAGE_H,
        leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
        id="back",
    )

    def on_cover(canvas, doc):
        _cover_page(canvas, doc, metadata, styles)

    def on_content(canvas, doc):
        _content_header_footer(canvas, doc, metadata, styles)

    def on_back(canvas, doc):
        _back_cover_page(canvas, doc, metadata, styles)

    def on_blank(canvas, doc):
        pass

    cover_template = PageTemplate(id="Cover", frames=[cover_frame], onPage=on_cover)
    content_template = PageTemplate(id="Content", frames=[content_frame], onPage=on_content)
    back_template = PageTemplate(id="Back", frames=[back_frame], onPage=on_back)
    blank_template = PageTemplate(id="Blank", frames=[cover_frame], onPage=on_blank)

    doc.addPageTemplates([cover_template, content_template, back_template, blank_template])

    story = []

    story.append(NextPageTemplate("Content"))
    story.append(Spacer(1, 1))
    story.append(PageBreak())

    in_references = False
    in_annexes = False
    in_credits = False
    in_toc = False
    in_presentation = False

    for i, block in enumerate(blocks):
        btype = block["type"]

        if btype == "section_divider":
            story.append(NextPageTemplate("Blank"))
            story.append(PageBreak())
            story.append(SectionDividerFlowable(block["title"]))
            story.append(NextPageTemplate("Content"))
            story.append(PageBreak())
            in_references = block["title"].lower().startswith("referência") or block["title"].lower() == "references"
            in_annexes = block["title"].lower().startswith("anexo")
            in_credits = False
            in_toc = False
            in_presentation = False

        elif btype == "references":
            story.append(NextPageTemplate("Blank"))
            story.append(PageBreak())
            story.append(SectionDividerFlowable(block["title"]))
            story.append(NextPageTemplate("Content"))
            story.append(PageBreak())
            in_references = True
            in_credits = False
            in_toc = False

        elif btype == "annexes":
            story.append(NextPageTemplate("Blank"))
            story.append(PageBreak())
            story.append(SectionDividerFlowable(block["title"]))
            story.append(NextPageTemplate("Content"))
            story.append(PageBreak())
            in_annexes = True
            in_references = False
            in_credits = False

        elif btype == "toc":
            story.append(Paragraph(block["title"], styles["SectionTitle"]))
            story.append(BlueDivider())
            story.append(Spacer(1, 8))
            in_toc = True
            in_credits = False

        elif btype in ("list_of_figures", "list_of_tables"):
            if not in_toc:
                story.append(PageBreak())
            story.append(Paragraph(block["title"], styles["SectionTitle"]))
            story.append(BlueDivider())
            story.append(Spacer(1, 8))

        elif btype == "abbreviations":
            story.append(PageBreak())
            story.append(Paragraph(block["title"], styles["SectionTitle"]))
            story.append(BlueDivider())
            story.append(Spacer(1, 8))

        elif btype == "presentation":
            story.append(PageBreak())
            story.append(Paragraph(block["title"], styles["SectionTitle"]))
            story.append(BlueDivider())
            story.append(Spacer(1, 12))
            in_presentation = True
            in_credits = False

        elif btype == "credits_heading":
            if not in_credits:
                story.append(PageBreak())
                in_credits = True
            story.append(Paragraph(block.get("text", block.get("title", "")), styles["CreditsLabel"]))
            story.append(Spacer(1, 2))

        elif btype == "credits_body":
            text = block["text"]
            if block.get("bold"):
                story.append(Paragraph(text, styles["CreditsLabel"]))
            else:
                story.append(Paragraph(text, styles["CreditsBody"]))

        elif btype == "heading":
            level = block.get("level", 1)
            text = block["text"]

            if in_toc:
                story.append(Paragraph(text, styles["BodyText2"]))
                continue

            if level == 1:
                story.append(Spacer(1, 4))
                story.append(Paragraph(text, styles["SectionTitle"]))
                story.append(BlueDivider())
                story.append(Spacer(1, 4))
            elif level == 2:
                story.append(Paragraph(text, styles["H2"]))
                story.append(GrayDivider())
                story.append(Spacer(1, 2))
            elif level == 3:
                story.append(Paragraph(text, styles["H3"]))
                story.append(Spacer(1, 2))
            else:
                story.append(Paragraph(text, styles["H4"]))
                story.append(Spacer(1, 2))

        elif btype == "paragraph":
            text = block["text"]

            if in_references:
                story.append(Paragraph(text, styles["ReferenceText"]))
                continue

            if block.get("bold"):
                story.append(Paragraph(text, styles["BodyBold"]))
            else:
                story.append(Paragraph(text, styles["BodyText2"]))

        elif btype == "table":
            _add_table(story, block, styles, i, blocks)

        elif btype == "image":
            _add_image(story, block, styles)

        elif btype == "caption":
            story.append(Paragraph(block["text"], styles["Caption"]))

    story.append(NextPageTemplate("Back"))
    story.append(PageBreak())
    story.append(Spacer(1, 1))

    doc.build(story)
    return output_path


def _add_table(story, block, styles, index, blocks):
    """Build and add a styled table."""
    rows = block.get("rows", [])
    if not rows:
        return

    caption_text = ""
    if index > 0:
        prev = blocks[index - 1]
        if prev["type"] == "caption":
            caption_text = prev["text"]

    if caption_text:
        story.append(Spacer(1, 6))
        story.append(Paragraph(f"<b>{caption_text}</b>", styles["Caption"]))
        story.append(Spacer(1, 4))

    num_cols = block.get("num_cols", len(rows[0]))
    if num_cols == 0:
        return

    avail_w = CONTENT_W - 4
    col_w = avail_w / num_cols

    table_data = []
    for ri, row in enumerate(rows):
        styled_row = []
        for ci, cell in enumerate(row):
            if ri == 0:
                styled_row.append(Paragraph(cell, styles["TableHeader"]))
            else:
                try:
                    float(cell.replace(",", ".").replace("%", "").strip())
                    styled_row.append(Paragraph(cell, styles["TableCellCenter"]))
                except (ValueError, AttributeError):
                    styled_row.append(Paragraph(cell, styles["TableCell"]))
        table_data.append(styled_row)

    col_widths = [col_w] * num_cols

    t = Table(table_data, colWidths=col_widths, repeatRows=1)

    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), DARK_NAVY),
        ("TEXTCOLOR", (0, 0), (-1, 0), WHITE),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("FONTSIZE", (0, 1), (-1, -1), 8.5),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, TABLE_BORDER),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]

    for ri in range(1, len(rows)):
        if ri % 2 == 0:
            style_cmds.append(("BACKGROUND", (0, ri), (-1, ri), VERY_LIGHT_BLUE))
        else:
            style_cmds.append(("BACKGROUND", (0, ri), (-1, ri), WHITE))

    t.setStyle(TableStyle(style_cmds))

    source_text = block.get("source", "")
    if source_text:
        story.append(KeepTogether([t, Spacer(1, 3), Paragraph(source_text, styles["Source"])]))
    else:
        story.append(t)

    story.append(Spacer(1, 10))


def _add_image(story, block, styles):
    """Add an image to the story with caption."""
    img_path = block.get("path", "")
    if not img_path or not os.path.exists(img_path):
        return

    try:
        pil_img = PILImage.open(img_path)
        iw, ih = pil_img.size
    except Exception:
        return

    max_w = CONTENT_W * 0.9
    max_h = PAGE_H * 0.4

    scale = min(max_w / iw, max_h / ih, 1.0)
    disp_w = iw * scale
    disp_h = ih * scale

    img = Image(img_path, width=disp_w, height=disp_h)

    elements = [Spacer(1, 6), img]

    caption = block.get("caption", "")
    if caption:
        elements.append(Spacer(1, 3))
        elements.append(Paragraph(caption, styles["Caption"]))

    source = block.get("source", "")
    if source:
        elements.append(Paragraph(source, styles["Source"]))

    elements.append(Spacer(1, 8))

    story.append(KeepTogether(elements))
