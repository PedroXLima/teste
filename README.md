# DOCX to PDF Designer

A web application that transforms Word documents (.docx) into professionally designed institutional PDFs with a complete corporate design system inspired by Brazilian technical intelligence reports.

## Features

- **Upload Interface** — Drag-and-drop or browse to upload `.docx` files
- **Full Design System** — Corporate blue palette, professional typography, structured layouts
- **Cover Page** — Gradient background with geometric patterns, title, subtitle, metadata badge
- **Section Dividers** — Full-page dividers for major sections (Introdução, Metodologia, etc.)
- **Styled Tables** — Dark blue headers, alternating row colors, proper borders
- **Content Preservation** — All text, headings, tables, figures, captions, and references preserved
- **Back Cover** — Institutional back cover with title and contact information
- **Print-Ready** — A4 format with proper margins, page numbers, and consistent typography

## Tech Stack

- **Backend**: Python, FastAPI, python-docx, WeasyPrint
- **Frontend**: Vanilla HTML/CSS/JS with a modern institutional UI
- **PDF Engine**: WeasyPrint (HTML/CSS → PDF)

## Quick Start

### Prerequisites

- Python 3.10+
- System libraries for WeasyPrint: `libpango`, `libcairo`, `libgdk-pixbuf`
- Fonts: Inter, Montserrat, Roboto (optional but recommended)

### Install

```bash
# System dependencies (Ubuntu/Debian)
sudo apt-get install -y libpango-1.0-0 libcairo2 libgdk-pixbuf2.0-0 \
    fonts-inter fonts-montserrat fonts-roboto

# Python dependencies
pip install -r requirements.txt
```

### Run

```bash
uvicorn app.main:app --host 0.0.0.0 --port 8000
```

Open http://localhost:8000 in your browser, upload a `.docx` file, and click **Generate PDF**.

## Project Structure

```
├── app/
│   ├── __init__.py
│   ├── main.py            # FastAPI application and endpoints
│   ├── docx_parser.py     # DOCX parsing and content extraction
│   ├── pdf_renderer.py    # HTML/CSS generation and PDF rendering
│   └── static/
│       └── index.html     # Frontend UI
├── requirements.txt
└── README.md
```

## Design System

The PDF design system applies:

| Element | Style |
|---|---|
| Primary Blue | `#005A9C` |
| Dark Navy | `#003B6F` |
| Headings | Montserrat, bold, blue |
| Body | Inter, 10.5pt, justified |
| Tables | Blue headers, alternating light-blue rows |
| Cover | Gradient background with geometric network pattern |
| Pages | A4 portrait, 20-22mm margins, page numbers |

## API

### `POST /api/convert`

Upload a `.docx` file and receive the designed PDF.

```bash
curl -o output.pdf -F "file=@document.docx" http://localhost:8000/api/convert
```

### `POST /api/preview`

Returns the intermediate HTML (useful for debugging).
