# DOCX to PDF Professional Designer

A web application that transforms Word documents (.docx) into professionally designed PDFs using an institutional design system inspired by Brazilian technical intelligence reports.

## Features

- **Institutional Cover Page** – Dark blue gradient with geometric network pattern, title panel, metadata badge
- **Section Dividers** – Full-page dividers with bold white titles and blue accent bars
- **Styled Tables** – Dark blue headers, alternating row colors, clean borders
- **Typography Hierarchy** – Consistent heading styles (H1–H4) with blue color system
- **Content Preservation** – All text, headings, tables, figures, captions and references are kept intact
- **Back Cover** – Professional back page with network pattern and institutional branding
- **Header/Footer** – Report title and page numbers on every content page

## Quick Start

### Requirements

- Python 3.10+
- pip

### Installation

```bash
pip install -r requirements.txt
```

### Running the App

```bash
python run.py
```

The app starts at `http://localhost:5000`.

### Usage

1. Open the app in your browser
2. Drag & drop a `.docx` file or click **Choose File**
3. Optionally fill in metadata (title, authors, institution, version, date)
4. Click **Generate PDF**
5. Download the professionally designed PDF

## Design System

The PDF follows a corporate institutional style with:

| Element | Value |
|---------|-------|
| Primary Blue | `#005A9C` |
| Dark Navy | `#003B6F` |
| Light Blue | `#D9EAF7` |
| Body Text | `#2B2B2B` |
| Page Format | A4 Portrait |
| Heading Font | Montserrat / Helvetica Bold |
| Body Font | Inter / Helvetica |

## Project Structure

```
├── run.py                  # Entry point
├── requirements.txt        # Python dependencies
├── app/
│   ├── main.py             # Flask app and API endpoint
│   ├── docx_parser.py      # DOCX structure extractor
│   ├── pdf_generator.py    # ReportLab PDF builder
│   ├── templates/
│   │   └── index.html      # Frontend UI
│   └── static/
│       ├── style.css        # Styles
│       └── app.js           # Upload and conversion logic
└── test_pipeline.py        # End-to-end test script
```

## API

### `POST /api/convert`

Upload a DOCX file and receive a designed PDF.

**Form fields:**
- `file` – The `.docx` file (required)
- `metadata` – JSON string with optional overrides: `title`, `subtitle`, `authors`, `institution`, `version`, `date`

**Response:** PDF file download

### `GET /health`

Returns `{"status": "ok"}`.
