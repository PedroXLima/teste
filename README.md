# Institutional DOCX to PDF Designer

A small FastAPI web app that accepts `.docx` uploads and returns a professionally styled PDF
with an institutional/analytical design system inspired by Brazilian industry intelligence reports.

## What it does

- Provides an upload button for DOCX files.
- Preserves extracted document order for paragraphs, headings, tables and embedded images.
- Generates an A4 portrait PDF with:
  - blue institutional cover and back cover;
  - page headers, dividers and page numbering;
  - styled heading hierarchy;
  - institutional table formatting with dark blue header rows and alternating light-blue rows;
  - figure/image placement with preserved proportions;
  - caption/source styling;
  - callout boxes for notes, recommendations, insights and risks.

## Run locally

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open <http://127.0.0.1:8000>, upload a DOCX file and download the generated PDF.

## API

```bash
curl -F "file=@report.docx" http://127.0.0.1:8000/convert --output report.pdf
```

## Notes

The converter intentionally avoids rewriting or summarizing content. It applies a consistent
editorial design to the content that can be extracted from the DOCX file.
