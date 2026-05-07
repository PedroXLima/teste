"""
HTML builder — converts parsed DOCX content into a fully-designed HTML document
that WeasyPrint will render as a professional institutional PDF.
"""

import html as _html
import re
from typing import Any, Dict, List, Optional

# ---------------------------------------------------------------------------
# SVG network pattern (encoded inline as data-URI)
# ---------------------------------------------------------------------------

_NETWORK_SVG = (
    "<svg xmlns='http://www.w3.org/2000/svg' width='800' height='1130' "
    "opacity='1'>"
    "<defs>"
    "<radialGradient id='ng' cx='50%' cy='50%' r='60%'>"
    "<stop offset='0%' stop-color='rgba(255,255,255,0.06)'/>"
    "<stop offset='100%' stop-color='rgba(255,255,255,0)'/>"
    "</radialGradient>"
    "</defs>"
    "<rect width='800' height='1130' fill='url(#ng)'/>"
    # Node dots
    "<circle cx='80' cy='120' r='3' fill='rgba(255,255,255,0.18)'/>"
    "<circle cx='210' cy='60' r='2' fill='rgba(255,255,255,0.14)'/>"
    "<circle cx='380' cy='90' r='3.5' fill='rgba(255,255,255,0.16)'/>"
    "<circle cx='560' cy='40' r='2' fill='rgba(255,255,255,0.12)'/>"
    "<circle cx='700' cy='130' r='3' fill='rgba(255,255,255,0.18)'/>"
    "<circle cx='150' cy='260' r='2.5' fill='rgba(255,255,255,0.14)'/>"
    "<circle cx='310' cy='220' r='3' fill='rgba(255,255,255,0.16)'/>"
    "<circle cx='480' cy='200' r='2' fill='rgba(255,255,255,0.12)'/>"
    "<circle cx='640' cy='280' r='3' fill='rgba(255,255,255,0.18)'/>"
    "<circle cx='50' cy='400' r='2.5' fill='rgba(255,255,255,0.14)'/>"
    "<circle cx='230' cy='380' r='3' fill='rgba(255,255,255,0.16)'/>"
    "<circle cx='420' cy='350' r='2' fill='rgba(255,255,255,0.12)'/>"
    "<circle cx='600' cy='420' r='3.5' fill='rgba(255,255,255,0.18)'/>"
    "<circle cx='750' cy='360' r='2' fill='rgba(255,255,255,0.14)'/>"
    "<circle cx='120' cy='550' r='3' fill='rgba(255,255,255,0.16)'/>"
    "<circle cx='350' cy='520' r='2.5' fill='rgba(255,255,255,0.12)'/>"
    "<circle cx='520' cy='580' r='3' fill='rgba(255,255,255,0.18)'/>"
    "<circle cx='700' cy='540' r='2' fill='rgba(255,255,255,0.14)'/>"
    "<circle cx='60' cy='700' r='3' fill='rgba(255,255,255,0.16)'/>"
    "<circle cx='260' cy='670' r='2' fill='rgba(255,255,255,0.12)'/>"
    "<circle cx='440' cy='720' r='3.5' fill='rgba(255,255,255,0.18)'/>"
    "<circle cx='620' cy='680' r='2.5' fill='rgba(255,255,255,0.14)'/>"
    "<circle cx='780' cy='740' r='2' fill='rgba(255,255,255,0.12)'/>"
    # Connecting lines
    "<line x1='80' y1='120' x2='210' y2='60' stroke='rgba(255,255,255,0.08)' stroke-width='0.8'/>"
    "<line x1='210' y1='60' x2='380' y2='90' stroke='rgba(255,255,255,0.08)' stroke-width='0.8'/>"
    "<line x1='380' y1='90' x2='560' y2='40' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='560' y1='40' x2='700' y2='130' stroke='rgba(255,255,255,0.08)' stroke-width='0.8'/>"
    "<line x1='80' y1='120' x2='150' y2='260' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='210' y1='60' x2='310' y2='220' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='380' y1='90' x2='480' y2='200' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='700' y1='130' x2='640' y2='280' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='150' y1='260' x2='310' y2='220' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='310' y1='220' x2='480' y2='200' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='480' y1='200' x2='640' y2='280' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='150' y1='260' x2='50' y2='400' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='310' y1='220' x2='230' y2='380' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='480' y1='200' x2='420' y2='350' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='640' y1='280' x2='600' y2='420' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='700' y1='130' x2='750' y2='360' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='50' y1='400' x2='230' y2='380' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='230' y1='380' x2='420' y2='350' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='420' y1='350' x2='600' y2='420' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='600' y1='420' x2='750' y2='360' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='50' y1='400' x2='120' y2='550' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='230' y1='380' x2='350' y2='520' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='420' y1='350' x2='520' y2='580' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='600' y1='420' x2='700' y2='540' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='120' y1='550' x2='350' y2='520' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='350' y1='520' x2='520' y2='580' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='520' y1='580' x2='700' y2='540' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "<line x1='120' y1='550' x2='60' y2='700' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='350' y1='520' x2='260' y2='670' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='520' y1='580' x2='440' y2='720' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='700' y1='540' x2='620' y2='680' stroke='rgba(255,255,255,0.07)' stroke-width='0.8'/>"
    "<line x1='700' y1='540' x2='780' y2='740' stroke='rgba(255,255,255,0.06)' stroke-width='0.8'/>"
    "</svg>"
)

import urllib.parse as _urlparse
_NETWORK_DATA_URI = (
    "data:image/svg+xml,"
    + _urlparse.quote(_NETWORK_SVG, safe="")
)

# ---------------------------------------------------------------------------
# CSS design system
# ---------------------------------------------------------------------------

_CSS = """
/* ---- Google Fonts ---- */
@import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800;900&family=Inter:wght@300;400;500;600&display=swap');

/* ---- CSS Variables ---- */
:root {
  --primary:     #005A9C;
  --navy:        #003B6F;
  --medium-blue: #0B63B6;
  --light-blue:  #D9EAF7;
  --vlight-blue: #EEF6FC;
  --white:       #FFFFFF;
  --text:        #2B2B2B;
  --secondary:   #6B7280;
  --border:      #D1D5DB;
  --font-head:   'Montserrat', 'Helvetica Neue', Arial, sans-serif;
  --font-body:   'Inter', 'Helvetica Neue', Arial, sans-serif;
}

/* ---- Page rules ---- */
@page {
  size: A4 portrait;
  margin: 2.2cm 2.4cm 2.2cm 2.4cm;
}

@page cover-page {
  size: A4 portrait;
  margin: 0;
}

@page back-cover-page {
  size: A4 portrait;
  margin: 0;
}

@page section-divider-page {
  size: A4 portrait;
  margin: 0;
}

@page credits-page {
  size: A4 portrait;
  margin: 2.2cm 2.4cm 2.2cm 2.4cm;
}

@page content-page {
  size: A4 portrait;
  margin: 2.6cm 2.4cm 2.0cm 2.4cm;

  @top-left {
    content: string(doc-title);
    font-family: var(--font-body);
    font-size: 7.5pt;
    color: var(--secondary);
    padding-top: 6pt;
    padding-bottom: 4pt;
    border-bottom: 1.2pt solid var(--primary);
    width: 100%;
  }

  @top-right {
    content: counter(page);
    font-family: var(--font-head);
    font-size: 8pt;
    font-weight: 700;
    color: var(--primary);
    padding-top: 6pt;
    padding-bottom: 4pt;
    border-bottom: 1.2pt solid var(--primary);
  }

  @bottom-center {
    content: '';
    border-top: 0.5pt solid var(--border);
    width: 100%;
    margin-top: 8pt;
  }
}

/* ---- Global reset ---- */
* {
  margin: 0;
  padding: 0;
  box-sizing: border-box;
}

body {
  font-family: var(--font-body);
  font-size: 10.5pt;
  color: var(--text);
  background: var(--white);
  line-height: 1.65;
}

/* ================================================================
   COVER PAGE
   ================================================================ */

.cover-page {
  page: cover-page;
  break-after: page;
  width: 210mm;
  height: 297mm;
  position: relative;
  background: linear-gradient(160deg, var(--navy) 0%, #004080 45%, var(--primary) 75%, var(--medium-blue) 100%);
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.cover-network {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background-image: url('__NETWORK_URI__');
  background-size: cover;
  background-position: center;
  pointer-events: none;
}

.cover-accent-bar {
  position: absolute;
  left: 0;
  bottom: 88mm;
  width: 18mm;
  height: 100%;
  background: rgba(255,255,255,0.06);
}

.cover-top {
  padding: 12mm 14mm 0 14mm;
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.cover-logo-label {
  font-family: var(--font-head);
  font-size: 8.5pt;
  font-weight: 700;
  color: rgba(255,255,255,0.75);
  letter-spacing: 2px;
  text-transform: uppercase;
}

.cover-logo-sub {
  font-family: var(--font-body);
  font-size: 7pt;
  color: rgba(255,255,255,0.5);
  letter-spacing: 1px;
  margin-top: 1mm;
}

.cover-main {
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding: 0 14mm;
}

.cover-tag {
  display: inline-block;
  background: rgba(255,255,255,0.12);
  border: 1pt solid rgba(255,255,255,0.25);
  border-radius: 2pt;
  padding: 1.5mm 4mm;
  font-family: var(--font-head);
  font-size: 7pt;
  font-weight: 600;
  color: rgba(255,255,255,0.8);
  letter-spacing: 2px;
  text-transform: uppercase;
  margin-bottom: 6mm;
}

.cover-title {
  font-family: var(--font-head);
  font-size: 28pt;
  font-weight: 900;
  color: var(--white);
  line-height: 1.15;
  letter-spacing: -0.5px;
  text-transform: uppercase;
  margin-bottom: 5mm;
}

.cover-subtitle {
  font-family: var(--font-body);
  font-size: 13pt;
  font-weight: 300;
  color: rgba(255,255,255,0.85);
  line-height: 1.4;
  max-width: 140mm;
  margin-bottom: 8mm;
}

.cover-divider {
  width: 20mm;
  height: 3pt;
  background: rgba(255,255,255,0.5);
  margin-bottom: 6mm;
}

.cover-meta {
  display: flex;
  flex-direction: column;
  gap: 1mm;
}

.cover-meta-item {
  font-family: var(--font-body);
  font-size: 8.5pt;
  color: rgba(255,255,255,0.7);
}

.cover-meta-item strong {
  color: rgba(255,255,255,0.9);
  font-weight: 600;
}

.cover-footer-band {
  background: rgba(0,0,0,0.30);
  padding: 7mm 14mm;
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.cover-footer-brand {
  font-family: var(--font-head);
  font-size: 8pt;
  font-weight: 700;
  color: rgba(255,255,255,0.9);
  letter-spacing: 1.5px;
  text-transform: uppercase;
}

.cover-badge {
  background: var(--primary);
  border: 1pt solid rgba(255,255,255,0.3);
  border-radius: 3pt;
  padding: 2mm 5mm;
  font-family: var(--font-head);
  font-size: 7.5pt;
  font-weight: 700;
  color: var(--white);
  text-align: center;
}

/* ================================================================
   CREDITS PAGE
   ================================================================ */

.credits-page {
  page: credits-page;
  break-after: page;
}

.credits-title {
  font-family: var(--font-head);
  font-size: 14pt;
  font-weight: 700;
  color: var(--primary);
  margin-bottom: 3mm;
}

.credits-divider {
  width: 100%;
  height: 1.5pt;
  background: var(--light-blue);
  margin-bottom: 8mm;
}

.credits-grid {
  display: flex;
  flex-wrap: wrap;
  gap: 6mm 10mm;
}

.credits-group {
  min-width: 60mm;
  flex: 1;
}

.credits-group-label {
  font-family: var(--font-head);
  font-size: 7.5pt;
  font-weight: 700;
  color: var(--primary);
  text-transform: uppercase;
  letter-spacing: 1.5px;
  margin-bottom: 2mm;
}

.credits-group-value {
  font-family: var(--font-body);
  font-size: 9.5pt;
  color: var(--text);
  line-height: 1.6;
}

/* ================================================================
   SECTION DIVIDER
   ================================================================ */

.section-divider-page {
  page: section-divider-page;
  break-before: page;
  break-after: page;
  width: 210mm;
  height: 297mm;
  position: relative;
  background: linear-gradient(150deg, var(--navy) 0%, var(--primary) 100%);
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  overflow: hidden;
}

.section-divider-network {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background-image: url('__NETWORK_URI__');
  background-size: cover;
  opacity: 0.5;
}

.section-divider-content {
  position: relative;
  padding: 0 14mm 18mm 14mm;
}

.section-divider-bar {
  width: 16mm;
  height: 4pt;
  background: var(--white);
  margin-bottom: 6mm;
}

.section-divider-number {
  font-family: var(--font-head);
  font-size: 9pt;
  font-weight: 600;
  color: rgba(255,255,255,0.6);
  letter-spacing: 3px;
  text-transform: uppercase;
  margin-bottom: 3mm;
}

.section-divider-title {
  font-family: var(--font-head);
  font-size: 30pt;
  font-weight: 900;
  color: var(--white);
  line-height: 1.1;
  text-transform: uppercase;
  letter-spacing: -0.5px;
}

.section-divider-accent {
  width: 100%;
  height: 3pt;
  background: rgba(255,255,255,0.3);
  margin-top: 5mm;
}

/* ================================================================
   CONTENT PAGES
   ================================================================ */

.content-section {
  page: content-page;
}

/* Hidden element used to populate the running header string */
.doc-title-runner {
  string-set: doc-title content();
  font-size: 0;
  color: transparent;
  display: block;
  height: 0;
  overflow: hidden;
}

/* ---- Headings ---- */

h1.doc-h1 {
  font-family: var(--font-head);
  font-size: 17pt;
  font-weight: 800;
  color: var(--navy);
  line-height: 1.2;
  margin-top: 8mm;
  margin-bottom: 2mm;
  break-after: avoid;
  letter-spacing: -0.3px;
}

.h1-underline {
  width: 100%;
  height: 2pt;
  background: linear-gradient(90deg, var(--primary) 0%, var(--light-blue) 100%);
  margin-bottom: 5mm;
}

h2.doc-h2 {
  font-family: var(--font-head);
  font-size: 13pt;
  font-weight: 700;
  color: var(--primary);
  margin-top: 6mm;
  margin-bottom: 3mm;
  break-after: avoid;
}

h3.doc-h3 {
  font-family: var(--font-head);
  font-size: 11pt;
  font-weight: 700;
  color: var(--medium-blue);
  margin-top: 5mm;
  margin-bottom: 2.5mm;
  break-after: avoid;
}

h4.doc-h4 {
  font-family: var(--font-head);
  font-size: 10.5pt;
  font-weight: 600;
  color: var(--primary);
  margin-top: 4mm;
  margin-bottom: 2mm;
  break-after: avoid;
}

h5.doc-h5, h6.doc-h6 {
  font-family: var(--font-body);
  font-size: 10.5pt;
  font-weight: 600;
  color: var(--text);
  margin-top: 3mm;
  margin-bottom: 1.5mm;
  break-after: avoid;
}

/* ---- Body text ---- */

p.body-text {
  font-family: var(--font-body);
  font-size: 10.5pt;
  color: var(--text);
  line-height: 1.7;
  margin-bottom: 3mm;
  text-align: justify;
}

p.body-text:last-child {
  margin-bottom: 0;
}

/* ---- Quote / Block text ---- */

blockquote.doc-quote {
  border-left: 3pt solid var(--primary);
  background: var(--vlight-blue);
  padding: 4mm 6mm;
  margin: 5mm 0;
  font-style: italic;
  color: var(--text);
  font-size: 10pt;
  line-height: 1.6;
}

/* ---- List items ---- */

.doc-list {
  margin: 3mm 0 3mm 7mm;
  padding: 0;
}

.doc-list li {
  font-family: var(--font-body);
  font-size: 10.5pt;
  color: var(--text);
  line-height: 1.65;
  margin-bottom: 1.5mm;
  padding-left: 2mm;
}

.doc-list li::marker {
  color: var(--primary);
  font-weight: 700;
}

/* ---- TOC entries ---- */

.toc-entry {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  font-family: var(--font-body);
  font-size: 10pt;
  color: var(--text);
  padding: 1.5mm 0;
  border-bottom: 0.5pt dotted var(--border);
}

.toc-entry.toc-h1 {
  font-weight: 700;
  color: var(--navy);
  font-size: 10.5pt;
  margin-top: 2mm;
}

.toc-entry.toc-h2 {
  padding-left: 5mm;
}

.toc-entry.toc-h3 {
  padding-left: 10mm;
  font-size: 9.5pt;
  color: var(--secondary);
}

/* ================================================================
   TABLES
   ================================================================ */

.table-wrapper {
  margin: 5mm 0;
  break-inside: avoid;
}

.table-title {
  font-family: var(--font-head);
  font-size: 9.5pt;
  font-weight: 700;
  color: var(--navy);
  margin-bottom: 2mm;
  text-transform: uppercase;
  letter-spacing: 0.5px;
}

table.doc-table {
  width: 100%;
  border-collapse: collapse;
  font-family: var(--font-body);
  font-size: 9.5pt;
}

table.doc-table thead tr {
  background: var(--navy);
}

table.doc-table thead th {
  color: var(--white);
  font-family: var(--font-head);
  font-weight: 700;
  font-size: 9pt;
  padding: 2.5mm 3mm;
  text-align: left;
  border: none;
  letter-spacing: 0.3px;
}

table.doc-table tbody tr:nth-child(odd) {
  background: var(--white);
}

table.doc-table tbody tr:nth-child(even) {
  background: var(--vlight-blue);
}

table.doc-table tbody tr:hover {
  background: var(--light-blue);
}

table.doc-table tbody td {
  padding: 2mm 3mm;
  border: 0.5pt solid var(--border);
  color: var(--text);
  vertical-align: top;
  line-height: 1.5;
}

.table-source {
  font-family: var(--font-body);
  font-size: 7.5pt;
  color: var(--secondary);
  margin-top: 1.5mm;
  text-align: center;
  font-style: italic;
}

/* ================================================================
   FIGURES
   ================================================================ */

.figure-wrapper {
  margin: 5mm 0;
  text-align: center;
  break-inside: avoid;
}

.figure-wrapper img {
  max-width: 100%;
  height: auto;
  border-radius: 2pt;
  display: block;
  margin: 0 auto;
}

.figure-caption {
  font-family: var(--font-body);
  font-size: 8.5pt;
  color: var(--secondary);
  margin-top: 2mm;
  text-align: center;
  font-style: italic;
  line-height: 1.45;
}

/* ================================================================
   CALLOUT / HIGHLIGHT BOXES
   ================================================================ */

.callout-box {
  background: var(--vlight-blue);
  border-left: 4pt solid var(--primary);
  border-radius: 0 3pt 3pt 0;
  padding: 4mm 5mm;
  margin: 5mm 0;
  break-inside: avoid;
}

.callout-title {
  font-family: var(--font-head);
  font-size: 9pt;
  font-weight: 700;
  color: var(--primary);
  text-transform: uppercase;
  letter-spacing: 1px;
  margin-bottom: 2mm;
}

.callout-body {
  font-family: var(--font-body);
  font-size: 10pt;
  color: var(--text);
  line-height: 1.6;
}

/* ================================================================
   CAPTION / SOURCE
   ================================================================ */

p.caption-text {
  font-family: var(--font-body);
  font-size: 8.5pt;
  color: var(--secondary);
  font-style: italic;
  text-align: center;
  margin-top: 1mm;
  margin-bottom: 3mm;
}

/* ================================================================
   BACK COVER
   ================================================================ */

.back-cover-page {
  page: back-cover-page;
  break-before: page;
  width: 210mm;
  height: 297mm;
  position: relative;
  background: linear-gradient(160deg, var(--navy) 0%, #00508C 60%, var(--primary) 100%);
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  overflow: hidden;
}

.back-cover-network {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  background-image: url('__NETWORK_URI__');
  background-size: cover;
  opacity: 0.6;
}

.back-cover-content {
  position: relative;
  padding: 0 14mm 14mm 14mm;
}

.back-cover-divider {
  width: 100%;
  height: 1pt;
  background: rgba(255,255,255,0.3);
  margin-bottom: 7mm;
}

.back-cover-footer {
  display: flex;
  justify-content: space-between;
  align-items: flex-end;
}

.back-cover-brand {
  font-family: var(--font-head);
  font-size: 10pt;
  font-weight: 800;
  color: var(--white);
  text-transform: uppercase;
  letter-spacing: 2px;
}

.back-cover-brand-sub {
  font-family: var(--font-body);
  font-size: 8pt;
  color: rgba(255,255,255,0.6);
  margin-top: 1mm;
  letter-spacing: 1px;
}

.back-cover-contact {
  text-align: right;
  font-family: var(--font-body);
  font-size: 8pt;
  color: rgba(255,255,255,0.7);
  line-height: 1.7;
}

.back-cover-title-block {
  margin-bottom: 8mm;
}

.back-cover-doc-title {
  font-family: var(--font-head);
  font-size: 13pt;
  font-weight: 700;
  color: var(--white);
  margin-bottom: 2mm;
}

.back-cover-doc-subtitle {
  font-family: var(--font-body);
  font-size: 10pt;
  color: rgba(255,255,255,0.75);
}

/* ================================================================
   PAGE BREAK UTILITIES
   ================================================================ */

.page-break-before { break-before: page; }
.page-break-after  { break-after:  page; }
.no-break          { break-inside: avoid; }
"""


# ---------------------------------------------------------------------------
# HtmlBuilder
# ---------------------------------------------------------------------------

class HtmlBuilder:
    def __init__(self, content: Dict[str, Any]):
        self.content = content
        self.meta = content.get('metadata', {})
        self.elements = content.get('elements', [])
        self._section_counter = 0
        self._table_counter = 0
        self._figure_counter = 0

    # ------------------------------------------------------------------
    # Public
    # ------------------------------------------------------------------

    def build(self) -> str:
        css = _CSS.replace('__NETWORK_URI__', _NETWORK_DATA_URI)

        sections = [
            self._build_cover(),
            self._build_content_section(),
            self._build_back_cover(),
        ]

        return (
            '<!DOCTYPE html>\n'
            '<html lang="pt-BR">\n'
            '<head>\n'
            '  <meta charset="UTF-8">\n'
            f'  <style>{css}</style>\n'
            '</head>\n'
            '<body>\n'
            + '\n'.join(sections)
            + '\n</body>\n</html>'
        )

    # ------------------------------------------------------------------
    # Cover
    # ------------------------------------------------------------------

    def _build_cover(self) -> str:
        title = self.meta.get('title') or 'Relatório Institucional'
        subtitle = self.meta.get('subtitle') or self.meta.get('subject') or ''
        author = self.meta.get('author') or ''
        date = self.meta.get('created') or self.meta.get('modified') or ''
        version = self.meta.get('version') or ''
        reviewer = self.meta.get('reviewer') or ''

        meta_items = ''
        if author:
            meta_items += f'<div class="cover-meta-item"><strong>Autores:</strong> {e(author)}</div>'
        if reviewer:
            meta_items += f'<div class="cover-meta-item"><strong>Revisão:</strong> {e(reviewer)}</div>'
        if date:
            meta_items += f'<div class="cover-meta-item"><strong>Data:</strong> {e(date)}</div>'

        badge_html = ''
        if version or date:
            badge_text = version or date
            badge_html = f'<div class="cover-badge">{e(badge_text)}</div>'

        return f"""
<div class="cover-page">
  <div class="cover-network"></div>
  <div class="cover-top">
    <div>
      <div class="cover-logo-label">Observatório da Indústria</div>
      <div class="cover-logo-sub">Sistema FIEA · Inteligência Estratégica</div>
    </div>
  </div>
  <div class="cover-main">
    <div class="cover-tag">Relatório Técnico</div>
    <h1 class="cover-title">{e(title)}</h1>
    {f'<p class="cover-subtitle">{e(subtitle)}</p>' if subtitle else ''}
    <div class="cover-divider"></div>
    <div class="cover-meta">
      {meta_items}
    </div>
  </div>
  <div class="cover-footer-band">
    <div class="cover-footer-brand">Sistema FIEA</div>
    {badge_html}
  </div>
</div>"""

    # ------------------------------------------------------------------
    # Main content section
    # ------------------------------------------------------------------

    def _build_content_section(self) -> str:
        title = self.meta.get('title') or 'Relatório Institucional'
        parts: List[str] = []

        # Invisible runner element to set the page header string
        parts.append(f'<span class="doc-title-runner">{e(title)}</span>')

        pending_list: List[str] = []

        def flush_list():
            if pending_list:
                parts.append('<ul class="doc-list">')
                parts.extend(pending_list)
                parts.append('</ul>')
                pending_list.clear()

        skip_first_title = True  # skip title/subtitle already used for cover

        for elem in self.elements:
            etype = elem.get('type', 'paragraph')
            text = elem.get('text') or ''
            runs = elem.get('runs') or []
            images = elem.get('images') or []

            # Skip doc-level title/subtitle already on cover
            if skip_first_title and etype in ('title', 'subtitle') and text:
                skip_first_title = False
                continue

            if etype != 'list_item':
                flush_list()

            if etype == 'heading1':
                self._section_counter += 1
                parts.append(
                    f'<h1 class="doc-h1">{render_runs(runs, text)}</h1>'
                    f'<div class="h1-underline"></div>'
                )

            elif etype == 'heading2':
                parts.append(f'<h2 class="doc-h2">{render_runs(runs, text)}</h2>')

            elif etype == 'heading3':
                parts.append(f'<h3 class="doc-h3">{render_runs(runs, text)}</h3>')

            elif etype in ('heading4', 'heading5', 'heading6'):
                tag = 'h4' if etype == 'heading4' else ('h5' if etype == 'heading5' else 'h6')
                cls = f'doc-{tag}'
                parts.append(f'<{tag} class="{cls}">{render_runs(runs, text)}</{tag}>')

            elif etype == 'paragraph':
                if text or images:
                    if images:
                        for img in images:
                            parts.append(self._render_inline_image(img))
                    if text:
                        parts.append(f'<p class="body-text">{render_runs(runs, text)}</p>')

            elif etype == 'list_item':
                pending_list.append(f'<li>{render_runs(runs, text)}</li>')

            elif etype == 'caption':
                parts.append(f'<p class="caption-text">{render_runs(runs, text)}</p>')

            elif etype == 'quote':
                parts.append(f'<blockquote class="doc-quote">{render_runs(runs, text)}</blockquote>')

            elif etype == 'callout':
                parts.append(
                    f'<div class="callout-box">'
                    f'<div class="callout-title">Destaque</div>'
                    f'<div class="callout-body">{render_runs(runs, text)}</div>'
                    f'</div>'
                )

            elif etype == 'toc_entry':
                parts.append(self._render_toc_entry(text, runs))

            elif etype == 'title':
                # Treat as a prominent paragraph
                parts.append(f'<p class="body-text"><strong>{render_runs(runs, text)}</strong></p>')

            elif etype == 'subtitle':
                parts.append(f'<p class="body-text"><em>{render_runs(runs, text)}</em></p>')

            elif etype == 'figure':
                for img in images:
                    parts.append(self._render_figure(img, text))

            elif etype == 'table':
                parts.append(self._render_table(elem))

        flush_list()

        return f'<div class="content-section">\n' + '\n'.join(parts) + '\n</div>'

    def _render_inline_image(self, img: Dict[str, str]) -> str:
        self._figure_counter += 1
        mime = img.get('content_type', 'image/png')
        data = img.get('data', '')
        return (
            f'<div class="figure-wrapper">'
            f'<img src="data:{mime};base64,{data}" alt="Figura {self._figure_counter}"/>'
            f'</div>'
        )

    def _render_figure(self, img: Dict[str, str], caption: str = '') -> str:
        self._figure_counter += 1
        mime = img.get('content_type', 'image/png')
        data = img.get('data', '')
        cap_html = (
            f'<p class="figure-caption">Figura {self._figure_counter}'
            + (f' – {e(caption)}' if caption else '')
            + '</p>'
        ) if True else ''
        return (
            f'<div class="figure-wrapper">'
            f'<img src="data:{mime};base64,{data}" alt="Figura {self._figure_counter}"/>'
            f'{cap_html}'
            f'</div>'
        )

    def _render_toc_entry(self, text: str, runs: List) -> str:
        # Detect indentation level by leading spaces
        stripped = text.lstrip()
        leading = len(text) - len(stripped)
        if leading >= 8:
            cls = 'toc-entry toc-h3'
        elif leading >= 4:
            cls = 'toc-entry toc-h2'
        else:
            cls = 'toc-entry toc-h1'
        return (
            f'<div class="{cls}">'
            f'<span>{render_runs(runs, stripped)}</span>'
            f'</div>'
        )

    def _render_table(self, elem: Dict[str, Any]) -> str:
        self._table_counter += 1
        rows = elem.get('rows', [])
        if not rows:
            return ''

        header_row = rows[0]
        body_rows = rows[1:]

        thead = '<thead><tr>'
        for cell in header_row:
            thead += f'<th>{e(cell["text"])}</th>'
        thead += '</tr></thead>'

        tbody = '<tbody>'
        for row in body_rows:
            tbody += '<tr>'
            for cell in row:
                tbody += f'<td>{e(cell["text"])}</td>'
            tbody += '</tr>'
        tbody += '</tbody>'

        return (
            f'<div class="table-wrapper">'
            f'<div class="table-title">Tabela {self._table_counter}</div>'
            f'<table class="doc-table">{thead}{tbody}</table>'
            f'</div>'
        )

    # ------------------------------------------------------------------
    # Back cover
    # ------------------------------------------------------------------

    def _build_back_cover(self) -> str:
        title = self.meta.get('title') or ''
        subtitle = self.meta.get('subtitle') or self.meta.get('subject') or ''

        return f"""
<div class="back-cover-page">
  <div class="back-cover-network"></div>
  <div class="back-cover-content">
    <div class="back-cover-title-block">
      {f'<div class="back-cover-doc-title">{e(title)}</div>' if title else ''}
      {f'<div class="back-cover-doc-subtitle">{e(subtitle)}</div>' if subtitle else ''}
    </div>
    <div class="back-cover-divider"></div>
    <div class="back-cover-footer">
      <div>
        <div class="back-cover-brand">Observatório da Indústria</div>
        <div class="back-cover-brand-sub">Sistema FIEA · Inteligência Estratégica</div>
      </div>
      <div class="back-cover-contact">
        observatorio.fiea.org.br<br>
        observatorio@fiea.org.br
      </div>
    </div>
  </div>
</div>"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def e(text: str) -> str:
    """HTML-escape a string."""
    return _html.escape(str(text))


def render_runs(runs: List[Dict[str, Any]], fallback: str = '') -> str:
    """Convert a list of run dicts to HTML with inline formatting."""
    if not runs:
        return e(fallback)
    parts: List[str] = []
    for run in runs:
        text = e(run.get('text', ''))
        if not text:
            continue
        if run.get('bold') and run.get('italic'):
            text = f'<strong><em>{text}</em></strong>'
        elif run.get('bold'):
            text = f'<strong>{text}</strong>'
        elif run.get('italic'):
            text = f'<em>{text}</em>'
        if run.get('underline'):
            text = f'<u>{text}</u>'
        parts.append(text)
    return ''.join(parts) if parts else e(fallback)
