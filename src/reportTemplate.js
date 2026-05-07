import * as cheerio from "cheerio";
import { slugify } from "./docxReport.js";

export function buildReportHtml(report) {
  const title = report.title || "Relatorio tecnico";
  const subtitle = report.subtitle || "Publicacao institucional de inteligencia estrategica";
  const bodyHtml = decorateBodyHtml(report.bodyHtml, title);

  return `<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>${escapeHtml(title)}</title>
    <style>${REPORT_CSS}</style>
  </head>
  <body>
    ${renderCover(report, title, subtitle)}
    ${renderCredits(report)}
    ${renderToc(report)}
    <main class="report-content">
      ${bodyHtml}
    </main>
    ${renderBackCover(report, title, subtitle)}
  </body>
</html>`;
}

export function decorateBodyHtml(html, title) {
  const $ = cheerio.load(`<main>${html}</main>`, { decodeEntities: false });
  const normalizedTitle = normalize(title);

  $("h1").each((index, element) => {
    const $heading = $(element);
    const headingText = normalizeWhitespace($heading.text());

    if (index === 0 && normalize(headingText) === normalizedTitle) {
      $heading.remove();
      return;
    }

    const id = $heading.attr("id") || slugify(headingText);
    const divider = `
      <section class="section-divider" id="divider-${escapeAttribute(id)}">
        <div class="network-pattern"></div>
        <div class="section-divider__content">
          <span>Inteligencia e analise</span>
          <h1>${escapeHtml(headingText)}</h1>
          <div class="section-divider__bar"></div>
        </div>
      </section>`;

    $heading.before(divider);
    $heading.remove();
  });

  $("table").each((_, element) => {
    const $table = $(element);

    if (!$table.parent().hasClass("table-shell")) {
      $table.wrap('<div class="table-shell"></div>');
    }
  });

  $("p").each((_, element) => {
    const $paragraph = $(element);

    if ($paragraph.find("img").length > 0) {
      $paragraph.addClass("image-card");
    }

    if ($paragraph.hasClass("callout-source")) {
      const text = normalizeWhitespace($paragraph.text());
      const [label, ...rest] = text.split(":");
      $paragraph.replaceWith(`
        <aside class="callout">
          <strong>${escapeHtml(label || "Destaque")}</strong>
          <span>${escapeHtml(rest.join(":").trim() || text)}</span>
        </aside>`);
    }
  });

  return $("main").html() ?? "";
}

function renderCover(report, title, subtitle) {
  const badge = findMetadataValue(report.metadata, /vers[aã]o|data/i);

  return `
    <section class="cover page-full">
      <div class="cover__gradient"></div>
      <div class="network-pattern"></div>
      <div class="cover__brand">
        <span>Observatorio da Industria</span>
        <strong>Sistema FIEA</strong>
      </div>
      <div class="cover__panel">
        <p class="eyebrow">Relatorio tecnico de inteligencia estrategica</p>
        <h1>${escapeHtml(title)}</h1>
        ${subtitle ? `<p class="cover__subtitle">${escapeHtml(subtitle)}</p>` : ""}
        ${renderCoverMeta(report)}
      </div>
      ${badge ? `<div class="cover__badge">${escapeHtml(badge)}</div>` : ""}
    </section>`;
}

function renderCoverMeta(report) {
  if (!report.metadata.length) {
    return `<dl class="cover__meta"><div><dt>Arquivo de origem</dt><dd>${escapeHtml(report.sourceName)}</dd></div></dl>`;
  }

  const items = report.metadata.slice(0, 4).map((entry) => `
    <div>
      <dt>${escapeHtml(entry.label)}</dt>
      <dd>${escapeHtml(entry.value)}</dd>
    </div>`);

  return `<dl class="cover__meta">${items.join("")}</dl>`;
}

function renderCredits(report) {
  if (!report.metadata.length) {
    return "";
  }

  const items = report.metadata.map((entry) => `
    <div class="credits__item">
      <dt>${escapeHtml(entry.label)}</dt>
      <dd>${escapeHtml(entry.value)}</dd>
    </div>`);

  return `
    <section class="credits page-break">
      <p class="eyebrow">Informacoes institucionais</p>
      <h2>Creditos institucionais</h2>
      <div class="rule"></div>
      <dl class="credits__grid">${items.join("")}</dl>
      ${report.warnings?.length ? renderWarnings(report.warnings) : ""}
    </section>`;
}

function renderWarnings(warnings) {
  const items = warnings.slice(0, 6).map((warning) => `<li>${escapeHtml(warning)}</li>`).join("");

  return `
    <aside class="conversion-notes">
      <strong>Notas de conversao</strong>
      <ul>${items}</ul>
    </aside>`;
}

function renderToc(report) {
  const headings = report.headings.filter((heading) => heading.level <= 3);

  if (!headings.length) {
    return "";
  }

  const items = headings.map((heading, index) => `
    <li class="toc__item toc__item--level-${heading.level}">
      <a href="#${escapeAttribute(heading.id)}">${escapeHtml(heading.text)}</a>
      <span class="toc__leader"></span>
      <span class="toc__number">${String(index + 1).padStart(2, "0")}</span>
    </li>`);

  return `
    <section class="toc page-break">
      <p class="eyebrow">Estrutura do documento</p>
      <h2>Sumario</h2>
      <div class="rule"></div>
      <ol>${items.join("")}</ol>
    </section>`;
}

function renderBackCover(report, title, subtitle) {
  const contactItems = [
    ...(report.contacts?.emails ?? []),
    ...(report.contacts?.websites ?? [])
  ];

  return `
    <section class="back-cover page-full">
      <div class="network-pattern"></div>
      <div class="back-cover__logo">
        <span>Observatorio da Industria</span>
        <strong>Sistema FIEA</strong>
      </div>
      <div class="back-cover__footer">
        <div>
          <h2>${escapeHtml(title)}</h2>
          ${subtitle ? `<p>${escapeHtml(subtitle)}</p>` : ""}
        </div>
        <address>
          ${contactItems.length ? contactItems.map((item) => `<span>${escapeHtml(item)}</span>`).join("") : "<span>Documento gerado automaticamente</span>"}
        </address>
      </div>
    </section>`;
}

function findMetadataValue(metadata, pattern) {
  const entry = metadata.find((item) => pattern.test(item.label) || pattern.test(item.value));

  return entry ? entry.value : "";
}

function normalize(value) {
  return normalizeWhitespace(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function normalizeWhitespace(value) {
  return value.replace(/\s+/g, " ").trim();
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/\s+/g, "-");
}

const REPORT_CSS = `
  :root {
    --primary-blue: #005A9C;
    --dark-navy: #003B6F;
    --medium-blue: #0B63B6;
    --light-blue: #D9EAF7;
    --very-light-blue: #EEF6FC;
    --white: #FFFFFF;
    --text: #2B2B2B;
    --secondary-text: #6B7280;
    --divider: #D1D5DB;
  }

  @page {
    size: A4;
    margin: 20mm 16mm 19mm;
  }

  * {
    box-sizing: border-box;
  }

  html,
  body {
    margin: 0;
    color: var(--text);
    font-family: Inter, Roboto, "Source Sans Pro", Arial, sans-serif;
    font-size: 10.5pt;
    line-height: 1.58;
    background: var(--white);
  }

  body {
    counter-reset: section;
  }

  h1,
  h2,
  h3,
  h4,
  h5,
  h6 {
    font-family: Montserrat, Poppins, Avenir Next, Inter, Arial, sans-serif;
    color: var(--dark-navy);
    line-height: 1.18;
    page-break-after: avoid;
  }

  h2 {
    margin: 0 0 6mm;
    font-size: 22pt;
    font-weight: 800;
  }

  h3 {
    margin: 9mm 0 3mm;
    color: var(--primary-blue);
    font-size: 15pt;
    font-weight: 800;
  }

  h4 {
    margin: 7mm 0 2mm;
    color: var(--medium-blue);
    font-size: 12pt;
  }

  p {
    margin: 0 0 4mm;
  }

  a {
    color: var(--primary-blue);
    text-decoration: none;
  }

  img {
    display: block;
    max-width: 100%;
    height: auto;
    margin: 5mm auto 2mm;
    border-radius: 10px;
  }

  blockquote {
    margin: 6mm 0;
    padding: 5mm 6mm;
    color: var(--dark-navy);
    background: var(--very-light-blue);
    border-left: 4px solid var(--primary-blue);
    border-radius: 8px;
    font-weight: 600;
  }

  .page-break {
    break-after: page;
    min-height: 238mm;
  }

  .page-full {
    break-after: page;
    min-height: 297mm;
    margin: -20mm -16mm -19mm;
    position: relative;
    overflow: hidden;
  }

  .eyebrow {
    margin: 0 0 3mm;
    color: var(--medium-blue);
    font-size: 8.5pt;
    font-weight: 800;
    letter-spacing: 0.12em;
    text-transform: uppercase;
  }

  .rule {
    width: 100%;
    height: 1px;
    margin: 0 0 10mm;
    background: linear-gradient(90deg, var(--primary-blue), var(--divider));
  }

  .network-pattern {
    position: absolute;
    inset: 0;
    opacity: 0.28;
    background-image:
      radial-gradient(circle at 18% 72%, rgba(255,255,255,0.9) 0 1.5px, transparent 1.8px),
      radial-gradient(circle at 34% 84%, rgba(255,255,255,0.8) 0 1.4px, transparent 1.7px),
      radial-gradient(circle at 66% 74%, rgba(255,255,255,0.75) 0 1.4px, transparent 1.7px),
      radial-gradient(circle at 80% 88%, rgba(255,255,255,0.85) 0 1.5px, transparent 1.8px),
      linear-gradient(34deg, transparent 24%, rgba(255,255,255,0.22) 24.2%, rgba(255,255,255,0.22) 24.7%, transparent 25%),
      linear-gradient(148deg, transparent 54%, rgba(255,255,255,0.16) 54.2%, rgba(255,255,255,0.16) 54.8%, transparent 55%);
  }

  .cover {
    color: var(--white);
    background:
      linear-gradient(145deg, rgba(0,59,111,0.92), rgba(0,90,156,0.78)),
      radial-gradient(circle at 82% 18%, rgba(217,234,247,0.38), transparent 30%),
      linear-gradient(120deg, #002b52, #005A9C 58%, #0B63B6);
  }

  .cover__gradient {
    position: absolute;
    inset: 0;
    background:
      linear-gradient(180deg, rgba(0,0,0,0.08), rgba(0,0,0,0.38)),
      repeating-linear-gradient(135deg, rgba(255,255,255,0.05) 0 1px, transparent 1px 16px);
  }

  .cover__brand,
  .back-cover__logo {
    position: absolute;
    top: 24mm;
    left: 24mm;
    display: flex;
    flex-direction: column;
    gap: 1mm;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }

  .cover__brand span,
  .back-cover__logo span {
    font-size: 9pt;
    opacity: 0.88;
  }

  .cover__brand strong,
  .back-cover__logo strong {
    font-size: 15pt;
  }

  .cover__panel {
    position: absolute;
    left: 24mm;
    right: 24mm;
    bottom: 36mm;
    padding: 13mm 14mm;
    background: rgba(0,59,111,0.72);
    border-left: 7px solid var(--light-blue);
    border-radius: 0 18px 18px 0;
    box-shadow: 0 16px 50px rgba(0,0,0,0.22);
  }

  .cover .eyebrow {
    color: var(--light-blue);
  }

  .cover h1 {
    margin: 0;
    color: var(--white);
    font-size: 33pt;
    font-weight: 900;
    letter-spacing: 0.02em;
    text-transform: uppercase;
  }

  .cover__subtitle {
    max-width: 150mm;
    margin: 5mm 0 0;
    color: rgba(255,255,255,0.9);
    font-size: 14pt;
    line-height: 1.35;
  }

  .cover__meta {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 4mm 9mm;
    margin: 10mm 0 0;
  }

  .cover__meta dt {
    color: var(--light-blue);
    font-size: 7.5pt;
    font-weight: 800;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }

  .cover__meta dd {
    margin: 1mm 0 0;
    color: var(--white);
    font-size: 9pt;
  }

  .cover__badge {
    position: absolute;
    right: 24mm;
    bottom: 18mm;
    padding: 2.5mm 5mm;
    color: var(--dark-navy);
    background: var(--light-blue);
    border-radius: 999px;
    font-weight: 800;
  }

  .credits,
  .toc {
    padding-top: 6mm;
  }

  .credits__grid {
    display: grid;
    grid-template-columns: repeat(2, minmax(0, 1fr));
    gap: 7mm 12mm;
  }

  .credits__item {
    padding-bottom: 4mm;
    border-bottom: 1px solid var(--divider);
  }

  .credits dt {
    color: var(--primary-blue);
    font-weight: 800;
  }

  .credits dd {
    margin: 1mm 0 0;
  }

  .conversion-notes {
    margin-top: 12mm;
    padding: 5mm;
    background: var(--very-light-blue);
    border-left: 4px solid var(--primary-blue);
    border-radius: 8px;
  }

  .conversion-notes ul {
    margin: 2mm 0 0;
    padding-left: 5mm;
  }

  .toc ol {
    margin: 0;
    padding: 0;
    list-style: none;
  }

  .toc__item {
    display: flex;
    align-items: baseline;
    gap: 3mm;
    padding: 2.3mm 0;
    border-bottom: 1px solid rgba(209,213,219,0.65);
  }

  .toc__item--level-2 {
    padding-left: 7mm;
    font-size: 9.8pt;
  }

  .toc__item--level-3 {
    padding-left: 14mm;
    color: var(--secondary-text);
    font-size: 9pt;
  }

  .toc__leader {
    flex: 1;
    border-bottom: 1px dotted var(--divider);
  }

  .toc__number {
    min-width: 11mm;
    color: var(--primary-blue);
    font-weight: 800;
    text-align: right;
  }

  .section-divider {
    break-before: page;
    break-after: page;
    min-height: 297mm;
    margin: -20mm -16mm -19mm;
    position: relative;
    overflow: hidden;
    color: var(--white);
    background:
      linear-gradient(145deg, rgba(0,59,111,0.95), rgba(0,90,156,0.78)),
      radial-gradient(circle at 12% 22%, rgba(217,234,247,0.34), transparent 30%),
      radial-gradient(circle at 85% 76%, rgba(11,99,182,0.45), transparent 28%),
      linear-gradient(120deg, #002b52, #005A9C);
  }

  .section-divider__content {
    position: absolute;
    left: 24mm;
    right: 24mm;
    bottom: 42mm;
  }

  .section-divider__content span {
    display: block;
    margin-bottom: 4mm;
    color: var(--light-blue);
    font-size: 9pt;
    font-weight: 800;
    letter-spacing: 0.14em;
    text-transform: uppercase;
  }

  .section-divider h1 {
    max-width: 150mm;
    margin: 0;
    color: var(--white);
    font-size: 32pt;
    font-weight: 900;
  }

  .section-divider__bar {
    width: 42mm;
    height: 2mm;
    margin-top: 7mm;
    background: var(--light-blue);
  }

  .report-content {
    break-before: page;
  }

  .report-content h2 {
    padding-bottom: 3mm;
    border-bottom: 1px solid var(--divider);
  }

  .report-content ul,
  .report-content ol {
    margin: 0 0 4mm;
    padding-left: 7mm;
  }

  .report-content li {
    margin-bottom: 1.6mm;
  }

  .image-card {
    margin: 7mm 0 2mm;
    padding: 4mm;
    background: var(--very-light-blue);
    border: 1px solid var(--divider);
    border-radius: 12px;
    page-break-inside: avoid;
  }

  .caption,
  .figure-caption,
  .source-note {
    color: var(--secondary-text);
    font-size: 8.3pt;
    line-height: 1.4;
  }

  .figure-caption {
    margin-top: 2mm;
    font-weight: 700;
    text-align: center;
  }

  .source-note {
    margin-top: -1mm;
    text-align: center;
  }

  .table-title {
    margin: 7mm 0 2mm;
    color: var(--dark-navy);
    font-weight: 800;
  }

  .table-shell {
    width: 100%;
    margin: 3mm 0 5mm;
    overflow: hidden;
    border: 1px solid var(--divider);
    border-radius: 10px;
    page-break-inside: avoid;
  }

  .institutional-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 8.8pt;
  }

  .institutional-table th {
    padding: 3mm 3.2mm;
    color: var(--white);
    background: var(--dark-navy);
    border: 1px solid var(--dark-navy);
    font-weight: 800;
    text-align: left;
  }

  .institutional-table td {
    padding: 2.7mm 3.2mm;
    border: 1px solid var(--divider);
    vertical-align: top;
  }

  .institutional-table tr:nth-child(even) td {
    background: var(--very-light-blue);
  }

  .callout {
    display: block;
    margin: 6mm 0;
    padding: 5mm 6mm;
    background: var(--very-light-blue);
    border-left: 5px solid var(--primary-blue);
    border-radius: 10px;
    page-break-inside: avoid;
  }

  .callout strong {
    display: block;
    margin-bottom: 1mm;
    color: var(--dark-navy);
    font-family: Montserrat, Poppins, Avenir Next, Inter, Arial, sans-serif;
  }

  .back-cover {
    color: var(--white);
    background:
      linear-gradient(135deg, #003B6F, #005A9C 58%, #0B63B6),
      var(--primary-blue);
  }

  .back-cover__logo {
    left: 50%;
    transform: translateX(-50%);
    align-items: center;
    text-align: center;
  }

  .back-cover__footer {
    position: absolute;
    left: 24mm;
    right: 24mm;
    bottom: 24mm;
    display: grid;
    grid-template-columns: 1.4fr 1fr;
    gap: 12mm;
    padding-top: 8mm;
    border-top: 1px solid rgba(255,255,255,0.75);
  }

  .back-cover h2 {
    margin: 0;
    color: var(--white);
    font-size: 18pt;
  }

  .back-cover p {
    margin-top: 3mm;
    color: rgba(255,255,255,0.82);
  }

  .back-cover address {
    display: flex;
    flex-direction: column;
    gap: 1.5mm;
    color: rgba(255,255,255,0.86);
    font-style: normal;
    text-align: right;
  }
`;
