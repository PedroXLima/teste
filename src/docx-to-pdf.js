const path = require("path");

const cheerio = require("cheerio");
const mammoth = require("mammoth");
const { PDFDocument } = require("pdf-lib");
const { chromium } = require("playwright");

const COLORS = {
  primaryBlue: "#005A9C",
  darkNavy: "#003B6F",
  mediumBlue: "#0B63B6",
  lightBlue: "#D9EAF7",
  veryLightBlue: "#EEF6FC",
  white: "#FFFFFF",
  mainText: "#2B2B2B",
  secondaryText: "#6B7280",
  border: "#D1D5DB",
};

const CREDIT_SECTION_RE =
  /\b(creditos|expediente|realizacao|execucao|coordenacao|coordenador|autores|autoria|revisao|projeto grafico|equipe tecnica)\b/;
const CATALOG_SECTION_RE =
  /\b(ficha catalografica|catalogacao|copyright|contato|endereco|publicacao|catalogo)\b/;
const MAJOR_SECTION_RE =
  /\b(introducao|apresentacao|metodologia|metodo|desenvolvimento|analise|resultados|conclusao|conclusoes|referencias|anexos|anexo)\b/;
const SOURCE_RE = /^(fonte|source|nota|observacao)\s*[:\-]/i;
const FIGURE_CAPTION_RE = /^(figura|grafico|quadro|mapa|ilustracao)\s+\d+/i;
const TABLE_TITLE_RE = /^(tabela|quadro)\s+\d+/i;
const EMAIL_RE = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
const URL_RE = /\b(?:https?:\/\/|www\.)[^\s<]+/gi;

async function convertDocxToDesignedPdf({ buffer, originalName }) {
  const conversion = await mammoth.convertToHtml(
    { buffer },
    {
      styleMap: [
        "p[style-name='Title'] => h1:fresh",
        "p[style-name='Subtitle'] => h2:fresh",
        "p[style-name='Heading 1'] => h1:fresh",
        "p[style-name='Heading 2'] => h2:fresh",
        "p[style-name='Heading 3'] => h3:fresh",
        "p[style-name='Heading 4'] => h4:fresh",
        "table => table:fresh",
      ],
      convertImage: mammoth.images.inline(async (image) => {
        const imageBuffer = await image.read();
        return {
          src: `data:${image.contentType};base64,${imageBuffer.toString("base64")}`,
        };
      }),
    },
  );

  const model = buildDocumentModel({
    html: conversion.value,
    originalName,
    warnings: conversion.messages.map((message) => message.message),
  });

  const browser = await chromium.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  try {
    const coverPdf = await renderPdfBuffer(browser, buildCoverHtml(model), {
      margin: { top: "0mm", right: "0mm", bottom: "0mm", left: "0mm" },
      displayHeaderFooter: false,
    });

    const bodyPage = await browser.newPage();
    await bodyPage.emulateMedia({ media: "print" });
    await bodyPage.setContent(buildBodyHtml(model), { waitUntil: "load" });

    await bodyPage.evaluate(generateTocEntries);
    await bodyPage.evaluate(generateTocEntries);
    await bodyPage.waitForTimeout(100);

    const bodyPdf = await bodyPage.pdf({
      format: "A4",
      printBackground: true,
      displayHeaderFooter: true,
      preferCSSPageSize: true,
      margin: { top: "20mm", right: "16mm", bottom: "20mm", left: "16mm" },
      headerTemplate: buildHeaderTemplate(model.runningTitle),
      footerTemplate: buildFooterTemplate(),
    });
    await bodyPage.close();

    const backPdf = await renderPdfBuffer(browser, buildBackCoverHtml(model), {
      margin: { top: "0mm", right: "0mm", bottom: "0mm", left: "0mm" },
      displayHeaderFooter: false,
    });

    const mergedPdf = await PDFDocument.create();
    for (const bufferPart of [coverPdf, bodyPdf, backPdf]) {
      const source = await PDFDocument.load(bufferPart);
      const pages = await mergedPdf.copyPages(source, source.getPageIndices());
      pages.forEach((page) => mergedPdf.addPage(page));
    }

    const pdfBuffer = await mergedPdf.save();
    return {
      pdfBuffer,
      warnings: model.warnings,
      outputFilename: `${toFileSafeStem(originalName)}.pdf`,
    };
  } finally {
    await browser.close();
  }
}

function buildDocumentModel({ html, originalName, warnings }) {
  const wrappedHtml = `<main id="doc-root">${html}</main>`;
  const $ = cheerio.load(wrappedHtml, { decodeEntities: false });
  const root = $("#doc-root");

  root.children().each((index, node) => {
    $(node).attr("data-node-id", String(index));
  });

  const consumedIds = new Set();
  const titleBundle = extractTopMatter($, root, consumedIds, originalName);
  const creditsSection = extractSectionByHeading($, root, consumedIds, CREDIT_SECTION_RE);
  const catalogSection = extractSectionByHeading($, root, consumedIds, CATALOG_SECTION_RE);
  const remainingNodes = root
    .children()
    .toArray()
    .filter((node) => !consumedIds.has(nodeId($(node))));

  const rendered = renderContentNodes($, remainingNodes);
  const contactLines = collectContactLines($, root);

  return {
    title: titleBundle.title,
    subtitle: titleBundle.subtitle,
    metaLines: titleBundle.metaLines,
    runningTitle: shortenTitle(titleBundle.title),
    creditsSection,
    catalogSection,
    bodyHtml: rendered.html,
    tocEntries: rendered.tocEntries,
    contactLines,
    warnings,
    reportDate: titleBundle.badgeDate,
    reportVersion: titleBundle.badgeVersion,
  };
}

function extractTopMatter($, root, consumedIds, originalName) {
  const nodes = root.children().toArray();
  const defaultTitle = humanizeFilename(originalName);

  let title = "";
  let subtitle = "";
  let badgeDate = "";
  let badgeVersion = "";
  const metaLines = [];

  let titleIndex = -1;
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (!isRenderableNode($, node)) {
      continue;
    }
    const text = getNodeText($, node);
    if (!text) {
      continue;
    }
    if (isHeadingTag(node) || (node.name === "p" && text.length <= 180)) {
      title = text;
      titleIndex = index;
      consumeNode($(node), consumedIds);
      break;
    }
  }

  if (!title) {
    return {
      title: defaultTitle,
      subtitle: "",
      metaLines,
      badgeDate,
      badgeVersion,
    };
  }

  for (let index = titleIndex + 1; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (!isRenderableNode($, node) || consumedIds.has(nodeId($(node)))) {
      continue;
    }

    const text = getNodeText($, node);
    if (!text) {
      continue;
    }

    if (!subtitle && isSubtitleCandidate(node, text)) {
      subtitle = text;
      consumeNode($(node), consumedIds);
      continue;
    }

    if (isMetadataCandidate(node, text, metaLines.length)) {
      metaLines.push(text);
      consumeNode($(node), consumedIds);
      continue;
    }

    break;
  }

  for (const line of metaLines) {
    if (!badgeDate) {
      const dateMatch = line.match(
        /\b(\d{1,2}[\/.-]\d{1,2}[\/.-]\d{2,4}|\d{4}|\w+\s+de\s+\d{4}|[A-Za-z]+\s+\d{4})\b/i,
      );
      if (dateMatch) {
        badgeDate = dateMatch[1];
      }
    }

    if (!badgeVersion) {
      const versionMatch = line.match(/\b(v(?:ersao)?\.?\s*\d+(?:\.\d+)*)\b/i);
      if (versionMatch) {
        badgeVersion = versionMatch[1];
      }
    }
  }

  return {
    title,
    subtitle,
    metaLines,
    badgeDate,
    badgeVersion,
  };
}

function extractSectionByHeading($, root, consumedIds, matcher) {
  const nodes = root.children().toArray();
  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    if (consumedIds.has(nodeId($(node))) || !isHeadingTag(node)) {
      continue;
    }

    const headingText = getNodeText($, node);
    if (!matcher.test(normalizeForMatch(headingText))) {
      continue;
    }

    const level = headingLevel(node);
    const sectionNodes = [node];
    consumeNode($(node), consumedIds);

    for (let innerIndex = index + 1; innerIndex < nodes.length; innerIndex += 1) {
      const nextNode = nodes[innerIndex];
      const nextWrapped = $(nextNode);
      if (consumedIds.has(nodeId(nextWrapped))) {
        continue;
      }

      if (isHeadingTag(nextNode) && headingLevel(nextNode) <= level) {
        break;
      }

      consumeNode(nextWrapped, consumedIds);
      sectionNodes.push(nextNode);
    }

    const bodyNodes = sectionNodes.slice(1);
    return {
      title: headingText,
      html: bodyNodes.map((sectionNode) => $.html(sectionNode)).join("\n"),
      lines: bodyNodes
        .map((sectionNode) => getNodeText($, sectionNode))
        .filter(Boolean),
    };
  }

  return null;
}

function renderContentNodes($, nodes) {
  const htmlParts = [];
  const tocEntries = [];
  let headingCount = 0;

  for (let index = 0; index < nodes.length; index += 1) {
    const node = nodes[index];
    const text = getNodeText($, node);
    if (!isRenderableNode($, node)) {
      continue;
    }

    if (node.name === "p" && TABLE_TITLE_RE.test(text) && nodes[index + 1]?.name === "table") {
      const sourceText = SOURCE_RE.test(getNodeText($, nodes[index + 2] || {}))
        ? getNodeText($, nodes[index + 2])
        : "";

      htmlParts.push(`
        <section class="table-block">
          <div class="table-title">${escapeHtml(text)}</div>
          <div class="table-frame">${decorateTable($.html(nodes[index + 1]))}</div>
          ${sourceText ? `<div class="source-note">${escapeHtml(sourceText)}</div>` : ""}
        </section>
      `);

      index += sourceText ? 2 : 1;
      continue;
    }

    if (isImageParagraph(node)) {
      const captionText = FIGURE_CAPTION_RE.test(getNodeText($, nodes[index + 1] || {}))
        ? getNodeText($, nodes[index + 1])
        : "";
      const sourceText = SOURCE_RE.test(getNodeText($, nodes[index + 2] || {}))
        ? getNodeText($, nodes[index + 2])
        : "";

      htmlParts.push(`
        <figure class="figure-card">
          <div class="figure-media">${$(node).html() || ""}</div>
          ${captionText ? `<figcaption class="figure-caption">${escapeHtml(captionText)}</figcaption>` : ""}
          ${sourceText ? `<div class="source-note">${escapeHtml(sourceText)}</div>` : ""}
        </figure>
      `);

      index += sourceText ? 2 : captionText ? 1 : 0;
      continue;
    }

    if (isHeadingTag(node)) {
      headingCount += 1;
      const level = headingLevel(node);
      const headingId = `${slugify(text) || "secao"}-${headingCount}`;
      tocEntries.push({ id: headingId, level, text });

      if (level === 1 || MAJOR_SECTION_RE.test(normalizeForMatch(text))) {
        htmlParts.push(renderSectionDivider(headingId, text));
        continue;
      }

      htmlParts.push(renderHeading(node.name, headingId, $(node).html() || ""));
      continue;
    }

    if (node.name === "table") {
      htmlParts.push(`
        <section class="table-block">
          <div class="table-frame">${decorateTable($.html(node))}</div>
        </section>
      `);
      continue;
    }

    if (node.name === "ul" || node.name === "ol") {
      htmlParts.push(`<${node.name} class="doc-list">${$(node).html() || ""}</${node.name}>`);
      continue;
    }

    if (node.name === "blockquote") {
      htmlParts.push(`<aside class="callout"><div class="callout-title">Nota</div>${$(node).html() || ""}</aside>`);
      continue;
    }

    if (node.name === "p" && SOURCE_RE.test(text)) {
      htmlParts.push(`<div class="source-note">${escapeHtml(text)}</div>`);
      continue;
    }

    if (node.name === "p" && looksLikeCallout(text)) {
      const { label, body } = splitCallout(text);
      htmlParts.push(`
        <aside class="callout">
          <div class="callout-title">${escapeHtml(label)}</div>
          <p>${escapeHtml(body)}</p>
        </aside>
      `);
      continue;
    }

    htmlParts.push(renderGenericNode(node, $(node).html() || ""));
  }

  return {
    html: htmlParts.join("\n"),
    tocEntries,
  };
}

function renderHeading(tagName, headingId, innerHtml) {
  return `<${tagName} id="${headingId}" data-toc-source="true" class="doc-heading ${tagName}">${innerHtml}</${tagName}>`;
}

function renderGenericNode(node, innerHtml) {
  if (node.name === "p") {
    return `<p class="doc-paragraph">${innerHtml}</p>`;
  }

  if (node.name === "h4" || node.name === "h5" || node.name === "h6") {
    return `<${node.name} class="doc-heading ${node.name}">${innerHtml}</${node.name}>`;
  }

  return `<${node.name}>${innerHtml}</${node.name}>`;
}

function renderSectionDivider(headingId, title) {
  return `
    <section class="section-divider">
      <div class="section-divider-panel">
        <div class="section-divider-kicker">Secao estrategica</div>
        <h1 id="${headingId}" data-toc-source="true">${escapeHtml(title)}</h1>
        <div class="section-divider-bar"></div>
      </div>
    </section>
  `;
}

function buildCoverHtml(model) {
  const badge = [model.reportVersion, model.reportDate].filter(Boolean).join(" | ");
  const metaHtml = model.metaLines
    .slice(0, 5)
    .map((line) => `<div class="cover-meta-line">${escapeHtml(line)}</div>`)
    .join("");

  return `
    <!DOCTYPE html>
    <html lang="pt-BR">
      <head>
        <meta charset="UTF-8" />
        <title>${escapeHtml(model.title)}</title>
        <style>
          ${baseCss()}

          @page { size: A4; margin: 0; }

          body {
            margin: 0;
            background: ${COLORS.white};
          }

          .cover {
            position: relative;
            width: 210mm;
            height: 297mm;
            overflow: hidden;
            color: ${COLORS.white};
            background:
              radial-gradient(circle at top right, rgba(217, 234, 247, 0.25), transparent 30%),
              linear-gradient(145deg, ${COLORS.darkNavy} 0%, ${COLORS.primaryBlue} 55%, ${COLORS.mediumBlue} 100%);
          }

          .cover::before {
            content: "";
            position: absolute;
            inset: 0;
            background-image: url("${networkPatternDataUri("rgba(255,255,255,0.18)")}");
            background-size: cover;
            opacity: 0.9;
          }

          .cover::after {
            content: "";
            position: absolute;
            inset: auto 0 0 0;
            height: 42%;
            background: linear-gradient(180deg, rgba(0, 59, 111, 0) 0%, rgba(0, 59, 111, 0.85) 30%, rgba(0, 59, 111, 0.95) 100%);
          }

          .cover-shell {
            position: relative;
            z-index: 1;
            display: flex;
            flex-direction: column;
            justify-content: flex-end;
            height: 100%;
            padding: 24mm 18mm 24mm;
          }

          .cover-panel {
            max-width: 160mm;
            padding: 18mm 16mm;
            border-radius: 18px 18px 0 0;
            background: linear-gradient(180deg, rgba(11, 99, 182, 0.22), rgba(0, 59, 111, 0.78));
            border: 1px solid rgba(255, 255, 255, 0.18);
            backdrop-filter: blur(6px);
          }

          .cover-kicker {
            margin-bottom: 8mm;
            font-size: 11pt;
            font-weight: 700;
            letter-spacing: 0.22em;
            text-transform: uppercase;
            color: rgba(255, 255, 255, 0.78);
          }

          .cover h1 {
            margin: 0;
            font-size: 27pt;
            line-height: 1.15;
            font-weight: 800;
            text-transform: uppercase;
          }

          .cover h2 {
            margin: 5mm 0 0;
            max-width: 135mm;
            font-size: 13pt;
            line-height: 1.45;
            font-weight: 500;
            color: rgba(255, 255, 255, 0.9);
          }

          .cover-meta {
            margin-top: 10mm;
            display: grid;
            gap: 2mm;
            max-width: 132mm;
            font-size: 9.5pt;
            color: rgba(255, 255, 255, 0.82);
          }

          .version-badge {
            position: absolute;
            z-index: 2;
            right: 18mm;
            bottom: 14mm;
            padding: 3mm 4.5mm;
            border-radius: 999px;
            border: 1px solid rgba(255, 255, 255, 0.24);
            background: rgba(255, 255, 255, 0.12);
            font-size: 9pt;
            letter-spacing: 0.08em;
            text-transform: uppercase;
          }
        </style>
      </head>
      <body>
        <section class="cover">
          <div class="cover-shell">
            <div class="cover-panel">
              <div class="cover-kicker">Relatorio tecnico institucional</div>
              <h1>${escapeHtml(model.title)}</h1>
              ${model.subtitle ? `<h2>${escapeHtml(model.subtitle)}</h2>` : ""}
              ${metaHtml ? `<div class="cover-meta">${metaHtml}</div>` : ""}
            </div>
          </div>
          ${badge ? `<div class="version-badge">${escapeHtml(badge)}</div>` : ""}
        </section>
      </body>
    </html>
  `;
}

function buildBodyHtml(model) {
  const frontMatter = [
    renderCreditsPage(model.creditsSection),
    renderCatalogPage(model.catalogSection),
    renderTocPage(model.title),
  ]
    .filter(Boolean)
    .join("\n");

  return `
    <!DOCTYPE html>
    <html lang="pt-BR">
      <head>
        <meta charset="UTF-8" />
        <title>${escapeHtml(model.title)}</title>
        <style>
          ${baseCss()}

          @page {
            size: A4;
            margin: 20mm 16mm 20mm 16mm;
          }

          :root {
            --page-content-height: 257mm;
          }

          html, body {
            margin: 0;
            padding: 0;
            background: ${COLORS.white};
          }

          body {
            font-size: 11pt;
            line-height: 1.65;
          }

          .body-document {
            width: 100%;
          }

          .front-page,
          .section-divider {
            position: relative;
            min-height: var(--page-content-height);
            box-sizing: border-box;
            break-after: page;
          }

          .front-page {
            padding: 2mm 0;
            display: flex;
          }

          .front-page .page-shell {
            width: 100%;
            display: flex;
            flex-direction: column;
          }

          .front-page .page-kicker {
            margin-bottom: 4mm;
            font-size: 9pt;
            font-weight: 700;
            letter-spacing: 0.18em;
            text-transform: uppercase;
            color: ${COLORS.mediumBlue};
          }

          .front-page h1 {
            margin: 0 0 4mm;
            color: ${COLORS.primaryBlue};
            font-size: 24pt;
            line-height: 1.15;
            font-weight: 800;
          }

          .front-page .divider {
            width: 48mm;
            height: 1.5mm;
            margin-bottom: 8mm;
            background: ${COLORS.primaryBlue};
            border-radius: 999px;
          }

          .credit-grid,
          .catalog-grid {
            display: grid;
            gap: 5mm;
          }

          .credit-item,
          .catalog-item {
            padding-bottom: 4mm;
            border-bottom: 1px solid ${COLORS.border};
          }

          .credit-label,
          .catalog-label {
            display: block;
            margin-bottom: 1.2mm;
            color: ${COLORS.darkNavy};
            font-size: 10pt;
            font-weight: 700;
          }

          .credit-value,
          .catalog-value {
            color: ${COLORS.mainText};
          }

          .toc-entries {
            display: grid;
            gap: 3mm;
            margin-top: 2mm;
          }

          .toc-entry {
            display: grid;
            grid-template-columns: auto 1fr auto;
            gap: 3mm;
            align-items: end;
            font-size: 10.5pt;
          }

          .toc-entry .toc-title {
            color: ${COLORS.mainText};
          }

          .toc-entry .toc-line {
            border-bottom: 1px dotted ${COLORS.border};
            transform: translateY(-1.5mm);
          }

          .toc-entry .toc-page {
            color: ${COLORS.secondaryText};
            font-weight: 600;
          }

          .toc-entry.level-2 { padding-left: 5mm; }
          .toc-entry.level-3 { padding-left: 10mm; }

          .section-divider {
            display: flex;
            align-items: flex-end;
            padding: 0;
            overflow: hidden;
            background:
              linear-gradient(180deg, rgba(0, 59, 111, 0.1), rgba(0, 59, 111, 0.82)),
              linear-gradient(135deg, ${COLORS.darkNavy} 0%, ${COLORS.primaryBlue} 60%, ${COLORS.mediumBlue} 100%);
            border-radius: 18px;
          }

          .section-divider::before {
            content: "";
            position: absolute;
            inset: 0;
            background-image: url("${networkPatternDataUri("rgba(255,255,255,0.16)")}");
            background-size: cover;
            opacity: 0.8;
          }

          .section-divider-panel {
            position: relative;
            z-index: 1;
            width: 100%;
            padding: 0 14mm 16mm;
            color: ${COLORS.white};
          }

          .section-divider-kicker {
            margin-bottom: 4mm;
            font-size: 9pt;
            letter-spacing: 0.18em;
            text-transform: uppercase;
            color: rgba(255, 255, 255, 0.75);
          }

          .section-divider h1 {
            margin: 0;
            font-size: 26pt;
            line-height: 1.15;
            font-weight: 800;
          }

          .section-divider-bar {
            margin-top: 6mm;
            width: 38mm;
            height: 2mm;
            border-radius: 999px;
            background: ${COLORS.lightBlue};
          }

          .section-anchor {
            position: absolute;
            inset: 0 auto auto 0;
          }

          .doc-heading {
            color: ${COLORS.primaryBlue};
            font-weight: 800;
            break-after: avoid-page;
          }

          h2.doc-heading {
            margin: 9mm 0 4mm;
            font-size: 18pt;
            padding-bottom: 2.5mm;
            border-bottom: 1px solid ${COLORS.border};
          }

          h3.doc-heading {
            margin: 7mm 0 3mm;
            font-size: 14pt;
          }

          h4.doc-heading,
          h5.doc-heading,
          h6.doc-heading {
            margin: 5mm 0 2mm;
            font-size: 12pt;
            color: ${COLORS.darkNavy};
          }

          .doc-paragraph {
            margin: 0 0 4mm;
            text-align: justify;
          }

          .doc-list {
            margin: 0 0 5mm 0;
            padding-left: 6mm;
          }

          .doc-list li {
            margin-bottom: 2mm;
          }

          .figure-card,
          .table-block,
          .callout {
            break-inside: avoid-page;
          }

          .figure-card {
            margin: 0 0 8mm;
            padding: 4mm;
            border: 1px solid rgba(11, 99, 182, 0.12);
            border-radius: 16px;
            background: linear-gradient(180deg, rgba(238, 246, 252, 0.6), rgba(255, 255, 255, 1));
          }

          .figure-media img {
            display: block;
            width: 100%;
            height: auto;
            border-radius: 12px;
          }

          .figure-caption,
          .source-note {
            margin-top: 3mm;
            font-size: 9pt;
            line-height: 1.45;
            color: ${COLORS.secondaryText};
          }

          .table-title {
            margin-bottom: 2.5mm;
            color: ${COLORS.darkNavy};
            font-size: 10.5pt;
            font-weight: 700;
          }

          .table-frame {
            overflow: hidden;
            border: 1px solid ${COLORS.border};
            border-radius: 14px;
          }

          table {
            width: 100%;
            border-collapse: collapse;
            table-layout: auto;
          }

          th,
          td {
            padding: 3.2mm;
            border: 1px solid ${COLORS.border};
            vertical-align: top;
            font-size: 9.8pt;
          }

          thead th {
            background: ${COLORS.darkNavy};
            color: ${COLORS.white};
            font-weight: 700;
            text-align: left;
          }

          tbody tr:nth-child(even) td {
            background: ${COLORS.veryLightBlue};
          }

          .callout {
            margin: 0 0 7mm;
            padding: 4.5mm 5mm;
            border-left: 4px solid ${COLORS.darkNavy};
            border-radius: 12px;
            background: ${COLORS.veryLightBlue};
          }

          .callout-title {
            margin-bottom: 1.5mm;
            color: ${COLORS.primaryBlue};
            font-size: 10pt;
            font-weight: 800;
          }

          .callout p {
            margin: 0;
          }
        </style>
      </head>
      <body>
        <main class="body-document">
          ${frontMatter}
          ${model.bodyHtml}
        </main>
      </body>
    </html>
  `;
}

function buildBackCoverHtml(model) {
  const contactBlock =
    model.contactLines.length > 0
      ? model.contactLines
          .slice(0, 4)
          .map((line) => `<div>${escapeHtml(line)}</div>`)
          .join("")
      : `<div>${escapeHtml(model.runningTitle)}</div>`;

  return `
    <!DOCTYPE html>
    <html lang="pt-BR">
      <head>
        <meta charset="UTF-8" />
        <title>${escapeHtml(model.title)}</title>
        <style>
          ${baseCss()}

          @page { size: A4; margin: 0; }

          body {
            margin: 0;
          }

          .back-cover {
            position: relative;
            width: 210mm;
            height: 297mm;
            overflow: hidden;
            color: ${COLORS.white};
            background: linear-gradient(145deg, ${COLORS.darkNavy} 0%, ${COLORS.primaryBlue} 100%);
          }

          .back-cover::before {
            content: "";
            position: absolute;
            inset: 0;
            background-image: url("${networkPatternDataUri("rgba(255,255,255,0.14)")}");
            background-size: cover;
            opacity: 0.9;
          }

          .back-cover-shell {
            position: relative;
            z-index: 1;
            display: flex;
            flex-direction: column;
            justify-content: flex-end;
            height: 100%;
            padding: 22mm 18mm;
            box-sizing: border-box;
          }

          .back-divider {
            width: 100%;
            height: 1px;
            margin-bottom: 10mm;
            background: rgba(255, 255, 255, 0.7);
          }

          .back-grid {
            display: grid;
            grid-template-columns: 1fr auto;
            gap: 8mm;
            align-items: end;
          }

          .back-title {
            max-width: 110mm;
          }

          .back-title h2 {
            margin: 0 0 2mm;
            font-size: 15pt;
            line-height: 1.3;
            font-weight: 800;
          }

          .back-title p,
          .back-contact {
            margin: 0;
            font-size: 9.5pt;
            line-height: 1.6;
            color: rgba(255, 255, 255, 0.84);
          }

          .back-contact {
            text-align: right;
          }
        </style>
      </head>
      <body>
        <section class="back-cover">
          <div class="back-cover-shell">
            <div class="back-divider"></div>
            <div class="back-grid">
              <div class="back-title">
                <h2>${escapeHtml(model.title)}</h2>
                ${model.subtitle ? `<p>${escapeHtml(model.subtitle)}</p>` : ""}
              </div>
              <div class="back-contact">${contactBlock}</div>
            </div>
          </div>
        </section>
      </body>
    </html>
  `;
}

function renderCreditsPage(section) {
  if (!section) {
    return "";
  }

  const items = renderLabeledRows(section.lines, "credit");

  return `
    <section class="front-page">
      <div class="page-shell">
        <div class="page-kicker">Informacoes institucionais</div>
        <h1>${escapeHtml(section.title)}</h1>
        <div class="divider"></div>
        <div class="credit-grid">${items}</div>
      </div>
    </section>
  `;
}

function renderCatalogPage(section) {
  if (!section) {
    return "";
  }

  const items = renderLabeledRows(section.lines, "catalog");

  return `
    <section class="front-page">
      <div class="page-shell">
        <div class="page-kicker">Catalogacao e contato</div>
        <h1>${escapeHtml(section.title)}</h1>
        <div class="divider"></div>
        <div class="catalog-grid">${items}</div>
      </div>
    </section>
  `;
}

function renderTocPage(title) {
  return `
    <section class="front-page">
      <div class="page-shell">
        <div class="page-kicker">Navegacao do relatorio</div>
        <h1>Sumario</h1>
        <div class="divider"></div>
        <p class="doc-paragraph">
          Estrutura principal do documento "${escapeHtml(title)}" com indicacao de paginas.
        </p>
        <div class="toc-entries" data-toc-container="true"></div>
      </div>
    </section>
  `;
}

function renderLabeledRows(lines, prefix) {
  const normalizedLines = lines.filter(Boolean);
  if (normalizedLines.length === 0) {
    return `<div class="${prefix}-item"><span class="${prefix}-value">Sem informacoes adicionais.</span></div>`;
  }

  return normalizedLines
    .map((line) => {
      const match = line.match(/^([^:]{2,60}):\s*(.+)$/);
      if (match) {
        return `
          <div class="${prefix}-item">
            <span class="${prefix}-label">${escapeHtml(match[1])}</span>
            <span class="${prefix}-value">${escapeHtml(match[2])}</span>
          </div>
        `;
      }

      return `
        <div class="${prefix}-item">
          <span class="${prefix}-value">${escapeHtml(line)}</span>
        </div>
      `;
    })
    .join("\n");
}

function buildHeaderTemplate(title) {
  return `
    <div style="width:100%; padding:0 16mm; box-sizing:border-box; font-family:Arial, sans-serif; color:${COLORS.darkNavy};">
      <div style="border-top:4px solid ${COLORS.primaryBlue}; padding-top:6px; display:flex; justify-content:space-between; align-items:center; font-size:9px; letter-spacing:0.08em; text-transform:uppercase;">
        <span>${escapeHtml(title)}</span>
        <span class="pageNumber"></span>
      </div>
    </div>
  `;
}

function buildFooterTemplate() {
  return `
    <div style="width:100%; padding:0 16mm 4mm; box-sizing:border-box; font-family:Arial, sans-serif; color:${COLORS.secondaryText};">
      <div style="border-top:1px solid ${COLORS.border}; padding-top:5px; font-size:8px; display:flex; justify-content:space-between; align-items:center;">
        <span>Relatorio institucional automatizado</span>
        <span><span class="pageNumber"></span> / <span class="totalPages"></span></span>
      </div>
    </div>
  `;
}

async function renderPdfBuffer(browser, html, options) {
  const page = await browser.newPage();
  try {
    await page.emulateMedia({ media: "print" });
    await page.setContent(html, { waitUntil: "load" });
    await page.waitForTimeout(50);
    return await page.pdf({
      format: "A4",
      printBackground: true,
      preferCSSPageSize: true,
      ...options,
    });
  } finally {
    await page.close();
  }
}

function generateTocEntries() {
  const escapeHtml = (value) =>
    String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");

  const container = document.querySelector("[data-toc-container='true']");
  if (!container) {
    return;
  }

  const a4ContentHeightMm = 257;
  const contentHeightPx = (a4ContentHeightMm * 96) / 25.4;
  const headings = Array.from(document.querySelectorAll("[data-toc-source='true']"))
    .map((element) => {
      const rawTitle = (element.textContent || "").replace(/\s+/g, " ").trim();
      if (!rawTitle) {
        return null;
      }

      const headingElement = element.matches("h2, h3, h4, h5, h6") ? element : element.closest(".section-divider, h1, h2, h3, h4, h5, h6") || element;
      const level = headingElement.tagName && /^H\d$/.test(headingElement.tagName)
        ? Number(headingElement.tagName.slice(1))
        : 1;
      const offsetTop = headingElement.getBoundingClientRect().top + window.scrollY;
      const pageNumber = Math.max(1, Math.floor(offsetTop / contentHeightPx) + 1);

      return {
        title: rawTitle,
        level,
        pageNumber,
      };
    })
    .filter(Boolean);

  container.innerHTML = headings
    .map(
      (heading) => `
        <div class="toc-entry level-${Math.min(heading.level, 3)}">
          <div class="toc-title">${escapeHtml(heading.title)}</div>
          <div class="toc-line"></div>
          <div class="toc-page">${heading.pageNumber}</div>
        </div>
      `,
    )
    .join("");
}

function decorateTable(tableHtml) {
  const $ = cheerio.load(tableHtml, { decodeEntities: false });
  const table = $("table");

  if (table.find("thead").length === 0) {
    const firstRow = table.find("tr").first();
    if (firstRow.length > 0) {
      firstRow.wrap("<thead></thead>");
      firstRow.find("td").each((_, cell) => {
        $(cell).replaceWith(`<th>${$(cell).html() || ""}</th>`);
      });
    }
  }

  return $.html(table);
}

function collectContactLines($, root) {
  const textPool = [];
  root.find("p, li").each((_, node) => {
    const text = getNodeText($, node);
    if (text) {
      textPool.push(text);
    }
  });

  const unique = new Set();
  for (const text of textPool) {
    const emails = text.match(EMAIL_RE) || [];
    const links = text.match(URL_RE) || [];
    const current = [...emails, ...links];
    if (current.length > 0) {
      current.forEach((item) => unique.add(item));
    }
  }

  return [...unique];
}

function looksLikeCallout(text) {
  const normalized = normalizeForMatch(text);
  return /^(insight|recomendacao|nota importante|oportunidade estrategica|risco|atencao)\b/.test(normalized);
}

function splitCallout(text) {
  const match = text.match(/^([^:]{2,60})[:\-]\s*(.+)$/);
  if (!match) {
    return { label: "Destaque", body: text };
  }

  return {
    label: match[1],
    body: match[2],
  };
}

function isImageParagraph(node) {
  return (
    node?.name === "p" &&
    Array.isArray(node.children) &&
    node.children.some((child) => child?.name === "img")
  );
}

function isRenderableNode($, node) {
  if (!node || !node.name) {
    return false;
  }

  if (["style", "script"].includes(node.name)) {
    return false;
  }

  if (node.name === "table") {
    return true;
  }

  const text = getNodeText($, node);
  return Boolean(text || $(node).find("img").length > 0);
}

function isHeadingTag(node) {
  return Boolean(node?.name && /^h[1-6]$/i.test(node.name));
}

function headingLevel(node) {
  return Number(node.name.slice(1));
}

function getNodeText($, node) {
  if (!node || !node.name) {
    return "";
  }

  return $(node).text().replace(/\s+/g, " ").trim();
}

function isSubtitleCandidate(node, text) {
  if (node.name === "h2" || node.name === "h3") {
    return text.length <= 220;
  }

  return node.name === "p" && text.length > 20 && text.length <= 220;
}

function isMetadataCandidate(node, text, count) {
  if (count >= 5 || node.name !== "p") {
    return false;
  }

  if (text.length > 180) {
    return false;
  }

  return /[:]|(versao|version|data|date|autor|revis|coordena|instituicao)/i.test(normalizeForMatch(text));
}

function consumeNode(wrappedNode, consumedIds) {
  consumedIds.add(nodeId(wrappedNode));
}

function nodeId(wrappedNode) {
  return wrappedNode.attr("data-node-id") || "";
}

function normalizeForMatch(text) {
  return (text || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function humanizeFilename(originalName) {
  return path
    .basename(originalName, path.extname(originalName))
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function toFileSafeStem(originalName) {
  return humanizeFilename(originalName).replace(/[^\w.-]+/g, "_");
}

function shortenTitle(title) {
  const compact = (title || "").replace(/\s+/g, " ").trim();
  if (compact.length <= 70) {
    return compact;
  }
  return `${compact.slice(0, 67).trim()}...`;
}

function slugify(value) {
  return normalizeForMatch(value)
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function escapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function networkPatternDataUri(strokeColor) {
  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 1200">
      <g fill="none" stroke="${strokeColor}" stroke-width="1.5">
        <path d="M70 890 L260 720 L520 860 L810 580 L1120 760" />
        <path d="M90 650 L250 520 L460 610 L690 390 L930 470 L1110 330" />
        <path d="M180 1020 L380 840 L640 980 L890 760 L1090 940" />
        <circle cx="260" cy="720" r="6" fill="${strokeColor}" />
        <circle cx="520" cy="860" r="6" fill="${strokeColor}" />
        <circle cx="810" cy="580" r="6" fill="${strokeColor}" />
        <circle cx="250" cy="520" r="6" fill="${strokeColor}" />
        <circle cx="690" cy="390" r="6" fill="${strokeColor}" />
        <circle cx="930" cy="470" r="6" fill="${strokeColor}" />
        <circle cx="380" cy="840" r="6" fill="${strokeColor}" />
        <circle cx="640" cy="980" r="6" fill="${strokeColor}" />
        <circle cx="890" cy="760" r="6" fill="${strokeColor}" />
      </g>
    </svg>
  `;

  return `data:image/svg+xml,${encodeURIComponent(svg)}`;
}

function baseCss() {
  return `
    * { box-sizing: border-box; }

    html, body {
      font-family: Inter, "Segoe UI", Roboto, Arial, sans-serif;
      color: ${COLORS.mainText};
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
  `;
}

module.exports = {
  convertDocxToDesignedPdf,
};
