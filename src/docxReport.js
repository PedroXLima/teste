import mammoth from "mammoth";
import * as cheerio from "cheerio";

const STYLE_MAP = [
  "p[style-name='Title'] => h1:fresh",
  "p[style-name='Título'] => h1:fresh",
  "p[style-name='Subtitle'] => p.subtitle:fresh",
  "p[style-name='Subtítulo'] => p.subtitle:fresh",
  "p[style-name='Heading 1'] => h1:fresh",
  "p[style-name='Título 1'] => h1:fresh",
  "p[style-name='Heading 2'] => h2:fresh",
  "p[style-name='Título 2'] => h2:fresh",
  "p[style-name='Heading 3'] => h3:fresh",
  "p[style-name='Título 3'] => h3:fresh",
  "p[style-name='Heading 4'] => h4:fresh",
  "p[style-name='Título 4'] => h4:fresh",
  "p[style-name='Caption'] => p.caption:fresh",
  "p[style-name='Legenda'] => p.caption:fresh",
  "p[style-name='Quote'] => blockquote:fresh",
  "p[style-name='Citação'] => blockquote:fresh"
];

const METADATA_LABELS = [
  "realização",
  "realizacao",
  "execução",
  "execucao",
  "coordenação",
  "coordenacao",
  "autores",
  "autor",
  "revisão",
  "revisao",
  "versão",
  "versao",
  "data",
  "projeto gráfico",
  "projeto grafico",
  "catalogação",
  "catalogacao",
  "copyright"
];

export async function convertDocxBuffer(buffer, originalName = "documento.docx") {
  const result = await mammoth.convertToHtml(
    { buffer },
    {
      styleMap: STYLE_MAP,
      includeDefaultStyleMap: true,
      convertImage: mammoth.images.imgElement(async (image) => {
        const imageBuffer = await image.read("base64");

        return {
          src: `data:${image.contentType};base64,${imageBuffer}`
        };
      })
    }
  );

  return analyzeDocument(result.value, {
    originalName,
    messages: result.messages ?? []
  });
}

export function analyzeDocument(html, { originalName = "documento.docx", messages = [] } = {}) {
  const $ = cheerio.load(`<main>${html}</main>`, { decodeEntities: false });
  const $main = $("main");

  $("script, style, iframe, object").remove();
  removeEmptyParagraphs($);
  decorateSemanticElements($);
  assignHeadingIds($);

  const title = inferTitle($, originalName);
  const subtitle = inferSubtitle($, title);
  const metadata = extractMetadata($);
  const headings = extractHeadings($);
  const contacts = extractContacts($);
  const filenameBase = stripExtension(originalName);

  return {
    title,
    subtitle,
    metadata,
    headings,
    contacts,
    sourceName: originalName,
    filenameBase,
    bodyHtml: $main.html() ?? "",
    warnings: messages.map((message) => message.message).filter(Boolean)
  };
}

function removeEmptyParagraphs($) {
  $("p").each((_, element) => {
    const text = normalizeWhitespace($(element).text());
    const hasImage = $(element).find("img").length > 0;

    if (!text && !hasImage) {
      $(element).remove();
    }
  });
}

function decorateSemanticElements($) {
  $("p").each((_, element) => {
    const $element = $(element);
    const text = normalizeWhitespace($element.text());

    if (/^(figura|figure)\s+\d+\s*[-–:]/i.test(text)) {
      $element.addClass("figure-caption");
    }

    if (/^(tabela|table|quadro)\s+\d+\s*[-–:]/i.test(text)) {
      $element.addClass("table-title");
    }

    if (/^(fonte|source|nota|observação|observacao)\s*:/i.test(text)) {
      $element.addClass("source-note");
    }

    if (/^(insight|recomendação|recomendacao|atenção|atencao|risco|oportunidade)\s*:/i.test(text)) {
      $element.addClass("callout-source");
    }
  });

  $("table").each((index, element) => {
    const $table = $(element);
    $table.addClass("institutional-table");
    $table.attr("data-table-index", String(index + 1));

    const firstRowCells = $table.find("tr").first().children("td, th");
    firstRowCells.each((_, cell) => {
      if (cell.tagName !== "th") {
        $(cell).replaceWith(`<th>${$(cell).html() ?? ""}</th>`);
      }
    });
  });
}

function assignHeadingIds($) {
  const seen = new Map();

  $("h1, h2, h3, h4, h5, h6").each((index, element) => {
    const $element = $(element);
    const base = slugify($element.text()) || `secao-${index + 1}`;
    const count = seen.get(base) ?? 0;
    seen.set(base, count + 1);
    $element.attr("id", count === 0 ? base : `${base}-${count + 1}`);
  });
}

function inferTitle($, originalName) {
  const firstHeading = normalizeWhitespace($("h1").first().text());

  if (firstHeading) {
    return firstHeading;
  }

  const firstStrongParagraph = $("p")
    .toArray()
    .map((element) => ({
      text: normalizeWhitespace($(element).text()),
      hasStrong: $(element).find("strong, b").length > 0
    }))
    .find((paragraph) => paragraph.hasStrong && paragraph.text.length >= 8);

  if (firstStrongParagraph) {
    return firstStrongParagraph.text;
  }

  const firstParagraph = $("p")
    .toArray()
    .map((element) => normalizeWhitespace($(element).text()))
    .find((text) => text.length >= 8);

  return firstParagraph || stripExtension(originalName);
}

function inferSubtitle($, title) {
  const explicitSubtitle = normalizeWhitespace($(".subtitle").first().text());

  if (explicitSubtitle && explicitSubtitle !== title) {
    return explicitSubtitle;
  }

  const candidates = $("h1, p")
    .toArray()
    .map((element) => normalizeWhitespace($(element).text()))
    .filter((text) => text && text !== title);

  return candidates.find((text) => text.length >= 12 && text.length <= 180) ?? "";
}

function extractMetadata($) {
  const metadata = [];

  $("p, li").slice(0, 40).each((_, element) => {
    const text = normalizeWhitespace($(element).text());
    const lower = removeAccents(text.toLowerCase());

    if (!text || text.length > 220) {
      return;
    }

    const matchesLabel = METADATA_LABELS.some((label) => lower.includes(removeAccents(label)));
    const hasLabelSyntax = /^[A-Za-zÀ-ÿ\s/.-]{3,35}:\s+/.test(text);

    if (matchesLabel || hasLabelSyntax) {
      const [label, ...rest] = text.split(":");
      metadata.push({
        label: rest.length ? label.trim() : "Informação",
        value: rest.length ? rest.join(":").trim() : text
      });
    }
  });

  return dedupeMetadata(metadata).slice(0, 14);
}

function extractHeadings($) {
  return $("h1, h2, h3")
    .toArray()
    .map((element) => ({
      id: $(element).attr("id") ?? "",
      level: Number(element.tagName.replace("h", "")),
      text: normalizeWhitespace($(element).text())
    }))
    .filter((heading) => heading.text);
}

function extractContacts($) {
  const bodyText = $("main").text();
  const emails = Array.from(new Set(bodyText.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi) ?? []));
  const websites = Array.from(new Set(bodyText.match(/https?:\/\/[^\s)]+|www\.[^\s)]+/gi) ?? []));

  return { emails: emails.slice(0, 4), websites: websites.slice(0, 4) };
}

function dedupeMetadata(metadata) {
  const seen = new Set();

  return metadata.filter((entry) => {
    const key = `${entry.label}:${entry.value}`.toLowerCase();

    if (seen.has(key)) {
      return false;
    }

    seen.add(key);
    return true;
  });
}

export function stripExtension(filename) {
  return filename.replace(/\.[^.]+$/, "");
}

export function slugify(value) {
  return removeAccents(value)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 80);
}

function normalizeWhitespace(value) {
  return value.replace(/\s+/g, " ").trim();
}

function removeAccents(value) {
  return value.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}
