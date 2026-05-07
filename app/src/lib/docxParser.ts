import mammoth from 'mammoth'
import type { ContentBlock, DocumentMeta, ParsedDocument } from './types'

const CALLOUT_PREFIX =
  /^(Insight-chave|Recomendação|Nota importante|Oportunidade estratégica|Risco|Ponto de atenção)\s*[:\-–]\s*(.+)$/i

const TABLE_TITLE_PREFIX = /^Tabela\s+\d*/i
const FIGURE_CAPTION_PREFIX = /^Figura\s+\d*/i
const SOURCE_PREFIX = /^(Fonte|Source)\s*[:\-–]?\s*(.*)$/i

const AUTHOR_LABEL = /^Autores?\s*[:\-–]\s*(.+)$/i
const REVIEWER_LABEL = /^(Revis[aã]o|Revisores?)\s*[:\-–]\s*(.+)$/i
const VERSION_LABEL = /^Vers[aã]o\s*[:\-–]\s*(.+)$/i
const DATE_LABEL = /^Data\s*[:\-–]\s*(.+)$/i
const DESCRIPTION_LABEL = /^(Descri[cç][aã]o|Escopo)\s*[:\-–]\s*(.+)$/i

const CREDITS_HINT =
  /(realiza[cç][aã]o|execu[cç][aã]o|coordena[cç][aã]o|projeto gr[aá]fico|revis[aã]o|equipe t[eé]cnica)/i
const CATALOG_HINT = /(copyright|isbn|catalog|endere[cç]o|cep|alagoas|brasil)/i
const CONTACT_HINT = /([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}|https?:\/\/|www\.)/i

const normalizeText = (value: string | null | undefined) =>
  (value ?? '').replace(/\s+/g, ' ').trim()

const extractRows = (table: HTMLTableElement) => {
  const rows: string[][] = []
  table.querySelectorAll('tr').forEach((row) => {
    const cells = Array.from(row.querySelectorAll('th, td')).map((cell) =>
      normalizeText(cell.textContent),
    )
    if (cells.length > 0) {
      rows.push(cells)
    }
  })
  return rows
}

const parseElement = (node: Element, blocks: ContentBlock[]) => {
  const tag = node.tagName.toLowerCase()

  if (/^h[1-6]$/.test(tag)) {
    const level = Number.parseInt(tag[1], 10)
    const text = normalizeText(node.textContent)
    if (text) {
      blocks.push({ type: 'heading', level, text })
    }
    return
  }

  if (tag === 'table') {
    const rows = extractRows(node as HTMLTableElement)
    if (rows.length > 0) {
      blocks.push({ type: 'table', rows })
    }
    return
  }

  if (tag === 'ul' || tag === 'ol') {
    const items = Array.from(node.querySelectorAll(':scope > li'))
      .map((item) => normalizeText(item.textContent))
      .filter(Boolean)
    if (items.length > 0) {
      blocks.push({ type: 'list', ordered: tag === 'ol', items })
    }
    return
  }

  if (tag === 'img') {
    const img = node as HTMLImageElement
    if (img.src) {
      blocks.push({ type: 'figure', src: img.src, alt: img.alt || undefined })
    }
    return
  }

  if (tag === 'p' || tag === 'figure') {
    const images = Array.from(node.querySelectorAll('img'))
    images.forEach((img) => {
      if (img.src) {
        blocks.push({ type: 'figure', src: img.src, alt: img.alt || undefined })
      }
    })

    const textOnly = node.cloneNode(true) as HTMLElement
    textOnly.querySelectorAll('img').forEach((img) => img.remove())
    const text = normalizeText(textOnly.textContent)
    if (text) {
      blocks.push({ type: 'paragraph', text })
    }
    return
  }

  const childElements = Array.from(node.children)
  if (childElements.length > 0) {
    childElements.forEach((child) => parseElement(child, blocks))
    return
  }

  const fallbackText = normalizeText(node.textContent)
  if (fallbackText) {
    blocks.push({ type: 'paragraph', text: fallbackText })
  }
}

const enrichBlocks = (raw: ContentBlock[]) => {
  const normalized: ContentBlock[] = []

  for (let i = 0; i < raw.length; i += 1) {
    const block = raw[i]

    if (block.type === 'paragraph') {
      const calloutMatch = block.text.match(CALLOUT_PREFIX)
      if (calloutMatch) {
        normalized.push({
          type: 'callout',
          title: calloutMatch[1],
          text: calloutMatch[2],
        })
        continue
      }
    }

    if (block.type === 'table') {
      const table = { ...block }
      const prev = normalized.at(-1)
      if (prev?.type === 'paragraph' && TABLE_TITLE_PREFIX.test(prev.text)) {
        table.title = prev.text
        normalized.pop()
      }
      const next = raw[i + 1]
      if (next?.type === 'paragraph') {
        const sourceMatch = next.text.match(SOURCE_PREFIX)
        if (sourceMatch) {
          table.source = next.text
          i += 1
        }
      }
      normalized.push(table)
      continue
    }

    if (block.type === 'figure') {
      const figure = { ...block }
      const next = raw[i + 1]
      if (next?.type === 'paragraph' && FIGURE_CAPTION_PREFIX.test(next.text)) {
        figure.caption = next.text
        i += 1
      }
      const sourceLine = raw[i + 1]
      if (sourceLine?.type === 'paragraph') {
        const sourceMatch = sourceLine.text.match(SOURCE_PREFIX)
        if (sourceMatch) {
          figure.source = sourceLine.text
          i += 1
        }
      }
      normalized.push(figure)
      continue
    }

    normalized.push(block)
  }

  return normalized
}

const collectMetadata = (blocks: ContentBlock[], fileName: string): DocumentMeta => {
  const headingBlocks = blocks.filter(
    (block): block is Extract<ContentBlock, { type: 'heading' }> => block.type === 'heading',
  )
  const title =
    headingBlocks.find((block) => block.level === 1)?.text ||
    fileName.replace(/\.docx$/i, '')

  const introParagraphs = blocks
    .filter((block): block is Extract<ContentBlock, { type: 'paragraph' }> => block.type === 'paragraph')
    .slice(0, 28)
    .map((block) => block.text)

  const subtitleCandidate =
    headingBlocks.find((block) => block.level === 2)?.text ||
    introParagraphs.find(
      (text) =>
        text.length > 24 &&
        text.length < 140 &&
        !AUTHOR_LABEL.test(text) &&
        !VERSION_LABEL.test(text) &&
        !DATE_LABEL.test(text),
    )

  const authors = introParagraphs.flatMap((line) => {
    const match = line.match(AUTHOR_LABEL)
    if (!match) {
      return []
    }
    return match[1].split(/,|;/).map((name) => name.trim())
  })

  const reviewers = introParagraphs.flatMap((line) => {
    const match = line.match(REVIEWER_LABEL)
    if (!match) {
      return []
    }
    return match[2].split(/,|;/).map((name) => name.trim())
  })

  const version = introParagraphs.find((line) => VERSION_LABEL.test(line))?.match(VERSION_LABEL)?.[1]
  const date = introParagraphs.find((line) => DATE_LABEL.test(line))?.match(DATE_LABEL)?.[1]
  const description = introParagraphs
    .find((line) => DESCRIPTION_LABEL.test(line))
    ?.match(DESCRIPTION_LABEL)?.[2]

  const credits = introParagraphs.filter((line) => CREDITS_HINT.test(line))
  const contacts = introParagraphs.filter((line) => CONTACT_HINT.test(line))
  const catalogInfo = introParagraphs.filter((line) => CATALOG_HINT.test(line))

  return {
    reportTitle: title,
    subtitle: subtitleCandidate,
    authors,
    reviewers,
    version,
    date,
    description,
    credits,
    contacts,
    catalogInfo,
  }
}

export const parseDocxFile = async (file: File): Promise<ParsedDocument> => {
  const buffer = await file.arrayBuffer()
  const result = await mammoth.convertToHtml({ arrayBuffer: buffer })

  const parser = new DOMParser()
  const html = parser.parseFromString(result.value, 'text/html')

  const rawBlocks: ContentBlock[] = []
  Array.from(html.body.children).forEach((node) => parseElement(node, rawBlocks))

  const blocks = enrichBlocks(rawBlocks)
  const headings = blocks
    .filter((block): block is Extract<ContentBlock, { type: 'heading' }> => block.type === 'heading')
    .map((heading) => ({ text: heading.text, level: heading.level }))

  return {
    meta: collectMetadata(blocks, file.name),
    headings,
    blocks,
  }
}
