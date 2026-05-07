export interface DocumentMeta {
  reportTitle: string
  subtitle?: string
  authors: string[]
  reviewers: string[]
  version?: string
  date?: string
  description?: string
  credits: string[]
  contacts: string[]
  catalogInfo: string[]
}

export interface HeadingEntry {
  text: string
  level: number
}

export type ContentBlock =
  | {
      type: 'heading'
      level: number
      text: string
    }
  | {
      type: 'paragraph'
      text: string
    }
  | {
      type: 'list'
      ordered: boolean
      items: string[]
    }
  | {
      type: 'table'
      rows: string[][]
      title?: string
      source?: string
    }
  | {
      type: 'figure'
      src: string
      alt?: string
      caption?: string
      source?: string
    }
  | {
      type: 'callout'
      title: string
      text: string
    }

export interface ParsedDocument {
  meta: DocumentMeta
  headings: HeadingEntry[]
  blocks: ContentBlock[]
}
