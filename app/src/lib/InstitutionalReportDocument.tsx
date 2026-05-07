import {
  Document,
  Image,
  Page,
  StyleSheet,
  Text,
  View,
} from '@react-pdf/renderer'
import type { ReactNode } from 'react'
import type { ContentBlock, ParsedDocument } from './types'

const COLORS = {
  primaryBlue: '#005A9C',
  darkNavy: '#003B6F',
  mediumBlue: '#0B63B6',
  lightBlue: '#D9EAF7',
  veryLightBlue: '#EEF6FC',
  white: '#FFFFFF',
  mainText: '#2B2B2B',
  secondaryText: '#6B7280',
  divider: '#D1D5DB',
}

const NETWORK_PATTERN = `data:image/svg+xml;utf8,${encodeURIComponent(`
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1200 1200">
  <rect width="1200" height="1200" fill="#003B6F"/>
  <g stroke="#4D83B8" stroke-width="1.2" opacity="0.35">
    <line x1="40" y1="900" x2="320" y2="740"/>
    <line x1="320" y1="740" x2="600" y2="840"/>
    <line x1="600" y1="840" x2="910" y2="730"/>
    <line x1="910" y1="730" x2="1160" y2="860"/>
    <line x1="120" y1="1090" x2="320" y2="740"/>
    <line x1="450" y1="1120" x2="600" y2="840"/>
    <line x1="790" y1="1120" x2="910" y2="730"/>
  </g>
  <g fill="#8CB7E0" opacity="0.6">
    <circle cx="40" cy="900" r="8"/>
    <circle cx="320" cy="740" r="8"/>
    <circle cx="600" cy="840" r="8"/>
    <circle cx="910" cy="730" r="8"/>
    <circle cx="1160" cy="860" r="8"/>
    <circle cx="120" cy="1090" r="8"/>
    <circle cx="450" cy="1120" r="8"/>
    <circle cx="790" cy="1120" r="8"/>
  </g>
</svg>
`)}`

const MAJOR_SECTION_REGEX =
  /^(introdu[cç][aã]o|metodologia|desenvolvimento|conclus(?:[aã]o|[oõ]es)|refer[eê]ncias|anexos?)$/i
const SOURCE_PREFIX = /^(Fonte|Source)\s*[:\-–]?/i

const styles = StyleSheet.create({
  coverPage: {
    backgroundColor: COLORS.darkNavy,
    color: COLORS.white,
    paddingTop: 34,
    paddingBottom: 34,
    paddingHorizontal: 42,
    fontFamily: 'Helvetica',
    position: 'relative',
  },
  coverBackground: {
    position: 'absolute',
    top: 0,
    left: 0,
    right: 0,
    bottom: 0,
  },
  coverOverlay: {
    position: 'absolute',
    bottom: 0,
    left: 0,
    right: 0,
    height: '46%',
    backgroundColor: COLORS.primaryBlue,
    opacity: 0.78,
  },
  coverLogoRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 36,
  },
  coverInstitution: {
    fontSize: 12,
    letterSpacing: 1.6,
    color: COLORS.lightBlue,
    textTransform: 'uppercase',
  },
  coverTag: {
    borderRadius: 12,
    paddingVertical: 5,
    paddingHorizontal: 10,
    backgroundColor: COLORS.mediumBlue,
    fontSize: 8,
    color: COLORS.white,
  },
  coverContent: {
    marginTop: 'auto',
    gap: 8,
  },
  coverTitle: {
    fontFamily: 'Helvetica-Bold',
    fontSize: 28,
    lineHeight: 1.25,
    textTransform: 'uppercase',
  },
  coverSubtitle: {
    fontSize: 13,
    color: COLORS.lightBlue,
    lineHeight: 1.4,
    marginTop: 4,
  },
  coverMetadata: {
    marginTop: 16,
    fontSize: 9.5,
    lineHeight: 1.45,
    color: COLORS.veryLightBlue,
  },
  versionBadge: {
    position: 'absolute',
    bottom: 34,
    right: 42,
    backgroundColor: COLORS.white,
    borderRadius: 3,
    paddingVertical: 4,
    paddingHorizontal: 8,
  },
  versionBadgeText: {
    color: COLORS.darkNavy,
    fontSize: 8.5,
    fontFamily: 'Helvetica-Bold',
  },
  simplePage: {
    backgroundColor: COLORS.white,
    paddingHorizontal: 52,
    paddingVertical: 58,
    color: COLORS.mainText,
    fontFamily: 'Helvetica',
  },
  pageTitle: {
    color: COLORS.primaryBlue,
    fontFamily: 'Helvetica-Bold',
    fontSize: 20,
    marginBottom: 8,
  },
  pageDivider: {
    height: 1,
    backgroundColor: COLORS.divider,
    marginBottom: 18,
  },
  simpleText: {
    fontSize: 11,
    lineHeight: 1.6,
    marginBottom: 6,
  },
  roleLabel: {
    fontFamily: 'Helvetica-Bold',
    color: COLORS.darkNavy,
    marginTop: 10,
    marginBottom: 2,
    fontSize: 11,
  },
  tocRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    borderBottomWidth: 1,
    borderBottomColor: '#EDF2F7',
    paddingVertical: 5,
  },
  tocEntry: {
    fontSize: 10.2,
    color: COLORS.mainText,
  },
  tocLevel2: {
    marginLeft: 10,
  },
  tocLevel3: {
    marginLeft: 20,
  },
  sectionDividerPage: {
    backgroundColor: COLORS.darkNavy,
    color: COLORS.white,
    position: 'relative',
    paddingHorizontal: 42,
    paddingBottom: 72,
    fontFamily: 'Helvetica',
  },
  sectionDividerTitle: {
    marginTop: 'auto',
    fontSize: 34,
    lineHeight: 1.2,
    fontFamily: 'Helvetica-Bold',
    textTransform: 'uppercase',
  },
  sectionDividerBar: {
    height: 4,
    width: 180,
    marginTop: 10,
    backgroundColor: COLORS.mediumBlue,
  },
  contentPage: {
    backgroundColor: COLORS.white,
    color: COLORS.mainText,
    paddingTop: 78,
    paddingBottom: 40,
    paddingHorizontal: 46,
    fontFamily: 'Helvetica',
    fontSize: 10.8,
  },
  contentHeader: {
    position: 'absolute',
    top: 24,
    left: 46,
    right: 46,
  },
  contentHeaderRow: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
  },
  contentHeaderTitle: {
    fontSize: 9.5,
    color: COLORS.darkNavy,
    letterSpacing: 0.5,
    textTransform: 'uppercase',
  },
  pageNumber: {
    fontSize: 9,
    color: COLORS.secondaryText,
  },
  contentHeaderLine: {
    marginTop: 5,
    height: 1.2,
    backgroundColor: COLORS.primaryBlue,
  },
  heading1: {
    fontSize: 19,
    color: COLORS.primaryBlue,
    fontFamily: 'Helvetica-Bold',
    marginTop: 4,
    marginBottom: 6,
  },
  heading2: {
    fontSize: 15,
    color: COLORS.mediumBlue,
    fontFamily: 'Helvetica-Bold',
    marginTop: 10,
    marginBottom: 5,
  },
  heading3: {
    fontSize: 12.5,
    color: COLORS.darkNavy,
    fontFamily: 'Helvetica-Bold',
    marginTop: 8,
    marginBottom: 4,
  },
  headingDivider: {
    height: 1,
    backgroundColor: '#E5E7EB',
    marginBottom: 9,
  },
  paragraph: {
    fontSize: 10.8,
    lineHeight: 1.66,
    marginBottom: 8,
    textAlign: 'justify',
  },
  bulletList: {
    marginBottom: 8,
    gap: 4,
  },
  listLine: {
    flexDirection: 'row',
    gap: 6,
  },
  listBullet: {
    width: 10,
    fontSize: 10.5,
    color: COLORS.darkNavy,
  },
  listText: {
    flex: 1,
    fontSize: 10.6,
    lineHeight: 1.55,
  },
  figureCard: {
    marginTop: 4,
    marginBottom: 11,
    borderRadius: 6,
    borderWidth: 1,
    borderColor: '#D6DEE8',
    padding: 8,
  },
  figureImage: {
    width: '100%',
    maxHeight: 280,
    objectFit: 'cover',
    borderRadius: 4,
    marginBottom: 5,
  },
  caption: {
    fontSize: 9,
    color: COLORS.secondaryText,
    marginBottom: 2,
  },
  source: {
    fontSize: 8.5,
    color: COLORS.secondaryText,
    textAlign: 'center',
  },
  tableWrapper: {
    marginBottom: 12,
  },
  tableTitle: {
    fontFamily: 'Helvetica-Bold',
    fontSize: 10.5,
    color: COLORS.darkNavy,
    marginBottom: 4,
  },
  table: {
    borderWidth: 1,
    borderColor: COLORS.divider,
  },
  tableRow: {
    flexDirection: 'row',
  },
  tableHeaderCell: {
    backgroundColor: COLORS.darkNavy,
    color: COLORS.white,
    fontFamily: 'Helvetica-Bold',
    fontSize: 9.2,
    borderRightWidth: 1,
    borderRightColor: COLORS.white,
    padding: 4,
    flex: 1,
  },
  tableCell: {
    fontSize: 9.5,
    color: COLORS.mainText,
    borderTopWidth: 1,
    borderTopColor: COLORS.divider,
    borderRightWidth: 1,
    borderRightColor: COLORS.divider,
    padding: 4,
    flex: 1,
  },
  tableEven: {
    backgroundColor: COLORS.veryLightBlue,
  },
  callout: {
    borderLeftWidth: 4,
    borderLeftColor: COLORS.darkNavy,
    backgroundColor: COLORS.veryLightBlue,
    paddingVertical: 8,
    paddingHorizontal: 10,
    marginBottom: 10,
  },
  calloutTitle: {
    color: COLORS.primaryBlue,
    fontFamily: 'Helvetica-Bold',
    fontSize: 10.5,
    marginBottom: 3,
  },
  calloutText: {
    color: COLORS.mainText,
    fontSize: 10.2,
    lineHeight: 1.5,
  },
  backCover: {
    backgroundColor: COLORS.primaryBlue,
    color: COLORS.white,
    paddingHorizontal: 42,
    paddingVertical: 40,
    fontFamily: 'Helvetica',
    position: 'relative',
  },
  backLine: {
    position: 'absolute',
    bottom: 78,
    left: 42,
    right: 42,
    height: 1,
    backgroundColor: '#EAF4FB',
  },
  backTitle: {
    marginTop: 'auto',
    fontSize: 18,
    fontFamily: 'Helvetica-Bold',
    lineHeight: 1.35,
  },
  backSubtitle: {
    marginTop: 4,
    fontSize: 11,
    color: COLORS.lightBlue,
  },
  backContact: {
    marginTop: 12,
    fontSize: 9.4,
    lineHeight: 1.5,
    color: COLORS.veryLightBlue,
  },
  backFooter: {
    position: 'absolute',
    right: 42,
    bottom: 32,
    textAlign: 'right',
    fontSize: 9,
    color: COLORS.veryLightBlue,
  },
})

interface GroupedContent {
  dividerTitle?: string
  blocks: ContentBlock[]
}

const buildGroups = (blocks: ContentBlock[]) => {
  const groups: GroupedContent[] = [{ blocks: [] }]

  blocks.forEach((block) => {
    const currentGroup = groups[groups.length - 1]
    if (
      block.type === 'heading' &&
      MAJOR_SECTION_REGEX.test(block.text) &&
      currentGroup.blocks.length > 0
    ) {
      groups.push({ dividerTitle: block.text, blocks: [block] })
      return
    }
    currentGroup.blocks.push(block)
  })

  return groups
}

const renderBlock = (block: ContentBlock, index: number): ReactNode => {
  if (block.type === 'heading') {
    const headingStyle =
      block.level === 1 ? styles.heading1 : block.level === 2 ? styles.heading2 : styles.heading3
    return (
      <View key={`heading-${index}`} wrap={false}>
        <Text style={headingStyle}>{block.text}</Text>
        <View style={styles.headingDivider} />
      </View>
    )
  }

  if (block.type === 'paragraph') {
    const paragraphStyle = SOURCE_PREFIX.test(block.text) ? styles.source : styles.paragraph
    return (
      <Text key={`paragraph-${index}`} style={paragraphStyle}>
        {block.text}
      </Text>
    )
  }

  if (block.type === 'list') {
    return (
      <View key={`list-${index}`} style={styles.bulletList}>
        {block.items.map((item, itemIndex) => (
          <View key={`${item}-${itemIndex}`} style={styles.listLine}>
            <Text style={styles.listBullet}>{block.ordered ? `${itemIndex + 1}.` : '•'}</Text>
            <Text style={styles.listText}>{item}</Text>
          </View>
        ))}
      </View>
    )
  }

  if (block.type === 'figure') {
    return (
      <View key={`figure-${index}`} style={styles.figureCard} wrap={false}>
        <Image style={styles.figureImage} src={block.src} />
        {block.caption ? <Text style={styles.caption}>{block.caption}</Text> : null}
        {block.source ? <Text style={styles.source}>{block.source}</Text> : null}
      </View>
    )
  }

  if (block.type === 'table') {
    const [header, ...rows] = block.rows
    return (
      <View key={`table-${index}`} style={styles.tableWrapper}>
        {block.title ? <Text style={styles.tableTitle}>{block.title}</Text> : null}
        <View style={styles.table}>
          {header ? (
            <View style={styles.tableRow}>
              {header.map((cell, cellIndex) => (
                <Text
                  key={`${cell}-${cellIndex}`}
                  style={[
                    styles.tableHeaderCell,
                    ...(cellIndex === header.length - 1 ? [{ borderRightWidth: 0 }] : []),
                  ]}
                >
                  {cell}
                </Text>
              ))}
            </View>
          ) : null}
          {rows.map((row, rowIndex) => (
            <View key={`row-${rowIndex}`} style={styles.tableRow}>
              {row.map((cell, cellIndex) => (
                <Text
                  key={`${cell}-${cellIndex}`}
                  style={[
                    styles.tableCell,
                    ...(rowIndex % 2 === 1 ? [styles.tableEven] : []),
                    ...(cellIndex === row.length - 1 ? [{ borderRightWidth: 0 }] : []),
                  ]}
                >
                  {cell}
                </Text>
              ))}
            </View>
          ))}
        </View>
        {block.source ? <Text style={[styles.source, { marginTop: 4 }]}>{block.source}</Text> : null}
      </View>
    )
  }

  return (
    <View key={`callout-${index}`} style={styles.callout}>
      <Text style={styles.calloutTitle}>{block.title}</Text>
      <Text style={styles.calloutText}>{block.text}</Text>
    </View>
  )
}

const hasCreditsPage = (doc: ParsedDocument) =>
  doc.meta.authors.length > 0 || doc.meta.reviewers.length > 0 || doc.meta.credits.length > 0

const hasCatalogPage = (doc: ParsedDocument) =>
  doc.meta.catalogInfo.length > 0 || doc.meta.contacts.length > 0

interface InstitutionalReportDocumentProps {
  parsed: ParsedDocument
  sourceFilename: string
}

export const InstitutionalReportDocument = ({
  parsed,
  sourceFilename,
}: InstitutionalReportDocumentProps) => {
  const groups = buildGroups(parsed.blocks)
  const generatedDate = new Date().toLocaleDateString('pt-BR')
  const contentPages: ReactNode[] = groups.flatMap((group, groupIndex) => {
    const pages: ReactNode[] = []
    if (group.dividerTitle) {
      pages.push(
        <Page key={`divider-${groupIndex}`} size="A4" style={styles.sectionDividerPage}>
          <Image src={NETWORK_PATTERN} style={styles.coverBackground} />
          <Text style={styles.sectionDividerTitle}>{group.dividerTitle}</Text>
          <View style={styles.sectionDividerBar} />
        </Page>,
      )
    }

    pages.push(
      <Page key={`content-${groupIndex}`} size="A4" style={styles.contentPage} wrap>
        <View style={styles.contentHeader} fixed>
          <View style={styles.contentHeaderRow}>
            <Text style={styles.contentHeaderTitle}>{parsed.meta.reportTitle}</Text>
            <Text
              style={styles.pageNumber}
              render={({ pageNumber, totalPages }) => `${pageNumber}/${totalPages}`}
            />
          </View>
          <View style={styles.contentHeaderLine} />
        </View>
        {group.blocks.map((block, index) => renderBlock(block, index))}
      </Page>,
    )

    return pages
  })

  return (
    <Document title={parsed.meta.reportTitle} author={parsed.meta.authors.join(', ') || 'Sistema FIEA'}>
      <Page size="A4" style={styles.coverPage}>
        <Image src={NETWORK_PATTERN} style={styles.coverBackground} />
        <View style={styles.coverOverlay} />
        <View style={styles.coverLogoRow}>
          <Text style={styles.coverInstitution}>Observatório da Indústria · Sistema FIEA</Text>
          <Text style={styles.coverTag}>RELATÓRIO TÉCNICO</Text>
        </View>
        <View style={styles.coverContent}>
          <Text style={styles.coverTitle}>{parsed.meta.reportTitle}</Text>
          {parsed.meta.subtitle ? <Text style={styles.coverSubtitle}>{parsed.meta.subtitle}</Text> : null}
          <Text style={styles.coverMetadata}>
            {parsed.meta.description || 'Publicação institucional para inteligência estratégica.'}
          </Text>
          {parsed.meta.authors.length > 0 ? (
            <Text style={styles.coverMetadata}>Autores: {parsed.meta.authors.join(', ')}</Text>
          ) : null}
          {parsed.meta.reviewers.length > 0 ? (
            <Text style={styles.coverMetadata}>Revisão: {parsed.meta.reviewers.join(', ')}</Text>
          ) : null}
        </View>
        {parsed.meta.version || parsed.meta.date ? (
          <View style={styles.versionBadge}>
            <Text style={styles.versionBadgeText}>
              {parsed.meta.version || 'Versão técnica'} · {parsed.meta.date || generatedDate}
            </Text>
          </View>
        ) : null}
      </Page>

      {hasCreditsPage(parsed) ? (
        <Page size="A4" style={styles.simplePage}>
          <Text style={styles.pageTitle}>CRÉDITOS INSTITUCIONAIS</Text>
          <View style={styles.pageDivider} />
          {parsed.meta.authors.length > 0 ? (
            <>
              <Text style={styles.roleLabel}>Autores</Text>
              <Text style={styles.simpleText}>{parsed.meta.authors.join(', ')}</Text>
            </>
          ) : null}
          {parsed.meta.reviewers.length > 0 ? (
            <>
              <Text style={styles.roleLabel}>Revisão</Text>
              <Text style={styles.simpleText}>{parsed.meta.reviewers.join(', ')}</Text>
            </>
          ) : null}
          {parsed.meta.credits.length > 0 ? (
            <>
              <Text style={styles.roleLabel}>Equipe e execução</Text>
              {parsed.meta.credits.map((entry, index) => (
                <Text key={`${entry}-${index}`} style={styles.simpleText}>
                  {entry}
                </Text>
              ))}
            </>
          ) : null}
        </Page>
      ) : null}

      {hasCatalogPage(parsed) ? (
        <Page size="A4" style={styles.simplePage}>
          <Text style={styles.pageTitle}>CATALOGAÇÃO E CONTATOS</Text>
          <View style={styles.pageDivider} />
          {parsed.meta.catalogInfo.map((entry, index) => (
            <Text key={`${entry}-${index}`} style={styles.simpleText}>
              {entry}
            </Text>
          ))}
          {parsed.meta.contacts.length > 0 ? (
            <>
              <Text style={styles.roleLabel}>Contato institucional</Text>
              {parsed.meta.contacts.map((entry, index) => (
                <Text key={`${entry}-${index}`} style={styles.simpleText}>
                  {entry}
                </Text>
              ))}
            </>
          ) : null}
        </Page>
      ) : null}

      <Page size="A4" style={styles.simplePage}>
        <Text style={styles.pageTitle}>SUMÁRIO</Text>
        <View style={styles.pageDivider} />
        {parsed.headings.map((heading, index) => (
          <View key={`${heading.text}-${index}`} style={styles.tocRow}>
            <Text
              style={[
                styles.tocEntry,
                ...(heading.level === 2 ? [styles.tocLevel2] : []),
                ...(heading.level >= 3 ? [styles.tocLevel3] : []),
              ]}
            >
              {heading.text}
            </Text>
            <Text style={styles.tocEntry}>{index + 1}</Text>
          </View>
        ))}
      </Page>

      {contentPages}

      <Page size="A4" style={styles.backCover}>
        <Image src={NETWORK_PATTERN} style={styles.coverBackground} />
        <Text style={styles.coverInstitution}>Sistema FIEA · Observatório da Indústria</Text>
        <View style={styles.backLine} />
        <Text style={styles.backTitle}>{parsed.meta.reportTitle}</Text>
        {parsed.meta.subtitle ? <Text style={styles.backSubtitle}>{parsed.meta.subtitle}</Text> : null}
        <Text style={styles.backContact}>
          {parsed.meta.contacts[0] || 'www.fiea.com.br · observatorio@fiea.com.br'}
        </Text>
        <Text style={styles.backFooter}>
          Documento gerado automaticamente a partir de {sourceFilename}. {'\n'}Formato A4 · PDF institucional
        </Text>
      </Page>
    </Document>
  )
}
