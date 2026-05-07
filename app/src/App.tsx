import { pdf } from '@react-pdf/renderer'
import { useMemo, useState, type ChangeEvent } from 'react'
import './App.css'
import { parseDocxFile } from './lib/docxParser'
import { InstitutionalReportDocument } from './lib/InstitutionalReportDocument'
import type { ParsedDocument } from './lib/types'

function App() {
  const [sourceFile, setSourceFile] = useState<File | null>(null)
  const [parsedDocument, setParsedDocument] = useState<ParsedDocument | null>(null)
  const [isParsing, setIsParsing] = useState(false)
  const [isGenerating, setIsGenerating] = useState(false)
  const [errorMessage, setErrorMessage] = useState<string | null>(null)

  const stats = useMemo(() => {
    if (!parsedDocument) {
      return null
    }
    const tables = parsedDocument.blocks.filter((block) => block.type === 'table').length
    const figures = parsedDocument.blocks.filter((block) => block.type === 'figure').length

    return {
      headings: parsedDocument.headings.length,
      tables,
      figures,
      blocks: parsedDocument.blocks.length,
    }
  }, [parsedDocument])

  const handleUpload = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }
    if (!file.name.toLowerCase().endsWith('.docx')) {
      setErrorMessage('Selecione um arquivo DOCX válido.')
      setSourceFile(null)
      setParsedDocument(null)
      return
    }

    try {
      setErrorMessage(null)
      setIsParsing(true)
      setSourceFile(file)
      const parsed = await parseDocxFile(file)
      setParsedDocument(parsed)
    } catch (error) {
      console.error(error)
      setParsedDocument(null)
      setErrorMessage(
        'Não foi possível ler o DOCX. Verifique se o arquivo não está corrompido e tente novamente.',
      )
    } finally {
      setIsParsing(false)
    }
  }

  const handleGeneratePdf = async () => {
    if (!parsedDocument || !sourceFile) {
      return
    }

    try {
      setErrorMessage(null)
      setIsGenerating(true)
      const pdfDocument = (
        <InstitutionalReportDocument parsed={parsedDocument} sourceFilename={sourceFile.name} />
      )
      const blob = await pdf(pdfDocument).toBlob()
      const outputName = sourceFile.name.replace(/\.docx$/i, '') || 'relatorio-institucional'
      const url = URL.createObjectURL(blob)
      const anchor = window.document.createElement('a')
      anchor.href = url
      anchor.download = `${outputName}.pdf`
      anchor.click()
      window.setTimeout(() => URL.revokeObjectURL(url), 1000)
    } catch (error) {
      console.error(error)
      setErrorMessage('Falha ao gerar PDF. Tente novamente com outro documento.')
    } finally {
      setIsGenerating(false)
    }
  }

  return (
    <main className="app">
      <header className="hero">
        <p className="kicker">Observatório da Indústria · Sistema FIEA</p>
        <h1>DOCX para PDF Institucional</h1>
        <p className="subtitle">
          Faça upload do relatório em DOCX para gerar automaticamente um PDF em padrão editorial
          corporativo com identidade analítica e institucional.
        </p>
      </header>

      <section className="card">
        <label className="upload">
          <span>Arquivo DOCX</span>
          <input
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            onChange={handleUpload}
          />
        </label>

        <div className="actions">
          <button type="button" onClick={handleGeneratePdf} disabled={!parsedDocument || isGenerating}>
            {isGenerating ? 'Gerando PDF...' : 'Gerar PDF'}
          </button>
          <span className="hint">
            {isParsing
              ? 'Lendo e estruturando o conteúdo do DOCX...'
              : parsedDocument
                ? 'Conteúdo estruturado com sucesso.'
                : 'Selecione um documento para habilitar a geração do PDF.'}
          </span>
        </div>

        {errorMessage ? <p className="error">{errorMessage}</p> : null}
      </section>

      {parsedDocument && sourceFile ? (
        <section className="card meta">
          <h2>Resumo da estrutura identificada</h2>
          <p>
            <strong>Título:</strong> {parsedDocument.meta.reportTitle}
          </p>
          {parsedDocument.meta.subtitle ? (
            <p>
              <strong>Subtítulo:</strong> {parsedDocument.meta.subtitle}
            </p>
          ) : null}
          <p>
            <strong>Arquivo:</strong> {sourceFile.name}
          </p>
          {stats ? (
            <ul>
              <li>{stats.headings} títulos e subtítulos</li>
              <li>{stats.tables} tabelas</li>
              <li>{stats.figures} figuras/imagens</li>
              <li>{stats.blocks} blocos de conteúdo</li>
            </ul>
          ) : null}
        </section>
      ) : null}
    </main>
  )
}

export default App
