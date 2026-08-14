import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { pdfMarkdownOutputFileName, pdfTextPagesMarkdownBytes, runPdfToolBytes } from '../src/index'

const pages = [
  {
    pageNumber: 1,
    text: 'Product launch\n\nEditable slides',
    blocks: [
      { kind: 'heading' as const, text: 'Product launch', level: 1 as const },
      { kind: 'paragraph' as const, text: 'Use *local* generation.' },
      { kind: 'listItem' as const, text: '• Editable slides' },
    ],
    links: [{ url: 'https://example.com/docs', label: 'Documentation' }],
  },
  {
    pageNumber: 2,
    text: 'Next steps',
    blocks: [{ kind: 'heading' as const, text: 'Next steps', level: 2 as const }],
    links: [],
  },
]

async function sourcePdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  document.addPage([100, 100])
  document.addPage([100, 100])
  return document.save()
}

describe('PDF to Markdown operation', () => {
  it('preserves page boundaries and shifts heading levels', () => {
    const markdown = new TextDecoder().decode(pdfTextPagesMarkdownBytes(pages, true))

    expect(markdown).toContain('## Page 1')
    expect(markdown).toContain('### Product launch')
    expect(markdown).toContain('- Editable slides')
    expect(markdown).toContain('### Links')
    expect(markdown).toContain('\n\n---\n\n## Page 2')
  })

  it('creates continuous Markdown with natural heading levels', () => {
    const markdown = new TextDecoder().decode(pdfTextPagesMarkdownBytes(pages, false))

    expect(markdown).toContain('# Product launch')
    expect(markdown).toContain('## Next steps')
    expect(markdown).not.toContain('Page 1')
    expect(markdown).not.toContain('\n---\n')
  })

  it('exports a named Markdown file with platform metadata', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'pdfToMarkdown',
      pageIndexes: [0, 1],
      includePageBreaks: false,
      baseName: '../Report.pdf',
      pages,
    })

    expect(output).toMatchObject({
      fileName: 'Report_converted.md',
      mimeType: 'text/markdown;charset=utf-8',
      extension: '.md',
    })
    expect(new TextDecoder().decode(output!.bytes)).toContain('# Product launch')
  })

  it('sanitizes names and rejects missing or empty extraction data', async () => {
    expect(pdfMarkdownOutputFileName('../../<bad>.pdf')).toBe('bad_converted.md')
    const source = await sourcePdf()
    await expect(
      runPdfToolBytes(source, {
        kind: 'pdfToMarkdown',
        pageIndexes: [0],
        includePageBreaks: true,
      }),
    ).rejects.toThrow('required')
    await expect(
      runPdfToolBytes(source, {
        kind: 'pdfToMarkdown',
        pageIndexes: [0],
        includePageBreaks: true,
        pages: [{ pageNumber: 1, text: '', blocks: [], links: [] }],
      }),
    ).rejects.toThrow('OCR')
  })
})
