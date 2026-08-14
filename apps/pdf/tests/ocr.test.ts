import type { PDFDocumentProxy } from 'pdfjs-dist'
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import { addPdfOcrTextLayersBytes } from '@genoffice/pdf-tools'
import { pdfOcrPagePlan } from '../src/renderer/ocr'

function fakeDocument(pageTexts: string[]): {
  document: PDFDocumentProxy
  cleanup: ReturnType<typeof vi.fn>
} {
  const cleanup = vi.fn()
  return {
    document: {
      numPages: pageTexts.length,
      getPage: vi.fn(async (pageNumber: number) => ({
        getTextContent: async () => ({ items: [{ str: pageTexts[pageNumber - 1] }] }),
        cleanup,
      })),
    } as unknown as PDFDocumentProxy,
    cleanup,
  }
}

describe('PDF OCR page planning', () => {
  it('skips pages that already contain text in automatic mode', async () => {
    const { document, cleanup } = fakeDocument(['Existing text', '', 'Scan'])
    await expect(pdfOcrPagePlan(document, 'skipText')).resolves.toEqual({
      pageIndexes: [1],
      skippedPageIndexes: [0, 2],
      existingPageTexts: [
        { pageIndex: 0, text: 'Existing text', source: 'existing' },
        { pageIndex: 2, text: 'Scan', source: 'existing' },
      ],
    })
    expect(cleanup).toHaveBeenCalledTimes(3)
  })

  it('processes every page in force mode', async () => {
    const { document } = fakeDocument(['Existing text', ''])
    await expect(pdfOcrPagePlan(document, 'force')).resolves.toEqual({
      pageIndexes: [0, 1],
      skippedPageIndexes: [],
      existingPageTexts: [{ pageIndex: 0, text: 'Existing text', source: 'existing' }],
    })
  })

  it('stops in strict mode when existing text is found', async () => {
    const { document, cleanup } = fakeDocument(['', 'Existing text'])
    await expect(pdfOcrPagePlan(document, 'strict')).rejects.toThrow(
      /page 2 already contains text/i,
    )
    expect(cleanup).toHaveBeenCalledTimes(2)
  })

  it('produces a text layer that PDF.js can search', async () => {
    const source = await PDFDocument.create()
    source.addPage([300, 200])
    const layer = await PDFDocument.create()
    const layerPage = layer.addPage([300, 200])
    const font = await layer.embedFont(StandardFonts.Helvetica)
    layerPage.drawText('SEARCHABLE LOCAL OCR', { x: 20, y: 150, size: 18, font })
    const output = await addPdfOcrTextLayersBytes(await source.save(), [
      { pageIndex: 0, bytes: await layer.save() },
    ])
    const loadingTask = getDocument({ data: output.slice(), useSystemFonts: true })
    const document = await loadingTask.promise
    try {
      const page = await document.getPage(1)
      const content = await page.getTextContent()
      expect(content.items.map((item) => ('str' in item ? item.str : '')).join(' ')).toContain(
        'SEARCHABLE LOCAL OCR',
      )
    } finally {
      await loadingTask.destroy()
    }
  })

  it('replaces the original text layer in force mode output', async () => {
    const source = await PDFDocument.create()
    const sourcePage = source.addPage([300, 200])
    const sourceFont = await source.embedFont(StandardFonts.Helvetica)
    sourcePage.drawText('OLD TEXT', { x: 20, y: 150, size: 18, font: sourceFont })
    const layer = await PDFDocument.create()
    const layerPage = layer.addPage([300, 200])
    const layerFont = await layer.embedFont(StandardFonts.Helvetica)
    layerPage.drawText('NEW OCR TEXT', { x: 20, y: 150, size: 18, font: layerFont })
    const output = await addPdfOcrTextLayersBytes(await source.save(), [
      { pageIndex: 0, bytes: await layer.save(), replacePage: true },
    ])
    const loadingTask = getDocument({ data: output.slice(), useSystemFonts: true })
    const document = await loadingTask.promise
    try {
      const page = await document.getPage(1)
      const content = await page.getTextContent()
      const text = content.items.map((item) => ('str' in item ? item.str : '')).join(' ')
      expect(text).toContain('NEW OCR TEXT')
      expect(text).not.toContain('OLD TEXT')
    } finally {
      await loadingTask.destroy()
    }
  })
})
