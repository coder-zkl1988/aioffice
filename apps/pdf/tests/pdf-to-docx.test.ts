import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'

vi.mock('../src/renderer/pdf-to-images', () => ({
  pdfPageImageRenderScale: vi.fn(() => 2),
  renderPdfPagesAsImages: vi.fn(),
}))

import { pdfDocxTextRun, preparePdfDocxPages } from '../src/renderer/pdf-to-docx'
import { renderPdfPagesAsImages } from '../src/renderer/pdf-to-images'

describe('PDF to Word and OpenDocument preparation', () => {
  beforeEach(() => vi.mocked(renderPdfPagesAsImages).mockReset())

  it('maps PDF text and keeps explicit line endings', () => {
    expect(
      pdfDocxTextRun(
        {
          str: 'Editable',
          transform: [18, 0, 0, 18, 40, 700],
          width: 72,
          height: 18,
          fontName: 'HeadingBoldItalic',
          hasEOL: true,
        },
        [1, 0, 0, -1, 0, 792],
        { fontFamily: 'Inter' },
      ),
    ).toMatchObject({
      text: 'Editable',
      x: 40,
      y: 74,
      fontFamily: 'Inter',
      bold: true,
      italic: true,
      hasEol: true,
    })
  })

  function documentFixture(cleanup = vi.fn()): PDFDocumentProxy {
    return {
      numPages: 1,
      getPage: vi.fn(async () => ({
        getViewport: () => ({ width: 612, height: 792, transform: [1, 0, 0, -1, 0, 792] }),
        getTextContent: async () => ({
          items: [
            {
              str: 'Local Word',
              transform: [20, 0, 0, 20, 48, 720],
              width: 120,
              height: 20,
              fontName: 'HeadingBold',
            },
          ],
          styles: { HeadingBold: { fontFamily: 'Arial' } },
        }),
        cleanup,
      })),
    } as unknown as PDFDocumentProxy
  }

  it('prepares editable text without rasterizing pages', async () => {
    const pages = await preparePdfDocxPages(documentFixture(), {
      pageIndexes: [0],
      mode: 'editableText',
      renderDpi: 150,
      includeAnnotations: true,
    })
    expect(pages).toMatchObject([
      { pageNumber: 1, width: 612, height: 792, textRuns: [{ text: 'Local Word' }] },
    ])
    expect(pages[0]).not.toHaveProperty('imageBytes')
    expect(renderPdfPagesAsImages).not.toHaveBeenCalled()
  })

  it('adds local page images and exact dimensions only for fidelity mode', async () => {
    vi.mocked(renderPdfPagesAsImages).mockResolvedValue([
      { pageNumber: 1, bytes: new Uint8Array([137, 80, 78, 71]) },
    ])
    const pages = await preparePdfDocxPages(documentFixture(), {
      pageIndexes: [0],
      mode: 'fidelity',
      renderDpi: 144,
      includeAnnotations: false,
    })
    expect(pages[0]).toMatchObject({ imageWidth: 1224, imageHeight: 1584 })
    expect(Array.from(pages[0]!.imageBytes!)).toEqual([137, 80, 78, 71])
    expect(renderPdfPagesAsImages).toHaveBeenCalledWith(
      expect.anything(),
      expect.objectContaining({ format: 'png', renderDpi: 144, includeAnnotations: false }),
    )
  })
})
