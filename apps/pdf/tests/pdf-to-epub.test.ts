import { beforeEach, describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'

vi.mock('../src/renderer/pdf-to-images', () => ({
  renderPdfPagesAsImages: vi.fn(),
}))

import { preparePdfEpubPages } from '../src/renderer/pdf-to-epub'
import { renderPdfPagesAsImages } from '../src/renderer/pdf-to-images'

describe('PDF to EPUB preparation', () => {
  beforeEach(() => vi.mocked(renderPdfPagesAsImages).mockReset())

  function documentFixture(cleanup = vi.fn()): PDFDocumentProxy {
    return {
      numPages: 1,
      getPage: vi.fn(async () => ({
        getViewport: () => ({ width: 612, height: 792 }),
        getTextContent: async () => ({
          items: [
            {
              str: 'Local EPUB',
              transform: [20, 0, 0, 20, 48, 720],
              width: 110,
              height: 20,
              fontName: 'HeadingBold',
            },
          ],
          styles: { HeadingBold: { fontFamily: 'Arial' } },
        }),
        getAnnotations: async () => [],
        cleanup,
      })),
    } as unknown as PDFDocumentProxy
  }

  it('prepares reflowable pages without rasterizing them', async () => {
    const cleanup = vi.fn()
    const pages = await preparePdfEpubPages(documentFixture(cleanup), {
      pageIndexes: [0],
      mode: 'reflowable',
      renderDpi: 150,
      includeAnnotations: true,
    })
    expect(pages).toMatchObject([{ pageNumber: 1, width: 612, height: 792, text: 'Local EPUB' }])
    expect(pages[0]).not.toHaveProperty('imageBytes')
    expect(renderPdfPagesAsImages).not.toHaveBeenCalled()
    expect(cleanup).toHaveBeenCalledTimes(2)
  })

  it('adds local PNG pages only for fixed layout', async () => {
    vi.mocked(renderPdfPagesAsImages).mockResolvedValue([
      { pageNumber: 1, bytes: new Uint8Array([137, 80, 78, 71]) },
    ])
    const pages = await preparePdfEpubPages(documentFixture(), {
      pageIndexes: [0],
      mode: 'fixed',
      renderDpi: 144,
      includeAnnotations: false,
    })
    expect(Array.from(pages[0]!.imageBytes!)).toEqual([137, 80, 78, 71])
    expect(renderPdfPagesAsImages).toHaveBeenCalledWith(
      expect.anything(),
      expect.objectContaining({ format: 'png', renderDpi: 144, includeAnnotations: false }),
    )
  })
})
