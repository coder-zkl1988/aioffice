import type { PdfEpubPage, PdfToEpubOptions } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { extractPdfTextPages } from './extract-text'
import { renderPdfPagesAsImages } from './pdf-to-images'

export async function preparePdfEpubPages(
  sourceDocument: PDFDocumentProxy,
  options: Pick<PdfToEpubOptions, 'pageIndexes' | 'mode' | 'renderDpi' | 'includeAnnotations'>,
): Promise<PdfEpubPage[]> {
  const semanticPages = await extractPdfTextPages(sourceDocument, options.pageIndexes)
  const images =
    options.mode === 'fixed'
      ? await renderPdfPagesAsImages(sourceDocument, {
          pageIndexes: options.pageIndexes,
          format: 'png',
          outputMode: 'multiple',
          renderDpi: options.renderDpi,
          imageQuality: 100,
          colorMode: 'color',
          includeAnnotations: options.includeAnnotations,
        })
      : undefined
  const pages: PdfEpubPage[] = []
  for (let index = 0; index < options.pageIndexes.length; index++) {
    const pageIndex = options.pageIndexes[index]!
    const page = await sourceDocument.getPage(pageIndex + 1)
    try {
      const viewport = page.getViewport({ scale: 1 })
      pages.push({
        ...semanticPages[index]!,
        width: viewport.width,
        height: viewport.height,
        ...(images ? { imageBytes: images[index]!.bytes } : {}),
      })
    } finally {
      page.cleanup()
    }
  }
  return pages
}
