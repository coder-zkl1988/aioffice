import type { PdfDocxPage, PdfDocxTextRun, PdfToDocxOptions } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfPptxTextRun } from './pdf-to-pptx'
import { pdfPageImageRenderScale, renderPdfPagesAsImages } from './pdf-to-images'

interface PdfDocxTextItem {
  str: string
  transform: number[]
  width: number
  height: number
  fontName: string
  hasEOL?: boolean
}

interface PdfTextStyle {
  fontFamily?: string
  fontSubstitution?: string
}

export function pdfDocxTextRun(
  item: PdfDocxTextItem,
  viewportTransform: number[],
  style?: PdfTextStyle,
): PdfDocxTextRun {
  return {
    ...pdfPptxTextRun(item, viewportTransform, style),
    ...(item.hasEOL ? { hasEol: true } : {}),
  }
}

export async function preparePdfDocxPages(
  sourceDocument: PDFDocumentProxy,
  options: Pick<PdfToDocxOptions, 'pageIndexes' | 'mode' | 'renderDpi' | 'includeAnnotations'>,
): Promise<PdfDocxPage[]> {
  const images =
    options.mode === 'fidelity'
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
  const pages: PdfDocxPage[] = []
  for (let index = 0; index < options.pageIndexes.length; index++) {
    const pageIndex = options.pageIndexes[index]!
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= sourceDocument.numPages) {
      throw new Error('Word export page indexes contain an invalid page')
    }
    const page = await sourceDocument.getPage(pageIndex + 1)
    try {
      const viewport = page.getViewport({ scale: 1 })
      const content = await page.getTextContent()
      const styles = content.styles as Record<string, PdfTextStyle>
      const items = content.items.filter(
        (item): item is typeof item & PdfDocxTextItem => 'str' in item && 'transform' in item,
      )
      const imageScale = pdfPageImageRenderScale(viewport.width, viewport.height, options.renderDpi)
      pages.push({
        pageNumber: pageIndex + 1,
        width: viewport.width,
        height: viewport.height,
        textRuns: items.map((item) =>
          pdfDocxTextRun(item, viewport.transform, styles[item.fontName]),
        ),
        ...(images
          ? {
              imageBytes: images[index]!.bytes,
              imageWidth: Math.max(1, Math.floor(viewport.width * imageScale)),
              imageHeight: Math.max(1, Math.floor(viewport.height * imageScale)),
            }
          : {}),
      })
    } finally {
      page.cleanup()
    }
  }
  return pages
}
