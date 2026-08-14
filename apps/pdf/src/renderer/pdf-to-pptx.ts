import type { PdfPptxPage, PdfPptxTextRun, PdfToPptxOptions } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfHtmlTextRun } from './pdf-to-html'
import { renderPdfPagesAsImages } from './pdf-to-images'

interface PdfTextItem {
  str: string
  transform: number[]
  width: number
  height: number
  fontName: string
}

interface PdfTextStyle {
  fontFamily?: string
  fontSubstitution?: string
}

export function pdfPptxTextRun(
  item: PdfTextItem,
  viewportTransform: number[],
  style?: PdfTextStyle,
): PdfPptxTextRun {
  const transformed = pdfHtmlTextRun(item, viewportTransform, style)
  return {
    text: transformed.text,
    x: transformed.x,
    y: transformed.y - transformed.fontSize,
    width: Math.max(0, item.width),
    height: Math.max(transformed.fontSize, item.height || 0),
    fontSize: transformed.fontSize,
    angle: transformed.angle,
    ...(transformed.fontFamily ? { fontFamily: transformed.fontFamily } : {}),
    bold: transformed.bold ?? false,
    italic: transformed.italic ?? false,
  }
}

export async function preparePdfPptxPages(
  sourceDocument: PDFDocumentProxy,
  options: Pick<PdfToPptxOptions, 'pageIndexes' | 'mode' | 'renderDpi' | 'includeAnnotations'>,
): Promise<PdfPptxPage[]> {
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
  const pages: PdfPptxPage[] = []
  for (let index = 0; index < options.pageIndexes.length; index++) {
    const pageIndex = options.pageIndexes[index]!
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= sourceDocument.numPages) {
      throw new Error('PowerPoint export page indexes contain an invalid page')
    }
    const page = await sourceDocument.getPage(pageIndex + 1)
    try {
      const viewport = page.getViewport({ scale: 1 })
      const content = await page.getTextContent()
      const styles = content.styles as Record<string, PdfTextStyle>
      const items = content.items.filter(
        (item): item is typeof item & PdfTextItem => 'str' in item && 'transform' in item,
      )
      pages.push({
        pageNumber: pageIndex + 1,
        width: viewport.width,
        height: viewport.height,
        textRuns: items.map((item) =>
          pdfPptxTextRun(item, viewport.transform, styles[item.fontName]),
        ),
        ...(images ? { imageBytes: images[index]!.bytes } : {}),
      })
    } finally {
      page.cleanup()
    }
  }
  return pages
}
