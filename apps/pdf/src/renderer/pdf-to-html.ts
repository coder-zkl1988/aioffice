import type { PdfHtmlPage, PdfHtmlTextRun } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { renderPdfPagesAsImages } from './pdf-to-images'

interface PdfTextItem {
  str: string
  transform: number[]
  fontName: string
  hasEOL?: boolean
}

interface PdfTextStyle {
  fontFamily?: string
  fontSubstitution?: string
}

export function pdfHtmlTextRun(
  item: PdfTextItem,
  viewportTransform: number[],
  style?: PdfTextStyle,
): PdfHtmlTextRun {
  const [a, b, c, d, e, f] = item.transform
  const [v0, v1, v2, v3, v4, v5] = viewportTransform
  const transformed = [
    v0! * a! + v2! * b!,
    v1! * a! + v3! * b!,
    v0! * c! + v2! * d!,
    v1! * c! + v3! * d!,
    v0! * e! + v2! * f! + v4!,
    v1! * e! + v3! * f! + v5!,
  ]
  const family = style?.fontFamily ?? style?.fontSubstitution
  const fontName = `${item.fontName} ${family ?? ''}`
  return {
    text: item.str,
    x: transformed[4]!,
    y: transformed[5]!,
    fontSize: Math.max(0.1, Math.hypot(transformed[2]!, transformed[3]!)),
    angle: (Math.atan2(transformed[1]!, transformed[0]!) * 180) / Math.PI,
    ...(family ? { fontFamily: family } : {}),
    ...(/(?:bold|black|heavy|semibold)/i.test(fontName) ? { bold: true } : {}),
    ...(/(?:italic|oblique)/i.test(fontName) ? { italic: true } : {}),
  }
}

export async function preparePdfHtmlPages(
  sourceDocument: PDFDocumentProxy,
  options: {
    pageIndexes: number[]
    renderDpi: number
    includeAnnotations: boolean
  },
): Promise<PdfHtmlPage[]> {
  const images = await renderPdfPagesAsImages(sourceDocument, {
    pageIndexes: options.pageIndexes,
    format: 'png',
    outputMode: 'multiple',
    renderDpi: options.renderDpi,
    imageQuality: 100,
    colorMode: 'color',
    includeAnnotations: options.includeAnnotations,
  })
  const pages: PdfHtmlPage[] = []
  for (let index = 0; index < options.pageIndexes.length; index++) {
    const pageIndex = options.pageIndexes[index]!
    const page = await sourceDocument.getPage(pageIndex + 1)
    try {
      const viewport = page.getViewport({ scale: 1 })
      const textContent = await page.getTextContent()
      const styles = textContent.styles as Record<string, PdfTextStyle>
      const items = textContent.items.filter(
        (item): item is typeof item & PdfTextItem => 'str' in item && 'transform' in item,
      )
      pages.push({
        pageNumber: pageIndex + 1,
        width: viewport.width,
        height: viewport.height,
        imageBytes: images[index]!.bytes,
        text: items.map((item) => item.str + (item.hasEOL ? '\n' : '')).join(''),
        textRuns: items.map((item) =>
          pdfHtmlTextRun(item, viewport.transform, styles[item.fontName]),
        ),
      })
    } finally {
      page.cleanup()
    }
  }
  return pages
}
