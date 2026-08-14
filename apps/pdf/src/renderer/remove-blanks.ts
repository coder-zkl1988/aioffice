import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfBlankPageDetectionOptions } from '@genoffice/pdf-tools'

const IMAGE_OPERATORS = new Set([
  OPS.paintImageMaskXObject,
  OPS.paintImageMaskXObjectGroup,
  OPS.paintImageXObject,
  OPS.paintInlineImageXObject,
  OPS.paintInlineImageXObjectGroup,
  OPS.paintImageXObjectRepeat,
  OPS.paintImageMaskXObjectRepeat,
  OPS.paintSolidColorImageMask,
])
const MAX_SCAN_PIXELS = 4_000_000

export function validateBlankPageDetectionOptions(options: PdfBlankPageDetectionOptions): void {
  if (!Number.isFinite(options.threshold) || options.threshold < 0 || options.threshold > 255) {
    throw new Error('threshold must be from 0 to 255')
  }
  if (
    !Number.isFinite(options.whitePercent) ||
    options.whitePercent <= 0 ||
    options.whitePercent > 100
  ) {
    throw new Error('whitePercent must be greater than 0 and at most 100')
  }
}

export function operatorListContainsImage(fnArray: readonly number[]): boolean {
  return fnArray.some((operator) => IMAGE_OPERATORS.has(operator))
}

export function isBlankPagePixels(
  data: Uint8ClampedArray,
  threshold: number,
  whitePercent: number,
): boolean {
  if (data.length === 0 || data.length % 4 !== 0) return false
  const minimumBlue = 255 - threshold
  let whitePixels = 0
  for (let index = 2; index < data.length; index += 4) {
    if (data[index]! >= minimumBlue) whitePixels++
  }
  return (whitePixels / (data.length / 4)) * 100 >= whitePercent
}

export async function detectBlankPdfPages(
  sourceDocument: PDFDocumentProxy,
  options: PdfBlankPageDetectionOptions,
): Promise<number[]> {
  validateBlankPageDetectionOptions(options)
  const blankPageIndexes: number[] = []

  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const textContent = await page.getTextContent()
    const hasText = textContent.items.some(
      (item) => 'str' in item && typeof item.str === 'string' && item.str.trim().length > 0,
    )
    if (hasText) continue

    const operatorList = await page.getOperatorList()
    if (!operatorListContainsImage(operatorList.fnArray)) {
      blankPageIndexes.push(pageIndex)
      continue
    }

    const baseViewport = page.getViewport({ scale: 1, rotation: 0 })
    const pagePixels = Math.max(1, baseViewport.width * baseViewport.height)
    const scale = Math.min(1, Math.sqrt(MAX_SCAN_PIXELS / pagePixels))
    const viewport = page.getViewport({ scale, rotation: 0 })
    const canvas = window.document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')
    await page.render({ canvas, viewport, background: '#ffffff' }).promise
    const imageData = context.getImageData(0, 0, canvas.width, canvas.height)
    if (isBlankPagePixels(imageData.data, options.threshold, options.whitePercent)) {
      blankPageIndexes.push(pageIndex)
    }
    canvas.width = 0
    canvas.height = 0
  }

  return blankPageIndexes
}
