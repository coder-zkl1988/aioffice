import jsQR from 'jsqr'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { PDF_AUTO_SPLIT_QR_CONTENTS } from '@genoffice/pdf-tools'

const MAX_SCAN_PIXELS = 4_000_000
const DEFAULT_RENDER_DPI = 150
const QR_DECODE_SCALES = [1, 0.75, 0.33, 0.2] as const
const VALID_DIVIDER_CONTENTS = new Set<string>(PDF_AUTO_SPLIT_QR_CONTENTS)

export function isPdfAutoSplitDividerData(value: string | null | undefined): boolean {
  return typeof value === 'string' && VALID_DIVIDER_CONTENTS.has(value)
}

export function decodePdfAutoSplitQr(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): string | null {
  if (width <= 0 || height <= 0 || data.length !== width * height * 4) return null
  return jsQR(data, width, height, { inversionAttempts: 'attemptBoth' })?.data ?? null
}

export function decodePdfAutoSplitCanvas(
  canvas: HTMLCanvasElement,
  decode: typeof decodePdfAutoSplitQr = decodePdfAutoSplitQr,
): string | null {
  for (const scale of QR_DECODE_SCALES) {
    const candidate = scale === 1 ? canvas : window.document.createElement('canvas')
    if (scale !== 1) {
      candidate.width = Math.max(1, Math.round(canvas.width * scale))
      candidate.height = Math.max(1, Math.round(canvas.height * scale))
      const candidateContext = candidate.getContext('2d', { willReadFrequently: true })
      if (!candidateContext) throw new Error('Canvas is unavailable')
      candidateContext.drawImage(canvas, 0, 0, candidate.width, candidate.height)
    }
    const context = candidate.getContext('2d', { willReadFrequently: true })
    if (!context) throw new Error('Canvas is unavailable')
    const pixels = context.getImageData(0, 0, candidate.width, candidate.height)
    const value = decode(pixels.data, candidate.width, candidate.height)
    if (scale !== 1) {
      candidate.width = 0
      candidate.height = 0
    }
    if (value) return value
  }
  return null
}

export async function detectPdfAutoSplitDividerPages(
  sourceDocument: PDFDocumentProxy,
  renderDpi = DEFAULT_RENDER_DPI,
  decode: typeof decodePdfAutoSplitQr = decodePdfAutoSplitQr,
): Promise<number[]> {
  if (!Number.isFinite(renderDpi) || renderDpi < 72 || renderDpi > 300) {
    throw new Error('QR scan DPI must be from 72 to 300')
  }
  const dividerPageIndexes: number[] = []
  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    let canvas: HTMLCanvasElement | null = null
    try {
      const baseViewport = page.getViewport({ scale: renderDpi / 72, rotation: 0 })
      const pagePixels = Math.max(1, baseViewport.width * baseViewport.height)
      const scale = Math.min(1, Math.sqrt(MAX_SCAN_PIXELS / pagePixels))
      const viewport = page.getViewport({ scale: (renderDpi / 72) * scale, rotation: 0 })
      canvas = window.document.createElement('canvas')
      canvas.width = Math.max(1, Math.floor(viewport.width))
      canvas.height = Math.max(1, Math.floor(viewport.height))
      const context = canvas.getContext('2d', { willReadFrequently: true })
      if (!context) throw new Error('Canvas is unavailable')
      await page.render({ canvas, canvasContext: context, viewport, background: '#ffffff' }).promise
      if (isPdfAutoSplitDividerData(decodePdfAutoSplitCanvas(canvas, decode))) {
        dividerPageIndexes.push(pageIndex)
      }
    } finally {
      if (canvas) {
        canvas.width = 0
        canvas.height = 0
      }
      page.cleanup()
    }
  }
  return dividerPageIndexes
}
