import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfPageCropBox } from '@genoffice/pdf-tools'

const RENDER_DPI = 150
const MAX_RENDER_PIXELS = 8_000_000
const MAX_RENDER_DIMENSION = 8192

export interface PixelContentBounds {
  left: number
  top: number
  right: number
  bottom: number
}

export interface PdfAutoCropDetectionOptions {
  whiteThreshold: number
  padding: number
}

function checkedAutoCropOptions(options: PdfAutoCropDetectionOptions): PdfAutoCropDetectionOptions {
  if (
    !Number.isInteger(options.whiteThreshold) ||
    options.whiteThreshold < 0 ||
    options.whiteThreshold > 255
  ) {
    throw new Error('White threshold must be a whole number from 0 to 255')
  }
  if (!Number.isFinite(options.padding) || options.padding < 0 || options.padding > 144) {
    throw new Error('Auto crop padding must be from 0 to 144 points')
  }
  return options
}

export function detectContentPixelBounds(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  whiteThreshold: number,
): PixelContentBounds | null {
  if (width < 1 || height < 1 || data.length !== width * height * 4) return null
  let left = width
  let top = height
  let right = -1
  let bottom = -1
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const offset = (y * width + x) * 4
      if (
        data[offset]! >= whiteThreshold &&
        data[offset + 1]! >= whiteThreshold &&
        data[offset + 2]! >= whiteThreshold
      ) {
        continue
      }
      left = Math.min(left, x)
      top = Math.min(top, y)
      right = Math.max(right, x)
      bottom = Math.max(bottom, y)
    }
  }
  return right < left || bottom < top ? null : { left, top, right, bottom }
}

export function autoCropRenderScale(pageWidth: number, pageHeight: number): number {
  const requestedScale = RENDER_DPI / 72
  const pagePixels = Math.max(1, pageWidth * pageHeight)
  return Math.min(
    requestedScale,
    Math.sqrt(MAX_RENDER_PIXELS / pagePixels),
    MAX_RENDER_DIMENSION / Math.max(1, pageWidth),
    MAX_RENDER_DIMENSION / Math.max(1, pageHeight),
  )
}

function convertedPdfBox(
  viewport: {
    width: number
    height: number
    viewBox: number[]
    convertToPdfPoint(x: number, y: number): number[]
  },
  bounds: PixelContentBounds | null,
  canvasWidth: number,
  canvasHeight: number,
  padding: number,
): PdfPageCropBox {
  if (!bounds) {
    const [firstX = 0, firstY = 0, secondX = firstX, secondY = firstY] = viewport.viewBox
    return {
      x: Math.min(firstX, secondX),
      y: Math.min(firstY, secondY),
      width: Math.abs(secondX - firstX),
      height: Math.abs(secondY - firstY),
    }
  }
  const pdfPoint = (x: number, y: number): [number, number] => {
    const point = viewport.convertToPdfPoint(
      (x * viewport.width) / canvasWidth,
      (y * viewport.height) / canvasHeight,
    )
    return [point[0] ?? 0, point[1] ?? 0]
  }
  const contentPoints = [
    pdfPoint(bounds.left, bounds.top),
    pdfPoint(bounds.right + 1, bounds.top),
    pdfPoint(bounds.left, bounds.bottom + 1),
    pdfPoint(bounds.right + 1, bounds.bottom + 1),
  ]
  const pagePoints = [
    pdfPoint(0, 0),
    pdfPoint(canvasWidth, 0),
    pdfPoint(0, canvasHeight),
    pdfPoint(canvasWidth, canvasHeight),
  ]
  const contentX = contentPoints.map(([x]) => x)
  const contentY = contentPoints.map(([, y]) => y)
  const pageX = pagePoints.map(([x]) => x)
  const pageY = pagePoints.map(([, y]) => y)
  const pageLeft = Math.min(...pageX)
  const pageBottom = Math.min(...pageY)
  const pageRight = Math.max(...pageX)
  const pageTop = Math.max(...pageY)
  const left = Math.max(pageLeft, Math.min(...contentX) - padding)
  const bottom = Math.max(pageBottom, Math.min(...contentY) - padding)
  const right = Math.min(pageRight, Math.max(...contentX) + padding)
  const top = Math.min(pageTop, Math.max(...contentY) + padding)
  return { x: left, y: bottom, width: right - left, height: top - bottom }
}

export async function detectPdfAutoCropBoxes(
  sourceDocument: PDFDocumentProxy,
  sourceOptions: PdfAutoCropDetectionOptions,
): Promise<PdfPageCropBox[]> {
  const options = checkedAutoCropOptions(sourceOptions)
  const pageBoxes: PdfPageCropBox[] = []
  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1 })
    const viewport = page.getViewport({
      scale: autoCropRenderScale(baseViewport.width, baseViewport.height),
    })
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')
    try {
      await page.render({
        canvas,
        viewport,
        intent: 'print',
        annotationMode: AnnotationMode.ENABLE_STORAGE,
        printAnnotationStorage: sourceDocument.annotationStorage.print,
        background: '#ffffff',
      }).promise
      const pixels = context.getImageData(0, 0, canvas.width, canvas.height)
      pageBoxes.push(
        convertedPdfBox(
          viewport,
          detectContentPixelBounds(
            pixels.data,
            canvas.width,
            canvas.height,
            options.whiteThreshold,
          ),
          canvas.width,
          canvas.height,
          options.padding,
        ),
      )
    } finally {
      canvas.width = 0
      canvas.height = 0
    }
  }
  return pageBoxes
}
