import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfLineArtOptions } from '@genoffice/pdf-tools'
import { convertImageDataToLineArt } from './line-art'

const MIN_RENDER_DPI = 72
const MAX_RENDER_DPI = 600
const MAX_RENDER_PIXELS = 32_000_000
const MAX_LINE_ART_RENDER_PIXELS = 16_000_000

function canvasJpegBytes(canvas: HTMLCanvasElement, quality: number): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) return reject(new Error('Could not encode flattened PDF page'))
        void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
      },
      'image/jpeg',
      quality,
    )
  })
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error('Could not encode line-art PDF page'))
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

function checkedImageQuality(imageQuality: number): number {
  if (!Number.isFinite(imageQuality) || imageQuality < 0.1 || imageQuality > 1) {
    throw new Error('Image quality must be from 10% to 100%')
  }
  return imageQuality
}

export function flattenRenderScale(
  pageWidth: number,
  pageHeight: number,
  renderDpi: number,
  maxRenderPixels = MAX_RENDER_PIXELS,
): number {
  if (!Number.isFinite(renderDpi) || !Number.isInteger(renderDpi)) {
    throw new Error('Rendering DPI must be a whole number')
  }
  if (renderDpi < MIN_RENDER_DPI || renderDpi > MAX_RENDER_DPI) {
    throw new Error(`Rendering DPI must be from ${MIN_RENDER_DPI} to ${MAX_RENDER_DPI}`)
  }
  if (!Number.isFinite(maxRenderPixels) || maxRenderPixels < 1) {
    throw new Error('Rendering pixel budget must be positive')
  }
  const requestedScale = renderDpi / 72
  const pagePixels = Math.max(1, pageWidth * pageHeight)
  return Math.min(requestedScale, Math.sqrt(maxRenderPixels / pagePixels))
}

export async function renderFlattenedPdfPages(
  sourceDocument: PDFDocumentProxy,
  renderDpi: number,
  imageQuality = 0.92,
  lineArt?: PdfLineArtOptions,
): Promise<Uint8Array[]> {
  const quality = checkedImageQuality(imageQuality)
  const images: Uint8Array[] = []

  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1 })
    const scale = flattenRenderScale(
      baseViewport.width,
      baseViewport.height,
      renderDpi,
      lineArt ? MAX_LINE_ART_RENDER_PIXELS : MAX_RENDER_PIXELS,
    )
    const viewport = page.getViewport({ scale })
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')

    await page.render({
      canvas,
      viewport,
      intent: 'print',
      annotationMode: AnnotationMode.ENABLE_STORAGE,
      printAnnotationStorage: sourceDocument.annotationStorage.print,
      background: '#ffffff',
    }).promise
    if (lineArt) {
      const imageData = context.getImageData(0, 0, canvas.width, canvas.height)
      convertImageDataToLineArt(imageData.data, canvas.width, canvas.height, lineArt)
      context.putImageData(imageData, 0, 0)
      images.push(await canvasPngBytes(canvas))
    } else {
      images.push(await canvasJpegBytes(canvas, quality))
    }
    canvas.width = 0
    canvas.height = 0
  }

  return images
}

export function compressionImageQuality(percent: number): number {
  if (!Number.isInteger(percent) || percent < 10 || percent > 100) {
    throw new Error('Image quality must be a whole percentage from 10 to 100')
  }
  return checkedImageQuality(percent / 100)
}
