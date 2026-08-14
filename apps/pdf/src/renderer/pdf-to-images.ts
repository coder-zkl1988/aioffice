import type {
  PdfPageImageColorMode,
  PdfPageImageFormat,
  PdfPageImageOutputMode,
  PdfRenderedPageImage,
} from '@genoffice/pdf-tools'
import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { encodeGifRgba } from './extract-images'
import { convertImagePixelsForPdf } from './images-to-pdf'

const MIN_RENDER_DPI = 72
const MAX_RENDER_DPI = 300
const MAX_RENDER_PIXELS = 32_000_000
const MAX_RENDER_DIMENSION = 16_384

export function pdfPageImageRenderScale(
  pageWidth: number,
  pageHeight: number,
  renderDpi: number,
): number {
  if (!Number.isInteger(renderDpi)) throw new Error('Rendering DPI must be a whole number')
  if (renderDpi < MIN_RENDER_DPI || renderDpi > MAX_RENDER_DPI) {
    throw new Error(`Rendering DPI must be from ${MIN_RENDER_DPI} to ${MAX_RENDER_DPI}`)
  }
  const requestedScale = renderDpi / 72
  const pagePixels = Math.max(1, pageWidth * pageHeight)
  return Math.min(requestedScale, Math.sqrt(MAX_RENDER_PIXELS / pagePixels))
}

export function pdfPageImageQuality(percent: number): number {
  if (!Number.isInteger(percent) || percent < 10 || percent > 100) {
    throw new Error('Image quality must be a whole percentage from 10 to 100')
  }
  return percent / 100
}

function canvasImageBytes(
  canvas: HTMLCanvasElement,
  format: PdfPageImageFormat,
  quality: number,
): Promise<Uint8Array> {
  if (format === 'gif') {
    const context = canvas.getContext('2d')
    if (!context) return Promise.reject(new Error('Canvas is unavailable'))
    return Promise.resolve(
      encodeGifRgba(
        context.getImageData(0, 0, canvas.width, canvas.height).data,
        canvas.width,
        canvas.height,
      ),
    )
  }
  const mimeType = format === 'jpg' ? 'image/jpeg' : `image/${format}`
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) {
          reject(new Error(`Could not encode PDF page as ${format.toUpperCase()}`))
          return
        }
        if (format === 'webp' && blob.type !== mimeType) {
          reject(new Error('This browser does not support WebP image export'))
          return
        }
        void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
      },
      mimeType,
      format === 'png' ? undefined : quality,
    )
  })
}

export function pdfLongImageRenderScale(
  pageSizes: Array<{ width: number; height: number }>,
  renderDpi: number,
): number {
  if (pageSizes.length === 0) throw new Error('Choose at least one page')
  if (!Number.isInteger(renderDpi)) throw new Error('Rendering DPI must be a whole number')
  if (renderDpi < MIN_RENDER_DPI || renderDpi > MAX_RENDER_DPI) {
    throw new Error(`Rendering DPI must be from ${MIN_RENDER_DPI} to ${MAX_RENDER_DPI}`)
  }
  const maximumWidth = Math.max(...pageSizes.map((page) => page.width))
  const totalHeight = pageSizes.reduce((total, page) => total + page.height, 0)
  if (maximumWidth <= 0 || totalHeight <= 0) throw new Error('PDF page size is invalid')
  return Math.min(
    renderDpi / 72,
    MAX_RENDER_DIMENSION / maximumWidth,
    MAX_RENDER_DIMENSION / totalHeight,
    Math.sqrt(MAX_RENDER_PIXELS / (maximumWidth * totalHeight)),
  )
}

async function renderPageCanvas(
  sourceDocument: PDFDocumentProxy,
  pageIndex: number,
  scale: number,
  colorMode: PdfPageImageColorMode,
  includeAnnotations: boolean,
): Promise<HTMLCanvasElement> {
  const page = await sourceDocument.getPage(pageIndex + 1)
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
    annotationMode: includeAnnotations ? AnnotationMode.ENABLE_STORAGE : AnnotationMode.DISABLE,
    ...(includeAnnotations
      ? { printAnnotationStorage: sourceDocument.annotationStorage.print }
      : {}),
    background: '#ffffff',
  }).promise
  if (colorMode !== 'color') {
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height)
    convertImagePixelsForPdf(pixels.data, colorMode)
    context.putImageData(pixels, 0, 0)
  }
  return canvas
}

export async function renderPdfPagesAsImages(
  sourceDocument: PDFDocumentProxy,
  options: {
    pageIndexes: number[]
    format: PdfPageImageFormat
    outputMode: PdfPageImageOutputMode
    renderDpi: number
    imageQuality: number
    colorMode: PdfPageImageColorMode
    includeAnnotations: boolean
  },
): Promise<PdfRenderedPageImage[]> {
  if (!['png', 'jpg', 'gif', 'webp'].includes(options.format)) {
    throw new Error('PDF page image format is invalid')
  }
  if (!['color', 'greyscale', 'blackwhite'].includes(options.colorMode)) {
    throw new Error('PDF page image color mode is invalid')
  }
  if (options.pageIndexes.length === 0) throw new Error('Choose at least one page')
  if (!['single', 'multiple'].includes(options.outputMode)) {
    throw new Error('PDF page image output mode is invalid')
  }
  const quality = pdfPageImageQuality(options.imageQuality)
  const images: PdfRenderedPageImage[] = []

  if (options.outputMode === 'single') {
    const pages = await Promise.all(
      options.pageIndexes.map(async (pageIndex) => {
        if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= sourceDocument.numPages) {
          throw new Error('Selected page is outside the PDF')
        }
        const page = await sourceDocument.getPage(pageIndex + 1)
        return { pageIndex, viewport: page.getViewport({ scale: 1 }) }
      }),
    )
    const scale = pdfLongImageRenderScale(
      pages.map(({ viewport }) => ({ width: viewport.width, height: viewport.height })),
      options.renderDpi,
    )
    const width = Math.max(...pages.map(({ viewport }) => Math.floor(viewport.width * scale)))
    const heights = pages.map(({ viewport }) => Math.max(1, Math.floor(viewport.height * scale)))
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, width)
    canvas.height = heights.reduce((total, height) => total + height, 0)
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')
    context.fillStyle = '#ffffff'
    context.fillRect(0, 0, canvas.width, canvas.height)
    let offsetY = 0
    for (let index = 0; index < pages.length; index++) {
      const pageCanvas = await renderPageCanvas(
        sourceDocument,
        pages[index]!.pageIndex,
        scale,
        options.colorMode,
        options.includeAnnotations,
      )
      context.drawImage(pageCanvas, Math.floor((canvas.width - pageCanvas.width) / 2), offsetY)
      offsetY += pageCanvas.height
      pageCanvas.width = 0
      pageCanvas.height = 0
    }
    images.push({
      pageNumber: options.pageIndexes[0]! + 1,
      bytes: await canvasImageBytes(canvas, options.format, quality),
    })
    canvas.width = 0
    canvas.height = 0
    return images
  }

  for (const pageIndex of options.pageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= sourceDocument.numPages) {
      throw new Error('Selected page is outside the PDF')
    }
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1 })
    const scale = pdfPageImageRenderScale(
      baseViewport.width,
      baseViewport.height,
      options.renderDpi,
    )
    const canvas = await renderPageCanvas(
      sourceDocument,
      pageIndex,
      scale,
      options.colorMode,
      options.includeAnnotations,
    )
    images.push({
      pageNumber: pageIndex + 1,
      bytes: await canvasImageBytes(canvas, options.format, quality),
    })
    canvas.width = 0
    canvas.height = 0
  }

  return images
}
