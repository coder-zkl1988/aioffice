import type { PdfComparisonPage } from '@genoffice/pdf-tools'
import pixelmatch from 'pixelmatch'
import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'

const CSS_DPI = 72
const MIN_RENDER_DPI = 72
const MAX_RENDER_DPI = 300
const MAX_RENDER_PIXELS = 16_777_216
const MAX_RENDER_DIMENSION = 8192
const REMOVAL_COLOR: [number, number, number] = [255, 59, 48]
const ADDITION_COLOR: [number, number, number] = [52, 199, 89]

interface RenderedComparisonSide {
  canvas: HTMLCanvasElement
  pixels: ImageData
}

export interface PdfComparisonSummary {
  pages: PdfComparisonPage[]
  diffPixels: number
  totalPixels: number
  pagesWithChanges: number
}

export function comparisonRenderScale(
  pageWidth: number,
  pageHeight: number,
  renderDpi: number,
): number {
  if (!Number.isInteger(renderDpi) || renderDpi < MIN_RENDER_DPI || renderDpi > MAX_RENDER_DPI) {
    throw new Error(
      `Rendering DPI must be a whole number from ${MIN_RENDER_DPI} to ${MAX_RENDER_DPI}`,
    )
  }
  const requestedScale = renderDpi / CSS_DPI
  const pixelScale = Math.sqrt(MAX_RENDER_PIXELS / Math.max(1, pageWidth * pageHeight))
  const dimensionScale = Math.min(
    MAX_RENDER_DIMENSION / Math.max(1, pageWidth),
    MAX_RENDER_DIMENSION / Math.max(1, pageHeight),
  )
  return Math.min(requestedScale, pixelScale, dimensionScale)
}

export function comparePixelBuffers(
  base: Uint8ClampedArray,
  comparison: Uint8ClampedArray,
  width: number,
  height: number,
  threshold: number,
): { pixels: Uint8ClampedArray; diffPixels: number } {
  if (!Number.isFinite(threshold) || threshold < 0 || threshold > 1) {
    throw new Error('Comparison threshold must be from 0 to 1')
  }
  const expectedLength = width * height * 4
  if (
    width < 1 ||
    height < 1 ||
    base.length !== expectedLength ||
    comparison.length !== expectedLength
  ) {
    throw new Error('Comparison pixel buffers have invalid dimensions')
  }
  const pixels = new Uint8ClampedArray(expectedLength)
  const diffPixels = pixelmatch(base, comparison, pixels, width, height, {
    threshold,
    includeAA: true,
    alpha: 0.3,
    diffColor: REMOVAL_COLOR,
    diffColorAlt: ADDITION_COLOR,
  })
  return { pixels, diffPixels }
}

function blankCanvas(width: number, height: number): RenderedComparisonSide {
  const canvas = document.createElement('canvas')
  canvas.width = width
  canvas.height = height
  const context = canvas.getContext('2d', { willReadFrequently: true })
  if (!context) throw new Error('Canvas is unavailable')
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, width, height)
  return { canvas, pixels: context.getImageData(0, 0, width, height) }
}

async function renderPage(
  page: PDFPageProxy | null,
  documentProxy: PDFDocumentProxy,
  scale: number,
  width: number,
  height: number,
): Promise<RenderedComparisonSide> {
  const rendered = blankCanvas(width, height)
  if (!page) return rendered
  const context = rendered.canvas.getContext('2d', { willReadFrequently: true })
  if (!context) throw new Error('Canvas is unavailable')
  const viewport = page.getViewport({ scale })
  context.save()
  context.translate(
    Math.round((width - viewport.width) / 2),
    Math.round((height - viewport.height) / 2),
  )
  await page.render({
    canvas: rendered.canvas,
    canvasContext: context,
    viewport,
    intent: 'print',
    annotationMode: AnnotationMode.ENABLE_STORAGE,
    printAnnotationStorage: documentProxy.annotationStorage.print,
    background: '#ffffff',
  }).promise
  context.restore()
  rendered.pixels = context.getImageData(0, 0, width, height)
  return rendered
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error('Could not encode comparison PDF page'))
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

export async function renderComparedPdfPages(
  baseDocument: PDFDocumentProxy,
  comparisonDocument: PDFDocumentProxy,
  renderDpi: number,
  threshold: number,
): Promise<PdfComparisonSummary> {
  const totalPages = Math.max(baseDocument.numPages, comparisonDocument.numPages)
  if (totalPages === 0) throw new Error('PDF has no pages')
  const pages: PdfComparisonPage[] = []
  let diffPixels = 0
  let totalPixels = 0
  let pagesWithChanges = 0

  for (let pageNumber = 1; pageNumber <= totalPages; pageNumber++) {
    const basePage =
      pageNumber <= baseDocument.numPages ? await baseDocument.getPage(pageNumber) : null
    const comparisonPage =
      pageNumber <= comparisonDocument.numPages
        ? await comparisonDocument.getPage(pageNumber)
        : null
    try {
      const baseViewport = basePage?.getViewport({ scale: 1 })
      const comparisonViewport = comparisonPage?.getViewport({ scale: 1 })
      const pageWidth = Math.max(baseViewport?.width ?? 1, comparisonViewport?.width ?? 1)
      const pageHeight = Math.max(baseViewport?.height ?? 1, comparisonViewport?.height ?? 1)
      const scale = comparisonRenderScale(pageWidth, pageHeight, renderDpi)
      const width = Math.max(1, Math.round(pageWidth * scale))
      const height = Math.max(1, Math.round(pageHeight * scale))
      const [base, comparison] = await Promise.all([
        renderPage(basePage, baseDocument, scale, width, height),
        renderPage(comparisonPage, comparisonDocument, scale, width, height),
      ])
      const result = comparePixelBuffers(
        base.pixels.data,
        comparison.pixels.data,
        width,
        height,
        threshold,
      )
      const outputCanvas = document.createElement('canvas')
      outputCanvas.width = width
      outputCanvas.height = height
      const outputContext = outputCanvas.getContext('2d')
      if (!outputContext) throw new Error('Canvas is unavailable')
      const outputPixels = outputContext.createImageData(width, height)
      outputPixels.data.set(result.pixels)
      outputContext.putImageData(outputPixels, 0, 0)
      pages.push({
        image: await canvasPngBytes(outputCanvas),
        width: width / scale,
        height: height / scale,
      })
      diffPixels += result.diffPixels
      totalPixels += width * height
      if (result.diffPixels > 0) pagesWithChanges += 1
      base.canvas.width = 0
      base.canvas.height = 0
      comparison.canvas.width = 0
      comparison.canvas.height = 0
      outputCanvas.width = 0
      outputCanvas.height = 0
    } finally {
      basePage?.cleanup()
      comparisonPage?.cleanup()
    }
  }

  return { pages, diffPixels, totalPixels, pagesWithChanges }
}
