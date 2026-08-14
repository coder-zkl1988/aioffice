import type { PdfDeskewOptions, PdfDeskewPage } from '@genoffice/pdf-tools'
import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfPageImageRenderScale } from './pdf-to-images'

const MAX_ANALYSIS_DIMENSION = 900
const MIN_FOREGROUND_PIXELS = 160
const MIN_ANGLE = 0.15
const MIN_SCORE_IMPROVEMENT = 1.02

function checkedMaxAngle(maxAngle: number): number {
  if (!Number.isFinite(maxAngle) || maxAngle < 0.5 || maxAngle > 15) {
    throw new Error('Deskew maximum angle must be from 0.5 to 15 degrees')
  }
  return maxAngle
}

function luminance(data: Uint8ClampedArray, offset: number): number {
  return data[offset]! * 0.299 + data[offset + 1]! * 0.587 + data[offset + 2]! * 0.114
}

function otsuThreshold(data: Uint8ClampedArray): number {
  const histogram = new Uint32Array(256)
  let total = 0
  let sum = 0
  for (let offset = 0; offset < data.length; offset += 4) {
    const value = Math.round(luminance(data, offset))
    histogram[value]++
    total++
    sum += value
  }
  let backgroundWeight = 0
  let backgroundSum = 0
  let bestVariance = -1
  let threshold = 180
  for (let value = 0; value < 256; value++) {
    backgroundWeight += histogram[value]!
    if (backgroundWeight === 0) continue
    const foregroundWeight = total - backgroundWeight
    if (foregroundWeight === 0) break
    backgroundSum += value * histogram[value]!
    const backgroundMean = backgroundSum / backgroundWeight
    const foregroundMean = (sum - backgroundSum) / foregroundWeight
    const variance = backgroundWeight * foregroundWeight * (backgroundMean - foregroundMean) ** 2
    if (variance > bestVariance) {
      bestVariance = variance
      threshold = value
    }
  }
  return Math.min(220, Math.max(80, threshold))
}

function foregroundPoints(
  data: Uint8ClampedArray,
  width: number,
  height: number,
): Array<readonly [number, number]> {
  const threshold = otsuThreshold(data)
  const marginX = Math.floor(width * 0.02)
  const marginY = Math.floor(height * 0.02)
  const points: Array<readonly [number, number]> = []
  for (let row = marginY; row < height - marginY; row++) {
    for (let column = marginX; column < width - marginX; column++) {
      const offset = (row * width + column) * 4
      if (data[offset + 3]! < 128 || luminance(data, offset) > threshold) continue
      points.push([column, row])
    }
  }
  return points
}

function projectionScore(
  points: Array<readonly [number, number]>,
  width: number,
  height: number,
  angle: number,
): number {
  const radians = (angle * Math.PI) / 180
  const sine = Math.sin(radians)
  const cosine = Math.cos(radians)
  const offset = Math.ceil(width * Math.abs(sine)) + 2
  const rows = new Uint32Array(Math.ceil(height * Math.abs(cosine) + offset * 2 + 4))
  for (const [column, row] of points) {
    const projected = Math.round(column * sine + row * cosine) + offset
    if (projected >= 0 && projected < rows.length) rows[projected]++
  }
  let score = 0
  for (let row = 1; row < rows.length; row++) {
    const difference = rows[row]! - rows[row - 1]!
    score += difference * difference
  }
  return score
}

function bestProjectionAngle(
  points: Array<readonly [number, number]>,
  width: number,
  height: number,
  minimum: number,
  maximum: number,
  step: number,
): { angle: number; score: number } {
  let best = { angle: 0, score: -1 }
  for (let angle = minimum; angle <= maximum + step / 2; angle += step) {
    const rounded = Math.round(angle * 1000) / 1000
    const score = projectionScore(points, width, height, rounded)
    if (score > best.score) best = { angle: rounded, score }
  }
  return best
}

export function detectDeskewAngle(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  sourceMaxAngle: number,
): number {
  const maxAngle = checkedMaxAngle(sourceMaxAngle)
  if (width < 1 || height < 1 || data.length !== width * height * 4) {
    throw new Error('Deskew image data is invalid')
  }
  const points = foregroundPoints(data, width, height)
  if (points.length < MIN_FOREGROUND_PIXELS) return 0
  const baseline = projectionScore(points, width, height, 0)
  const coarse = bestProjectionAngle(points, width, height, -maxAngle, maxAngle, 0.5)
  const fine = bestProjectionAngle(
    points,
    width,
    height,
    Math.max(-maxAngle, coarse.angle - 0.5),
    Math.min(maxAngle, coarse.angle + 0.5),
    0.1,
  )
  if (Math.abs(fine.angle) < MIN_ANGLE || fine.score < baseline * MIN_SCORE_IMPROVEMENT) return 0
  return Math.round(fine.angle * 10) / 10
}

function analysisPixels(canvas: HTMLCanvasElement): ImageData {
  const scale = Math.min(1, MAX_ANALYSIS_DIMENSION / Math.max(canvas.width, canvas.height))
  if (scale === 1) {
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')
    return context.getImageData(0, 0, canvas.width, canvas.height)
  }
  const analysis = document.createElement('canvas')
  analysis.width = Math.max(1, Math.round(canvas.width * scale))
  analysis.height = Math.max(1, Math.round(canvas.height * scale))
  const context = analysis.getContext('2d')
  if (!context) throw new Error('Canvas is unavailable')
  context.drawImage(canvas, 0, 0, analysis.width, analysis.height)
  const pixels = context.getImageData(0, 0, analysis.width, analysis.height)
  analysis.width = 0
  analysis.height = 0
  return pixels
}

function correctedCanvas(source: HTMLCanvasElement, angle: number): HTMLCanvasElement {
  const output = document.createElement('canvas')
  output.width = source.width
  output.height = source.height
  const context = output.getContext('2d')
  if (!context) throw new Error('Canvas is unavailable')
  context.fillStyle = '#ffffff'
  context.fillRect(0, 0, output.width, output.height)
  context.translate(output.width / 2, output.height / 2)
  context.rotate((angle * Math.PI) / 180)
  context.drawImage(source, -source.width / 2, -source.height / 2)
  return output
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error('Could not encode deskewed PDF page'))
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

export async function prepareDeskewPdfPages(
  sourceDocument: PDFDocumentProxy,
  options: Pick<PdfDeskewOptions, 'pageIndexes' | 'maxAngle' | 'renderDpi' | 'includeAnnotations'>,
): Promise<PdfDeskewPage[]> {
  const maxAngle = checkedMaxAngle(options.maxAngle)
  const pages: PdfDeskewPage[] = []
  for (const pageIndex of options.pageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= sourceDocument.numPages) {
      throw new Error('Deskew page indexes contain an invalid page')
    }
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1 })
    const viewport = page.getViewport({
      scale: pdfPageImageRenderScale(baseViewport.width, baseViewport.height, options.renderDpi),
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
        annotationMode: options.includeAnnotations
          ? AnnotationMode.ENABLE_STORAGE
          : AnnotationMode.DISABLE,
        ...(options.includeAnnotations
          ? { printAnnotationStorage: sourceDocument.annotationStorage.print }
          : {}),
        background: '#ffffff',
      }).promise
      const pixels = analysisPixels(canvas)
      const angle = detectDeskewAngle(pixels.data, pixels.width, pixels.height, maxAngle)
      if (angle === 0) continue
      const corrected = correctedCanvas(canvas, angle)
      pages.push({ pageIndex, angle, image: await canvasPngBytes(corrected) })
      corrected.width = 0
      corrected.height = 0
    } finally {
      page.cleanup()
      canvas.width = 0
      canvas.height = 0
    }
  }
  return pages
}
