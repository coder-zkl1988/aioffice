import type { PdfScannerEffectOptions } from '@genoffice/pdf-tools'
import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'

const MAX_RENDER_DPI = 500
const MIN_RENDER_DPI = 72
const MAX_RENDER_PIXELS = 16_777_216
const MAX_RENDER_DIMENSION = 8192

export const SCANNER_EFFECT_PRESETS = {
  high: { blur: 0.1, noise: 1, brightness: 1.03, contrast: 1.06, renderDpi: 150 },
  medium: { blur: 0.1, noise: 1, brightness: 1.06, contrast: 1.12, renderDpi: 100 },
  low: { blur: 0.9, noise: 2.5, brightness: 1.08, contrast: 1.15, renderDpi: 75 },
} as const

const ROTATION_DEGREES = { none: 0, slight: 2, moderate: 5, severe: 8 } as const

interface ScannerGradient {
  vertical: boolean
  start: number
  end: number
}

function checkedScannerEffectOptions(options: PdfScannerEffectOptions): PdfScannerEffectOptions {
  const resolved =
    options.quality === 'custom'
      ? options
      : { ...options, ...SCANNER_EFFECT_PRESETS[options.quality] }
  const checks = [
    ['border', resolved.border, 0, 200],
    ['rotate', resolved.rotate, -15, 15],
    ['rotateVariance', resolved.rotateVariance, 0, 10],
    ['brightness', resolved.brightness, 0.5, 2],
    ['contrast', resolved.contrast, 0.5, 2],
    ['blur', resolved.blur, 0, 10],
    ['noise', resolved.noise, 0, 50],
    ['renderDpi', resolved.renderDpi, MIN_RENDER_DPI, MAX_RENDER_DPI],
  ] as const
  for (const [name, value, minimum, maximum] of checks) {
    if (!Number.isFinite(value) || value < minimum || value > maximum) {
      throw new Error(`${name} must be from ${minimum} to ${maximum}`)
    }
  }
  if (!Number.isInteger(resolved.border)) throw new Error('border must be a whole number')
  if (!Number.isInteger(resolved.renderDpi)) throw new Error('renderDpi must be a whole number')
  if (!Number.isInteger(resolved.seed)) throw new Error('seed must be a whole number')
  return resolved
}

function mulberry32(seed: number): () => number {
  let value = seed >>> 0
  return () => {
    value = (value + 0x6d2b79f5) | 0
    let next = Math.imul(value ^ (value >>> 15), 1 | value)
    next ^= next + Math.imul(next ^ (next >>> 7), 61 | next)
    return ((next ^ (next >>> 14)) >>> 0) / 4294967296
  }
}

function gaussian(random: () => number): number {
  const first = Math.max(Number.EPSILON, random())
  const second = random()
  return Math.sqrt(-2 * Math.log(first)) * Math.cos(2 * Math.PI * second)
}

const clampByte = (value: number): number => Math.min(255, Math.max(0, Math.round(value)))

function gradientAt(
  gradient: ScannerGradient,
  x: number,
  y: number,
  width: number,
  height: number,
) {
  const position = gradient.vertical ? y / Math.max(1, height - 1) : x / Math.max(1, width - 1)
  return clampByte(gradient.start + (gradient.end - gradient.start) * position)
}

function fillGradient(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  gradient: ScannerGradient,
): void {
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const value = gradientAt(gradient, x, y, width, height)
      const offset = (y * width + x) * 4
      data[offset] = value
      data[offset + 1] = value
      data[offset + 2] = value
      data[offset + 3] = 255
    }
  }
}

export function convertScannerColorspace(
  data: Uint8ClampedArray,
  colorspace: PdfScannerEffectOptions['colorspace'],
): Uint8ClampedArray {
  if (colorspace === 'color') return data
  for (let offset = 0; offset < data.length; offset += 4) {
    const gray = Math.round((data[offset]! + data[offset + 1]! + data[offset + 2]!) / 3)
    data[offset] = gray
    data[offset + 1] = gray
    data[offset + 2] = gray
  }
  return data
}

export function applyScannerPixelEffects(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  options: Pick<PdfScannerEffectOptions, 'brightness' | 'contrast' | 'noise' | 'yellowish'>,
  random: () => number,
): Uint8ClampedArray {
  const noiseStrength = (options.noise * Math.min(width, height)) / 1000
  const contrastOffset = 128 - 128 * options.contrast
  for (let offset = 0; offset < data.length; offset += 4) {
    let red = clampByte((data[offset]! * options.contrast + contrastOffset) * options.brightness)
    let green = clampByte(
      (data[offset + 1]! * options.contrast + contrastOffset) * options.brightness,
    )
    let blue = clampByte(
      (data[offset + 2]! * options.contrast + contrastOffset) * options.brightness,
    )

    if (options.yellowish) {
      const brightness = (red + green + blue) / 765
      red = clampByte(red + (255 - red) * 0.18 * brightness)
      green = clampByte(green + (255 - green) * 0.12 * brightness)
      blue = clampByte(blue * (1 - 0.25 * brightness))
    }

    if (noiseStrength > 0) {
      red = clampByte(red + gaussian(random) * noiseStrength)
      green = clampByte(green + gaussian(random) * noiseStrength)
      blue = clampByte(blue + gaussian(random) * noiseStrength)
    }

    data[offset] = red
    data[offset + 1] = green
    data[offset + 2] = blue
    data[offset + 3] = 255
  }
  return data
}

export function softenScannerEdges(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  radius: number,
  gradient: ScannerGradient,
): Uint8ClampedArray {
  if (radius <= 0) return data
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const distance = Math.min(x, width - 1 - x, y, height - 1 - y)
      if (distance >= radius) continue
      const alpha = Math.max(0, distance / radius)
      const background = gradientAt(gradient, x, y, width, height)
      const offset = (y * width + x) * 4
      data[offset] = clampByte(data[offset]! * alpha + background * (1 - alpha))
      data[offset + 1] = clampByte(data[offset + 1]! * alpha + background * (1 - alpha))
      data[offset + 2] = clampByte(data[offset + 2]! * alpha + background * (1 - alpha))
    }
  }
  return data
}

function boxBlurPass(
  source: Uint8ClampedArray,
  target: Uint8ClampedArray,
  width: number,
  height: number,
  radius: number,
  horizontal: boolean,
): void {
  const lineCount = horizontal ? height : width
  const lineLength = horizontal ? width : height
  const diameter = radius * 2 + 1
  for (let line = 0; line < lineCount; line++) {
    const sums = [0, 0, 0]
    const indexAt = (position: number) =>
      (horizontal ? line * width + position : position * width + line) * 4
    for (let position = -radius; position <= radius; position++) {
      const offset = indexAt(Math.max(0, Math.min(lineLength - 1, position)))
      sums[0] += source[offset]!
      sums[1] += source[offset + 1]!
      sums[2] += source[offset + 2]!
    }
    for (let position = 0; position < lineLength; position++) {
      const offset = indexAt(position)
      target[offset] = Math.round(sums[0]! / diameter)
      target[offset + 1] = Math.round(sums[1]! / diameter)
      target[offset + 2] = Math.round(sums[2]! / diameter)
      target[offset + 3] = 255
      const removed = indexAt(Math.max(0, position - radius))
      const added = indexAt(Math.min(lineLength - 1, position + radius + 1))
      sums[0] += source[added]! - source[removed]!
      sums[1] += source[added + 1]! - source[removed + 1]!
      sums[2] += source[added + 2]! - source[removed + 2]!
    }
  }
}

export function blurScannerPixels(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  sigma: number,
): Uint8ClampedArray {
  const scaledSigma = (sigma * Math.min(width, height)) / 1000
  if (scaledSigma <= 0) return data
  const radius = Math.max(1, Math.ceil(scaledSigma * 2))
  let current = new Uint8ClampedArray(data)
  let intermediate = new Uint8ClampedArray(data.length)
  for (let pass = 0; pass < 2; pass++) {
    boxBlurPass(current, intermediate, width, height, radius, true)
    const vertical = new Uint8ClampedArray(data.length)
    boxBlurPass(intermediate, vertical, width, height, radius, false)
    current = vertical
    intermediate = new Uint8ClampedArray(data.length)
  }
  return current
}

export function scannerRenderScale(
  pageWidth: number,
  pageHeight: number,
  renderDpi: number,
): number {
  if (!Number.isInteger(renderDpi) || renderDpi < MIN_RENDER_DPI || renderDpi > MAX_RENDER_DPI) {
    throw new Error(
      `Rendering DPI must be a whole number from ${MIN_RENDER_DPI} to ${MAX_RENDER_DPI}`,
    )
  }
  const requestedScale = renderDpi / 72
  const pixelScale = Math.sqrt(MAX_RENDER_PIXELS / Math.max(1, pageWidth * pageHeight))
  const dimensionScale = Math.min(
    MAX_RENDER_DIMENSION / Math.max(1, pageWidth),
    MAX_RENDER_DIMENSION / Math.max(1, pageHeight),
  )
  return Math.min(requestedScale, pixelScale, dimensionScale)
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error('Could not encode scanner-effect PDF page'))
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

function scannerGradient(random: () => number): ScannerGradient {
  return {
    vertical: random() >= 0.5,
    start: (0.6 + random() * 0.3) * 255,
    end: (0.6 + random() * 0.3) * 255,
  }
}

export async function renderScannerEffectPdfPages(
  sourceDocument: PDFDocumentProxy,
  sourceOptions: PdfScannerEffectOptions,
): Promise<Uint8Array[]> {
  const options = checkedScannerEffectOptions(sourceOptions)
  const images: Uint8Array[] = []

  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const pageRandom = mulberry32((options.seed + Math.imul(pageIndex + 1, 0x9e3779b1)) >>> 0)
    const gradient = scannerGradient(pageRandom)
    const rotation =
      ROTATION_DEGREES[options.rotation] +
      options.rotate +
      (pageRandom() * 2 - 1) * options.rotateVariance
    const baseViewport = page.getViewport({ scale: 1 })
    const scale = scannerRenderScale(baseViewport.width, baseViewport.height, options.renderDpi)
    const viewport = page.getViewport({ scale })
    const sourceCanvas = document.createElement('canvas')
    sourceCanvas.width = Math.max(1, Math.floor(viewport.width))
    sourceCanvas.height = Math.max(1, Math.floor(viewport.height))
    const sourceContext = sourceCanvas.getContext('2d')
    if (!sourceContext) throw new Error('Canvas is unavailable')
    await page.render({
      canvas: sourceCanvas,
      viewport,
      intent: 'print',
      annotationMode: AnnotationMode.ENABLE_STORAGE,
      printAnnotationStorage: sourceDocument.annotationStorage.print,
      background: '#ffffff',
    }).promise
    const sourceData = sourceContext.getImageData(0, 0, sourceCanvas.width, sourceCanvas.height)
    convertScannerColorspace(sourceData.data, options.colorspace)
    sourceContext.putImageData(sourceData, 0, 0)

    const borderedCanvas = document.createElement('canvas')
    borderedCanvas.width = sourceCanvas.width + options.border * 2
    borderedCanvas.height = sourceCanvas.height + options.border * 2
    const borderedContext = borderedCanvas.getContext('2d')
    if (!borderedContext) throw new Error('Canvas is unavailable')
    const borderData = borderedContext.createImageData(borderedCanvas.width, borderedCanvas.height)
    fillGradient(borderData.data, borderedCanvas.width, borderedCanvas.height, gradient)
    borderedContext.putImageData(borderData, 0, 0)
    borderedContext.drawImage(sourceCanvas, options.border, options.border)

    const radians = (rotation * Math.PI) / 180
    const cosine = Math.abs(Math.cos(radians))
    const sine = Math.abs(Math.sin(radians))
    const rotatedCanvas = document.createElement('canvas')
    rotatedCanvas.width = Math.max(
      1,
      Math.floor(borderedCanvas.width * cosine + borderedCanvas.height * sine),
    )
    rotatedCanvas.height = Math.max(
      1,
      Math.floor(borderedCanvas.height * cosine + borderedCanvas.width * sine),
    )
    const rotatedContext = rotatedCanvas.getContext('2d')
    if (!rotatedContext) throw new Error('Canvas is unavailable')
    const rotatedBackground = rotatedContext.createImageData(
      rotatedCanvas.width,
      rotatedCanvas.height,
    )
    fillGradient(rotatedBackground.data, rotatedCanvas.width, rotatedCanvas.height, gradient)
    rotatedContext.putImageData(rotatedBackground, 0, 0)
    rotatedContext.translate(rotatedCanvas.width / 2, rotatedCanvas.height / 2)
    rotatedContext.rotate(radians)
    rotatedContext.drawImage(borderedCanvas, -borderedCanvas.width / 2, -borderedCanvas.height / 2)
    rotatedContext.setTransform(1, 0, 0, 1, 0, 0)

    let processed = rotatedContext.getImageData(0, 0, rotatedCanvas.width, rotatedCanvas.height)
    const featherRadius = Math.max(
      10,
      Math.round(Math.min(rotatedCanvas.width, rotatedCanvas.height) * 0.02),
    )
    softenScannerEdges(
      processed.data,
      rotatedCanvas.width,
      rotatedCanvas.height,
      featherRadius,
      gradient,
    )
    const blurredPixels = blurScannerPixels(
      processed.data,
      rotatedCanvas.width,
      rotatedCanvas.height,
      options.blur,
    )
    processed = rotatedContext.createImageData(rotatedCanvas.width, rotatedCanvas.height)
    processed.data.set(blurredPixels)
    applyScannerPixelEffects(
      processed.data,
      rotatedCanvas.width,
      rotatedCanvas.height,
      options,
      pageRandom,
    )
    rotatedContext.putImageData(processed, 0, 0)
    images.push(await canvasPngBytes(rotatedCanvas))

    sourceCanvas.width = 0
    sourceCanvas.height = 0
    borderedCanvas.width = 0
    borderedCanvas.height = 0
    rotatedCanvas.width = 0
    rotatedCanvas.height = 0
  }

  return images
}
