import type { PdfScannedImagePage, PdfScannerImageSplitOptions } from '@genoffice/pdf-tools'
import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'

const MAX_RENDER_PIXELS = 16_777_216
const MAX_RENDER_DIMENSION = 8192
const DILATION_RADIUS = 4

export interface ScannedPhotoRegion {
  x: number
  y: number
  width: number
  height: number
  foregroundArea: number
}

interface ScanRun {
  y: number
  start: number
  end: number
  parent: number
}

type Rgb = readonly [number, number, number]

function checkedOptions(options: PdfScannerImageSplitOptions): PdfScannerImageSplitOptions {
  const checks = [
    ['angleThreshold', options.angleThreshold, 0, 45],
    ['tolerance', options.tolerance, 0, 255],
    ['minArea', options.minArea, 1, 100_000_000],
    ['minContourArea', options.minContourArea, 1, 100_000_000],
    ['borderSize', options.borderSize, 0, 200],
    ['renderDpi', options.renderDpi, 72, 300],
  ] as const
  for (const [name, value, minimum, maximum] of checks) {
    if (!Number.isInteger(value) || value < minimum || value > maximum) {
      throw new Error(`${name} must be a whole number from ${minimum} to ${maximum}`)
    }
  }
  return options
}

export function scannerImageRenderScale(
  pageWidth: number,
  pageHeight: number,
  renderDpi: number,
): number {
  if (!Number.isInteger(renderDpi) || renderDpi < 72 || renderDpi > 300) {
    throw new Error('Rendering DPI must be a whole number from 72 to 300')
  }
  const requestedScale = renderDpi / 72
  const pixelScale = Math.sqrt(MAX_RENDER_PIXELS / Math.max(1, pageWidth * pageHeight))
  const dimensionScale = Math.min(
    MAX_RENDER_DIMENSION / Math.max(1, pageWidth),
    MAX_RENDER_DIMENSION / Math.max(1, pageHeight),
  )
  return Math.min(requestedScale, pixelScale, dimensionScale)
}

function median(values: number[]): number {
  values.sort((first, second) => first - second)
  return values[Math.floor(values.length / 2)]!
}

export function estimateScanBackground(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
): Rgb {
  if (width < 1 || height < 1 || pixels.length !== width * height * 4) {
    throw new Error('Scan pixels have invalid dimensions')
  }
  const points = [
    [0, 0],
    [width - 1, 0],
    [width - 1, height - 1],
    [0, height - 1],
    [Math.floor(width / 2), Math.floor(height / 2)],
  ] as const
  const channels = [[], [], []] as [number[], number[], number[]]
  for (const [x, y] of points) {
    const offset = (y * width + x) * 4
    channels[0].push(pixels[offset]!)
    channels[1].push(pixels[offset + 1]!)
    channels[2].push(pixels[offset + 2]!)
  }
  return [median(channels[0]), median(channels[1]), median(channels[2])]
}

export function scanForegroundMask(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  background: Rgb,
  tolerance: number,
): Uint8Array {
  if (!Number.isInteger(tolerance) || tolerance < 0 || tolerance > 255) {
    throw new Error('Scan tolerance must be a whole number from 0 to 255')
  }
  if (pixels.length !== width * height * 4) throw new Error('Scan pixels have invalid dimensions')
  const mask = new Uint8Array(width * height)
  for (let index = 0; index < mask.length; index++) {
    const offset = index * 4
    mask[index] =
      Math.abs(pixels[offset]! - background[0]) > tolerance ||
      Math.abs(pixels[offset + 1]! - background[1]) > tolerance ||
      Math.abs(pixels[offset + 2]! - background[2]) > tolerance
        ? 1
        : 0
  }
  return mask
}

export function dilateScanMask(
  mask: Uint8Array,
  width: number,
  height: number,
  radius = DILATION_RADIUS,
): Uint8Array {
  if (mask.length !== width * height || width < 1 || height < 1 || radius < 0) {
    throw new Error('Scan mask has invalid dimensions')
  }
  if (radius === 0) return new Uint8Array(mask)
  const horizontal = new Uint8Array(mask.length)
  for (let y = 0; y < height; y++) {
    const row = y * width
    let count = 0
    for (let x = 0; x <= Math.min(radius, width - 1); x++) count += mask[row + x]!
    for (let x = 0; x < width; x++) {
      horizontal[row + x] = count > 0 ? 1 : 0
      const removed = x - radius
      const added = x + radius + 1
      if (removed >= 0) count -= mask[row + removed]!
      if (added < width) count += mask[row + added]!
    }
  }

  const output = new Uint8Array(mask.length)
  for (let x = 0; x < width; x++) {
    let count = 0
    for (let y = 0; y <= Math.min(radius, height - 1); y++) count += horizontal[y * width + x]!
    for (let y = 0; y < height; y++) {
      output[y * width + x] = count > 0 ? 1 : 0
      const removed = y - radius
      const added = y + radius + 1
      if (removed >= 0) count -= horizontal[removed * width + x]!
      if (added < height) count += horizontal[added * width + x]!
    }
  }
  return output
}

function rootOf(runs: ScanRun[], index: number): number {
  let root = index
  while (runs[root]!.parent !== root) root = runs[root]!.parent
  while (runs[index]!.parent !== index) {
    const next = runs[index]!.parent
    runs[index]!.parent = root
    index = next
  }
  return root
}

function unionRuns(runs: ScanRun[], first: number, second: number): void {
  const firstRoot = rootOf(runs, first)
  const secondRoot = rootOf(runs, second)
  if (firstRoot !== secondRoot) runs[secondRoot]!.parent = firstRoot
}

export function connectedScanRegions(
  mask: Uint8Array,
  width: number,
  height: number,
  minArea: number,
  minContourArea: number,
): ScannedPhotoRegion[] {
  if (mask.length !== width * height) throw new Error('Scan mask has invalid dimensions')
  const runs: ScanRun[] = []
  let previous: number[] = []
  for (let y = 0; y < height; y++) {
    const current: number[] = []
    let x = 0
    while (x < width) {
      while (x < width && mask[y * width + x] === 0) x++
      if (x >= width) break
      const start = x
      while (x + 1 < width && mask[y * width + x + 1] !== 0) x++
      const end = x
      const runIndex = runs.length
      runs.push({ y, start, end, parent: runIndex })
      current.push(runIndex)
      for (const previousIndex of previous) {
        const previousRun = runs[previousIndex]!
        if (previousRun.end < start - 1) continue
        if (previousRun.start > end + 1) break
        unionRuns(runs, runIndex, previousIndex)
      }
      x++
    }
    previous = current
  }

  const components = new Map<
    number,
    { minX: number; minY: number; maxX: number; maxY: number; area: number }
  >()
  for (let index = 0; index < runs.length; index++) {
    const run = runs[index]!
    const root = rootOf(runs, index)
    const existing = components.get(root)
    if (existing) {
      existing.minX = Math.min(existing.minX, run.start)
      existing.maxX = Math.max(existing.maxX, run.end)
      existing.maxY = run.y
      existing.area += run.end - run.start + 1
    } else {
      components.set(root, {
        minX: run.start,
        minY: run.y,
        maxX: run.end,
        maxY: run.y,
        area: run.end - run.start + 1,
      })
    }
  }

  return [...components.values()]
    .map((component) => ({
      x: component.minX,
      y: component.minY,
      width: component.maxX - component.minX + 1,
      height: component.maxY - component.minY + 1,
      foregroundArea: component.area,
    }))
    .filter(
      (region) =>
        region.width * region.height >= minArea && region.foregroundArea >= minContourArea,
    )
    .sort((first, second) => first.y - second.y || first.x - second.x)
}

export function findScannedPhotoRegions(
  pixels: Uint8ClampedArray,
  width: number,
  height: number,
  tolerance: number,
  minArea: number,
  minContourArea: number,
): ScannedPhotoRegion[] {
  const background = estimateScanBackground(pixels, width, height)
  const mask = scanForegroundMask(pixels, width, height, background, tolerance)
  return connectedScanRegions(
    dilateScanMask(mask, width, height),
    width,
    height,
    minArea,
    minContourArea,
  )
}

export function estimateMaskRotation(
  mask: Uint8Array,
  width: number,
  height: number,
  threshold: number,
): number {
  if (mask.length !== width * height) throw new Error('Scan mask has invalid dimensions')
  let count = 0
  let sumX = 0
  let sumY = 0
  let sumXX = 0
  let sumYY = 0
  let sumXY = 0
  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      if (mask[y * width + x] === 0) continue
      count++
      sumX += x
      sumY += y
      sumXX += x * x
      sumYY += y * y
      sumXY += x * y
    }
  }
  if (count < 2) return 0
  const meanX = sumX / count
  const meanY = sumY / count
  const covarianceX = sumXX / count - meanX * meanX
  const covarianceY = sumYY / count - meanY * meanY
  const covarianceXY = sumXY / count - meanX * meanY
  const discriminant = Math.sqrt((covarianceX - covarianceY) ** 2 + 4 * covarianceXY * covarianceXY)
  const largest = (covarianceX + covarianceY + discriminant) / 2
  const smallest = (covarianceX + covarianceY - discriminant) / 2
  if (largest <= 0 || (largest - smallest) / largest < 0.05) return 0
  let angle = (Math.atan2(2 * covarianceXY, covarianceX - covarianceY) * 90) / Math.PI
  while (angle > 45) angle -= 90
  while (angle < -45) angle += 90
  return Math.abs(angle) >= threshold ? angle : 0
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) return reject(new Error('Could not encode detected scanned image'))
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

function backgroundCss(background: Rgb): string {
  return `rgb(${background[0]}, ${background[1]}, ${background[2]})`
}

export async function renderScannedImagePages(
  sourceDocument: PDFDocumentProxy,
  sourceOptions: PdfScannerImageSplitOptions,
): Promise<PdfScannedImagePage[]> {
  const options = checkedOptions(sourceOptions)
  const output: PdfScannedImagePage[] = []

  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1 })
    const scale = scannerImageRenderScale(
      baseViewport.width,
      baseViewport.height,
      options.renderDpi,
    )
    const viewport = page.getViewport({ scale })
    const sourceCanvas = document.createElement('canvas')
    sourceCanvas.width = Math.max(1, Math.round(viewport.width))
    sourceCanvas.height = Math.max(1, Math.round(viewport.height))
    const sourceContext = sourceCanvas.getContext('2d', { willReadFrequently: true })
    if (!sourceContext) throw new Error('Canvas is unavailable')
    await page.render({
      canvas: sourceCanvas,
      viewport,
      intent: 'print',
      annotationMode: AnnotationMode.ENABLE_STORAGE,
      printAnnotationStorage: sourceDocument.annotationStorage.print,
      background: '#ffffff',
    }).promise
    const sourcePixels = sourceContext.getImageData(0, 0, sourceCanvas.width, sourceCanvas.height)
    const background = estimateScanBackground(
      sourcePixels.data,
      sourceCanvas.width,
      sourceCanvas.height,
    )

    const paddedCanvas = document.createElement('canvas')
    paddedCanvas.width = sourceCanvas.width + options.borderSize * 2
    paddedCanvas.height = sourceCanvas.height + options.borderSize * 2
    const paddedContext = paddedCanvas.getContext('2d', { willReadFrequently: true })
    if (!paddedContext) throw new Error('Canvas is unavailable')
    paddedContext.fillStyle = backgroundCss(background)
    paddedContext.fillRect(0, 0, paddedCanvas.width, paddedCanvas.height)
    paddedContext.drawImage(sourceCanvas, options.borderSize, options.borderSize)
    const paddedPixels = paddedContext.getImageData(0, 0, paddedCanvas.width, paddedCanvas.height)
    const rawMask = scanForegroundMask(
      paddedPixels.data,
      paddedCanvas.width,
      paddedCanvas.height,
      background,
      options.tolerance,
    )
    const mask = dilateScanMask(rawMask, paddedCanvas.width, paddedCanvas.height)
    const regions = connectedScanRegions(
      mask,
      paddedCanvas.width,
      paddedCanvas.height,
      options.minArea,
      options.minContourArea,
    )

    for (const region of regions) {
      const cropCanvas = document.createElement('canvas')
      cropCanvas.width = region.width
      cropCanvas.height = region.height
      const cropContext = cropCanvas.getContext('2d')
      if (!cropContext) throw new Error('Canvas is unavailable')
      cropContext.drawImage(
        paddedCanvas,
        region.x,
        region.y,
        region.width,
        region.height,
        0,
        0,
        region.width,
        region.height,
      )
      const regionMask = new Uint8Array(region.width * region.height)
      for (let y = 0; y < region.height; y++) {
        const sourceStart = (region.y + y) * paddedCanvas.width + region.x
        regionMask.set(mask.subarray(sourceStart, sourceStart + region.width), y * region.width)
      }
      const angle = estimateMaskRotation(
        regionMask,
        region.width,
        region.height,
        options.angleThreshold,
      )
      const rotatedCanvas = document.createElement('canvas')
      rotatedCanvas.width = region.width
      rotatedCanvas.height = region.height
      const rotatedContext = rotatedCanvas.getContext('2d')
      if (!rotatedContext) throw new Error('Canvas is unavailable')
      rotatedContext.fillStyle = backgroundCss(background)
      rotatedContext.fillRect(0, 0, rotatedCanvas.width, rotatedCanvas.height)
      rotatedContext.translate(rotatedCanvas.width / 2, rotatedCanvas.height / 2)
      rotatedContext.rotate((-angle * Math.PI) / 180)
      rotatedContext.drawImage(cropCanvas, -cropCanvas.width / 2, -cropCanvas.height / 2)
      rotatedContext.setTransform(1, 0, 0, 1, 0, 0)

      const trim =
        options.borderSize > 0 &&
        rotatedCanvas.width > options.borderSize * 2 &&
        rotatedCanvas.height > options.borderSize * 2
          ? options.borderSize
          : 0
      const finalCanvas = document.createElement('canvas')
      finalCanvas.width = rotatedCanvas.width - trim * 2
      finalCanvas.height = rotatedCanvas.height - trim * 2
      const finalContext = finalCanvas.getContext('2d')
      if (!finalContext) throw new Error('Canvas is unavailable')
      finalContext.drawImage(
        rotatedCanvas,
        trim,
        trim,
        finalCanvas.width,
        finalCanvas.height,
        0,
        0,
        finalCanvas.width,
        finalCanvas.height,
      )
      output.push({
        image: await canvasPngBytes(finalCanvas),
        width: finalCanvas.width / scale,
        height: finalCanvas.height / scale,
      })
      cropCanvas.width = 0
      cropCanvas.height = 0
      rotatedCanvas.width = 0
      rotatedCanvas.height = 0
      finalCanvas.width = 0
      finalCanvas.height = 0
    }

    sourceCanvas.width = 0
    sourceCanvas.height = 0
    paddedCanvas.width = 0
    paddedCanvas.height = 0
    page.cleanup()
  }

  if (output.length === 0) throw new Error('No scanned images were detected')
  return output
}
