import type { PdfColorAdjustments } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'

const MAX_RENDER_PIXELS = 16_000_000
const DEFAULT_RENDER_SCALE = 2

function validateAdjustments(adjustments: PdfColorAdjustments): void {
  const values = [
    ['contrast', adjustments.contrast],
    ['brightness', adjustments.brightness],
    ['saturation', adjustments.saturation],
    ['red', adjustments.red],
    ['green', adjustments.green],
    ['blue', adjustments.blue],
  ] as const
  for (const [name, value] of values) {
    if (!Number.isFinite(value) || value < 0 || value > 200) {
      throw new Error(`${name} must be from 0 to 200`)
    }
  }
}

const clampByte = (value: number): number => Math.min(255, Math.max(0, value))

export function adjustImageDataPixels(
  data: Uint8ClampedArray,
  adjustments: PdfColorAdjustments,
): Uint8ClampedArray {
  validateAdjustments(adjustments)
  const contrast = adjustments.contrast / 100
  const brightness = adjustments.brightness / 100
  const saturation = adjustments.saturation / 100
  const redMultiplier = adjustments.red / 100
  const greenMultiplier = adjustments.green / 100
  const blueMultiplier = adjustments.blue / 100

  const hueToRgb = (p: number, q: number, sourceT: number): number => {
    let t = sourceT
    if (t < 0) t += 1
    if (t > 1) t -= 1
    if (t < 1 / 6) return p + (q - p) * 6 * t
    if (t < 1 / 2) return q
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6
    return p
  }

  for (let index = 0; index < data.length; index += 4) {
    let red = data[index]! * redMultiplier
    let green = data[index + 1]! * greenMultiplier
    let blue = data[index + 2]! * blueMultiplier

    red = clampByte((red - 128) * contrast + 128)
    green = clampByte((green - 128) * contrast + 128)
    blue = clampByte((blue - 128) * contrast + 128)

    red = clampByte(red * brightness)
    green = clampByte(green * brightness)
    blue = clampByte(blue * brightness)

    const normalizedRed = red / 255
    const normalizedGreen = green / 255
    const normalizedBlue = blue / 255
    const maximum = Math.max(normalizedRed, normalizedGreen, normalizedBlue)
    const minimum = Math.min(normalizedRed, normalizedGreen, normalizedBlue)
    const lightness = (maximum + minimum) / 2
    let hue = 0
    let channelSaturation = 0

    if (maximum !== minimum) {
      const delta = maximum - minimum
      channelSaturation =
        lightness > 0.5 ? delta / (2 - maximum - minimum) : delta / (maximum + minimum)
      if (maximum === normalizedRed) {
        hue =
          (normalizedGreen - normalizedBlue) / delta + (normalizedGreen < normalizedBlue ? 6 : 0)
      } else if (maximum === normalizedGreen) {
        hue = (normalizedBlue - normalizedRed) / delta + 2
      } else {
        hue = (normalizedRed - normalizedGreen) / delta + 4
      }
      hue /= 6
    }

    channelSaturation = Math.min(1, Math.max(0, channelSaturation * saturation))
    let adjustedRed: number
    let adjustedGreen: number
    let adjustedBlue: number
    if (channelSaturation === 0) {
      adjustedRed = adjustedGreen = adjustedBlue = lightness
    } else {
      const q =
        lightness < 0.5
          ? lightness * (1 + channelSaturation)
          : lightness + channelSaturation - lightness * channelSaturation
      const p = 2 * lightness - q
      adjustedRed = hueToRgb(p, q, hue + 1 / 3)
      adjustedGreen = hueToRgb(p, q, hue)
      adjustedBlue = hueToRgb(p, q, hue - 1 / 3)
    }

    data[index] = clampByte(Math.round(adjustedRed * 255))
    data[index + 1] = clampByte(Math.round(adjustedGreen * 255))
    data[index + 2] = clampByte(Math.round(adjustedBlue * 255))
  }

  return data
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('Could not encode adjusted PDF page'))
        return
      }
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

export async function renderAdjustedPdfPages(
  sourceDocument: PDFDocumentProxy,
  pageIndexes: number[],
  adjustments: PdfColorAdjustments,
): Promise<Uint8Array[]> {
  validateAdjustments(adjustments)
  const images: Uint8Array[] = []

  for (const pageIndex of pageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= sourceDocument.numPages) {
      throw new Error('Selected page is outside the PDF')
    }
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1, rotation: 0 })
    const pagePixels = Math.max(1, baseViewport.width * baseViewport.height)
    const scale = Math.min(DEFAULT_RENDER_SCALE, Math.sqrt(MAX_RENDER_PIXELS / pagePixels))
    const viewport = page.getViewport({ scale, rotation: 0 })
    const canvas = window.document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')

    await page.render({ canvas, viewport, background: '#ffffff' }).promise
    const imageData = context.getImageData(0, 0, canvas.width, canvas.height)
    adjustImageDataPixels(imageData.data, adjustments)
    context.putImageData(imageData, 0, 0)
    images.push(await canvasPngBytes(canvas))
    canvas.width = 0
    canvas.height = 0
  }

  return images
}
