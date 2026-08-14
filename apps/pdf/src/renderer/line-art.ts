import type { PdfLineArtOptions } from '@genoffice/pdf-tools'

function checkedLineArtOptions(options: PdfLineArtOptions): PdfLineArtOptions {
  if (!Number.isFinite(options.threshold) || options.threshold < 0 || options.threshold > 100) {
    throw new Error('Line-art threshold must be from 0 to 100')
  }
  if (!Number.isInteger(options.edgeLevel) || options.edgeLevel < 1 || options.edgeLevel > 3) {
    throw new Error('Line-art edge level must be a whole number from 1 to 3')
  }
  return options
}

const pixelIndex = (x: number, y: number, width: number): number => y * width + x

export function convertImageDataToLineArt(
  data: Uint8ClampedArray,
  width: number,
  height: number,
  sourceOptions: PdfLineArtOptions,
): Uint8ClampedArray {
  const options = checkedLineArtOptions(sourceOptions)
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    throw new Error('Line-art image dimensions must be positive whole numbers')
  }
  if (data.length !== width * height * 4) {
    throw new Error('Line-art pixels do not match the image dimensions')
  }

  const grayscale = new Uint8Array(width * height)
  for (let index = 0; index < grayscale.length; index++) {
    const offset = index * 4
    grayscale[index] = Math.round(
      data[offset]! * 0.299 + data[offset + 1]! * 0.587 + data[offset + 2]! * 0.114,
    )
  }

  const magnitudes = new Uint16Array(width * height)
  let maximumMagnitude = 0
  const radius = options.edgeLevel
  const sample = (x: number, y: number): number =>
    grayscale[
      pixelIndex(Math.max(0, Math.min(width - 1, x)), Math.max(0, Math.min(height - 1, y)), width)
    ]!

  for (let y = 0; y < height; y++) {
    for (let x = 0; x < width; x++) {
      const topLeft = sample(x - radius, y - radius)
      const top = sample(x, y - radius)
      const topRight = sample(x + radius, y - radius)
      const left = sample(x - radius, y)
      const right = sample(x + radius, y)
      const bottomLeft = sample(x - radius, y + radius)
      const bottom = sample(x, y + radius)
      const bottomRight = sample(x + radius, y + radius)
      const horizontal = topRight + 2 * right + bottomRight - topLeft - 2 * left - bottomLeft
      const vertical = bottomLeft + 2 * bottom + bottomRight - topLeft - 2 * top - topRight
      const magnitude = Math.round(Math.hypot(horizontal, vertical))
      magnitudes[pixelIndex(x, y, width)] = magnitude
      maximumMagnitude = Math.max(maximumMagnitude, magnitude)
    }
  }

  const cutoff = maximumMagnitude * (1 - options.threshold / 100)
  for (let index = 0; index < magnitudes.length; index++) {
    const line = maximumMagnitude > 0 && magnitudes[index]! > 0 && magnitudes[index]! >= cutoff
    const value = line ? 0 : 255
    const offset = index * 4
    data[offset] = value
    data[offset + 1] = value
    data[offset + 2] = value
    data[offset + 3] = 255
  }
  return data
}
