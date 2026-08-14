import { describe, expect, it } from 'vitest'
import { detectDeskewAngle } from '../src/renderer/deskew'

function whitePixels(width: number, height: number): Uint8ClampedArray {
  const pixels = new Uint8ClampedArray(width * height * 4)
  pixels.fill(255)
  return pixels
}

function setBlackPixel(
  pixels: Uint8ClampedArray,
  width: number,
  column: number,
  row: number,
): void {
  if (column < 0 || row < 0 || column >= width || row * width * 4 >= pixels.length) return
  const offset = (row * width + column) * 4
  pixels[offset] = 0
  pixels[offset + 1] = 0
  pixels[offset + 2] = 0
  pixels[offset + 3] = 255
}

function tiltedTextPixels(angle: number): Uint8ClampedArray {
  const width = 240
  const height = 160
  const pixels = whitePixels(width, height)
  const slope = Math.tan((angle * Math.PI) / 180)
  for (const baseline of [32, 62, 92, 122]) {
    for (const [start, end] of [
      [24, 82],
      [94, 146],
      [160, 216],
    ] as const) {
      for (let column = start; column <= end; column++) {
        const center = Math.round(baseline + slope * (column - width / 2))
        for (let thickness = -2; thickness <= 2; thickness++) {
          setBlackPixel(pixels, width, column, center + thickness)
        }
      }
    }
  }
  return pixels
}

describe('PDF deskew detection', () => {
  it('finds the inverse correction for clockwise and counter-clockwise text', () => {
    expect(detectDeskewAngle(tiltedTextPixels(5), 240, 160, 8)).toBeCloseTo(-5, 0)
    expect(detectDeskewAngle(tiltedTextPixels(-3), 240, 160, 8)).toBeCloseTo(3, 0)
  })

  it('leaves blank and already-level pages unchanged', () => {
    expect(detectDeskewAngle(whitePixels(240, 160), 240, 160, 8)).toBe(0)
    expect(detectDeskewAngle(tiltedTextPixels(0), 240, 160, 8)).toBe(0)
  })

  it('validates dimensions and the supported search angle', () => {
    expect(() => detectDeskewAngle(whitePixels(10, 10), 10, 10, 0)).toThrow('0.5 to 15')
    expect(() => detectDeskewAngle(whitePixels(10, 10), 9, 10, 8)).toThrow('image data')
  })
})
