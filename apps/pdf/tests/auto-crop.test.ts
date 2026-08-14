import { describe, expect, it } from 'vitest'
import { autoCropRenderScale, detectContentPixelBounds } from '../src/renderer/auto-crop'

function pixels(width: number, height: number): Uint8ClampedArray {
  const data = new Uint8ClampedArray(width * height * 4)
  data.fill(255)
  return data
}

function setPixel(
  data: Uint8ClampedArray,
  width: number,
  x: number,
  y: number,
  value: number,
): void {
  const offset = (y * width + x) * 4
  data[offset] = value
  data[offset + 1] = value
  data[offset + 2] = value
  data[offset + 3] = 255
}

describe('detectContentPixelBounds', () => {
  it('finds the outer non-white pixel rectangle', () => {
    const data = pixels(8, 6)
    setPixel(data, 8, 2, 1, 0)
    setPixel(data, 8, 6, 4, 120)
    expect(detectContentPixelBounds(data, 8, 6, 250)).toEqual({
      left: 2,
      top: 1,
      right: 6,
      bottom: 4,
    })
  })

  it('returns null for blank pages and honors the white threshold', () => {
    const data = pixels(3, 2)
    setPixel(data, 3, 1, 1, 245)
    expect(detectContentPixelBounds(data, 3, 2, 240)).toBeNull()
    expect(detectContentPixelBounds(data, 3, 2, 250)).toEqual({
      left: 1,
      top: 1,
      right: 1,
      bottom: 1,
    })
  })

  it('caps render scale for very large pages', () => {
    expect(autoCropRenderScale(612, 792)).toBeCloseTo(150 / 72)
    expect(autoCropRenderScale(20_000, 20_000)).toBeLessThan(1)
  })
})
