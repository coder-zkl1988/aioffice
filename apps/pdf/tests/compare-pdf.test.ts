import { describe, expect, it } from 'vitest'
import { comparePixelBuffers, comparisonRenderScale } from '../src/renderer/compare-pdf'

function pixel(red: number, green: number, blue: number): Uint8ClampedArray {
  return new Uint8ClampedArray([red, green, blue, 255])
}

describe('PDF pixel comparison', () => {
  it('uses red for content removed from the base and green for additions', () => {
    const removed = comparePixelBuffers(pixel(0, 0, 0), pixel(255, 255, 255), 1, 1, 0.1)
    expect(Array.from(removed.pixels)).toEqual([255, 59, 48, 255])
    expect(removed.diffPixels).toBe(1)

    const added = comparePixelBuffers(pixel(255, 255, 255), pixel(0, 0, 0), 1, 1, 0.1)
    expect(Array.from(added.pixels)).toEqual([52, 199, 89, 255])
    expect(added.diffPixels).toBe(1)
  })

  it('renders unchanged content as grayscale context', () => {
    const result = comparePixelBuffers(pixel(40, 80, 120), pixel(40, 80, 120), 1, 1, 0.1)
    expect(result.diffPixels).toBe(0)
    expect(result.pixels[0]).toBe(result.pixels[1])
    expect(result.pixels[1]).toBe(result.pixels[2])
    expect(result.pixels[3]).toBe(255)
  })

  it('applies tolerance and rejects invalid buffers', () => {
    expect(
      comparePixelBuffers(pixel(250, 250, 250), pixel(255, 255, 255), 1, 1, 1).diffPixels,
    ).toBe(0)
    expect(() => comparePixelBuffers(pixel(0, 0, 0), pixel(0, 0, 0), 2, 1, 0.1)).toThrow(
      'invalid dimensions',
    )
    expect(() => comparePixelBuffers(pixel(0, 0, 0), pixel(0, 0, 0), 1, 1, 1.1)).toThrow(
      'from 0 to 1',
    )
  })

  it('caps render dimensions and validates the supported DPI range', () => {
    expect(comparisonRenderScale(612, 792, 144)).toBe(2)
    expect(comparisonRenderScale(10_000, 10_000, 300)).toBeLessThan(1)
    expect(() => comparisonRenderScale(612, 792, 71)).toThrow('72 to 300')
    expect(() => comparisonRenderScale(612, 792, 100.5)).toThrow('whole number')
  })
})
