import { describe, expect, it } from 'vitest'
import { convertImageDataToLineArt } from '../src/renderer/line-art'

function pixels(values: number[]): Uint8ClampedArray {
  return new Uint8ClampedArray(values.flatMap((value) => [value, value, value, 255]))
}

describe('line-art conversion', () => {
  it('turns a uniform image into white paper', () => {
    const result = convertImageDataToLineArt(pixels([80, 80, 80, 80]), 2, 2, {
      threshold: 55,
      edgeLevel: 1,
    })
    expect(Array.from(result)).toEqual(new Array(16).fill(255))
  })

  it('keeps high-contrast boundaries as black lines', () => {
    const result = convertImageDataToLineArt(
      pixels([0, 0, 255, 255, 255, 0, 0, 255, 255, 255, 0, 0, 255, 255, 255]),
      5,
      3,
      { threshold: 55, edgeLevel: 1 },
    )
    const firstChannel = Array.from({ length: 15 }, (_, index) => result[index * 4])
    expect(firstChannel).toEqual([255, 0, 0, 255, 255, 255, 0, 0, 255, 255, 255, 0, 0, 255, 255])
  })

  it('uses edge level as the sampling radius', () => {
    const source = pixels([0, 0, 0, 255, 255, 255, 255])
    const light = convertImageDataToLineArt(new Uint8ClampedArray(source), 7, 1, {
      threshold: 55,
      edgeLevel: 1,
    })
    const strong = convertImageDataToLineArt(new Uint8ClampedArray(source), 7, 1, {
      threshold: 55,
      edgeLevel: 3,
    })
    expect(Array.from(strong)).not.toEqual(Array.from(light))
  })

  it('does not turn flat regions black at maximum sensitivity', () => {
    const result = convertImageDataToLineArt(pixels([0, 0, 255, 255, 255]), 5, 1, {
      threshold: 100,
      edgeLevel: 1,
    })
    expect(result[0]).toBe(255)
    expect(result[4 * 4]).toBe(255)
    expect(Array.from(result).some((value, index) => index % 4 !== 3 && value === 0)).toBe(true)
  })

  it('validates thresholds, edge levels, and pixel dimensions', () => {
    expect(() =>
      convertImageDataToLineArt(pixels([0]), 1, 1, { threshold: 101, edgeLevel: 1 }),
    ).toThrow('0 to 100')
    expect(() =>
      convertImageDataToLineArt(pixels([0]), 1, 1, { threshold: 55, edgeLevel: 4 as 1 }),
    ).toThrow('1 to 3')
    expect(() =>
      convertImageDataToLineArt(pixels([0]), 2, 1, { threshold: 55, edgeLevel: 1 }),
    ).toThrow('dimensions')
  })
})
