import { describe, expect, it } from 'vitest'
import { adjustImageDataPixels } from '../src/renderer/adjust-colors'

const NEUTRAL = {
  contrast: 100,
  brightness: 100,
  saturation: 100,
  red: 100,
  green: 100,
  blue: 100,
}

describe('adjustImageDataPixels', () => {
  it('leaves pixels and alpha unchanged at neutral settings', () => {
    const pixels = new Uint8ClampedArray([10, 20, 30, 77, 255, 128, 0, 255])
    expect(Array.from(adjustImageDataPixels(pixels, NEUTRAL))).toEqual([
      10, 20, 30, 77, 255, 128, 0, 255,
    ])
  })

  it('applies RGB channel multipliers before the other adjustments', () => {
    const pixels = new Uint8ClampedArray([100, 100, 100, 255])
    expect(
      Array.from(
        adjustImageDataPixels(pixels, {
          ...NEUTRAL,
          red: 200,
          green: 50,
          blue: 0,
        }),
      ),
    ).toEqual([200, 50, 0, 255])
  })

  it('applies contrast and brightness around the same neutral point as Stirling', () => {
    const contrastPixels = new Uint8ClampedArray([20, 128, 240, 255])
    expect(Array.from(adjustImageDataPixels(contrastPixels, { ...NEUTRAL, contrast: 0 }))).toEqual([
      128, 128, 128, 255,
    ])

    const brightnessPixels = new Uint8ClampedArray([200, 200, 200, 255])
    expect(
      Array.from(adjustImageDataPixels(brightnessPixels, { ...NEUTRAL, brightness: 50 })),
    ).toEqual([100, 100, 100, 255])
  })

  it('uses HSL lightness when saturation is removed', () => {
    const pixels = new Uint8ClampedArray([200, 100, 50, 255])
    expect(Array.from(adjustImageDataPixels(pixels, { ...NEUTRAL, saturation: 0 }))).toEqual([
      125, 125, 125, 255,
    ])
  })

  it('rejects values outside the supported range', () => {
    expect(() => adjustImageDataPixels(new Uint8ClampedArray(4), { ...NEUTRAL, red: 201 })).toThrow(
      'red must be from 0 to 200',
    )
  })
})
