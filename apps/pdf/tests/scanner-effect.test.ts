import { describe, expect, it } from 'vitest'
import {
  SCANNER_EFFECT_PRESETS,
  applyScannerPixelEffects,
  blurScannerPixels,
  convertScannerColorspace,
  scannerRenderScale,
  softenScannerEdges,
} from '../src/renderer/scanner-effect'

describe('scanner effect pixels', () => {
  it('uses the same quality presets as Stirling', () => {
    expect(SCANNER_EFFECT_PRESETS).toEqual({
      high: { blur: 0.1, noise: 1, brightness: 1.03, contrast: 1.06, renderDpi: 150 },
      medium: { blur: 0.1, noise: 1, brightness: 1.06, contrast: 1.12, renderDpi: 100 },
      low: { blur: 0.9, noise: 2.5, brightness: 1.08, contrast: 1.15, renderDpi: 75 },
    })
  })

  it('converts color pixels to arithmetic grayscale without changing alpha', () => {
    const pixels = new Uint8ClampedArray([30, 60, 90, 77])
    expect(Array.from(convertScannerColorspace(pixels, 'grayscale'))).toEqual([60, 60, 60, 77])
  })

  it('applies contrast, brightness, and yellow paper tone deterministically', () => {
    const pixels = new Uint8ClampedArray([200, 200, 200, 255])
    const result = applyScannerPixelEffects(
      pixels,
      1,
      1,
      { brightness: 1, contrast: 1, noise: 0, yellowish: true },
      () => 0.5,
    )
    expect(result[0]).toBeGreaterThan(result[2]!)
    expect(result[1]).toBeGreaterThan(result[2]!)
    expect(result[3]).toBe(255)
  })

  it('softens outer pixels into the paper gradient while preserving the center', () => {
    const pixels = new Uint8ClampedArray(5 * 5 * 4).fill(255)
    for (let offset = 0; offset < pixels.length; offset += 4) {
      pixels[offset] = pixels[offset + 1] = pixels[offset + 2] = 0
      pixels[offset + 3] = 255
    }
    softenScannerEdges(pixels, 5, 5, 2, { vertical: true, start: 180, end: 180 })
    expect(pixels[0]).toBe(180)
    const center = (2 * 5 + 2) * 4
    expect(pixels[center]).toBe(0)
  })

  it('spreads a bright center pixel using the scanner blur', () => {
    const pixels = new Uint8ClampedArray(5 * 5 * 4)
    for (let offset = 0; offset < pixels.length; offset += 4) pixels[offset + 3] = 255
    const center = (2 * 5 + 2) * 4
    pixels[center] = pixels[center + 1] = pixels[center + 2] = 255
    const blurred = blurScannerPixels(pixels, 5, 5, 100)
    expect(blurred[center]).toBeLessThan(255)
    expect(blurred[(2 * 5 + 1) * 4]).toBeGreaterThan(0)
  })

  it('caps render dimensions and validates the supported DPI range', () => {
    expect(scannerRenderScale(612, 792, 144)).toBe(2)
    expect(scannerRenderScale(10_000, 10_000, 500)).toBeLessThan(1)
    expect(() => scannerRenderScale(612, 792, 71)).toThrow('72 to 500')
    expect(() => scannerRenderScale(612, 792, 100.5)).toThrow('whole number')
  })
})
