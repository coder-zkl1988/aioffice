import { describe, expect, it } from 'vitest'
import {
  connectedScanRegions,
  dilateScanMask,
  estimateMaskRotation,
  estimateScanBackground,
  findScannedPhotoRegions,
  scanForegroundMask,
  scannerImageRenderScale,
} from '../src/renderer/scanner-image-split'

function rgbaPixels(width: number, height: number, color = 255): Uint8ClampedArray {
  const pixels = new Uint8ClampedArray(width * height * 4)
  for (let offset = 0; offset < pixels.length; offset += 4) {
    pixels[offset] = color
    pixels[offset + 1] = color
    pixels[offset + 2] = color
    pixels[offset + 3] = 255
  }
  return pixels
}

function fillRectangle(
  pixels: Uint8ClampedArray,
  width: number,
  x: number,
  y: number,
  rectangleWidth: number,
  rectangleHeight: number,
  color: readonly [number, number, number],
): void {
  for (let row = y; row < y + rectangleHeight; row++) {
    for (let column = x; column < x + rectangleWidth; column++) {
      const offset = (row * width + column) * 4
      pixels[offset] = color[0]
      pixels[offset + 1] = color[1]
      pixels[offset + 2] = color[2]
    }
  }
}

describe('scanned photo detection', () => {
  it('estimates the background from the median of four corners and the center', () => {
    const pixels = rgbaPixels(7, 7)
    fillRectangle(pixels, 7, 2, 2, 3, 3, [0, 0, 0])
    expect(estimateScanBackground(pixels, 7, 7)).toEqual([255, 255, 255])
  })

  it('classifies color channels independently using the configured tolerance', () => {
    const pixels = rgbaPixels(2, 1)
    fillRectangle(pixels, 2, 1, 0, 1, 1, [255, 224, 255])
    expect(Array.from(scanForegroundMask(pixels, 2, 1, [255, 255, 255], 30))).toEqual([0, 1])
    expect(Array.from(scanForegroundMask(pixels, 2, 1, [255, 255, 255], 31))).toEqual([0, 0])
  })

  it('merges nearby fragments with the Stirling-compatible dilation radius', () => {
    const mask = new Uint8Array(20 * 8)
    mask[3 * 20 + 3] = 1
    mask[3 * 20 + 11] = 1
    const regions = connectedScanRegions(dilateScanMask(mask, 20, 8), 20, 8, 1, 1)
    expect(regions).toHaveLength(1)
    expect(regions[0]).toMatchObject({ x: 0, y: 0, width: 16, height: 8 })
  })

  it('filters small regions and sorts photos in reading order', () => {
    const pixels = rgbaPixels(80, 60)
    fillRectangle(pixels, 80, 45, 5, 20, 15, [20, 40, 80])
    fillRectangle(pixels, 80, 5, 35, 25, 15, [80, 30, 20])
    fillRectangle(pixels, 80, 70, 52, 2, 2, [0, 0, 0])
    const regions = findScannedPhotoRegions(pixels, 80, 60, 10, 150, 50)
    expect(regions).toHaveLength(2)
    expect(regions[0]!.x).toBeGreaterThan(regions[1]!.x)
    expect(regions[0]!.y).toBeLessThan(regions[1]!.y)
  })

  it('estimates skew with PCA and respects the correction threshold', () => {
    const width = 80
    const height = 60
    const mask = new Uint8Array(width * height)
    const radians = (15 * Math.PI) / 180
    const cosine = Math.cos(radians)
    const sine = Math.sin(radians)
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const centeredX = x - width / 2
        const centeredY = y - height / 2
        const localX = centeredX * cosine + centeredY * sine
        const localY = -centeredX * sine + centeredY * cosine
        if (Math.abs(localX) <= 25 && Math.abs(localY) <= 8) mask[y * width + x] = 1
      }
    }
    expect(estimateMaskRotation(mask, width, height, 10)).toBeCloseTo(15, 0)
    expect(estimateMaskRotation(mask, width, height, 20)).toBe(0)
  })

  it('caps render dimensions and validates the supported DPI range', () => {
    expect(scannerImageRenderScale(612, 792, 144)).toBe(2)
    expect(scannerImageRenderScale(10_000, 10_000, 300)).toBeLessThan(1)
    expect(() => scannerImageRenderScale(612, 792, 71)).toThrow('72 to 300')
    expect(() => scannerImageRenderScale(612, 792, 100.5)).toThrow('whole number')
  })
})
