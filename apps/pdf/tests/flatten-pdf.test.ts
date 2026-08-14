import { describe, expect, it } from 'vitest'
import { compressionImageQuality, flattenRenderScale } from '../src/renderer/flatten-pdf'

describe('PDF flatten rendering scale', () => {
  it('converts DPI to PDF.js scale', () => {
    expect(flattenRenderScale(612, 792, 72)).toBe(1)
    expect(flattenRenderScale(612, 792, 144)).toBe(2)
  })

  it('caps oversized pages by the pixel budget', () => {
    expect(flattenRenderScale(10_000, 10_000, 600)).toBeCloseTo(Math.sqrt(0.32))
    expect(flattenRenderScale(10_000, 10_000, 600, 16_000_000)).toBeCloseTo(0.4)
  })

  it('rejects fractional and out-of-range DPI values', () => {
    expect(() => flattenRenderScale(612, 792, 71)).toThrow('72 to 600')
    expect(() => flattenRenderScale(612, 792, 601)).toThrow('72 to 600')
    expect(() => flattenRenderScale(612, 792, 100.5)).toThrow('whole number')
    expect(() => flattenRenderScale(612, 792, 100, 0)).toThrow('pixel budget')
  })
})

describe('PDF compression image quality', () => {
  it('converts a whole percentage to the canvas quality scale', () => {
    expect(compressionImageQuality(10)).toBe(0.1)
    expect(compressionImageQuality(72)).toBe(0.72)
    expect(compressionImageQuality(100)).toBe(1)
  })

  it('rejects fractional and out-of-range values', () => {
    expect(() => compressionImageQuality(9)).toThrow('10 to 100')
    expect(() => compressionImageQuality(101)).toThrow('10 to 100')
    expect(() => compressionImageQuality(72.5)).toThrow('whole percentage')
  })
})
