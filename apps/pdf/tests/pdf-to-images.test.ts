import { describe, expect, it } from 'vitest'
import {
  pdfLongImageRenderScale,
  pdfPageImageQuality,
  pdfPageImageRenderScale,
} from '../src/renderer/pdf-to-images'

describe('PDF page image rendering settings', () => {
  it('converts DPI to a PDF.js render scale', () => {
    expect(pdfPageImageRenderScale(612, 792, 72)).toBe(1)
    expect(pdfPageImageRenderScale(612, 792, 144)).toBe(2)
  })

  it('caps oversized pages by the per-page pixel budget', () => {
    expect(pdfPageImageRenderScale(10_000, 10_000, 300)).toBeCloseTo(Math.sqrt(0.32))
  })

  it('rejects fractional and out-of-range DPI values', () => {
    expect(() => pdfPageImageRenderScale(612, 792, 71)).toThrow('72 to 300')
    expect(() => pdfPageImageRenderScale(612, 792, 301)).toThrow('72 to 300')
    expect(() => pdfPageImageRenderScale(612, 792, 150.5)).toThrow('whole number')
  })

  it('converts and validates image quality percentages', () => {
    expect(pdfPageImageQuality(10)).toBe(0.1)
    expect(pdfPageImageQuality(92)).toBe(0.92)
    expect(pdfPageImageQuality(100)).toBe(1)
    expect(() => pdfPageImageQuality(9)).toThrow('10 to 100')
    expect(() => pdfPageImageQuality(100.5)).toThrow('whole percentage')
  })

  it('caps a combined long image by dimensions and total pixels', () => {
    expect(
      pdfLongImageRenderScale(
        [
          { width: 360, height: 240 },
          { width: 360, height: 240 },
        ],
        144,
      ),
    ).toBe(2)
    expect(pdfLongImageRenderScale([{ width: 1000, height: 100_000 }], 300)).toBeCloseTo(
      16_384 / 100_000,
    )
    expect(pdfLongImageRenderScale([{ width: 10_000, height: 10_000 }], 300)).toBeCloseTo(
      Math.sqrt(0.32),
    )
  })
})
