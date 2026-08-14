import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { describe, expect, it } from 'vitest'
import {
  isBlankPagePixels,
  operatorListContainsImage,
  validateBlankPageDetectionOptions,
} from '../src/renderer/remove-blanks'

describe('blank page detection', () => {
  it('recognizes every PDF.js image painting operator', () => {
    expect(operatorListContainsImage([OPS.save, OPS.paintImageXObject, OPS.restore])).toBe(true)
    expect(operatorListContainsImage([OPS.paintInlineImageXObject])).toBe(true)
    expect(operatorListContainsImage([OPS.paintImageMaskXObject])).toBe(true)
    expect(operatorListContainsImage([OPS.moveTo, OPS.lineTo, OPS.stroke])).toBe(false)
  })

  it('uses Stirling blue-channel threshold semantics', () => {
    expect(isBlankPagePixels(new Uint8ClampedArray([0, 0, 245, 255]), 10, 100)).toBe(true)
    expect(isBlankPagePixels(new Uint8ClampedArray([255, 255, 244, 255]), 10, 100)).toBe(false)
    expect(isBlankPagePixels(new Uint8ClampedArray([0, 0, 255, 255]), 0, 100)).toBe(true)
  })

  it('accepts the configured white percentage at its exact boundary', () => {
    const pixels = new Uint8ClampedArray([
      255, 255, 255, 255, 255, 255, 255, 255, 0, 0, 0, 255, 0, 0, 0, 255,
    ])
    expect(isBlankPagePixels(pixels, 0, 50)).toBe(true)
    expect(isBlankPagePixels(pixels, 0, 50.1)).toBe(false)
  })

  it('rejects malformed pixel buffers and unsupported parameter values', () => {
    expect(isBlankPagePixels(new Uint8ClampedArray(), 10, 99.9)).toBe(false)
    expect(isBlankPagePixels(new Uint8ClampedArray([255, 255, 255]), 10, 99.9)).toBe(false)
    expect(() =>
      validateBlankPageDetectionOptions({
        threshold: -1,
        whitePercent: 99.9,
        includeBlankPages: false,
      }),
    ).toThrow('threshold must be from 0 to 255')
    expect(() =>
      validateBlankPageDetectionOptions({
        threshold: 10,
        whitePercent: 0,
        includeBlankPages: false,
      }),
    ).toThrow('whitePercent must be greater than 0 and at most 100')
  })
})
