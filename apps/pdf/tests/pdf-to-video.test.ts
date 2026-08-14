import { describe, expect, it } from 'vitest'
import { pdfVideoDimensions, validatePdfVideoSettings } from '../src/renderer/pdf-to-video'

describe('PDF video settings', () => {
  it('maps the supported 16:9 output resolutions', () => {
    expect(pdfVideoDimensions('480p')).toEqual({ width: 854, height: 480 })
    expect(pdfVideoDimensions('720p')).toEqual({ width: 1280, height: 720 })
    expect(pdfVideoDimensions('1080p')).toEqual({ width: 1920, height: 1080 })
  })

  it('accepts bounded page selections and timing', () => {
    expect(() =>
      validatePdfVideoSettings({
        pageCount: 4,
        pageIndexes: [0, 2, 3],
        secondsPerPage: 2,
        transitionSeconds: 0.4,
        resolution: '720p',
      }),
    ).not.toThrow()
  })

  it('rejects invalid pages, timing, and oversized videos', () => {
    const valid = {
      pageCount: 4,
      pageIndexes: [0],
      secondsPerPage: 2,
      transitionSeconds: 0.4,
      resolution: '720p' as const,
    }
    expect(() => validatePdfVideoSettings({ ...valid, pageIndexes: [4] })).toThrow(
      'outside the PDF',
    )
    expect(() => validatePdfVideoSettings({ ...valid, secondsPerPage: 0 })).toThrow('1 to 10')
    expect(() => validatePdfVideoSettings({ ...valid, transitionSeconds: 1.1 })).toThrow('0 to 1')
    expect(() =>
      validatePdfVideoSettings({
        ...valid,
        pageCount: 101,
        pageIndexes: Array.from({ length: 101 }, (_, index) => index),
      }),
    ).toThrow('100 pages')
  })
})
