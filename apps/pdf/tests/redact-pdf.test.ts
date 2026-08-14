import { describe, expect, it } from 'vitest'
import {
  matchesByPage,
  paddedViewportRectangle,
  redactionAreasByPage,
  validateRedactionArea,
  validateRedactionOptions,
} from '../src/renderer/redact-pdf'

const validOptions = {
  patterns: ['secret'],
  useRegex: false,
  wholeWord: false,
  color: '#000000',
  padding: 2,
  renderDpi: 144,
}

describe('PDF redaction rendering helpers', () => {
  it('validates text patterns, colors, padding, DPI, and regex syntax', () => {
    expect(() => validateRedactionOptions(validOptions)).not.toThrow()
    expect(() => validateRedactionOptions({ ...validOptions, patterns: [' '] })).toThrow(
      /at least one/i,
    )
    expect(() => validateRedactionOptions({ ...validOptions, color: 'black' })).toThrow('color')
    expect(() => validateRedactionOptions({ ...validOptions, padding: 73 })).toThrow('0 to 72')
    expect(() => validateRedactionOptions({ ...validOptions, renderDpi: 71 })).toThrow('72 to 600')
    expect(() =>
      validateRedactionOptions({ ...validOptions, patterns: ['['], useRegex: true }),
    ).toThrow()
  })

  it('groups every matched rectangle by source page', () => {
    const grouped = matchesByPage([
      { pageIndex: 1, rects: [[10, 20, 30, 40]] },
      { pageIndex: 0, rects: [[1, 2, 3, 4]] },
      { pageIndex: 1, rects: [[50, 60, 70, 80]] },
    ])
    expect(grouped.get(0)).toEqual([[1, 2, 3, 4]])
    expect(grouped.get(1)).toEqual([
      [10, 20, 30, 40],
      [50, 60, 70, 80],
    ])
  })

  it('pads PDF rectangles before converting and normalizes viewport direction', () => {
    expect(
      paddedViewportRectangle([10, 20, 30, 40], 2, ([x1, y1, x2, y2]) => [x2, y2, x1, y1]),
    ).toEqual([8, 18, 32, 42])
  })

  it('validates normalized manual areas and groups them by page', () => {
    const first = { pageIndex: 1, x: 0.1, y: 0.2, width: 0.4, height: 0.3 }
    const second = { pageIndex: 1, x: 0.6, y: 0.1, width: 0.2, height: 0.2 }
    expect(() => validateRedactionArea(first)).not.toThrow()
    expect(() => validateRedactionArea({ ...first, x: 0.8 })).toThrow(/inside/i)
    expect(() => validateRedactionArea({ ...first, pageIndex: -1 })).toThrow(/page/i)
    expect(redactionAreasByPage([first, second]).get(1)).toEqual([first, second])
  })

  it('accepts manual areas without requiring text patterns', () => {
    expect(() =>
      validateRedactionOptions({
        ...validOptions,
        mode: 'areas',
        patterns: [],
        areas: [{ pageIndex: 0, x: 0, y: 0, width: 1, height: 1 }],
      }),
    ).not.toThrow()
    expect(() =>
      validateRedactionOptions({ ...validOptions, mode: 'areas', patterns: [], areas: [] }),
    ).toThrow(/area/i)
  })
})
