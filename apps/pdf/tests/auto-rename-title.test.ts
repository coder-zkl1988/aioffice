import { describe, expect, it } from 'vitest'
import { inferPdfAutoRenameTitleFromLines } from '../src/renderer/auto-rename'

describe('inferPdfAutoRenameTitleFromLines', () => {
  it('uses the largest heading and joins consecutive lines with the same font size', () => {
    expect(
      inferPdfAutoRenameTitleFromLines(
        [
          { text: 'Quarterly', fontSize: 26, pageNumber: 1 },
          { text: 'Business Review', fontSize: 26, pageNumber: 1 },
          { text: 'Prepared for the board', fontSize: 12, pageNumber: 1 },
        ],
        'largestHeading',
      ),
    ).toBe('Quarterly Business Review')
  })

  it('supports first-text mode and keeps CJK titles intact', () => {
    const lines = [
      { text: '山东文旅年度报告', fontSize: 18, pageNumber: 1 },
      { text: '附录', fontSize: 30, pageNumber: 2 },
    ]
    expect(inferPdfAutoRenameTitleFromLines(lines, 'firstText')).toBe('山东文旅年度报告')
  })

  it('returns undefined when no usable text exists', () => {
    expect(inferPdfAutoRenameTitleFromLines([], 'largestHeading')).toBeUndefined()
  })
})
