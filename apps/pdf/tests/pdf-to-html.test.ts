import { describe, expect, it } from 'vitest'
import { pdfHtmlTextRun } from '../src/renderer/pdf-to-html'

describe('pdfHtmlTextRun', () => {
  it('maps PDF text coordinates through the viewport transform', () => {
    expect(
      pdfHtmlTextRun(
        { str: 'Hello', transform: [12, 0, 0, 12, 30, 40], fontName: 'BodyBold' },
        [1, 0, 0, -1, 0, 200],
        { fontFamily: 'Inter' },
      ),
    ).toEqual({
      text: 'Hello',
      x: 30,
      y: 160,
      fontSize: 12,
      angle: 0,
      fontFamily: 'Inter',
      bold: true,
    })
  })

  it('preserves rotated and italic text metadata', () => {
    const run = pdfHtmlTextRun(
      { str: 'Vertical', transform: [0, 9, -9, 0, 15, 25], fontName: 'BodyItalic' },
      [1, 0, 0, 1, 0, 0],
    )
    expect(run.angle).toBe(90)
    expect(run.fontSize).toBe(9)
    expect(run.italic).toBe(true)
  })
})
