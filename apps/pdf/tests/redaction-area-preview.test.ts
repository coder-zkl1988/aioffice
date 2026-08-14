import { describe, expect, it } from 'vitest'
import {
  drawRedactionArea,
  moveRedactionArea,
  resizeRedactionArea,
} from '../src/renderer/RedactionAreaPreview'

const area = { pageIndex: 1, x: 0.2, y: 0.3, width: 0.4, height: 0.2 }

describe('redaction area preview geometry', () => {
  it('moves an area while keeping it inside the page', () => {
    const moved = moveRedactionArea(area, 0.1, -0.1)
    expect(moved.x).toBeCloseTo(0.3)
    expect(moved.y).toBeCloseTo(0.2)
    expect(moveRedactionArea(area, 1, 1)).toEqual({ ...area, x: 0.6, y: 0.8 })
  })

  it('resizes each edge and enforces a minimum size', () => {
    const resizedNorthWest = resizeRedactionArea(area, 'nw', 0.1, 0.2)
    expect(resizedNorthWest.x).toBeCloseTo(0.1)
    expect(resizedNorthWest.y).toBeCloseTo(0.2)
    expect(resizedNorthWest.width).toBeCloseTo(0.5)
    expect(resizedNorthWest.height).toBeCloseTo(0.3)
    const resizedSouthEast = resizeRedactionArea(area, 'se', 0.1, 0.1)
    expect(resizedSouthEast.width).toBeCloseTo(0.01)
    expect(resizedSouthEast.height).toBeCloseTo(0.01)
  })

  it('draws in either direction and clamps to page bounds', () => {
    const drawn = drawRedactionArea(2, 0.8, 0.7, 0.2, 0.1)
    expect(drawn.pageIndex).toBe(2)
    expect(drawn.x).toBeCloseTo(0.2)
    expect(drawn.y).toBeCloseTo(0.1)
    expect(drawn.width).toBeCloseTo(0.6)
    expect(drawn.height).toBeCloseTo(0.6)
    expect(drawRedactionArea(0, -1, -1, 2, 2)).toEqual({
      pageIndex: 0,
      x: 0,
      y: 0,
      width: 1,
      height: 1,
    })
  })
})
