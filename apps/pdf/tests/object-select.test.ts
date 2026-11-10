import { describe, expect, it } from 'vitest'
import {
  OBJECT_HIT_TOLERANCE_PX,
  hitTestObject,
  isTextEntryTarget,
  objectSelectTolerance,
  pointRectDistance,
  rectsOverlap,
  selectableTextRuns,
  textRunsForPage,
  type SelectableImage,
  type SelectableTextRun,
} from '../src/renderer/object-select'
import type { PageEntry } from '../src/renderer/search'

const image = (rect: [number, number, number, number], pageIndex = 0): SelectableImage => ({
  kind: 'image',
  pageIndex,
  rect,
})

const run = (
  rect: [number, number, number, number],
  text = 'word',
  pageIndex = 0,
): SelectableTextRun => ({ kind: 'text', pageIndex, rect, text, fontSize: rect[3] - rect[1] })

describe('object hit testing', () => {
  it('requires the click to land within the tolerance of an object edge', () => {
    const target = run([10, 20, 30, 40])
    expect(hitTestObject([35, 30], [target], 8)).toBe(target) // 5 units away
    expect(hitTestObject([39, 30], [target], 8)).toBeNull() // 9 units — outside
    expect(hitTestObject([19, 30], [target], 8)).toBe(target) // inside
    expect(hitTestObject([100, 100], [target], 8)).toBeNull()
  })

  it('picks the nearest object among candidates', () => {
    const near = run([0, 0, 10, 10], 'near')
    const far = run([30, 0, 40, 10], 'far')
    expect(hitTestObject([12, 5], [far, near], 8)).toBe(near)
    expect(hitTestObject([27, 5], [far, near], 8)).toBe(far)
  })

  it('prefers the image when the pointer is inside both an image and a text run', () => {
    const picture = image([0, 0, 100, 100])
    const caption = run([10, 10, 30, 20])
    expect(hitTestObject([20, 15], [caption, picture], 8)).toBe(picture)
    // Also when the image arrives after the text in the candidate list
    expect(hitTestObject([20, 15], [picture, caption], 8)).toBe(picture)
  })

  it('keeps a text run under the pointer over a merely nearby image', () => {
    const picture = image([0, 0, 10, 10])
    const text = run([12, 0, 60, 20])
    // 2 units from the image, inside the text run
    expect(hitTestObject([12, 5], [picture, text], 8)).toBe(text)
  })

  it('breaks an exact distance tie toward the image', () => {
    const picture = image([0, 0, 10, 10])
    const text = run([0, 5, 10, 15])
    // Both rects are 5 units from this point; the image wins the tie
    expect(hitTestObject([15, 5], [text, picture], 8)).toBe(picture)
  })
})

describe('object select tolerance', () => {
  it('converts the display-pixel tolerance with the current zoom', () => {
    expect(OBJECT_HIT_TOLERANCE_PX).toBe(8)
    expect(objectSelectTolerance(1)).toBe(8)
    expect(objectSelectTolerance(2)).toBe(4)
    expect(objectSelectTolerance(0.5)).toBe(16)
  })

  it('accounts for the PDF user unit and survives a bad scale', () => {
    expect(objectSelectTolerance(1, 2)).toBe(4)
    expect(objectSelectTolerance(4, 2)).toBe(1)
    expect(objectSelectTolerance(0)).toBe(Infinity)
  })

  it('the same PDF offset is inside at one zoom and outside at another', () => {
    const target = run([0, 0, 100, 100])
    // 6 PDF units from the right edge: within 8 at 1x, outside 4 at 2x
    const point: [number, number] = [106, 50]
    expect(hitTestObject(point, [target], objectSelectTolerance(1))).toBe(target)
    expect(hitTestObject(point, [target], objectSelectTolerance(2))).toBeNull()
    // 3 units out is inside the 2x tolerance
    expect(hitTestObject([103, 50], [target], objectSelectTolerance(2))).toBe(target)
  })
})

describe('geometry helpers', () => {
  it('pointRectDistance is zero inside and edge distance outside', () => {
    expect(pointRectDistance(5, 5, [0, 0, 10, 10])).toBe(0)
    expect(pointRectDistance(13, 5, [0, 0, 10, 10])).toBe(3)
    expect(pointRectDistance(13, 14, [0, 0, 10, 10])).toBe(5)
  })

  it('rectsOverlap ignores shared edges', () => {
    expect(rectsOverlap([0, 0, 10, 10], [5, 5, 15, 15])).toBe(true)
    expect(rectsOverlap([0, 0, 10, 10], [10, 0, 20, 10])).toBe(false)
    expect(rectsOverlap([0, 0, 10, 10], [-5, -5, 15, 15])).toBe(true)
  })
})

describe('textRunsForPage', () => {
  const entry: PageEntry = {
    text: 'Hello world',
    lower: 'hello world',
    items: [
      { start: 0, end: 5, x: 10, y: 20, w: 30, h: 12 },
      { start: 6, end: 11, x: 45, y: 20, w: 28, h: 12 },
    ],
  }

  it('turns index items into PDF-space runs with their text and font size', () => {
    expect(textRunsForPage(3, entry)).toEqual([
      { kind: 'text', pageIndex: 3, rect: [10, 20, 40, 32], text: 'Hello', fontSize: 12 },
      { kind: 'text', pageIndex: 3, rect: [45, 20, 73, 32], text: 'world', fontSize: 12 },
    ])
  })

  it('skips rotated, blank and degenerate runs', () => {
    const noisy: PageEntry = {
      text: 'ab cd',
      lower: 'ab cd',
      items: [
        { start: 0, end: 2, x: 0, y: 0, w: 10, h: 10, rot: true },
        { start: 2, end: 3, x: 0, y: 20, w: 5, h: 10 }, // space
        { start: 4, end: 6, x: 0, y: 40, w: 0, h: 10 }, // zero width
      ],
    }
    expect(textRunsForPage(0, noisy)).toEqual([])
    expect(textRunsForPage(0, undefined)).toEqual([])
  })
})

describe('isTextEntryTarget (Backspace guard)', () => {
  it('treats form controls and contenteditable as text entry', () => {
    expect(isTextEntryTarget({ tagName: 'INPUT' } as unknown as EventTarget)).toBe(true)
    expect(isTextEntryTarget({ tagName: 'TEXTAREA' } as unknown as EventTarget)).toBe(true)
    expect(isTextEntryTarget({ tagName: 'SELECT' } as unknown as EventTarget)).toBe(true)
    expect(
      isTextEntryTarget({ tagName: 'DIV', isContentEditable: true } as unknown as EventTarget),
    ).toBe(true)
  })

  it('lets the viewer handle the key everywhere else', () => {
    expect(isTextEntryTarget({ tagName: 'DIV' } as unknown as EventTarget)).toBe(false)
    expect(isTextEntryTarget({ tagName: 'CANVAS' } as unknown as EventTarget)).toBe(false)
    expect(isTextEntryTarget(null)).toBe(false)
    expect(isTextEntryTarget({} as unknown as EventTarget)).toBe(false)
  })
})

describe('selectableTextRuns (pending-edit filtering)', () => {
  it('drops runs overlapping a pending text edit rect', () => {
    const runs = [run([0, 0, 50, 10], 'kept'), run([60, 0, 100, 10], 'owned')]
    const result = selectableTextRuns(runs, [[60, 0, 100, 10]], [])
    expect(result.map((r) => r.text)).toEqual(['kept'])
  })

  it('drops runs covered by a pending above-text image (bridge write-back)', () => {
    const runs = [run([300, 300, 350, 310], 'outside'), run([10, 2, 40, 8], 'hidden')]
    const result = selectableTextRuns(runs, [], [[5, -5, 200, 100]])
    expect(result.map((r) => r.text)).toEqual(['outside'])
  })

  it('keeps runs outside both rect lists untouched', () => {
    const runs = [run([0, 0, 50, 10]), run([100, 100, 150, 110])]
    expect(selectableTextRuns(runs, [[300, 300, 400, 400]], [[-50, -50, -10, -10]])).toEqual(runs)
  })

  it('returns everything when no pending rects exist', () => {
    const runs = [run([0, 0, 50, 10]), run([100, 100, 150, 110])]
    expect(selectableTextRuns(runs, [], [])).toEqual(runs)
  })
})
