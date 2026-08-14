import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  buildSearchIndex,
  planBulkTextReplacements,
  searchInIndex,
  searchPatternsInIndex,
  type SearchIndex,
} from '../src/renderer/search'

interface FakeItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
}

function fakeDoc(pages: FakeItem[][]): PDFDocumentProxy {
  return {
    numPages: pages.length,
    getPage: async (n: number) => ({
      getTextContent: async () => ({ items: pages[n - 1] }),
    }),
  } as unknown as PDFDocumentProxy
}

const item = (str: string, x: number, y: number, w: number, h: number): FakeItem => ({
  str,
  transform: [1, 0, 0, 1, x, y],
  width: w,
  height: h,
})

describe('buildSearchIndex', () => {
  it('concatenates item text per page and records char ranges', async () => {
    const doc = fakeDoc([[item('Hello ', 10, 700, 60, 12), item('World', 70, 700, 50, 12)]])
    const index = await buildSearchIndex(doc)
    expect(index).toHaveLength(1)
    expect(index[0]!.text).toBe('Hello World')
    expect(index[0]!.lower).toBe('hello world')
    expect(index[0]!.items).toEqual([
      { start: 0, end: 6, x: 10, y: 700, w: 60, h: 12 },
      { start: 6, end: 11, x: 70, y: 700, w: 50, h: 12 },
    ])
  })

  it('inserts newlines for hasEOL and skips empty/invalid items', async () => {
    const doc = fakeDoc([
      [
        { ...item('line1', 0, 0, 10, 10), hasEOL: true },
        { str: '', hasEOL: true }, // empty text still contributes the EOL
        { transform: [1, 0, 0, 1, 0, 0] }, // no str -> skipped entirely
        item('line2', 0, 0, 10, 10),
      ],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.text).toBe('line1\n\nline2')
  })

  it('derives height from the transform when height is missing', async () => {
    const doc = fakeDoc([[{ str: 'x', transform: [1, 0, 3, 4, 0, 0], width: 5 }]])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.h).toBe(5) // hypot(3, 4)
  })

  it('flags rotated runs and leaves upright ones unflagged', async () => {
    const doc = fakeDoc([
      [
        item('upright', 10, 700, 60, 12),
        { str: 'rotated', transform: [0, 12, -12, 0, 200, 400], width: 60, height: 12 },
      ],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.rot).toBeUndefined()
    expect(index[0]!.items[1]).toMatchObject({ rot: true, ux: 0, uy: 1 })
  })

  it('synthetic italic shear (c ≠ 0, horizontal baseline) is not flagged as rotated', async () => {
    // ~12° shear as writers emit for fake italics: b = 0, c = tan(12°) × size
    const doc = fakeDoc([
      [{ str: 'emphasis', transform: [12, 0, 2.55, 12, 100, 700], width: 48, height: 12 }],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.rot).toBeUndefined()
  })
})

describe('searchInIndex', () => {
  const entry = (text: string, items: SearchIndex[number]['items']): SearchIndex[number] => ({
    text,
    lower: text.toLowerCase(),
    items,
  })

  it('returns empty for an empty query', () => {
    const index = [entry('abc', [{ start: 0, end: 3, x: 0, y: 0, w: 30, h: 10 }])]
    expect(searchInIndex(index, '')).toEqual([])
  })

  it('finds case-insensitive matches with interpolated rects', () => {
    const index = [entry('Hello World', [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }])]
    const matches = searchInIndex(index, 'WORLD')
    expect(matches).toHaveLength(1)
    expect(matches[0]!.pageIndex).toBe(0)
    // 'World' spans chars 6..11 of 11 -> x from 60 to 110
    expect(matches[0]!.rects).toHaveLength(1)
    const [x1, y1, x2, y2] = matches[0]!.rects[0]!
    expect(x1).toBeCloseTo(60)
    expect(y1).toBe(700)
    expect(x2).toBeCloseTo(110)
    expect(y2).toBe(712)
  })

  it('spans multiple items with one rect per item', () => {
    const index = [
      entry('abcdef', [
        { start: 0, end: 3, x: 0, y: 0, w: 30, h: 10 },
        { start: 3, end: 6, x: 30, y: 0, w: 30, h: 10 },
      ]),
    ]
    const matches = searchInIndex(index, 'cd')
    expect(matches).toHaveLength(1)
    expect(matches[0]!.rects).toHaveLength(2)
    expect(matches[0]!.rects[0]).toEqual([20, 0, 30, 10])
    expect(matches[0]!.rects[1]).toEqual([30, 0, 40, 10])
  })

  it('reports every occurrence and the correct page index', () => {
    const index = [
      entry('nothing here', [{ start: 0, end: 12, x: 0, y: 0, w: 120, h: 10 }]),
      entry('foo bar foo', [{ start: 0, end: 11, x: 0, y: 0, w: 110, h: 10 }]),
    ]
    const matches = searchInIndex(index, 'foo')
    expect(matches).toHaveLength(2)
    expect(matches.every((m) => m.pageIndex === 1)).toBe(true)
  })

  it('skips matches falling in EOL-only gaps with no item coverage', () => {
    // '\n' at chars 5..6 belongs to no item -> no rects -> match dropped
    const index = [
      entry('hello\nworld', [
        { start: 0, end: 5, x: 0, y: 0, w: 50, h: 10 },
        { start: 6, end: 11, x: 0, y: -20, w: 50, h: 10 },
      ]),
    ]
    expect(searchInIndex(index, 'hello')).toHaveLength(1)
    expect(searchInIndex(index, 'world')).toHaveLength(1)
    // The match itself spans the newline; rects come from both surrounding items
    expect(searchInIndex(index, 'hello\nworld')[0]!.rects).toHaveLength(2)
  })

  it('caps results at 1000 matches', () => {
    const text = 'a'.repeat(2000)
    const index = [entry(text, [{ start: 0, end: 2000, x: 0, y: 0, w: 2000, h: 10 }])]
    expect(searchInIndex(index, 'a')).toHaveLength(1000)
  })
})

describe('searchPatternsInIndex', () => {
  const entry = (text: string, items: SearchIndex[number]['items']): SearchIndex[number] => ({
    text,
    lower: text.toLowerCase(),
    items,
  })

  it('matches multiple literal patterns case-insensitively across text items', () => {
    const index = [
      entry('Account 123 SECRET', [
        { start: 0, end: 8, x: 0, y: 0, w: 80, h: 10 },
        { start: 8, end: 18, x: 80, y: 0, w: 100, h: 10 },
      ]),
    ]
    const matches = searchPatternsInIndex(index, ['account 123', 'secret'], false, false)
    expect(matches).toHaveLength(2)
    expect(matches[0]!.rects).toHaveLength(2)
    expect(matches[1]!.rects).toEqual([[80, 0, 180, 10]])
  })

  it('supports regular expressions and whole-word matching', () => {
    const index = [
      entry('ID-123 ID-999 identity', [{ start: 0, end: 22, x: 0, y: 0, w: 220, h: 10 }]),
    ]
    expect(searchPatternsInIndex(index, ['ID-\\d{3}'], true, false)).toHaveLength(2)
    expect(searchPatternsInIndex(index, ['ID'], false, true)).toHaveLength(2)
    expect(searchPatternsInIndex(index, ['ident'], false, true)).toEqual([])
  })

  it('ignores blank patterns', () => {
    const index = [entry('secret', [{ start: 0, end: 6, x: 0, y: 0, w: 60, h: 10 }])]
    expect(searchPatternsInIndex(index, [' ', ''], false, false)).toEqual([])
  })

  it('uses the complete text item and rotation-aware bounds to avoid partial leaks', () => {
    const index = [
      entry('secret', [
        { start: 0, end: 6, x: 100, y: 200, w: 60, h: 12, rot: true, ux: 0, uy: 1 },
      ]),
    ]
    expect(searchPatternsInIndex(index, ['cret'], false, false)[0]!.rects).toEqual([
      [88, 200, 100, 260],
    ])
  })
})

describe('planBulkTextReplacements', () => {
  const entry = (text: string, items: SearchIndex[number]['items']): SearchIndex[number] => ({
    text,
    lower: text.toLowerCase(),
    items,
  })

  it('applies ordered rules so later replacements see earlier results', () => {
    const index = [entry('foo', [{ start: 0, end: 3, x: 10, y: 20, w: 30, h: 12 }])]
    const plan = planBulkTextReplacements(
      index,
      [
        { find: 'foo', replace: 'foos' },
        { find: 'foos', replace: 'bars' },
      ],
      [0],
      { caseSensitive: true, wholeWord: false },
    )
    expect(plan.matchCount).toBe(2)
    expect(plan.edits).toEqual([
      {
        pageIndex: 0,
        rect: [10, 20, 40, 32],
        oldText: 'foo',
        newText: 'bars',
        fontSize: 12,
        allowEmpty: true,
      },
    ])
  })

  it('coalesces multiple hits sharing one PDF text item', () => {
    const index = [entry('foo and foo', [{ start: 0, end: 11, x: 0, y: 0, w: 110, h: 10 }])]
    const plan = planBulkTextReplacements(index, [{ find: 'foo', replace: 'bar' }], [0], {
      caseSensitive: true,
      wholeWord: false,
    })
    expect(plan.matchCount).toBe(2)
    expect(plan.edits).toHaveLength(1)
    expect(plan.edits[0]).toMatchObject({ oldText: 'foo and foo', newText: 'bar and bar' })
  })

  it('matches across text items and supports empty replacements', () => {
    const index = [
      entry('Hello World', [
        { start: 0, end: 6, x: 0, y: 50, w: 60, h: 10 },
        { start: 6, end: 11, x: 60, y: 50, w: 50, h: 10 },
      ]),
    ]
    const plan = planBulkTextReplacements(index, [{ find: 'lo Wo', replace: '' }], [0], {
      caseSensitive: true,
      wholeWord: false,
    })
    expect(plan.edits).toHaveLength(1)
    expect(plan.edits[0]).toMatchObject({ oldText: 'lo Wo', newText: '', allowEmpty: true })
    expect(plan.edits[0]!.rect).toEqual([30, 50, 80, 60])
  })

  it('supports match case, ASCII whole words, and selected pages', () => {
    const make = (text: string) =>
      entry(text, [{ start: 0, end: text.length, x: 0, y: 0, w: text.length * 10, h: 10 }])
    const index = [make('cat catalog Cat'), make('cat')]
    const plan = planBulkTextReplacements(index, [{ find: 'cat', replace: 'dog' }], [0], {
      caseSensitive: true,
      wholeWord: true,
    })
    expect(plan.matchCount).toBe(1)
    expect(plan.edits).toHaveLength(1)
    expect(plan.edits[0]).toMatchObject({ pageIndex: 0, oldText: 'cat', newText: 'dog' })
  })
})
