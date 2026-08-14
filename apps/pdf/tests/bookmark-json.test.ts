import { describe, expect, it } from 'vitest'
import {
  MAX_BOOKMARK_COUNT,
  MAX_BOOKMARK_DEPTH,
  MAX_BOOKMARK_JSON_BYTES,
  bookmarkJsonText,
  bookmarkTreeCount,
  parseBookmarkJson,
} from '../src/renderer/bookmark-json'

describe('bookmark JSON exchange', () => {
  it('parses nested bookmarks and trims titles', () => {
    expect(
      parseBookmarkJson(
        JSON.stringify([
          {
            title: ' Chapter 1 ',
            pageNumber: 1,
            children: [{ title: 'Section', pageNumber: 2 }],
          },
        ]),
        3,
      ),
    ).toEqual([
      {
        title: 'Chapter 1',
        pageNumber: 1,
        children: [{ title: 'Section', pageNumber: 2, children: [] }],
      },
    ])
  })

  it('round-trips the public Stirling-compatible shape', () => {
    const bookmarks = [
      {
        title: 'Overview',
        pageNumber: 1,
        children: [{ title: 'Details', pageNumber: 2, children: [] }],
      },
    ]
    const json = bookmarkJsonText(bookmarks)
    expect(json.endsWith('\n')).toBe(true)
    expect(parseBookmarkJson(json, 2)).toEqual(bookmarks)
    expect(bookmarkTreeCount(bookmarks)).toBe(2)
  })

  it('accepts exactly twenty nested levels and rejects the twenty-first', () => {
    const nested = (levels: number): unknown[] => {
      let children: unknown[] = []
      for (let level = levels; level >= 1; level -= 1) {
        children = [{ title: `Level ${level}`, pageNumber: 1, children }]
      }
      return children
    }

    expect(parseBookmarkJson(JSON.stringify(nested(MAX_BOOKMARK_DEPTH)), 1)).toHaveLength(1)
    expect(() => parseBookmarkJson(JSON.stringify(nested(MAX_BOOKMARK_DEPTH + 1)), 1)).toThrow(
      'invalid-bookmarks',
    )
  })

  it('rejects too many bookmarks, invalid pages, blank titles, and non-array roots', () => {
    const tooMany = Array.from({ length: MAX_BOOKMARK_COUNT + 1 }, (_, index) => ({
      title: String(index),
      pageNumber: 1,
    }))
    expect(() => parseBookmarkJson(JSON.stringify(tooMany), 1)).toThrow('invalid-bookmarks')
    expect(() => parseBookmarkJson('[{"title":"A","pageNumber":2}]', 1)).toThrow(
      'invalid-bookmarks',
    )
    expect(() => parseBookmarkJson('[{"title":" ","pageNumber":1}]', 1)).toThrow(
      'invalid-bookmarks',
    )
    expect(() => parseBookmarkJson('{}', 1)).toThrow('invalid-bookmarks')
  })

  it('rejects JSON payloads over two megabytes before parsing', () => {
    expect(() => parseBookmarkJson(' '.repeat(MAX_BOOKMARK_JSON_BYTES + 1), 1)).toThrow(
      'invalid-bookmarks',
    )
  })
})
