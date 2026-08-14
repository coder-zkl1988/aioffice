import type { PdfBookmark } from '@genoffice/pdf-tools'

export const MAX_BOOKMARK_JSON_BYTES = 2 * 1024 * 1024
export const MAX_BOOKMARK_COUNT = 1_000
export const MAX_BOOKMARK_DEPTH = 20
export const MAX_BOOKMARK_TITLE_LENGTH = 1_000

function parseBookmarkLevel(
  value: unknown,
  pageCount: number,
  depth: number,
  counter: { count: number },
): PdfBookmark[] {
  if (!Array.isArray(value)) throw new Error('invalid-bookmarks')
  if (value.length === 0) return []
  if (depth > MAX_BOOKMARK_DEPTH) throw new Error('invalid-bookmarks')
  return value.map((item) => {
    if (!item || typeof item !== 'object' || Array.isArray(item)) {
      throw new Error('invalid-bookmarks')
    }
    counter.count += 1
    if (counter.count > MAX_BOOKMARK_COUNT) throw new Error('invalid-bookmarks')
    const candidate = item as Record<string, unknown>
    const title = typeof candidate.title === 'string' ? candidate.title.trim() : ''
    const pageNumber = candidate.pageNumber
    if (
      !title ||
      title.length > MAX_BOOKMARK_TITLE_LENGTH ||
      !Number.isInteger(pageNumber) ||
      (pageNumber as number) < 1 ||
      (pageNumber as number) > pageCount
    ) {
      throw new Error('invalid-bookmarks')
    }
    const children = candidate.children === undefined ? [] : candidate.children
    return {
      title,
      pageNumber: pageNumber as number,
      children: parseBookmarkLevel(children, pageCount, depth + 1, counter),
    }
  })
}

export function parseBookmarkJson(text: string, pageCount: number): PdfBookmark[] {
  if (
    !Number.isInteger(pageCount) ||
    pageCount < 1 ||
    new TextEncoder().encode(text).byteLength > MAX_BOOKMARK_JSON_BYTES
  ) {
    throw new Error('invalid-bookmarks')
  }
  let value: unknown
  try {
    value = JSON.parse(text)
  } catch {
    throw new Error('invalid-bookmarks')
  }
  return parseBookmarkLevel(value, pageCount, 1, { count: 0 })
}

export function bookmarkJsonText(bookmarks: PdfBookmark[]): string {
  return `${JSON.stringify(bookmarks, null, 2)}\n`
}

export function bookmarkTreeCount(bookmarks: PdfBookmark[]): number {
  return bookmarks.reduce((total, bookmark) => total + 1 + bookmarkTreeCount(bookmark.children), 0)
}
