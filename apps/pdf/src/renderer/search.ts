import type { PDFDocumentProxy } from 'pdfjs-dist'
import { foldCase, isWordChar } from '@genoffice/ui'
import type { TextEditInput } from '../shared/ipc'

/** One hit: original page + PDF user-space rects (multiple when spanning several text items) */
export interface SearchMatch {
  pageIndex: number
  rects: [number, number, number, number][]
}

export interface IndexedItem {
  start: number
  end: number
  x: number
  y: number
  w: number
  h: number
  /** Rotated run (tilted baseline) — excluded from block grouping */
  rot?: boolean
  /** Baseline unit vector for conservative rotated-run redaction bounds. */
  ux?: number
  uy?: number
  /** pdf.js font id (e.g. 'g_d0_f7'); resolves to the run's font for edit previews */
  font?: string
}

export interface PageEntry {
  /** Original text (same length as lower; used for context excerpts) */
  text: string
  lower: string
  items: IndexedItem[]
}

export type SearchIndex = PageEntry[]

export interface PdfTextReplacementRule {
  find: string
  replace: string
}

export interface PdfBulkTextReplacementPlan {
  edits: TextEditInput[]
  matchCount: number
  skippedSpans: number
}

const MAX_MATCHES = 1000
const MAX_REPLACEMENT_RULES = 50

/** Match modifiers for `searchInIndex`; both default off (plain case-insensitive scan) */
export interface SearchOptions {
  /** Case-sensitive matching — compare characters without folding case */
  matchCase?: boolean
  /** Whole-word matching — the hit must not continue a longer word on either side */
  wholeWord?: boolean
}

/** CJK scripts (kana, hangul, ideographs, and the punctuation/fullwidth blocks
    that travel with them): text written with these has no inter-word spaces */
const CJK_RE =
  /[\u1100-\u11FF\u2E80-\u318F\u31A0-\u31BF\u3400-\u4DBF\u4E00-\u9FFF\uA960-\uA97F\uAC00-\uD7AF\uF900-\uFAFF\uFE30-\uFE4F\uFF00-\uFFEF\u{20000}-\u{2FA1F}]/u

/** Whole-word boundary check for the hit at [start, start + len) of `text`.
    CJK characters never count as word characters here: CJK scripts do not
    separate words with spaces, so there is no word boundary to assert inside
    CJK runs — a CJK neighbour acts like a delimiter and a CJK match edge
    imposes no boundary at all. That degrades whole-word matching to a plain
    search for CJK sequences (and for the CJK side of mixed text) while Latin
    runs glued into a longer word are still rejected. */
function wordBoundaryOk(text: string, start: number, len: number): boolean {
  const wordChar = (ch: string | undefined): boolean => !!ch && !CJK_RE.test(ch) && isWordChar(ch)
  const left = !wordChar(text[start - 1]) || !wordChar(text[start])
  const right = !wordChar(text[start + len]) || !wordChar(text[start + len - 1])
  return left && right
}

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
  fontName?: string
}

/** Concatenate text per page + record each item's char range and PDF-space box (built once, cached per doc by caller) */
export async function buildSearchIndex(doc: PDFDocumentProxy): Promise<SearchIndex> {
  const entries: PageEntry[] = []
  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    let text = ''
    const items: IndexedItem[] = []
    for (const it of content.items as RawTextItem[]) {
      if (typeof it.str !== 'string') continue
      if (it.str.length > 0) {
        // Geometry is optional; without a transform the item still contributes its
        // characters, or it is unfindable while its hasEOL newline still lands.
        if (it.transform) {
          const h = it.height || Math.hypot(it.transform[2] ?? 0, it.transform[3] ?? 0)
          // Rotation tilts the baseline (b ≠ 0). A non-zero c alone is horizontal
          // shear — synthetic italics — which stays horizontally set and must keep
          // participating in block grouping.
          const rot = Math.abs(it.transform[1] ?? 0) > h * 1e-3
          const baselineLength = Math.hypot(it.transform[0] ?? 0, it.transform[1] ?? 0)
          items.push({
            start: text.length,
            end: text.length + it.str.length,
            x: it.transform[4] ?? 0,
            y: it.transform[5] ?? 0,
            w: it.width ?? 0,
            h,
            ...(rot ? { rot: true } : {}),
            ...(rot && baselineLength > 0
              ? {
                  ux: (it.transform[0] ?? 0) / baselineLength,
                  uy: (it.transform[1] ?? 0) / baselineLength,
                }
              : {}),
            ...(typeof it.fontName === 'string' ? { font: it.fontName } : {}),
          })
        }
        text += it.str
      }
      if (it.hasEOL) text += '\n'
    }
    entries.push({ text, lower: foldCase(text), items })
  }
  return entries
}

/** One text extraction per loaded document, shared by every consumer */
export interface SearchIndexCache {
  /** Base index of `doc`; the same promise for the same document until cleared */
  get(doc: PDFDocumentProxy): Promise<SearchIndex>
  /** Drop the cached extraction (a save-reload replaces the document content) */
  clear(): void
}

/** Building the index walks every page of the document, so it must happen once per
    loaded document: search, paragraph boxes, the AI tools and the auto-OCR pass all
    read this cache instead of extracting the text a second time. */
export function createSearchIndexCache(): SearchIndexCache {
  let entry: { doc: PDFDocumentProxy; promise: Promise<SearchIndex> } | null = null
  return {
    get(doc) {
      if (entry?.doc !== doc) entry = { doc, promise: buildSearchIndex(doc) }
      return entry.promise
    },
    clear() {
      entry = null
    },
  }
}

/** Full-text search with optional case/word modifiers; rects linearly interpolated within items by char ratio (approximate; bounding box for rotated glyphs) */
export function searchInIndex(
  index: SearchIndex,
  query: string,
  opts: SearchOptions = {},
): SearchMatch[] {
  const matchCase = opts.matchCase ?? false
  const wholeWord = opts.wholeWord ?? false
  const q = matchCase ? query : foldCase(query)
  if (!q) return []
  const matches: SearchMatch[] = []
  for (let pageIndex = 0; pageIndex < index.length; pageIndex++) {
    const { text, lower, items } = index[pageIndex]!
    // The sensitive path reads the original text; the insensitive path reuses
    // the precomputed fold so a search never re-folds the whole page. foldCase
    // is length-preserving, so hit offsets are valid in `text` either way.
    const hay = matchCase ? text : lower
    let from = 0
    while (matches.length < MAX_MATCHES) {
      const s = hay.indexOf(q, from)
      if (s < 0) break
      if (wholeWord && !wordBoundaryOk(text, s, q.length)) {
        from = s + 1
        continue
      }
      const e = s + q.length
      from = e
      const rects: [number, number, number, number][] = []
      for (const it of items) {
        if (it.end <= s || it.start >= e) continue
        const len = it.end - it.start
        const lo = (Math.max(s, it.start) - it.start) / len
        const hi = (Math.min(e, it.end) - it.start) / len
        const x1 = it.x + it.w * lo
        const x2 = it.x + it.w * hi
        if (x2 - x1 < 0.01) continue
        rects.push([x1, it.y, x2, it.y + it.h])
      }
      if (rects.length > 0) matches.push({ pageIndex, rects })
    }
    if (matches.length >= MAX_MATCHES) break
  }
  return matches
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function matchRanges(
  text: string,
  pattern: string,
  useRegex: boolean,
  wholeWord: boolean,
  caseSensitive = false,
): Array<readonly [number, number]> {
  const source = useRegex ? pattern : escapeRegExp(pattern)
  const expression = wholeWord ? `\\b(?:${source})\\b` : source
  const regex = new RegExp(expression, caseSensitive ? 'gu' : 'giu')
  const ranges: Array<readonly [number, number]> = []
  let match: RegExpExecArray | null
  while ((match = regex.exec(text)) && ranges.length < MAX_MATCHES) {
    if (match[0].length === 0) {
      regex.lastIndex++
      continue
    }
    ranges.push([match.index, match.index + match[0].length])
  }
  return ranges
}

function rectsForRange(entry: PageEntry, start: number, end: number): SearchMatch['rects'] {
  const rects: SearchMatch['rects'] = []
  for (const item of entry.items) {
    if (item.end <= start || item.start >= end) continue
    if (item.rot && item.ux !== undefined && item.uy !== undefined) {
      const baselineX = item.ux * item.w
      const baselineY = item.uy * item.w
      const heightX = -item.uy * item.h
      const heightY = item.ux * item.h
      const xs = [item.x, item.x + baselineX, item.x + heightX, item.x + baselineX + heightX]
      const ys = [item.y, item.y + baselineY, item.y + heightY, item.y + baselineY + heightY]
      rects.push([Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)])
    } else if (item.w > 0 && item.h > 0) {
      rects.push([item.x, item.y, item.x + item.w, item.y + item.h])
    }
  }
  return rects
}

export function searchPatternsInIndex(
  index: SearchIndex,
  patterns: string[],
  useRegex: boolean,
  wholeWord: boolean,
  caseSensitive = false,
): SearchMatch[] {
  const cleanedPatterns = patterns.map((pattern) => pattern.trim()).filter(Boolean)
  if (cleanedPatterns.length === 0) return []
  const matches: SearchMatch[] = []
  for (let pageIndex = 0; pageIndex < index.length && matches.length < MAX_MATCHES; pageIndex++) {
    const entry = index[pageIndex]!
    for (const pattern of cleanedPatterns) {
      for (const [start, end] of matchRanges(
        entry.text,
        pattern,
        useRegex,
        wholeWord,
        caseSensitive,
      )) {
        const rects = rectsForRange(entry, start, end)
        if (rects.length > 0) matches.push({ pageIndex, rects })
        if (matches.length >= MAX_MATCHES) break
      }
      if (matches.length >= MAX_MATCHES) break
    }
  }
  return matches
}

interface ReplacementAtom {
  text: string
  sourceStart: number
  sourceEnd: number
  changed: boolean
}

interface ReplacementSpan {
  start: number
  end: number
}

/** ASCII word characters for bulk-replacement whole-word matching (pdf text, not CJK) */
const isReplacementWordChar = (value: string | undefined): boolean =>
  value !== undefined && /[A-Za-z0-9_]/.test(value)

function literalMatchRanges(
  value: string,
  find: string,
  caseSensitive: boolean,
  wholeWord: boolean,
): ReplacementSpan[] {
  const haystack = caseSensitive ? value : value.toLocaleLowerCase()
  const needle = caseSensitive ? find : find.toLocaleLowerCase()
  const ranges: ReplacementSpan[] = []
  let from = 0
  while (ranges.length < MAX_MATCHES) {
    const start = haystack.indexOf(needle, from)
    if (start < 0) break
    const end = start + needle.length
    from = end
    if (
      wholeWord &&
      (isReplacementWordChar(value[start - 1]) || isReplacementWordChar(value[end]))
    ) {
      continue
    }
    ranges.push({ start, end })
  }
  return ranges
}

function splitAtomsAt(atoms: ReplacementAtom[], offset: number): void {
  let cursor = 0
  for (let index = 0; index < atoms.length; index++) {
    const atom = atoms[index]!
    const next = cursor + atom.text.length
    if (offset > cursor && offset < next) {
      const split = offset - cursor
      atoms.splice(
        index,
        1,
        { ...atom, text: atom.text.slice(0, split) },
        { ...atom, text: atom.text.slice(split) },
      )
      return
    }
    cursor = next
  }
}

function atomBoundary(atoms: ReplacementAtom[], offset: number, fromEnd = false): number {
  let cursor = 0
  for (let index = 0; index < atoms.length; index++) {
    const next = cursor + atoms[index]!.text.length
    if (offset < next) return index
    if (offset === next) {
      let boundary = index + 1
      if (fromEnd) {
        while (boundary < atoms.length && atoms[boundary]!.text.length === 0) boundary++
      }
      return boundary
    }
    cursor = next
  }
  return atoms.length
}

function sequentialReplacementAtoms(
  source: string,
  rules: PdfTextReplacementRule[],
  caseSensitive: boolean,
  wholeWord: boolean,
): { atoms: ReplacementAtom[]; matchCount: number } {
  const atoms: ReplacementAtom[] = []
  for (let offset = 0; offset < source.length;) {
    const codePoint = source.codePointAt(offset)
    if (codePoint === undefined) break
    const text = String.fromCodePoint(codePoint)
    atoms.push({ text, sourceStart: offset, sourceEnd: offset + text.length, changed: false })
    offset += text.length
  }
  let matchCount = 0
  for (const rule of rules.slice(0, MAX_REPLACEMENT_RULES)) {
    if (!rule.find) continue
    const current = atoms.map((atom) => atom.text).join('')
    const matches = literalMatchRanges(current, rule.find, caseSensitive, wholeWord)
    matchCount += matches.length
    for (let index = matches.length - 1; index >= 0; index--) {
      const match = matches[index]!
      splitAtomsAt(atoms, match.end)
      splitAtomsAt(atoms, match.start)
      const startIndex = atomBoundary(atoms, match.start, true)
      const endIndex = atomBoundary(atoms, match.end)
      const consumed = atoms.slice(startIndex, endIndex)
      if (consumed.length === 0) continue
      atoms.splice(startIndex, consumed.length, {
        text: rule.replace,
        sourceStart: Math.min(...consumed.map((atom) => atom.sourceStart)),
        sourceEnd: Math.max(...consumed.map((atom) => atom.sourceEnd)),
        changed: true,
      })
    }
  }
  return { atoms, matchCount }
}

function itemsForRange(entry: PageEntry, start: number, end: number): IndexedItem[] {
  return entry.items.filter((item) => item.end > start && item.start < end)
}

function replacementRect(
  entry: PageEntry,
  start: number,
  end: number,
): { rect: [number, number, number, number]; fontSize: number } | null {
  let x1 = Infinity
  let y1 = Infinity
  let x2 = -Infinity
  let y2 = -Infinity
  let fontSize = 0
  for (const item of itemsForRange(entry, start, end)) {
    if (item.rot && item.ux !== undefined && item.uy !== undefined) {
      const baselineX = item.ux * item.w
      const baselineY = item.uy * item.w
      const heightX = -item.uy * item.h
      const heightY = item.ux * item.h
      const xs = [item.x, item.x + baselineX, item.x + heightX, item.x + baselineX + heightX]
      const ys = [item.y, item.y + baselineY, item.y + heightY, item.y + baselineY + heightY]
      x1 = Math.min(x1, ...xs)
      y1 = Math.min(y1, ...ys)
      x2 = Math.max(x2, ...xs)
      y2 = Math.max(y2, ...ys)
    } else {
      const length = item.end - item.start
      const from = (Math.max(start, item.start) - item.start) / length
      const to = (Math.min(end, item.end) - item.start) / length
      x1 = Math.min(x1, item.x + item.w * from)
      y1 = Math.min(y1, item.y)
      x2 = Math.max(x2, item.x + item.w * to)
      y2 = Math.max(y2, item.y + item.h)
    }
    fontSize = Math.max(fontSize, item.h)
  }
  if (!Number.isFinite(x1) || !Number.isFinite(y1) || x2 - x1 < 0.01 || fontSize <= 0) {
    return null
  }
  return { rect: [x1, y1, x2, y2], fontSize }
}

function changedSourceSpans(atoms: ReplacementAtom[], entry: PageEntry): ReplacementSpan[] {
  const spans = atoms
    .filter((atom) => atom.changed)
    .map((atom) => ({ start: atom.sourceStart, end: atom.sourceEnd }))
    .sort((left, right) => left.start - right.start || left.end - right.end)
  const merged: ReplacementSpan[] = []
  for (const span of spans) {
    const previous = merged[merged.length - 1]
    if (previous && span.start <= previous.end) previous.end = Math.max(previous.end, span.end)
    else merged.push({ ...span })
  }

  let didMerge = true
  while (didMerge) {
    didMerge = false
    outer: for (let left = 0; left < merged.length; left++) {
      const leftItems = new Set(itemsForRange(entry, merged[left]!.start, merged[left]!.end))
      for (let right = left + 1; right < merged.length; right++) {
        if (
          !itemsForRange(entry, merged[right]!.start, merged[right]!.end).some((item) =>
            leftItems.has(item),
          )
        ) {
          continue
        }
        merged[left] = {
          start: Math.min(merged[left]!.start, merged[right]!.start),
          end: Math.max(merged[left]!.end, merged[right]!.end),
        }
        merged.splice(right, 1)
        didMerge = true
        break outer
      }
    }
  }
  return merged
}

/**
 * Apply ordered literal replacements to the search index and map the final result back
 * to non-overlapping PDFium text edits. Later rules see earlier replacements, matching
 * Stirling's edit-text semantics; edits sharing one underlying text item are coalesced.
 */
export function planBulkTextReplacements(
  index: SearchIndex,
  rules: PdfTextReplacementRule[],
  pageIndexes: number[],
  options: { caseSensitive: boolean; wholeWord: boolean },
): PdfBulkTextReplacementPlan {
  const edits: TextEditInput[] = []
  let matchCount = 0
  let skippedSpans = 0
  for (const pageIndex of [...new Set(pageIndexes)]) {
    const entry = index[pageIndex]
    if (!entry) continue
    const planned = sequentialReplacementAtoms(
      entry.text,
      rules,
      options.caseSensitive,
      options.wholeWord,
    )
    matchCount += planned.matchCount
    for (const span of changedSourceSpans(planned.atoms, entry)) {
      const geometry = replacementRect(entry, span.start, span.end)
      if (!geometry) {
        skippedSpans++
        continue
      }
      const newText = planned.atoms
        .filter((atom) => atom.sourceEnd > span.start && atom.sourceStart < span.end)
        .map((atom) => atom.text)
        .join('')
      edits.push({
        pageIndex,
        rect: geometry.rect,
        oldText: entry.text.slice(span.start, span.end),
        newText,
        fontSize: geometry.fontSize,
        allowEmpty: true,
      })
    }
  }
  return { edits, matchCount, skippedSpans }
}
