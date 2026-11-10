import { AnnotationMode, getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfRedactedPage, PdfRedactionArea, PdfRedactionOptions } from '@genoffice/pdf-tools'
import { flattenRenderScale } from './flatten-pdf'
import {
  buildSearchIndex,
  searchPatternsInIndex,
  type IndexedItem,
  type SearchMatch,
} from './search'
import { DOC_OPTS } from './view-config'

const JPEG_QUALITY = 0.94

function canvasJpegBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) return reject(new Error('Could not encode redacted PDF page'))
        void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
      },
      'image/jpeg',
      JPEG_QUALITY,
    )
  })
}

export function validateRedactionOptions(options: PdfRedactionOptions): void {
  if ((options.mode ?? 'text') === 'text') {
    if (options.patterns.map((pattern) => pattern.trim()).filter(Boolean).length === 0) {
      throw new Error('At least one redaction pattern is required')
    }
    if (options.useRegex) {
      const flags = options.matchCase ? 'gu' : 'giu'
      for (const pattern of options.patterns) new RegExp(pattern, flags)
    }
  } else {
    if (!options.areas || options.areas.length === 0) {
      throw new Error('At least one redaction area is required')
    }
    for (const area of options.areas) validateRedactionArea(area)
  }
  if (!/^#[\da-f]{6}$/i.test(options.color)) throw new Error('Redaction color is invalid')
  if (!Number.isFinite(options.padding) || options.padding < 0 || options.padding > 72) {
    throw new Error('Redaction padding must be from 0 to 72 points')
  }
  flattenRenderScale(612, 792, options.renderDpi)
}

export function validateRedactionArea(area: PdfRedactionArea): void {
  if (!Number.isInteger(area.pageIndex) || area.pageIndex < 0) {
    throw new Error('Redaction area page is invalid')
  }
  if (
    !Number.isFinite(area.x) ||
    !Number.isFinite(area.y) ||
    !Number.isFinite(area.width) ||
    !Number.isFinite(area.height) ||
    area.x < 0 ||
    area.y < 0 ||
    area.width <= 0 ||
    area.height <= 0 ||
    area.x + area.width > 1 ||
    area.y + area.height > 1
  ) {
    throw new Error('Redaction area must fit inside the page')
  }
}

export function redactionAreasByPage(areas: PdfRedactionArea[]): Map<number, PdfRedactionArea[]> {
  const result = new Map<number, PdfRedactionArea[]>()
  for (const area of areas) {
    const current = result.get(area.pageIndex)
    if (current) current.push(area)
    else result.set(area.pageIndex, [area])
  }
  return result
}

export function matchesByPage(matches: SearchMatch[]): Map<number, SearchMatch['rects']> {
  const result = new Map<number, SearchMatch['rects']>()
  for (const match of matches) {
    const current = result.get(match.pageIndex)
    if (current) current.push(...match.rects)
    else result.set(match.pageIndex, [...match.rects])
  }
  return result
}

export function paddedViewportRectangle(
  rectangle: readonly [number, number, number, number],
  padding: number,
  convert: (rectangle: readonly [number, number, number, number]) => number[],
): readonly [number, number, number, number] {
  const converted = convert([
    rectangle[0] - padding,
    rectangle[1] - padding,
    rectangle[2] + padding,
    rectangle[3] + padding,
  ])
  const x1 = Math.min(converted[0] ?? 0, converted[2] ?? 0)
  const y1 = Math.min(converted[1] ?? 0, converted[3] ?? 0)
  const x2 = Math.max(converted[0] ?? 0, converted[2] ?? 0)
  const y2 = Math.max(converted[1] ?? 0, converted[3] ?? 0)
  return [x1, y1, x2, y2]
}

export async function renderRedactedPdfPages(
  sourceDocument: PDFDocumentProxy,
  options: PdfRedactionOptions,
): Promise<PdfRedactedPage[]> {
  validateRedactionOptions(options)
  const mode = options.mode ?? 'text'
  const pageMatches = new Map<number, SearchMatch['rects']>()
  const pageAreas = redactionAreasByPage(options.areas ?? [])
  if (mode === 'text') {
    const index = await buildSearchIndex(sourceDocument)
    const matches = searchPatternsInIndex(
      index,
      options.patterns,
      options.useRegex,
      options.wholeWord,
      options.matchCase ?? false,
    )
    if (matches.length === 0) throw new Error('No matching text was found for redaction')
    for (const [pageIndex, rectangles] of matchesByPage(matches)) {
      pageMatches.set(pageIndex, rectangles)
    }
  }
  const affectedPageIndexes = [...(mode === 'text' ? pageMatches.keys() : pageAreas.keys())].sort(
    (left, right) => left - right,
  )
  if (affectedPageIndexes.some((pageIndex) => pageIndex >= sourceDocument.numPages)) {
    throw new Error('Redaction area page is outside the PDF')
  }
  const pages: PdfRedactedPage[] = []

  for (const pageIndex of affectedPageIndexes) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const baseViewport = page.getViewport({ scale: 1 })
    const scale = flattenRenderScale(baseViewport.width, baseViewport.height, options.renderDpi)
    const viewport = page.getViewport({ scale })
    const canvas = window.document.createElement('canvas')
    canvas.width = Math.max(1, Math.floor(viewport.width))
    canvas.height = Math.max(1, Math.floor(viewport.height))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('Canvas is unavailable')

    await page.render({
      canvas,
      viewport,
      intent: 'print',
      annotationMode: AnnotationMode.ENABLE_STORAGE,
      printAnnotationStorage: sourceDocument.annotationStorage.print,
      background: '#ffffff',
    }).promise

    if (mode === 'text') {
      context.fillStyle = options.color
      for (const rectangle of pageMatches.get(pageIndex) ?? []) {
        const [x1, y1, x2, y2] = paddedViewportRectangle(
          rectangle,
          options.padding,
          ([left, bottom, right, top]) => [
            ...viewport.convertToViewportPoint(left, bottom),
            ...viewport.convertToViewportPoint(right, top),
          ],
        )
        context.fillRect(x1, y1, x2 - x1, y2 - y1)
      }
    } else {
      context.fillStyle = options.color
      for (const area of pageAreas.get(pageIndex) ?? []) {
        context.fillRect(
          area.x * canvas.width,
          area.y * canvas.height,
          area.width * canvas.width,
          area.height * canvas.height,
        )
      }
    }
    pages.push({ pageIndex, image: await canvasJpegBytes(canvas) })
    canvas.width = 0
    canvas.height = 0
  }

  return pages
}

/** One-based page list, keeping Stirling's "page 3" / "pages 2, 5" phrasing. */
function describePages(pageNumbers: number[]): string {
  const oneBased = pageNumbers.map((pageIndex) => pageIndex + 1)
  return oneBased.length === 1 ? `page ${oneBased[0]}` : `pages ${oneBased.join(', ')}`
}

function redactionVerificationError(pageNumbers: number[], detail: string): Error {
  return new Error(
    `Redaction could not be verified on ${describePages(pageNumbers)}: ${detail}, so no file was returned.`,
  )
}

/** Axis-aligned PDF-space bounds of one indexed text item (rotated runs included). */
function itemPdfBounds(item: IndexedItem): [number, number, number, number] | null {
  if (!(item.w > 0 && item.h > 0)) return null
  if (item.rot && item.ux !== undefined && item.uy !== undefined) {
    const baselineX = item.ux * item.w
    const baselineY = item.uy * item.w
    const heightX = -item.uy * item.h
    const heightY = item.ux * item.h
    const xs = [item.x, item.x + baselineX, item.x + heightX, item.x + baselineX + heightX]
    const ys = [item.y, item.y + baselineY, item.y + heightY, item.y + baselineY + heightY]
    return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
  }
  return [item.x, item.y, item.x + item.w, item.y + item.h]
}

/** Axis-aligned PDF-space rectangle overlap (touching edges do not count as containing text). */
function rectanglesOverlap(
  left: number,
  bottom: number,
  right: number,
  top: number,
  other: readonly [number, number, number, number],
): boolean {
  return other[0] < right && other[2] > left && other[1] < top && other[3] > bottom
}

/**
 * Re-read a redacted document and prove the redaction actually took effect.
 *
 * The redacted copy is checked the way the redaction was requested: in text
 * mode every pattern is searched again with the same regex/whole-word/case
 * semantics and any surviving match fails the whole operation; in areas mode
 * every area rectangle is mapped back to PDF space and must no longer intersect
 * a text run. The returned error names the offending one-based pages, so a
 * pipeline that silently returned the original (or skipped a page while baking)
 * can never be presented as a successful redaction.
 *
 * Accepts either a loaded `PDFDocumentProxy` or the copy's bytes; bytes are
 * loaded with the renderer's usual pdf.js options and released afterwards.
 */
export async function verifyRedactedCopy(
  redacted: Uint8Array | PDFDocumentProxy,
  options: PdfRedactionOptions,
): Promise<void> {
  // A verification with nothing to look for would pass on any input; refuse it the
  // same way the redaction itself does instead of reporting a meaningless success.
  if ((options.mode ?? 'text') === 'text') {
    if (options.patterns.map((pattern) => pattern.trim()).filter(Boolean).length === 0) {
      throw new Error('At least one redaction pattern is required')
    }
  } else if (!options.areas || options.areas.length === 0) {
    throw new Error('At least one redaction area is required')
  }
  if (redacted instanceof Uint8Array) {
    // pdf.js may take ownership of the buffer it is handed; keep the caller's alive.
    const loadingTask = getDocument({ data: redacted.slice(), ...DOC_OPTS })
    try {
      await verifyRedactedCopy(await loadingTask.promise, options)
    } finally {
      await loadingTask.destroy()
    }
    return
  }

  const document = redacted
  const index = await buildSearchIndex(document)

  if ((options.mode ?? 'text') === 'text') {
    const remaining = searchPatternsInIndex(
      index,
      options.patterns,
      options.useRegex,
      options.wholeWord,
      options.matchCase ?? false,
    )
    const failedPages = [...new Set(remaining.map((match) => match.pageIndex))].sort(
      (left, right) => left - right,
    )
    if (failedPages.length > 0) {
      throw redactionVerificationError(failedPages, 'matching text is still present')
    }
    return
  }

  const failedPages: number[] = []
  for (const [pageIndex, areas] of redactionAreasByPage(options.areas ?? [])) {
    // An area whose page vanished from the output cannot be vouched for.
    if (pageIndex < 0 || pageIndex >= document.numPages) {
      failedPages.push(pageIndex)
      continue
    }
    const entry = index[pageIndex]
    if (!entry || entry.items.length === 0) continue
    const page = await document.getPage(pageIndex + 1)
    const viewport = page.getViewport({ scale: 1 })
    let hit = false
    for (const area of areas) {
      // Normalized area -> viewport pixels -> PDF user space, min/max over the
      // four corners so page rotation does not flip the comparison.
      const corners = [
        viewport.convertToPdfPoint(area.x * viewport.width, area.y * viewport.height),
        viewport.convertToPdfPoint(
          (area.x + area.width) * viewport.width,
          area.y * viewport.height,
        ),
        viewport.convertToPdfPoint(
          area.x * viewport.width,
          (area.y + area.height) * viewport.height,
        ),
        viewport.convertToPdfPoint(
          (area.x + area.width) * viewport.width,
          (area.y + area.height) * viewport.height,
        ),
      ]
      const xs = corners.map(([x]) => x ?? 0)
      const ys = corners.map(([, y]) => y ?? 0)
      const left = Math.min(...xs)
      const right = Math.max(...xs)
      const bottom = Math.min(...ys)
      const top = Math.max(...ys)
      if (
        entry.items.some((item) => {
          const bounds = itemPdfBounds(item)
          return bounds !== null && rectanglesOverlap(left, bottom, right, top, bounds)
        })
      ) {
        hit = true
        break
      }
    }
    if (hit) failedPages.push(pageIndex)
  }
  if (failedPages.length > 0) {
    throw redactionVerificationError(failedPages, 'a redaction area still contains text')
  }
}
