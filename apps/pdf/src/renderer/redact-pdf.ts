import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfRedactedPage, PdfRedactionArea, PdfRedactionOptions } from '@genoffice/pdf-tools'
import { flattenRenderScale } from './flatten-pdf'
import { buildSearchIndex, searchPatternsInIndex, type SearchMatch } from './search'

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
      for (const pattern of options.patterns) new RegExp(pattern, 'giu')
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
