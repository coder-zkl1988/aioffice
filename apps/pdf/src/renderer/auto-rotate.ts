import type { PdfPageRotation, PdfPageRotationCorrection } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'

const MIN_GLYPHS = 30
const MIN_DOMINANCE = 0.95
const MIN_GLYPHS_UNANIMOUS = 8
const UNANIMOUS_DOMINANCE = 0.99

interface RawTextItem {
  str?: string
  transform?: number[]
}

export interface PdfTextDirection {
  dominantDirection: number
  dominance: number
  glyphCount: number
  conclusive: boolean
}

export interface PdfAutoRotatePageResult {
  pageIndex: number
  currentRotation: number
  correction: number
  method: 'text' | 'inferred' | 'none'
}

export interface PdfAutoRotateAnalysis {
  pages: PdfAutoRotatePageResult[]
  pageRotations: PdfPageRotationCorrection[]
}

function normalizedQuarterTurn(value: number): number {
  return ((value % 360) + 360) % 360
}

function textDirection(transform: number[]): number | null {
  const horizontal = transform[0] ?? 0
  const vertical = transform[1] ?? 0
  if (Math.hypot(horizontal, vertical) < 1e-6) return null
  if (Math.abs(horizontal) >= Math.abs(vertical)) return horizontal >= 0 ? 0 : 180
  return vertical >= 0 ? 90 : 270
}

export function detectPdfTextDirection(items: RawTextItem[]): PdfTextDirection {
  const counts = [0, 0, 0, 0]
  for (const item of items) {
    if (!item.transform || typeof item.str !== 'string') continue
    const glyphCount = Array.from(item.str).filter((character) => !/\s/u.test(character)).length
    if (glyphCount === 0) continue
    const direction = textDirection(item.transform)
    if (direction === null) continue
    counts[direction / 90]! += glyphCount
  }
  const total = counts.reduce((sum, count) => sum + count, 0)
  let bestIndex = 0
  for (let index = 1; index < counts.length; index++) {
    if (counts[index]! > counts[bestIndex]!) bestIndex = index
  }
  const dominance = total === 0 ? 0 : counts[bestIndex]! / total
  return {
    dominantDirection: bestIndex * 90,
    dominance,
    glyphCount: total,
    conclusive:
      (total >= MIN_GLYPHS && dominance >= MIN_DOMINANCE) ||
      (total >= MIN_GLYPHS_UNANIMOUS && dominance >= UNANIMOUS_DOMINANCE),
  }
}

function inferUndetectedPages(results: PdfAutoRotatePageResult[]): void {
  const consensus = new Map<number, number>()
  const conflicted = new Set<number>()
  for (const result of results) {
    if (result.method !== 'text' || conflicted.has(result.currentRotation)) continue
    const existing = consensus.get(result.currentRotation)
    if (existing === undefined) consensus.set(result.currentRotation, result.correction)
    else if (existing !== result.correction) {
      consensus.delete(result.currentRotation)
      conflicted.add(result.currentRotation)
    }
  }
  for (const result of results) {
    if (result.method !== 'none') continue
    const correction = consensus.get(result.currentRotation)
    if (correction === undefined) continue
    result.correction = correction
    result.method = 'inferred'
  }
}

export async function analyzePdfAutoRotation(
  document: PDFDocumentProxy,
  inferUndetected: boolean,
): Promise<PdfAutoRotateAnalysis> {
  const pages: PdfAutoRotatePageResult[] = []
  for (let pageIndex = 0; pageIndex < document.numPages; pageIndex++) {
    const page = await document.getPage(pageIndex + 1)
    const content = await page.getTextContent()
    const direction = detectPdfTextDirection(content.items as RawTextItem[])
    const currentRotation = normalizedQuarterTurn(page.rotate)
    pages.push({
      pageIndex,
      currentRotation,
      correction: direction.conclusive
        ? normalizedQuarterTurn(direction.dominantDirection - currentRotation)
        : 0,
      method: direction.conclusive ? 'text' : 'none',
    })
  }
  if (inferUndetected) inferUndetectedPages(pages)
  return {
    pages,
    pageRotations: pages.flatMap((page) =>
      page.correction === 0
        ? []
        : [
            {
              pageIndex: page.pageIndex,
              angle: page.correction as PdfPageRotation,
            },
          ],
    ),
  }
}
