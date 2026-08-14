import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfContentFilterCriterion } from '@genoffice/pdf-tools'

const IMAGE_OPERATORS = new Set<number>([
  OPS.paintImageXObject,
  OPS.paintImageXObjectRepeat,
  OPS.paintInlineImageXObject,
  OPS.paintInlineImageXObjectGroup,
  OPS.paintImageMaskXObject,
  OPS.paintImageMaskXObjectGroup,
  OPS.paintImageMaskXObjectRepeat,
  OPS.paintSolidColorImageMask,
])

export interface PdfContentFilterAnalysisOptions {
  criterion: PdfContentFilterCriterion
  pageIndexes: number[]
  text?: string
  caseSensitive: boolean
  wholeWord: boolean
}

function escapedRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

export function pdfPageTextMatches(
  pageText: string,
  query: string,
  caseSensitive: boolean,
  wholeWord: boolean,
): boolean {
  const needle = query.trim()
  if (!needle) throw new Error('Enter text to find')
  const expression = wholeWord
    ? `(?<![\\p{L}\\p{N}_])${escapedRegExp(needle)}(?![\\p{L}\\p{N}_])`
    : escapedRegExp(needle)
  return new RegExp(expression, caseSensitive ? 'u' : 'iu').test(pageText)
}

export function pdfOperatorListHasImage(functions: ArrayLike<number>): boolean {
  for (let index = 0; index < functions.length; index++) {
    if (IMAGE_OPERATORS.has(functions[index]!)) return true
  }
  return false
}

export async function analyzePdfContentFilter(
  document: PDFDocumentProxy,
  options: PdfContentFilterAnalysisOptions,
): Promise<number[]> {
  if (options.criterion !== 'text' && options.criterion !== 'image') return []
  const uniquePages = [...new Set(options.pageIndexes)]
  if (
    uniquePages.length === 0 ||
    uniquePages.some(
      (pageIndex) =>
        !Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= document.numPages,
    )
  ) {
    throw new Error('Content filter page indexes contain an invalid page')
  }
  const matched: number[] = []
  for (const pageIndex of uniquePages) {
    const page = await document.getPage(pageIndex + 1)
    try {
      if (options.criterion === 'image') {
        const operators = await page.getOperatorList()
        if (pdfOperatorListHasImage(operators.fnArray)) matched.push(pageIndex)
      } else {
        const content = await page.getTextContent()
        const pageText = content.items
          .map((item) => ('str' in item ? item.str : ''))
          .filter(Boolean)
          .join(' ')
        if (
          pdfPageTextMatches(pageText, options.text ?? '', options.caseSensitive, options.wholeWord)
        ) {
          matched.push(pageIndex)
        }
      }
    } finally {
      page.cleanup()
    }
  }
  return matched
}
