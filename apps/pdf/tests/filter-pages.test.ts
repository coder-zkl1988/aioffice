import { OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { describe, expect, it, vi } from 'vitest'
import {
  analyzePdfContentFilter,
  pdfOperatorListHasImage,
  pdfPageTextMatches,
} from '../src/renderer/filter-pages'

describe('PDF content filter analysis', () => {
  it('matches literal text with case and whole-word controls', () => {
    expect(pdfPageTextMatches('Launch A+B now', 'a+b', false, false)).toBe(true)
    expect(pdfPageTextMatches('Launch A+B now', 'a+b', true, false)).toBe(false)
    expect(pdfPageTextMatches('cat category', 'cat', false, true)).toBe(true)
    expect(pdfPageTextMatches('category', 'cat', false, true)).toBe(false)
    expect(pdfPageTextMatches('山东文化旅行', '文化', false, true)).toBe(false)
    expect(pdfPageTextMatches('山东 文化 旅行', '文化', false, true)).toBe(true)
    expect(() => pdfPageTextMatches('anything', '  ', false, false)).toThrow('Enter text')
  })

  it('recognizes raster, inline, repeat, and mask image operators', () => {
    expect(pdfOperatorListHasImage([OPS.save, OPS.paintImageXObject, OPS.restore])).toBe(true)
    expect(pdfOperatorListHasImage([OPS.paintInlineImageXObject])).toBe(true)
    expect(pdfOperatorListHasImage([OPS.paintImageMaskXObjectRepeat])).toBe(true)
    expect(pdfOperatorListHasImage([OPS.paintSolidColorImageMask])).toBe(true)
    expect(pdfOperatorListHasImage([OPS.save, OPS.paintPath, OPS.restore])).toBe(false)
  })

  it('analyzes only requested pages and cleans each page', async () => {
    const cleanups = [vi.fn(), vi.fn(), vi.fn()]
    const pages = [
      { text: 'Cover', operators: [OPS.paintImageXObject] },
      { text: 'Launch plan', operators: [OPS.paintPath] },
      { text: 'Appendix', operators: [OPS.paintInlineImageXObject] },
    ]
    const document = {
      numPages: 3,
      getPage: vi.fn(async (pageNumber: number) => {
        const page = pages[pageNumber - 1]!
        return {
          getTextContent: async () => ({ items: [{ str: page.text }] }),
          getOperatorList: async () => ({ fnArray: page.operators, argsArray: [] }),
          cleanup: cleanups[pageNumber - 1],
        }
      }),
    } as unknown as PDFDocumentProxy

    await expect(
      analyzePdfContentFilter(document, {
        criterion: 'text',
        pageIndexes: [1, 2],
        text: 'launch',
        caseSensitive: false,
        wholeWord: false,
      }),
    ).resolves.toEqual([1])
    expect(document.getPage).toHaveBeenNthCalledWith(1, 2)
    expect(document.getPage).toHaveBeenNthCalledWith(2, 3)
    expect(cleanups[1]).toHaveBeenCalledOnce()
    expect(cleanups[2]).toHaveBeenCalledOnce()

    await expect(
      analyzePdfContentFilter(document, {
        criterion: 'image',
        pageIndexes: [0, 2],
        caseSensitive: false,
        wholeWord: false,
      }),
    ).resolves.toEqual([0, 2])
  })

  it('leaves geometric criteria to the shared PDF engine', async () => {
    const document = {
      numPages: 1,
      getPage: vi.fn(),
    } as unknown as PDFDocumentProxy
    await expect(
      analyzePdfContentFilter(document, {
        criterion: 'orientation',
        pageIndexes: [0],
        caseSensitive: false,
        wholeWord: false,
      }),
    ).resolves.toEqual([])
    expect(document.getPage).not.toHaveBeenCalled()
  })

  it('validates pages and cleans resources after a failure', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 1,
      getPage: async () => ({
        getTextContent: async () => {
          throw new Error('text failure')
        },
        cleanup,
      }),
    } as unknown as PDFDocumentProxy
    await expect(
      analyzePdfContentFilter(document, {
        criterion: 'text',
        pageIndexes: [0],
        text: 'x',
        caseSensitive: false,
        wholeWord: false,
      }),
    ).rejects.toThrow('text failure')
    expect(cleanup).toHaveBeenCalledOnce()
    await expect(
      analyzePdfContentFilter(document, {
        criterion: 'image',
        pageIndexes: [1],
        caseSensitive: false,
        wholeWord: false,
      }),
    ).rejects.toThrow('invalid page')
  })
})
