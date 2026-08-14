import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { resolvePdfCommentAnchors } from '../src/renderer/pdf-comments'

function fakeDocument(): PDFDocumentProxy {
  return {
    numPages: 2,
    getPage: async (pageNumber: number) => ({
      getTextContent: async () => ({
        items:
          pageNumber === 2
            ? [
                {
                  str: 'Quarterly revenue',
                  transform: [1, 0, 0, 1, 120, 560],
                  width: 144,
                  height: 12,
                },
              ]
            : [],
      }),
    }),
  } as unknown as PDFDocumentProxy
}

describe('resolvePdfCommentAnchors', () => {
  it('places a comment beside the first matching text on its target page', async () => {
    const [comment] = await resolvePdfCommentAnchors(fakeDocument(), [
      {
        pageIndex: 1,
        x: 24,
        y: 24,
        width: 20,
        height: 20,
        text: 'Check the number',
        anchorText: 'REVENUE',
      },
    ])
    expect(comment?.x).toBeCloseTo(204.71, 2)
    expect(comment?.y).toBe(552)
  })

  it('keeps fallback coordinates when the anchor is not found', async () => {
    const [comment] = await resolvePdfCommentAnchors(fakeDocument(), [
      {
        pageIndex: 0,
        x: 30,
        y: 40,
        width: 20,
        height: 20,
        text: 'Fallback',
        anchorText: 'missing',
      },
    ])
    expect(comment).toMatchObject({ x: 30, y: 40 })
  })
})
