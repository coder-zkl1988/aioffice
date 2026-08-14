import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { analyzePdfAutoRotation, detectPdfTextDirection } from '../src/renderer/auto-rotate'

interface FakeTextItem {
  str: string
  transform: number[]
}

function textItem(text: string, angle: 0 | 90 | 180 | 270): FakeTextItem {
  const transform =
    angle === 90
      ? [0, 12, -12, 0, 0, 0]
      : angle === 180
        ? [-12, 0, 0, -12, 0, 0]
        : angle === 270
          ? [0, -12, 12, 0, 0, 0]
          : [12, 0, 0, 12, 0, 0]
  return { str: text, transform }
}

function fakeDocument(pages: Array<{ rotate: number; items: FakeTextItem[] }>): PDFDocumentProxy {
  return {
    numPages: pages.length,
    getPage: async (pageNumber: number) => ({
      rotate: pages[pageNumber - 1]!.rotate,
      getTextContent: async () => ({ items: pages[pageNumber - 1]!.items }),
    }),
  } as unknown as PDFDocumentProxy
}

describe('detectPdfTextDirection', () => {
  it.each([0, 90, 180, 270] as const)('detects %s degree text', (angle) => {
    const result = detectPdfTextDirection([
      textItem('The quick brown fox jumps over the lazy dog', angle),
    ])
    expect(result).toMatchObject({ dominantDirection: angle, conclusive: true })
    expect(result.dominance).toBe(1)
  })

  it('trusts sparse but unanimous text and rejects mixed directions', () => {
    expect(detectPdfTextDirection([textItem('york.gov', 90)]).conclusive).toBe(true)
    expect(
      detectPdfTextDirection([
        textItem('The quick brown fox jumps over the lazy dog', 0),
        textItem('The quick brown fox jumps over the lazy dog', 90),
      ]).conclusive,
    ).toBe(false)
  })
})

describe('analyzePdfAutoRotation', () => {
  it('computes additive corrections from text direction and current rotation', async () => {
    const analysis = await analyzePdfAutoRotation(
      fakeDocument([
        { rotate: 90, items: [textItem('The quick brown fox jumps over the lazy dog', 0)] },
        { rotate: 0, items: [textItem('The quick brown fox jumps over the lazy dog', 90)] },
        { rotate: 180, items: [textItem('The quick brown fox jumps over the lazy dog', 180)] },
      ]),
      true,
    )
    expect(analysis.pages.map((page) => page.correction)).toEqual([270, 90, 0])
    expect(analysis.pageRotations).toEqual([
      { pageIndex: 0, angle: 270 },
      { pageIndex: 1, angle: 90 },
    ])
  })

  it('infers a textless page only when same-rotation pages agree', async () => {
    const pages = [
      { rotate: 90, items: [textItem('The quick brown fox jumps over the lazy dog', 0)] },
      { rotate: 90, items: [] },
    ]
    const inferred = await analyzePdfAutoRotation(fakeDocument(pages), true)
    expect(inferred.pages[1]).toMatchObject({ correction: 270, method: 'inferred' })
    const disabled = await analyzePdfAutoRotation(fakeDocument(pages), false)
    expect(disabled.pages[1]).toMatchObject({ correction: 0, method: 'none' })
  })

  it('does not infer when decided pages with the same rotation conflict', async () => {
    const analysis = await analyzePdfAutoRotation(
      fakeDocument([
        { rotate: 0, items: [textItem('The quick brown fox jumps over the lazy dog', 0)] },
        { rotate: 0, items: [textItem('The quick brown fox jumps over the lazy dog', 90)] },
        { rotate: 0, items: [] },
      ]),
      true,
    )
    expect(analysis.pages[2]).toMatchObject({ correction: 0, method: 'none' })
  })
})
