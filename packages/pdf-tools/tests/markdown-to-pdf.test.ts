import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { markdownPdfOutputFileName, runPdfToolBytes } from '../src/index'

const onePixelPng = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z2S8AAAAASUVORK5CYII=',
    'base64',
  ),
)

async function sourcePdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  document.addPage([100, 100])
  return document.save()
}

describe('Markdown to PDF operation', () => {
  it('creates standalone A4 pages and preserves the document title', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'markdownToPdf',
      baseName: '../Guide.markdown',
      title: 'Guide title',
      pages: [
        { image: onePixelPng, width: 595.28, height: 841.89 },
        { image: onePixelPng, width: 595.28, height: 841.89 },
      ],
    })

    expect(output?.fileName).toBe('Guide_converted.pdf')
    const document = await PDFDocument.load(output!.bytes)
    expect(document.getPageCount()).toBe(2)
    expect(document.getPage(0).getSize().height).toBeCloseTo(841.89, 1)
    expect(document.getTitle()).toBe('Guide title')
    expect(document.getCreator()).toBe('GenOffice PDF')
  })

  it('sanitizes names and requires prepared pages', async () => {
    expect(markdownPdfOutputFileName('../../<bad>.md')).toBe('bad_converted.pdf')
    await expect(
      runPdfToolBytes(await sourcePdf(), { kind: 'markdownToPdf', pages: [] }),
    ).rejects.toThrow('pages')
  })
})
