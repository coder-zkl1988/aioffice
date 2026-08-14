import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { epubPdfOutputFileName, runPdfToolBytes } from '../src/index'

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

describe('EPUB to PDF operation', () => {
  it('creates A4 pages and preserves ebook metadata', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'epubToPdf',
      baseName: '../Novel.epub',
      title: 'Novel title',
      author: 'Author name',
      pages: [
        { image: onePixelPng, width: 595.28, height: 841.89 },
        { image: onePixelPng, width: 595.28, height: 841.89 },
      ],
    })

    expect(output?.fileName).toBe('Novel_converted.pdf')
    const document = await PDFDocument.load(output!.bytes)
    expect(document.getPageCount()).toBe(2)
    expect(document.getPage(0).getSize().width).toBeCloseTo(595.28, 1)
    expect(document.getTitle()).toBe('Novel title')
    expect(document.getAuthor()).toBe('Author name')
    expect(document.getCreator()).toBe('GenOffice PDF')
  })

  it('sanitizes names and requires prepared pages', async () => {
    expect(epubPdfOutputFileName('../../<bad>.epub')).toBe('bad_converted.pdf')
    await expect(
      runPdfToolBytes(await sourcePdf(), { kind: 'epubToPdf', pages: [] }),
    ).rejects.toThrow('pages')
  })
})
