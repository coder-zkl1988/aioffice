import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { createdPdfOutputFileName, htmlPdfOutputFileName, runPdfToolBytes } from '../src/index'

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

describe('HTML to PDF operation', () => {
  it('creates standalone A4 pages and preserves the page title', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'htmlToPdf',
      baseName: '../Website.zip',
      title: 'Website title',
      pages: [
        { image: onePixelPng, width: 595.28, height: 841.89 },
        { image: onePixelPng, width: 595.28, height: 841.89 },
      ],
    })

    expect(output?.fileName).toBe('Website_converted.pdf')
    const document = await PDFDocument.load(output!.bytes)
    expect(document.getPageCount()).toBe(2)
    expect(document.getPage(0).getSize().width).toBeCloseTo(595.28, 1)
    expect(document.getTitle()).toBe('Website title')
    expect(document.getCreator()).toBe('GenOffice PDF')
  })

  it('sanitizes names and requires prepared pages', async () => {
    expect(htmlPdfOutputFileName('../../<bad>.html')).toBe('bad_converted.pdf')
    await expect(
      runPdfToolBytes(await sourcePdf(), { kind: 'htmlToPdf', pages: [] }),
    ).rejects.toThrow('pages')
  })
})

describe('structured document PDF operation', () => {
  it('creates a standalone PDF with a clean requested name and metadata', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'createPdf',
      baseName: '../../Quarterly review.pdf',
      title: 'Quarterly review',
      pages: [{ image: onePixelPng, width: 595.28, height: 841.89 }],
    })

    expect(output?.fileName).toBe('Quarterly review.pdf')
    const document = await PDFDocument.load(output!.bytes)
    expect(document.getPageCount()).toBe(1)
    expect(document.getTitle()).toBe('Quarterly review')
    expect(document.getCreator()).toBe('GenOffice PDF')
  })

  it('sanitizes generated file names and requires prepared pages', async () => {
    expect(createdPdfOutputFileName('../../<bad>.pdf')).toBe('bad.pdf')
    await expect(
      runPdfToolBytes(await sourcePdf(), { kind: 'createPdf', pages: [] }),
    ).rejects.toThrow('pages')
  })
})
