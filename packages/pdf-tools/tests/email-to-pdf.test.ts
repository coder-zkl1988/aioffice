import { describe, expect, it } from 'vitest'
import { PDFDocument } from 'pdf-lib'
import { emailDocumentOutputFileName, listPdfAttachmentsBytes, runPdfToolBytes } from '../src/index'

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

describe('email document PDF operation', () => {
  it('creates an A4 PDF and embeds requested attachments', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'emailToPdf',
      outputFormat: 'pdf',
      baseName: '../Quarterly.eml',
      pages: [{ image: onePixelPng, width: 595.28, height: 841.89 }],
      attachments: [
        {
          name: 'notes.txt',
          bytes: new TextEncoder().encode('notes'),
          mimeType: 'text/plain',
        },
      ],
    })

    expect(output?.fileName).toBe('Quarterly_converted.pdf')
    const document = await PDFDocument.load(output!.bytes)
    expect(document.getPageCount()).toBe(1)
    expect(document.getPage(0).getSize().width).toBeCloseTo(595.28, 1)
    expect(await listPdfAttachmentsBytes(output!.bytes)).toMatchObject([
      { name: 'notes.txt', size: 5, mimeType: 'text/plain' },
    ])
  })

  it('returns a standalone offline HTML output', async () => {
    const htmlBytes = new TextEncoder().encode('<!doctype html><title>Mail</title>')
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'emailToPdf',
      outputFormat: 'html',
      baseName: 'mail.eml',
      htmlBytes,
    })

    expect(output).toMatchObject({
      fileName: 'mail_converted.html',
      mimeType: 'text/html',
      extension: '.html',
    })
    expect(output?.bytes).toEqual(htmlBytes)
  })

  it('sanitizes output names and requires prepared content', async () => {
    expect(emailDocumentOutputFileName('../../<bad>.eml', 'pdf')).toBe('bad_converted.pdf')
    await expect(
      runPdfToolBytes(await sourcePdf(), { kind: 'emailToPdf', outputFormat: 'pdf' }),
    ).rejects.toThrow('pages')
    await expect(
      runPdfToolBytes(await sourcePdf(), { kind: 'emailToPdf', outputFormat: 'html' }),
    ).rejects.toThrow('HTML')
  })
})
