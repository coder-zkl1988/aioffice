import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { verifyRedactedCopy } from '../src/renderer/redact-pdf'
import type { PdfRedactionOptions } from '@genoffice/pdf-tools'

const baseOptions: PdfRedactionOptions = {
  patterns: ['secret'],
  useRegex: false,
  wholeWord: false,
  color: '#000000',
  padding: 2,
  renderDpi: 144,
}

interface FakeItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
}

function fakeDoc(pages: FakeItem[][], size = { width: 100, height: 100 }): PDFDocumentProxy {
  return {
    numPages: pages.length,
    getPage: async (pageNumber: number) => ({
      getTextContent: async () => ({ items: pages[pageNumber - 1] }),
      getViewport: () => ({
        width: size.width,
        height: size.height,
        // pdf.js default viewport: top-left origin, y flipped from PDF space.
        convertToPdfPoint: (x: number, y: number) => [x, size.height - y],
      }),
    }),
  } as unknown as PDFDocumentProxy
}

const item = (str: string, x: number, y: number, w: number, h: number): FakeItem => ({
  str,
  transform: [1, 0, 0, 1, x, y],
  width: w,
  height: h,
})

const areaOptions = (areas: PdfRedactionOptions['areas']): PdfRedactionOptions => ({
  ...baseOptions,
  mode: 'areas',
  patterns: [],
  areas,
})

describe('verifyRedactedCopy', () => {
  it('passes when the redacted document no longer contains a pattern', async () => {
    await expect(
      verifyRedactedCopy(fakeDoc([[item('public text', 10, 10, 80, 12)]]), baseOptions),
    ).resolves.toBeUndefined()
  })

  it('names the page that still contains a pattern', async () => {
    const document = fakeDoc([
      [item('public', 10, 10, 40, 12)],
      [item('top secret memo', 10, 10, 80, 12)],
    ])
    await expect(verifyRedactedCopy(document, baseOptions)).rejects.toThrow(
      'Redaction could not be verified on page 2: matching text is still present, so no file was returned.',
    )
  })

  it('lists every failing page, one-based and sorted', async () => {
    const document = fakeDoc([
      [item('secret one', 10, 10, 60, 12)],
      [item('clean', 10, 10, 40, 12)],
      [item('secret three', 10, 10, 70, 12)],
    ])
    await expect(verifyRedactedCopy(document, baseOptions)).rejects.toThrow(
      /could not be verified on pages 1, 3/,
    )
  })

  it('honours whole-word, case and regex matching exactly like the redactor', async () => {
    const glued = fakeDoc([[item('secretive', 10, 10, 60, 12)]])
    await expect(verifyRedactedCopy(glued, baseOptions)).rejects.toThrow(/page 1/)
    await expect(
      verifyRedactedCopy(glued, { ...baseOptions, wholeWord: true }),
    ).resolves.toBeUndefined()

    const lower = fakeDoc([[item('secret', 10, 10, 60, 12)]])
    await expect(
      verifyRedactedCopy(lower, { ...baseOptions, patterns: ['Secret'] }),
    ).rejects.toThrow(/page 1/)
    await expect(
      verifyRedactedCopy(lower, { ...baseOptions, patterns: ['Secret'], matchCase: true }),
    ).resolves.toBeUndefined()

    const coded = fakeDoc([[item('S42 project', 10, 10, 60, 12)]])
    const regexOptions = { ...baseOptions, patterns: ['S\\d+'], useRegex: true }
    await expect(verifyRedactedCopy(coded, regexOptions)).rejects.toThrow(/page 1/)
    await expect(
      verifyRedactedCopy(coded, { ...regexOptions, patterns: ['S\\d{4}'] }),
    ).resolves.toBeUndefined()
  })

  it('fails an area that still intersects a text run', async () => {
    const document = fakeDoc([[item('secret', 10, 10, 50, 10)]])
    // Viewport y 70..90 maps to PDF y 10..30, covering the run at x 10..60, y 10..20.
    await expect(
      verifyRedactedCopy(
        document,
        areaOptions([{ pageIndex: 0, x: 0.1, y: 0.7, width: 0.5, height: 0.2 }]),
      ),
    ).rejects.toThrow(
      'Redaction could not be verified on page 1: a redaction area still contains text, so no file was returned.',
    )
  })

  it('accepts an area that cleared the text run', async () => {
    const document = fakeDoc([[item('secret', 10, 10, 50, 10)]])
    // Viewport y 10..20 maps to PDF y 80..90 — nothing there.
    await expect(
      verifyRedactedCopy(
        document,
        areaOptions([{ pageIndex: 0, x: 0.1, y: 0.1, width: 0.5, height: 0.1 }]),
      ),
    ).resolves.toBeUndefined()
  })

  it('fails an area whose page is missing from the copy', async () => {
    const document = fakeDoc([[item('clean', 10, 10, 40, 12)]])
    await expect(
      verifyRedactedCopy(
        document,
        areaOptions([{ pageIndex: 3, x: 0.1, y: 0.1, width: 0.2, height: 0.2 }]),
      ),
    ).rejects.toThrow(/could not be verified on page 4/)
  })

  it('checks the copy bytes with the same rules and releases the loading task', async () => {
    const source = await PDFDocument.create()
    const page = source.addPage([300, 200])
    const font = await source.embedFont(StandardFonts.Helvetica)
    page.drawText('secret budget', { x: 20, y: 150, size: 18, font })
    const clean = await PDFDocument.create()
    clean.addPage([300, 200])

    await expect(verifyRedactedCopy(await source.save(), baseOptions)).rejects.toThrow(/page 1/)
    await expect(verifyRedactedCopy(await clean.save(), baseOptions)).resolves.toBeUndefined()
  })

  it('refuses a verification with nothing to look for', async () => {
    const document = fakeDoc([[item('anything', 10, 10, 50, 10)]])
    await expect(
      verifyRedactedCopy(document, { ...baseOptions, patterns: ['  '] }),
    ).rejects.toThrow(/at least one redaction pattern/i)
    await expect(
      verifyRedactedCopy(document, { ...baseOptions, mode: 'areas', patterns: [], areas: [] }),
    ).rejects.toThrow(/at least one redaction area/i)
  })

  it('passes once the matched page has been replaced by a text-free bitmap', async () => {
    const source = await PDFDocument.create()
    const first = source.addPage([300, 200])
    const second = source.addPage([300, 200])
    const font = await source.embedFont(StandardFonts.Helvetica)
    first.drawText('secret budget', { x: 20, y: 150, size: 18, font })
    second.drawText('public notes', { x: 20, y: 150, size: 18, font })

    const redacted = await PDFDocument.create()
    redacted.addPage([300, 200]) // rasterized stand-in: no text layer
    const [copiedSecond] = await redacted.copyPages(source, [1])
    redacted.addPage(copiedSecond)

    await expect(verifyRedactedCopy(await redacted.save(), baseOptions)).resolves.toBeUndefined()
  })
})
