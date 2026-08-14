import { PDFDict, PDFDocument, PDFName, StandardFonts, rgb } from 'pdf-lib'
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { addPdfOcrTextLayersBytes, pdfOcrSidecarText, runPdfToolBytes } from '../src/index'

async function pdfWithText(text: string, width = 300, height = 200): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  const page = document.addPage([width, height])
  const font = await document.embedFont(StandardFonts.Helvetica)
  page.drawText(text, { x: 24, y: height - 40, size: 18, font, color: rgb(0, 0, 0) })
  return document.save({ useObjectStreams: false })
}

async function blankPdf(pageCount = 2): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  for (let index = 0; index < pageCount; index++) document.addPage([300, 200])
  return document.save({ useObjectStreams: false })
}

describe('PDF OCR text layers', () => {
  it('adds a recognized text layer to the selected page', async () => {
    const output = await addPdfOcrTextLayersBytes(await blankPdf(), [
      { pageIndex: 1, bytes: await pdfWithText('SEARCHABLE OCR') },
    ])
    const document = await PDFDocument.load(output)
    expect(document.getPageCount()).toBe(2)
    const secondResources = document.getPage(1).node.Resources()
    expect(secondResources?.lookupMaybe(PDFName.of('XObject'), PDFDict)).toBeDefined()
  })

  it('validates OCR layer pages and duplicate inputs', async () => {
    const layer = await pdfWithText('OCR')
    await expect(
      addPdfOcrTextLayersBytes(await blankPdf(1), [{ pageIndex: 1, bytes: layer }]),
    ).rejects.toThrow(/page is invalid/i)
    await expect(
      addPdfOcrTextLayersBytes(await blankPdf(1), [
        { pageIndex: 0, bytes: layer },
        { pageIndex: 0, bytes: layer },
      ]),
    ).rejects.toThrow(/duplicated/i)
  })

  it('runs the shared OCR tool contract', async () => {
    const [output] = await runPdfToolBytes(await blankPdf(1), {
      kind: 'ocr',
      mode: 'skipText',
      languages: ['eng'],
      renderDpi: 200,
      clean: true,
      sidecar: false,
      textLayers: [{ pageIndex: 0, bytes: await pdfWithText('LOCAL OCR') }],
      skippedPageIndexes: [],
    })
    expect(output.suffix).toBe('_ocr.pdf')
    await expect(PDFDocument.load(output.bytes)).resolves.toBeInstanceOf(PDFDocument)
  })

  it('packages the searchable PDF and page-ordered sidecar text', async () => {
    const [output] = await runPdfToolBytes(await blankPdf(2), {
      kind: 'ocr',
      mode: 'skipText',
      languages: ['eng', 'chi_sim'],
      renderDpi: 200,
      clean: true,
      sidecar: true,
      baseName: 'Quarterly/Scan.pdf',
      textLayers: [{ pageIndex: 0, bytes: await pdfWithText('LOCAL OCR') }],
      pageTexts: [
        { pageIndex: 1, text: 'Existing page text', source: 'existing' },
        { pageIndex: 0, text: 'Recognized page text', source: 'ocr' },
      ],
      skippedPageIndexes: [1],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_ocr.zip',
        fileName: 'Quarterly_Scan_OCR.zip',
        mimeType: 'application/zip',
        extension: '.zip',
      }),
    )
    const archive = await JSZip.loadAsync(output.bytes)
    expect(Object.keys(archive.files).sort()).toEqual([
      'Quarterly_Scan_OCR.pdf',
      'Quarterly_Scan_OCR.txt',
    ])
    await expect(
      PDFDocument.load(await archive.file('Quarterly_Scan_OCR.pdf')!.async('uint8array')),
    ).resolves.toBeInstanceOf(PDFDocument)
    expect(await archive.file('Quarterly_Scan_OCR.txt')!.async('string')).toBe(
      '===== Page 1 =====\nRecognized page text\n\n\f\n\n===== Page 2 =====\nExisting page text\n',
    )
  })

  it('marks empty pages and validates duplicate sidecar entries', () => {
    expect(pdfOcrSidecarText(2, [{ pageIndex: 0, text: 'OCR', source: 'ocr' }])).toBe(
      '===== Page 1 =====\nOCR\n\n\f\n\n===== Page 2 =====\n[No text recognized]\n',
    )
    expect(() =>
      pdfOcrSidecarText(1, [
        { pageIndex: 0, text: 'First', source: 'ocr' },
        { pageIndex: 0, text: 'Second', source: 'existing' },
      ]),
    ).toThrow(/duplicated/i)
  })
})
