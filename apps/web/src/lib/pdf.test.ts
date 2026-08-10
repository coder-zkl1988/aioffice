import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { applyWebPdfSave, extractWebPdf, insertWebPdf } from './pdf'

async function onePagePdf(): Promise<ArrayBuffer> {
  const document = await PDFDocument.create()
  document.addPage([300, 500])
  return (await document.save()).buffer as ArrayBuffer
}

describe('web PDF operations', () => {
  it('persists rotation and markup into PDF bytes', async () => {
    const result = await applyWebPdfSave(await onePagePdf(), {
      path: 'test.pdf',
      markups: [
        {
          pageIndex: 0,
          type: 'highlight',
          color: [1, 1, 0],
          quads: [[20, 80, 120, 80, 20, 60, 120, 60]],
        },
      ],
      drawings: [],
      formValues: [],
      stamps: [],
      rotations: [{ pageIndex: 0, delta: 90 }],
    })
    const saved = await PDFDocument.load(result.bytes)
    expect(saved.getPage(0).getRotation().angle).toBe(90)
    expect(result.bytes.byteLength).toBeGreaterThan(500)
  })

  it('extracts and inserts pages', async () => {
    const source = await onePagePdf()
    const extracted = await extractWebPdf(source, [0])
    expect((await PDFDocument.load(extracted)).getPageCount()).toBe(1)

    const inserted = await insertWebPdf(source, source, 0)
    expect(inserted.count).toBe(1)
    expect((await PDFDocument.load(inserted.bytes)).getPageCount()).toBe(2)
  })
})
