import { PDFArray, PDFDict, PDFDocument, PDFHexString, PDFName, PDFWidgetAnnotation } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { readPdfClassificationMetadataBytes } from '@genoffice/pdf-tools'
import { applyWebPdfSave, extractWebPdf, insertWebPdf, insertWebPdfBlankPage } from './pdf'

const TINY_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

async function onePagePdf(): Promise<ArrayBuffer> {
  const document = await PDFDocument.create()
  document.addPage([300, 500])
  return (await document.save()).buffer as ArrayBuffer
}

function addSignatureField(
  document: PDFDocument,
  name: string,
  signed = false,
  withAppearance = true,
): void {
  const page = document.getPage(0)
  const form = document.getForm()
  const signatureDictionary = document.context.obj({
    FT: 'Sig',
    T: PDFHexString.fromText(name),
    ...(signed
      ? { V: { Type: 'Sig', Filter: 'Adobe.PPKLite', SubFilter: 'adbe.pkcs7.detached' } }
      : {}),
  })
  const signatureReference = document.context.register(signatureDictionary)
  const widget = PDFWidgetAnnotation.create(document.context, signatureReference)
  widget.setRectangle({ x: 40, y: 80, width: 160, height: 48 })
  widget.setP(page.ref)
  if (withAppearance) {
    const appearanceReference = document.context.register(
      document.context.formXObject([], { BBox: [0, 0, 160, 48], Resources: {} }),
    )
    widget.setNormalAppearance(appearanceReference)
  }
  const widgetReference = document.context.register(widget.dict)
  signatureDictionary.set(PDFName.of('Kids'), document.context.obj([widgetReference]))
  form.acroForm.addField(signatureReference)
  page.node.addAnnot(widgetReference)
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

  it('inserts a local blank page with the current page size', async () => {
    const inserted = await insertWebPdfBlankPage(await onePagePdf(), 0)
    const output = await PDFDocument.load(inserted.bytes)

    expect(inserted.count).toBe(1)
    expect(output.getPageCount()).toBe(2)
    expect(output.getPage(1).getSize()).toEqual({ width: 300, height: 500 })
  })

  it('inserts configured blank pages in the browser path', async () => {
    const inserted = await insertWebPdfBlankPage(await onePagePdf(), -1, {
      count: 2,
      pageSize: 'LETTER',
      orientation: 'landscape',
    })
    const output = await PDFDocument.load(inserted.bytes)

    expect(inserted.count).toBe(2)
    expect(output.getPageCount()).toBe(3)
    expect(output.getPage(0).getSize()).toEqual({ width: 792, height: 612 })
  })

  it('persists deduplicated center-rotated stamps', async () => {
    const result = await applyWebPdfSave(await onePagePdf(), {
      path: 'test.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      stampImages: [TINY_PNG],
      stamps: [
        { pageIndex: 0, image: '', imageIndex: 0, rect: [50, 100, 150, 150], rotation: 30 },
        { pageIndex: 0, image: '', imageIndex: 0, rect: [150, 300, 250, 350], rotation: -30 },
      ],
    })
    const saved = await PDFDocument.load(result.bytes)
    expect(saved.getPageCount()).toBe(1)
    expect(result.bytes.byteLength).toBeGreaterThan(500)
  })

  it('persists classification metadata in browser-generated PDF bytes', async () => {
    const result = await applyWebPdfSave(await onePagePdf(), {
      path: 'test.pdf',
      markups: [],
      drawings: [],
      formValues: [],
      stamps: [],
      classification: {
        labels: [{ id: 'invoice', name: 'Invoice' }],
        sensitivity: 'internal',
      },
    })

    expect(await readPdfClassificationMetadataBytes(result.bytes)).toEqual({
      labels: [{ id: 'invoice', name: 'Invoice' }],
      sensitivity: 'internal',
    })
  })

  it('persists multi-select lists and editable dropdown values in browser saves', async () => {
    const source = await PDFDocument.create()
    const page = source.addPage([300, 300])
    const form = source.getForm()
    const topics = form.createOptionList('topics')
    topics.addOptions(['Design', 'Engineering', 'Research'])
    topics.enableMultiselect()
    topics.addToPage(page, { x: 20, y: 180, width: 160, height: 60 })
    const department = form.createDropdown('department')
    department.addOptions(['Sales', 'Support'])
    department.enableEditing()
    department.addToPage(page, { x: 20, y: 130, width: 160, height: 24 })

    const result = await applyWebPdfSave((await source.save()).buffer as ArrayBuffer, {
      path: 'form.pdf',
      markups: [],
      drawings: [],
      formValues: [
        { name: 'topics', kind: 'choice', value: ['Design', 'Research'] },
        { name: 'department', kind: 'choice', value: 'Customer Success' },
      ],
      stamps: [],
    })
    const saved = await PDFDocument.load(result.bytes)

    expect(saved.getForm().getOptionList('topics').getSelected()).toEqual(['Design', 'Research'])
    expect(saved.getForm().getDropdown('department').getSelected()).toEqual(['Customer Success'])
  })

  it('removes completed empty signature fields in browser saves', async () => {
    const source = await PDFDocument.create()
    source.addPage([300, 300])
    addSignatureField(source, 'approval', false, false)
    addSignatureField(source, 'existing_digital_signature', true)

    const result = await applyWebPdfSave(
      (await source.save({ useObjectStreams: false, updateFieldAppearances: false }))
        .buffer as ArrayBuffer,
      {
        path: 'signature-field.pdf',
        markups: [],
        drawings: [],
        removeSignatureFields: ['approval', 'existing_digital_signature'],
        formValues: [],
        stamps: [],
      },
    )
    const saved = await PDFDocument.load(result.bytes)

    expect(
      saved
        .getForm()
        .getFields()
        .map((field) => field.getName()),
    ).toEqual(['existing_digital_signature'])
  })

  it('persists multilingual review notes as PDF annotations', async () => {
    const result = await applyWebPdfSave(await onePagePdf(), {
      path: 'test.pdf',
      markups: [],
      drawings: [
        {
          kind: 'note',
          pageIndex: 0,
          color: [1, 0.82, 0.2],
          at: [40, 420],
          contents: '请确认共享范围',
          subject: '敏感信息',
          author: 'GenOffice 文档审查',
        },
      ],
      formValues: [],
      stamps: [],
    })
    const document = await PDFDocument.load(result.bytes)
    const annotations = document.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const annotation = annotations.lookup(0, PDFDict)

    expect(annotation.lookup(PDFName.of('Subtype'), PDFName).decodeText()).toBe('Text')
    expect(annotation.lookup(PDFName.of('Contents'), PDFHexString).decodeText()).toBe(
      '请确认共享范围',
    )
    expect(annotation.lookup(PDFName.of('Subj'), PDFHexString).decodeText()).toBe('敏感信息')
  })
})
