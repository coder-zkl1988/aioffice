import { createHash } from 'node:crypto'
import { mkdtempSync, readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFWidgetAnnotation,
  degrees,
} from 'pdf-lib'
import { readPdfClassificationMetadataBytes } from '@genoffice/pdf-tools'
import {
  applySaveRequest,
  extractPagesBytes,
  insertBlankPageBytes,
  insertPdfBytes,
  savePdfToPath,
} from '../src/main/save-pdf'
import type { SavePdfRequest } from '../src/shared/ipc'

/** 1x1 red pixel PNG */
const TINY_PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

async function makePdf(sizes: [number, number][]): Promise<Uint8Array> {
  const doc = await PDFDocument.create()
  for (const size of sizes) doc.addPage(size)
  return doc.save({ useObjectStreams: false })
}

/** applySaveRequest with the skipped-text-edit channel unwrapped (none expected here) */
async function apply(bytes: Uint8Array, req: SavePdfRequest): Promise<Uint8Array> {
  const result = await applySaveRequest(bytes, req)
  expect(result.skippedTextEdits).toEqual([])
  return result.bytes
}

const request = (over: Partial<SavePdfRequest> = {}): SavePdfRequest => ({
  path: '/tmp/test.pdf',
  markups: [],
  drawings: [],
  formValues: [],
  stamps: [],
  ...over,
})

function pageAnnots(doc: PDFDocument, pageIndex: number): PDFDict[] {
  const annots = doc.getPage(pageIndex).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (!annots) return []
  return Array.from({ length: annots.size() }, (_, i) => annots.lookup(i, PDFDict))
}

function addSignatureField(
  document: PDFDocument,
  name: string,
  y: number,
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
  widget.setRectangle({ x: 40, y, width: 160, height: 48 })
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

const subtypeOf = (annot: PDFDict) => annot.lookup(PDFName.of('Subtype'), PDFName).decodeText()

describe('extractPagesBytes', () => {
  it('extracts the requested pages in the given order', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const out = await PDFDocument.load(await extractPagesBytes(bytes, [2, 0]))
    expect(out.getPageCount()).toBe(2)
    expect(out.getPage(0).getWidth()).toBe(300)
    expect(out.getPage(1).getWidth()).toBe(100)
  })

  it('silently drops out-of-range page indices', async () => {
    const bytes = await makePdf([[100, 100]])
    const out = await PDFDocument.load(await extractPagesBytes(bytes, [-1, 0, 5]))
    expect(out.getPageCount()).toBe(1)
  })
})

describe('insertPdfBytes', () => {
  it('inserts all pages at the front when afterPageIndex is -1', async () => {
    const dst = await makePdf([[100, 100]])
    const src = await makePdf([
      [200, 200],
      [300, 300],
    ])
    const { merged, count } = await insertPdfBytes(dst, src, -1)
    expect(count).toBe(2)
    const out = await PDFDocument.load(merged)
    expect(out.getPageCount()).toBe(3)
    expect(out.getPage(0).getWidth()).toBe(200)
    expect(out.getPage(2).getWidth()).toBe(100)
  })

  it('clamps an out-of-range afterPageIndex to the document end', async () => {
    const dst = await makePdf([[100, 100]])
    const src = await makePdf([[200, 200]])
    const { merged } = await insertPdfBytes(dst, src, 99)
    const out = await PDFDocument.load(merged)
    expect(out.getPage(1).getWidth()).toBe(200)
  })
})

describe('insertBlankPageBytes', () => {
  it('inserts a blank page after the requested page with matching geometry', async () => {
    const source = await PDFDocument.create()
    source.addPage([300, 500]).setRotation(degrees(90))
    source.addPage([600, 800])

    const { merged, count } = await insertBlankPageBytes(await source.save(), 0)
    const output = await PDFDocument.load(merged)

    expect(count).toBe(1)
    expect(output.getPageCount()).toBe(3)
    expect(output.getPage(1).getSize()).toEqual({ width: 300, height: 500 })
    expect(output.getPage(1).getRotation().angle).toBe(90)
    expect(output.getPage(2).getSize()).toEqual({ width: 600, height: 800 })
  })

  it('inserts at the front using the first page geometry for index -1', async () => {
    const bytes = await makePdf([[320, 480]])
    const { merged } = await insertBlankPageBytes(bytes, -1)
    const output = await PDFDocument.load(merged)

    expect(output.getPageCount()).toBe(2)
    expect(output.getPage(0).getSize()).toEqual({ width: 320, height: 480 })
  })

  it('inserts multiple landscape pages with a selected paper size', async () => {
    const bytes = await makePdf([[320, 480]])
    const { merged, count } = await insertBlankPageBytes(bytes, 0, {
      count: 3,
      pageSize: 'A5',
      orientation: 'landscape',
    })
    const output = await PDFDocument.load(merged)

    expect(count).toBe(3)
    expect(output.getPageCount()).toBe(4)
    expect(output.getPage(1).getSize()).toEqual({ width: 595.28, height: 419.53 })
    expect(output.getPage(3).getSize()).toEqual({ width: 595.28, height: 419.53 })
  })

  it('rejects an unsafe blank page count', async () => {
    const bytes = await makePdf([[320, 480]])
    await expect(insertBlankPageBytes(bytes, 0, { count: 101 })).rejects.toThrow('Blank page count')
  })
})

describe('savePdfToPath', () => {
  const sha256 = (path: string) => createHash('sha256').update(readFileSync(path)).digest('hex')
  const highlight = {
    pageIndex: 0,
    type: 'highlight' as const,
    color: [1, 0.87, 0.35] as [number, number, number],
    quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
  }

  it('Save As writes the edits to the target only and never mutates the source', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'original.pdf')
    const dst = join(dir, 'copy.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))
    const srcHash = sha256(src)
    const srcInode = statSync(src).ino

    await savePdfToPath(src, dst, request({ path: src, targetPath: dst, markups: [highlight] }))

    // Source: same inode, same bytes
    expect(sha256(src)).toBe(srcHash)
    expect(statSync(src).ino).toBe(srcInode)
    // Target: valid PDF containing the new annotation
    const out = await PDFDocument.load(new Uint8Array(readFileSync(dst)))
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
    // No temp files left behind
    expect(readdirSync(dir).sort()).toEqual(['copy.pdf', 'original.pdf'])
  })

  it('in-place save (target === source) replaces the file atomically', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'doc.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))

    await savePdfToPath(src, src, request({ path: src, markups: [highlight] }))

    const out = await PDFDocument.load(new Uint8Array(readFileSync(src)))
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Highlight'])
    expect(readdirSync(dir)).toEqual(['doc.pdf'])
  })

  it('a failed save leaves the source and target untouched and cleans up temp files', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'gen-pdf-'))
    const src = join(dir, 'original.pdf')
    writeFileSync(src, await makePdf([[612, 792]]))
    const srcHash = sha256(src)

    // Apply failure (unknown form field): nothing may be written anywhere
    await expect(
      savePdfToPath(
        src,
        join(dir, 'copy.pdf'),
        request({ path: src, formValues: [{ name: 'missing', kind: 'text', value: 'x' }] }),
      ),
    ).rejects.toThrow()
    expect(sha256(src)).toBe(srcHash)
    expect(readdirSync(dir)).toEqual(['original.pdf'])

    // Write failure (target directory does not exist): source intact, temp cleaned up
    await expect(
      savePdfToPath(src, join(dir, 'no-such-dir', 'copy.pdf'), request({ path: src })),
    ).rejects.toThrow()
    expect(sha256(src)).toBe(srcHash)
    expect(readdirSync(dir)).toEqual(['original.pdf'])
  })
})

describe('applySaveRequest', () => {
  it('applies page rotation deltas on top of the existing rotation', async () => {
    const bytes = await makePdf([[100, 100]])
    const saved = await apply(bytes, request({ rotations: [{ pageIndex: 0, delta: 90 }] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPage(0).getRotation().angle).toBe(90)
  })

  it('writes markup annotations with an appearance stream', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        markups: [
          {
            pageIndex: 0,
            type: 'highlight',
            color: [1, 0.87, 0.35],
            quads: [[10, 100, 60, 100, 10, 88, 60, 88]],
          },
          {
            pageIndex: 0,
            type: 'underline',
            color: [0.17, 0.4, 1],
            quads: [[10, 80, 60, 80, 10, 68, 60, 68]],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Highlight', 'Underline'])
    // Every markup gets a hand-written /AP /N appearance
    for (const a of annots) {
      expect(a.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
    }
  })

  it('writes note and shape drawing annotations', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        drawings: [
          {
            kind: 'note',
            pageIndex: 0,
            color: [1, 0, 0],
            at: [50, 700],
            contents: 'hello note',
            author: 'AI reviewer',
            subject: 'Review finding',
          },
          {
            kind: 'ink',
            pageIndex: 0,
            color: [0, 0, 1],
            width: 2,
            paths: [[10, 10, 20, 20, 30, 15]],
          },
          { kind: 'rect', pageIndex: 0, color: [0, 1, 0], width: 1, rect: [40, 40, 90, 80] },
          {
            kind: 'arrow',
            pageIndex: 0,
            color: [0, 0, 0],
            width: 2,
            from: [100, 100],
            to: [200, 150],
          },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annotations = pageAnnots(out, 0)
    expect(annotations.map(subtypeOf)).toEqual(['Text', 'Ink', 'Square', 'Line'])
    expect(annotations[0]!.lookup(PDFName.of('T'), PDFHexString).decodeText()).toBe('AI reviewer')
    expect(annotations[0]!.lookup(PDFName.of('Subj'), PDFHexString).decodeText()).toBe(
      'Review finding',
    )
  })

  it('ignores markups and drawings addressing missing pages', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        markups: [
          { pageIndex: 9, type: 'highlight', color: [1, 1, 0], quads: [[0, 1, 1, 1, 0, 0, 1, 0]] },
        ],
        drawings: [{ kind: 'note', pageIndex: 9, color: [1, 0, 0], at: [0, 0], contents: 'x' }],
      }),
    )
    expect(pageAnnots(await PDFDocument.load(saved), 0)).toHaveLength(0)
  })

  it('writes an image drawing as a Stamp annotation with an image appearance', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        drawings: [{ kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [100, 500, 300, 600] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    const annots = pageAnnots(out, 0)
    expect(annots.map(subtypeOf)).toEqual(['Stamp'])
    const rect = annots[0]!.lookup(PDFName.of('Rect'), PDFArray)
    expect(String(rect)).toBe('[ 100 500 300 600 ]')
    expect(annots[0]!.lookup(PDFName.of('AP'), PDFDict).has(PDFName.of('N'))).toBe(true)
  })

  it('counter-rotates the image appearance on rotated pages', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        rotations: [{ pageIndex: 0, delta: 90 }],
        drawings: [{ kind: 'image', pageIndex: 0, image: TINY_PNG, rect: [100, 500, 300, 600] }],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(pageAnnots(out, 0).map(subtypeOf)).toEqual(['Stamp'])
  })

  it('embeds PNG stamps without failing', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        stamps: [{ pageIndex: 0, image: TINY_PNG, rect: [0, 0, 612, 792], opacity: 0.2 }],
      }),
    )
    expect((await PDFDocument.load(saved)).getPageCount()).toBe(1)
  })

  it('embeds deduplicated center-rotated stamps', async () => {
    const bytes = await makePdf([[612, 792]])
    const saved = await apply(
      bytes,
      request({
        stampImages: [TINY_PNG],
        stamps: [
          { pageIndex: 0, image: '', imageIndex: 0, rect: [100, 300, 200, 350], rotation: 35 },
          { pageIndex: 0, image: '', imageIndex: 0, rect: [300, 500, 400, 550], rotation: -35 },
        ],
      }),
    )
    const document = await PDFDocument.load(saved)
    const resources = document.getPage(0).node.Resources()
    const xObjects = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)
    const imageRefs = xObjects?.keys().map((key) => String(xObjects.get(key))) ?? []
    expect(imageRefs).toHaveLength(2)
    expect(new Set(imageRefs).size).toBe(1)
    expect(document.getPageCount()).toBe(1)
  })

  it('applies metadata and splits keywords on mixed separators', async () => {
    const bytes = await makePdf([[100, 100]])
    const saved = await apply(
      bytes,
      request({ metadata: { title: 'My Title', author: 'Me', keywords: 'a, b；c，d' } }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getTitle()).toBe('My Title')
    expect(out.getAuthor()).toBe('Me')
    expect(out.getKeywords()).toContain('a')
    expect(out.getKeywords()).toContain('d')
  })

  it('persists document classification without replacing standard metadata', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 100])
    document.setTitle('Existing title')
    const saved = await apply(
      await document.save({ useObjectStreams: false }),
      request({
        classification: {
          labels: [{ id: 'resume', name: 'Resume' }],
          sensitivity: 'restricted',
        },
      }),
    )

    expect((await PDFDocument.load(saved)).getTitle()).toBe('Existing title')
    expect(await readPdfClassificationMetadataBytes(saved)).toEqual({
      labels: [{ id: 'resume', name: 'Resume' }],
      sensitivity: 'restricted',
    })
  })

  it('deletes pages by original index but never removes the last page', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(bytes, request({ deletedPages: [0, 2] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(1)
    expect(out.getPage(0).getWidth()).toBe(200)

    const savedAll = await apply(bytes, request({ deletedPages: [0, 1, 2] }))
    expect((await PDFDocument.load(savedAll)).getPageCount()).toBe(1)
  })

  it('reorders pages by original index', async () => {
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(bytes, request({ pageOrder: [2, 0, 1] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([300, 100, 200])
  })

  it('applies the reorder when combined with deletions', async () => {
    // Regression: pdf-lib's removePage does not invalidate its page cache,
    // so a getPages() call after the deletion loop returns the stale
    // pre-deletion list. applySaveRequest must derive the remaining pages
    // from the pre-deletion snapshot instead of re-reading them.
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
      [300, 300],
    ])
    const saved = await apply(bytes, request({ deletedPages: [1], pageOrder: [2, 0] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(2)
    expect(out.getPages().map((p) => p.getWidth())).toEqual([300, 100])
  })

  it('skips the reorder when deleting every page leaves the guarded last page', async () => {
    // Deleting all pages keeps one via the last-page guard; the order list
    // then matches nothing alive, so the reorder must be skipped safely.
    const bytes = await makePdf([
      [100, 100],
      [200, 200],
    ])
    const saved = await apply(bytes, request({ deletedPages: [0, 1], pageOrder: [1, 0] }))
    const out = await PDFDocument.load(saved)
    expect(out.getPageCount()).toBe(1)
  })

  it('fills text fields and checkboxes', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    const form = doc.getForm()
    form.createTextField('user.name').addToPage(page, { x: 20, y: 200, width: 200, height: 20 })
    form.createCheckBox('user.agree').addToPage(page, { x: 20, y: 150, width: 16, height: 16 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({
        formValues: [
          { name: 'user.name', kind: 'text', value: 'Alice' },
          { name: 'user.agree', kind: 'checkbox', checked: true },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getTextField('user.name').getText()).toBe('Alice')
    expect(out.getForm().getCheckBox('user.agree').isChecked()).toBe(true)
  })

  it('preserves multi-select choices and editable dropdown values', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    const form = doc.getForm()
    const topics = form.createOptionList('topics')
    topics.addOptions(['Design', 'Engineering', 'Research'])
    topics.enableMultiselect()
    topics.addToPage(page, { x: 20, y: 180, width: 160, height: 60 })
    const department = form.createDropdown('department')
    department.addOptions(['Sales', 'Support'])
    department.enableEditing()
    department.addToPage(page, { x: 20, y: 130, width: 160, height: 24 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({
        formValues: [
          { name: 'topics', kind: 'choice', value: ['Design', 'Research'] },
          { name: 'department', kind: 'choice', value: 'Customer Success' },
        ],
      }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getOptionList('topics').getSelected()).toEqual(['Design', 'Research'])
    expect(out.getForm().getOptionList('topics').isMultiselect()).toBe(true)
    expect(out.getForm().getDropdown('department').getSelected()).toEqual(['Customer Success'])
    expect(out.getForm().getDropdown('department').isEditable()).toBe(true)
  })

  it('removes completed empty signature fields without touching signed fields', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    doc.getForm().createTextField('customer').addToPage(page, {
      x: 20,
      y: 220,
      width: 160,
      height: 24,
    })
    addSignatureField(doc, 'approval', 140, false, false)
    addSignatureField(doc, 'existing_digital_signature', 70, true)

    const saved = await apply(
      await doc.save({ useObjectStreams: false, updateFieldAppearances: false }),
      request({ removeSignatureFields: ['approval', 'existing_digital_signature'] }),
    )
    const output = await PDFDocument.load(saved)

    expect(
      output
        .getForm()
        .getFields()
        .map((field) => field.getName()),
    ).toEqual(['customer', 'existing_digital_signature'])
  })

  it('falls back to NeedAppearances when form values cannot be WinAnsi-encoded', async () => {
    const doc = await PDFDocument.create()
    const page = doc.addPage([300, 300])
    doc.getForm().createTextField('cjk').addToPage(page, { x: 20, y: 200, width: 200, height: 20 })
    const bytes = await doc.save({ useObjectStreams: false })

    const saved = await apply(
      bytes,
      request({ formValues: [{ name: 'cjk', kind: 'text', value: '中文测试' }] }),
    )
    const out = await PDFDocument.load(saved)
    expect(out.getForm().getTextField('cjk').getText()).toBe('中文测试')
    const needAppearances = out.getForm().acroForm.dict.get(PDFName.of('NeedAppearances'))
    expect(String(needAppearances)).toBe('true')
  })
})
