import { PDFDict, PDFDocument, PDFHexString, PDFName } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import {
  PDF_CLASSIFICATION_METADATA_KEY,
  applyPdfClassificationMetadata,
  readPdfClassificationMetadata,
  readPdfClassificationMetadataBytes,
  updatePdfMetadataBytes,
} from '../src/index'

describe('PDF classification metadata', () => {
  it('round-trips labels without replacing standard or unrelated custom metadata', async () => {
    const document = await PDFDocument.create()
    document.addPage([300, 500])
    document.setTitle('Existing title')
    const info = document.context.lookup(document.context.trailerInfo.Info, PDFDict)
    info.set(PDFName.of('Department'), PDFHexString.fromText('Engineering'))

    applyPdfClassificationMetadata(document, {
      labels: [
        { id: 'resume', name: 'Resume' },
        { id: 'contract', name: 'Contract' },
      ],
      sensitivity: 'confidential',
    })
    const bytes = await document.save({ useObjectStreams: false })
    const loaded = await PDFDocument.load(bytes, { updateMetadata: false })
    const loadedInfo = loaded.context.lookup(loaded.context.trailerInfo.Info, PDFDict)

    expect(loaded.getTitle()).toBe('Existing title')
    expect(loadedInfo.lookup(PDFName.of('Department'), PDFHexString).decodeText()).toBe(
      'Engineering',
    )
    expect(readPdfClassificationMetadata(loaded)).toEqual({
      labels: [
        { id: 'resume', name: 'Resume' },
        { id: 'contract', name: 'Contract' },
      ],
      sensitivity: 'confidential',
    })
    expect(loadedInfo.has(PDFName.of(PDF_CLASSIFICATION_METADATA_KEY))).toBe(true)
  })

  it('rejects invalid or duplicated labels before writing', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 100])

    expect(() =>
      applyPdfClassificationMetadata(document, {
        labels: [{ id: '../resume', name: 'Resume' }],
        sensitivity: 'standard',
      }),
    ).toThrow('invalid id')
    expect(() =>
      applyPdfClassificationMetadata(document, {
        labels: [
          { id: 'resume', name: 'Resume' },
          { id: 'resume', name: 'CV' },
        ],
        sensitivity: 'standard',
      }),
    ).toThrow('duplicated')
  })

  it('ignores malformed or unsupported metadata when reading recent files', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 100])
    const info = document.context.lookup(document.context.trailerInfo.Info, PDFDict)
    info.set(
      PDFName.of(PDF_CLASSIFICATION_METADATA_KEY),
      PDFHexString.fromText(JSON.stringify({ version: 2, labels: [], sensitivity: 'standard' })),
    )

    expect(await readPdfClassificationMetadataBytes(await document.save())).toBeNull()
    expect(await readPdfClassificationMetadataBytes(new Uint8Array([1, 2, 3]))).toBeNull()
  })

  it('keeps owned classification metadata out of custom fields and preserves it on metadata edits', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 100])
    applyPdfClassificationMetadata(document, {
      labels: [{ id: 'invoice', name: 'Invoice' }],
      sensitivity: 'internal',
    })

    const updated = await updatePdfMetadataBytes(await document.save(), {
      deleteAll: false,
      metadata: {
        title: 'Updated',
        author: '',
        subject: '',
        keywords: '',
        creator: '',
        producer: '',
        creationDate: '',
        modificationDate: '',
        trapped: '',
        custom: [],
      },
    })

    expect(await readPdfClassificationMetadataBytes(updated)).toEqual({
      labels: [{ id: 'invoice', name: 'Invoice' }],
      sensitivity: 'internal',
    })
    await expect(
      updatePdfMetadataBytes(updated, {
        deleteAll: false,
        metadata: {
          title: '',
          author: '',
          subject: '',
          keywords: '',
          creator: '',
          producer: '',
          creationDate: '',
          modificationDate: '',
          trapped: '',
          custom: [{ key: 'genofficeclassification', value: '{}' }],
        },
      }),
    ).rejects.toThrow('reserved')
  })
})
