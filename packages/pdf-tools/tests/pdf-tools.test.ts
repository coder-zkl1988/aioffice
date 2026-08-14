import { readFile } from 'node:fs/promises'
import {
  PDFArray,
  PDFBool,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRawStream,
  PDFSignature,
  PDFString,
  PDFWidgetAnnotation,
  decodePDFRawStream,
  degrees,
} from 'pdf-lib'
import JSZip from 'jszip'
import { XMLValidator } from 'fast-xml-parser'
import { describe, expect, it } from 'vitest'
import {
  addPdfPageNumbersBytes,
  addPdfCommentsBytes,
  addPdfAttachmentsBytes,
  analyzePdfAnnotationsBytes,
  analyzePdfBytes,
  analyzePdfFontsBytes,
  autoSplitPdfPageRanges,
  autoSplitPdfZipBytes,
  auditPdfJavaScriptBytes,
  autoRotatePdfPagesBytes,
  bookletPagePairs,
  bookletPdfBytes,
  compressPdfPagesBytes,
  contentFilterOutputPageIndexes,
  comparisonPdfPagesBytes,
  createPdfAutoSplitDividerBytes,
  cbzPdfOutputFileName,
  cropPdfMarginsBytes,
  cropPdfPageBoxesBytes,
  decompressPdfBytes,
  deskewPdfPagesBytes,
  deletePdfAttachmentBytes,
  documentMatchesFilterBytes,
  extractCbzImageEntries,
  extractPagesBytes,
  flattenPdfPagesBytes,
  geometricFilterMatchedPageIndexes,
  extractPdfImagesZipBytes,
  extractPdfAttachmentsZipBytes,
  appendImagesToPdfBytes,
  imagesToPdfBytes,
  insertPdfBytes,
  invertPdfColorsBytes,
  jsonToPdfBytes,
  listPdfAttachmentsBytes,
  listPdfBookmarksBytes,
  listPdfFormFieldsBytes,
  mergePdfBytes,
  nUpPdfBytes,
  overlayPageAssignments,
  overlayAdjustedPdfPagesBytes,
  overlayImagePdfBytes,
  overlayPdfBytes,
  pdfImageOverlayPlacement,
  pdfPageImagesOutput,
  pdfToCbzBytes,
  pdfToJsonOutput,
  pdfToXmlOutput,
  pdfToVideoOutput,
  pdfToEpubBytes,
  pdfToHtmlZipBytes,
  pdfToDocxBytes,
  pdfToOdtBytes,
  pdfToRtfBytes,
  pdfToPptxBytes,
  pdfToPdfaBytes,
  pdfaPreservationReportBytes,
  pdfAutoRenameFileName,
  pdfPageNumberLabels,
  pdfFormFieldsCsvBytes,
  pdfFormFieldsJsonBytes,
  pdfFormFieldsXlsxBytes,
  pdfAutoSplitDividerQrModules,
  pdfTablesCsvOutput,
  pdfTablesXlsxBytes,
  pdfTextPagesMarkdownBytes,
  pdfTextPagesTxtBytes,
  posterPdfBytes,
  preflightPdfBytes,
  processPdfFormBytes,
  redactPdfPagesBytes,
  redactSelectedPdfPagesBytes,
  rearrangePageIndexes,
  rearrangePdfPagesBytes,
  repairPdfBytes,
  replacePdfColorsBytes,
  removePdfAnnotationsBytes,
  removePdfImagesBytes,
  removePdfPagesBytes,
  removePdfSignaturesBytes,
  renamePdfAttachmentBytes,
  rotatePdfPagesBytes,
  runPdfToolBytes,
  scannerEffectPdfPagesBytes,
  scannerImageSplitPdfBytes,
  sanitizePdfBytes,
  scalePdfPagesBytes,
  setPdfBookmarksBytes,
  singlePagePdfBytes,
  splitPdfByChaptersBytes,
  splitPdfByDocumentCountBytes,
  splitPdfByPageCountBytes,
  splitPdfBySizeBytes,
  splitPdfSectionsBytes,
  splitPdfBytes,
  updatePdfMetadataBytes,
  type PdfJsonImportFonts,
} from '../src/index'
import { pdfLibFontkit } from '../src/pdf-lib-fontkit'

const decodeUtf8 = (bytes: Uint8Array): string => new TextDecoder().decode(bytes)

async function pdfJsonImportFonts(): Promise<PdfJsonImportFonts> {
  const fontUrl = (name: string) =>
    new URL(`../../../apps/docs/src/renderer/fonts/${name}`, import.meta.url)
  const [regular, bold, italic, boldItalic, unicode] = await Promise.all([
    readFile(fontUrl('LiberationSans-Regular.ttf')),
    readFile(fontUrl('LiberationSans-Bold.ttf')),
    readFile(fontUrl('LiberationSans-Italic.ttf')),
    readFile(fontUrl('LiberationSans-BoldItalic.ttf')),
    readFile(fontUrl('NotoSansSC-Regular-subset.ttf')),
  ])
  return { regular, bold, italic, boldItalic, unicode }
}

async function pdfWithWidths(widths: number[]): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  for (const width of widths) document.addPage([width, 200])
  return document.save()
}

async function pageWidths(bytes: Uint8Array): Promise<number[]> {
  const document = await PDFDocument.load(bytes)
  return document.getPages().map((page) => page.getWidth())
}

async function pageSizes(bytes: Uint8Array): Promise<Array<readonly [number, number]>> {
  const document = await PDFDocument.load(bytes)
  return document.getPages().map((page) => [page.getWidth(), page.getHeight()] as const)
}

async function pageRotations(bytes: Uint8Array): Promise<number[]> {
  const document = await PDFDocument.load(bytes)
  return document.getPages().map((page) => page.getRotation().angle)
}

function embeddedTrueTypeFontPrograms(document: PDFDocument, resources: PDFDict): Uint8Array[] {
  const fonts = resources.lookupMaybe(PDFName.of('Font'), PDFDict)
  if (!fonts) return []
  const programs: Uint8Array[] = []
  for (const name of fonts.keys()) {
    const font = fonts.lookupMaybe(name, PDFDict)
    if (!font) continue
    const descendants = font.lookupMaybe(PDFName.of('DescendantFonts'), PDFArray)
    const candidates = [font]
    if (descendants) {
      const descendant = descendants.lookupMaybe(0, PDFDict)
      if (descendant) candidates.push(descendant)
    }
    for (const candidate of candidates) {
      const descriptor = candidate.lookupMaybe(PDFName.of('FontDescriptor'), PDFDict)
      if (!descriptor) continue
      const fontFile = document.context.lookup(descriptor.get(PDFName.of('FontFile2')))
      if (fontFile instanceof PDFRawStream) programs.push(decodePDFRawStream(fontFile).decode())
    }
  }
  return programs
}

function trueTypeGlyphCount(program: Uint8Array): number {
  const view = new DataView(program.buffer, program.byteOffset, program.byteLength)
  if (![0x0001_0000, 0x7472_7565].includes(view.getUint32(0))) return 0
  const tableCount = view.getUint16(4)
  for (let index = 0; index < tableCount; index++) {
    const recordOffset = 12 + index * 16
    const tag = String.fromCharCode(
      view.getUint8(recordOffset),
      view.getUint8(recordOffset + 1),
      view.getUint8(recordOffset + 2),
      view.getUint8(recordOffset + 3),
    )
    if (tag === 'maxp') return view.getUint16(view.getUint32(recordOffset + 8) + 4)
  }
  return 0
}

function countImagesInResources(
  document: PDFDocument,
  resources: PDFDict | undefined,
  visited = new Set<PDFDict>(),
): number {
  if (!resources || visited.has(resources)) return 0
  visited.add(resources)
  const xObjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xObjects) return 0
  let count = 0
  for (const name of xObjects.keys()) {
    const stream = document.context.lookup(xObjects.get(name))
    if (!(stream instanceof PDFRawStream)) continue
    const subtype = stream.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
    if (subtype === 'Image') count++
    if (subtype === 'Form') {
      count += countImagesInResources(
        document,
        stream.dict.lookupMaybe(PDFName.of('Resources'), PDFDict),
        visited,
      )
    }
  }
  return count
}

function tinyPngBytes(): Uint8Array {
  return Uint8Array.from(
    atob(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=',
    ),
    (character) => character.charCodeAt(0),
  )
}

describe('analyzePdfFontsBytes', () => {
  it('aggregates standard font usage across pages', async () => {
    const document = await PDFDocument.create()
    const font = await document.embedFont('Helvetica')
    document.addPage([240, 320]).drawText('First page', { font })
    document.addPage([240, 320]).drawText('Second page', { font })

    const report = await analyzePdfFontsBytes(await document.save())
    expect(report).toEqual({
      fontCount: 1,
      embeddedCount: 0,
      subsetCount: 0,
      fonts: [
        {
          name: 'Helvetica',
          subtype: 'Type1',
          embedded: false,
          subset: false,
          encoding: 'WinAnsiEncoding',
          hasToUnicode: false,
          pages: [1, 2],
        },
      ],
    })
  })

  it('finds embedded subset fonts in nested form resources', async () => {
    const document = await PDFDocument.create()
    const fontProgram = document.context.register(
      document.context.flateStream(new Uint8Array([0, 1, 2, 3])),
    )
    const descriptor = document.context.register(
      document.context.obj({
        Type: 'FontDescriptor',
        FontName: 'ABCDEF+ArchiveSans',
        Flags: 4,
        FontBBox: [0, 0, 1000, 1000],
        ItalicAngle: 0,
        Ascent: 800,
        Descent: -200,
        CapHeight: 700,
        StemV: 80,
        FontFile2: fontProgram,
      }),
    )
    const toUnicode = document.context.register(
      document.context.flateStream('/CIDInit /ProcSet findresource begin'),
    )
    const font = document.context.register(
      document.context.obj({
        Type: 'Font',
        Subtype: 'TrueType',
        BaseFont: 'ABCDEF+ArchiveSans',
        Encoding: 'Identity-H',
        FontDescriptor: descriptor,
        ToUnicode: toUnicode,
      }),
    )
    const nestedResources = document.context.register(document.context.obj({ Font: { F1: font } }))
    const form = document.context.register(
      document.context.flateStream('', {
        Type: 'XObject',
        Subtype: 'Form',
        BBox: [0, 0, 10, 10],
        Resources: nestedResources,
      }),
    )
    const page = document.addPage([240, 320])
    page.node.set(PDFName.of('Resources'), document.context.obj({ XObject: { Form1: form } }))

    const report = await analyzePdfFontsBytes(await document.save())
    expect(report).toEqual({
      fontCount: 1,
      embeddedCount: 1,
      subsetCount: 1,
      fonts: [
        {
          name: 'ArchiveSans',
          subtype: 'TrueType',
          embedded: true,
          subset: true,
          encoding: 'Identity-H',
          hasToUnicode: true,
          pages: [1],
        },
      ],
    })
  })

  it('returns an empty report for image-only PDFs', async () => {
    expect(await analyzePdfFontsBytes(await pdfWithWidths([100]))).toEqual({
      fontCount: 0,
      embeddedCount: 0,
      subsetCount: 0,
      fonts: [],
    })
  })
})

describe('analyzePdfAnnotationsBytes', () => {
  it('lists annotation metadata, positions, flags, and type counts', async () => {
    const document = await PDFDocument.create()
    const firstPage = document.addPage([300, 200])
    const secondPage = document.addPage([400, 250])
    const note = document.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [40, 80, 10, 20],
      T: PDFHexString.fromText('Reviewer'),
      Subj: PDFHexString.fromText('Approval'),
      Contents: PDFHexString.fromText('Check this section.'),
      M: PDFString.fromDate(new Date('2026-08-14T01:02:03.000Z')),
      NM: PDFHexString.fromText('note-1'),
      F: 4,
    })
    const highlight = document.context.obj({
      Type: 'Annot',
      Subtype: 'Highlight',
      Rect: [50, 60, 150, 80],
      Contents: PDFHexString.fromText('Important'),
    })
    firstPage.node.set(PDFName.of('Annots'), document.context.obj([note]))
    secondPage.node.set(
      PDFName.of('Annots'),
      document.context.obj([document.context.register(highlight)]),
    )

    expect(await analyzePdfAnnotationsBytes(await document.save())).toEqual({
      totalCount: 2,
      typeBreakdown: { Text: 1, Highlight: 1 },
      annotations: [
        {
          pageNumber: 1,
          annotationNumber: 1,
          subtype: 'Text',
          author: 'Reviewer',
          subject: 'Approval',
          contents: 'Check this section.',
          modifiedAt: '2026-08-14T01:02:03.000Z',
          name: 'note-1',
          flags: 4,
          rectangle: { x: 10, y: 20, width: 30, height: 60 },
        },
        {
          pageNumber: 2,
          annotationNumber: 1,
          subtype: 'Highlight',
          contents: 'Important',
          rectangle: { x: 50, y: 60, width: 100, height: 20 },
        },
      ],
    })
  })

  it('returns an empty annotation report', async () => {
    expect(await analyzePdfAnnotationsBytes(await pdfWithWidths([100]))).toEqual({
      totalCount: 0,
      typeBreakdown: {},
      annotations: [],
    })
  })
})

describe('extractPagesBytes', () => {
  it('copies valid pages in the requested order', async () => {
    const result = await extractPagesBytes(await pdfWithWidths([100, 200, 300]), [2, -1, 0, 5])
    expect(await pageWidths(result)).toEqual([300, 100])
  })

  it('keeps source metadata and repeated pages', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Selected pages')
    source.addPage([100, 100])
    source.addPage([200, 100])
    source.addPage([300, 100])
    const result = await extractPagesBytes(await source.save(), [2, 0, 2])
    const extracted = await PDFDocument.load(result)
    expect(extracted.getTitle()).toBe('Selected pages')
    expect(extracted.getPages().map((page) => page.getWidth())).toEqual([300, 100, 300])
  })
})

describe('PDF text export', () => {
  const pages = [
    {
      pageNumber: 1,
      text: 'Product launch\r\n\r\nLocal generation  ',
      blocks: [
        { kind: 'heading' as const, level: 1 as const, text: 'Product [launch]' },
        { kind: 'paragraph' as const, text: 'Use *local* generation.' },
        { kind: 'listItem' as const, text: '• Editable slides' },
      ],
      links: [
        { url: 'https://example.com/docs', label: 'Product docs' },
        { url: 'https://example.com/docs', label: 'Duplicate' },
      ],
    },
    {
      pageNumber: 3,
      text: '第三页',
      blocks: [{ kind: 'paragraph' as const, text: '第三页' }],
      links: [],
    },
  ]

  it('serializes plain text with stable page breaks', () => {
    expect(decodeUtf8(pdfTextPagesTxtBytes(pages))).toBe(
      'Product launch\n\nLocal generation\n\n\f\n\n第三页\n',
    )
  })

  it('serializes page structure, escapes content, and deduplicates links', () => {
    const markdown = decodeUtf8(pdfTextPagesMarkdownBytes(pages))
    expect(markdown).toContain('## Page 1')
    expect(markdown).toContain('### Product \\[launch\\]')
    expect(markdown).toContain('Use \\*local\\* generation.')
    expect(markdown).toContain('- Editable slides')
    expect(markdown.match(/https:\/\/example\.com\/docs/g)).toHaveLength(1)
    expect(markdown).toContain('## Page 3')
  })

  it('returns TXT and Markdown outputs with platform metadata', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'extractText',
      format: 'both',
      pageIndexes: [0, 2],
      pages,
    })
    expect(
      outputs.map(({ suffix, mimeType, extension }) => ({ suffix, mimeType, extension })),
    ).toEqual([
      {
        suffix: '_text.txt',
        mimeType: 'text/plain;charset=utf-8',
        extension: '.txt',
      },
      {
        suffix: '_text.md',
        mimeType: 'text/markdown;charset=utf-8',
        extension: '.md',
      },
    ])
  })

  it('rejects missing, mismatched, empty, and out-of-range prepared data', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      runPdfToolBytes(source, { kind: 'extractText', format: 'txt', pageIndexes: [0] }),
    ).rejects.toThrow('Extracted PDF text pages are required')
    await expect(
      runPdfToolBytes(source, {
        kind: 'extractText',
        format: 'txt',
        pageIndexes: [0],
        pages: [{ pageNumber: 2, text: 'wrong page', blocks: [], links: [] }],
      }),
    ).rejects.toThrow('do not match')
    await expect(
      runPdfToolBytes(source, {
        kind: 'extractText',
        format: 'txt',
        pageIndexes: [0],
        pages: [{ pageNumber: 1, text: '', blocks: [], links: [] }],
      }),
    ).rejects.toThrow('require OCR')
    await expect(
      runPdfToolBytes(source, {
        kind: 'extractText',
        format: 'txt',
        pageIndexes: [1],
        pages: [{ pageNumber: 2, text: 'out of range', blocks: [], links: [] }],
      }),
    ).rejects.toThrow('invalid page')
  })
})

describe('PDF structured JSON export', () => {
  const page = {
    pageNumber: 1,
    width: 612,
    height: 792,
    rotation: 0,
    text: 'Product launch',
    blocks: [{ kind: 'heading' as const, text: 'Product launch', level: 1 as const }],
    links: [{ url: 'https://example.com', label: 'Example' }],
    textRuns: [
      {
        text: 'Product launch',
        x: 48,
        y: 720,
        width: 130,
        height: 24,
        fontSize: 24,
        fontFamily: 'Inter',
        bold: true,
        italic: false,
      },
    ],
  }

  it('exports a versioned layout document with source metadata', async () => {
    const output = await pdfToJsonOutput(await pdfWithWidths([612]), {
      pageIndexes: [0],
      lightweight: false,
      baseName: '../Launch:Plan.pdf',
      pages: [page],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_structured.json',
        fileName: '_Launch_Plan_structured.json',
        mimeType: 'application/json;charset=utf-8',
        extension: '.json',
      }),
    )
    const document = JSON.parse(decodeUtf8(output.bytes))
    expect(document).toMatchObject({
      schema: 'genoffice.pdf.json',
      version: 1,
      mode: 'layout',
      source: { fileName: '_Launch_Plan', pageCount: 1, selectedPages: [1] },
      pages: [{ pageNumber: 1, text: 'Product launch', textRuns: page.textRuns }],
      bookmarks: [],
      formFields: [],
    })
  })

  it('omits layout runs in lightweight mode and validates prepared pages', async () => {
    const source = await pdfWithWidths([612])
    const [output] = await runPdfToolBytes(source, {
      kind: 'pdfToJson',
      pageIndexes: [0],
      lightweight: true,
      pages: [page],
    })
    const document = JSON.parse(decodeUtf8(output!.bytes))
    expect(document.mode).toBe('semantic')
    expect(document.pages[0]).not.toHaveProperty('textRuns')
    await expect(
      pdfToJsonOutput(source, { pageIndexes: [0], lightweight: false, pages: undefined }),
    ).rejects.toThrow('Structured PDF pages are required')
    await expect(
      pdfToJsonOutput(source, {
        pageIndexes: [0],
        lightweight: false,
        pages: [{ ...page, pageNumber: 2 }],
      }),
    ).rejects.toThrow('order is invalid')
    await expect(
      pdfToJsonOutput(source, {
        pageIndexes: [0],
        lightweight: false,
        pages: [{ ...page, textRuns: undefined }],
      }),
    ).rejects.toThrow('requires text layout data')
  })
})

describe('PDF structured JSON import', () => {
  it('restores layout text, Unicode fonts, metadata, and remapped bookmarks', async () => {
    const structured = {
      schema: 'genoffice.pdf.json',
      version: 1,
      mode: 'layout',
      source: { fileName: '发布计划' },
      metadata: {
        title: '产品发布计划',
        author: 'GenOffice',
        subject: '本地结构化还原',
        keywords: 'PDF, JSON',
      },
      pages: [
        {
          pageNumber: 3,
          width: 612,
          height: 792,
          rotation: 0,
          text: 'Product launch 产品发布',
          blocks: [{ kind: 'heading', text: 'Product launch 产品发布', level: 1 }],
          links: [],
          textRuns: [
            {
              text: '产品发布',
              x: 48,
              y: 680,
              width: 96,
              height: 20,
              fontSize: 20,
              fontFamily: 'Noto Sans CJK SC',
              bold: false,
              italic: false,
            },
          ],
        },
      ],
      bookmarks: [{ title: '发布', pageNumber: 3, children: [] }],
      formFields: [],
    }
    const output = await jsonToPdfBytes({
      jsonBytes: new TextEncoder().encode(JSON.stringify(structured)),
      fonts: await pdfJsonImportFonts(),
      baseName: 'Launch:Plan.json',
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_restored.pdf',
        fileName: 'Launch_Plan_restored.pdf',
        mimeType: 'application/pdf',
        extension: '.pdf',
      }),
    )
    const restored = await PDFDocument.load(output.bytes)
    expect(output.bytes.length).toBeLessThan(200_000)
    expect(restored.getPageCount()).toBe(1)
    expect(restored.getPage(0).getSize()).toEqual({ width: 612, height: 792 })
    expect(restored.getTitle()).toBe('产品发布计划')
    expect(restored.getAuthor()).toBe('GenOffice')
    expect(await listPdfBookmarksBytes(output.bytes)).toEqual([
      { title: '发布', pageNumber: 1, children: [] },
    ])
    const programs = embeddedTrueTypeFontPrograms(restored, restored.getPage(0).node.Resources()!)
    expect(programs).toHaveLength(1)
    expect(trueTypeGlyphCount(programs[0]!)).toBeGreaterThanOrEqual(5)
  })

  it('round-trips semantic export through the shared adapter with local reflow', async () => {
    const exported = await pdfToJsonOutput(await pdfWithWidths([320]), {
      pageIndexes: [0],
      lightweight: true,
      baseName: 'Notes.pdf',
      pages: [
        {
          pageNumber: 1,
          width: 320,
          height: 200,
          rotation: 0,
          text: 'Local semantic content',
          blocks: [
            { kind: 'heading', text: 'Local notes', level: 2 },
            { kind: 'paragraph', text: 'Local semantic content remains\neditable.' },
          ],
          links: [],
        },
      ],
    })
    const [restored] = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'jsonToPdf',
      jsonBytes: exported.bytes,
      baseName: 'Notes_structured.json',
    })
    expect(restored!.suffix).toBe('_restored.pdf')
    expect(await pageSizes(restored!.bytes)).toEqual([[320, 200]])
  })

  it('rejects malformed, unsupported, incomplete, and oversized JSON', async () => {
    await expect(
      jsonToPdfBytes({ jsonBytes: new TextEncoder().encode('{not-json') }),
    ).rejects.toThrow('malformed')
    await expect(
      jsonToPdfBytes({
        jsonBytes: new TextEncoder().encode(
          JSON.stringify({ schema: 'genoffice.pdf.json', version: 2, pages: [] }),
        ),
      }),
    ).rejects.toThrow('version 1')
    await expect(
      jsonToPdfBytes({
        jsonBytes: new TextEncoder().encode(
          JSON.stringify({
            schema: 'genoffice.pdf.json',
            version: 1,
            mode: 'layout',
            pages: [
              {
                pageNumber: 1,
                width: 612,
                height: 792,
                rotation: 0,
                text: '',
                blocks: [],
                links: [],
              },
            ],
            bookmarks: [],
          }),
        ),
      }),
    ).rejects.toThrow('requires text runs')
    await expect(
      jsonToPdfBytes({ jsonBytes: new Uint8Array(25 * 1024 * 1024 + 1) }),
    ).rejects.toThrow('too large')
  })
})

describe('PDF structured XML export', () => {
  const page = {
    pageNumber: 1,
    width: 612,
    height: 792,
    rotation: 0,
    text: 'Launch <Plan> & roadmap',
    blocks: [{ kind: 'heading' as const, level: 1 as const, text: 'Launch <Plan>' }],
    links: [{ url: 'https://example.com/?a=1&b=2', label: 'Details "now"' }],
    textRuns: [
      {
        text: 'Launch & roadmap',
        x: 48,
        y: 720,
        width: 130,
        height: 24,
        fontSize: 24,
        fontFamily: 'Inter',
        bold: true,
        italic: false,
      },
    ],
  }

  it('exports escaped layout XML through the shared adapter', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612]), {
      kind: 'pdfToXml',
      pageIndexes: [0],
      lightweight: false,
      baseName: '../Launch:Plan.pdf',
      pages: [page],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_structured.xml',
        fileName: '_Launch_Plan_structured.xml',
        mimeType: 'application/xml;charset=utf-8',
        extension: '.xml',
      }),
    )
    const xml = decodeUtf8(output!.bytes)
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8"?>')
    expect(xml).toContain('schema="genoffice.pdf.xml" version="1" mode="layout"')
    expect(xml).toContain('<text>Launch &lt;Plan&gt; &amp; roadmap</text>')
    expect(xml).toContain('url="https://example.com/?a=1&amp;b=2"')
    expect(xml).toContain('label="Details &quot;now&quot;"')
    expect(xml).toContain('<textRuns>')
    expect(xml).toContain('fontFamily="Inter" bold="true" italic="false"')
  })

  it('omits layout runs in semantic mode and validates prepared pages', async () => {
    const source = await pdfWithWidths([612])
    const output = await pdfToXmlOutput(source, {
      pageIndexes: [0],
      lightweight: true,
      pages: [page],
    })
    const xml = decodeUtf8(output.bytes)
    expect(xml).toContain('mode="semantic"')
    expect(xml).not.toContain('<textRuns>')
    await expect(
      pdfToXmlOutput(source, {
        pageIndexes: [0],
        lightweight: false,
        pages: undefined,
      }),
    ).rejects.toThrow('Structured PDF pages are required')
  })
})

describe('PDF to video output', () => {
  const videoBytes = new Uint8Array([
    0x1a, 0x45, 0xdf, 0xa3, 0x9f, 0x42, 0x86, 0x81, 0x01, 0x42, 0xf7, 0x81, 0x01, 0x42, 0xf2, 0x81,
  ])

  it('returns a named WebM through the shared adapter', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612, 612]), {
      kind: 'pdfToVideo',
      pageIndexes: [0, 1],
      secondsPerPage: 2,
      resolution: '720p',
      transitionSeconds: 0.4,
      includeAnnotations: true,
      baseName: '../Launch:Plan.pdf',
      videoBytes,
    })
    expect(output).toEqual({
      suffix: '_slideshow.webm',
      fileName: '_Launch_Plan_slideshow.webm',
      bytes: videoBytes,
      mimeType: 'video/webm',
      extension: '.webm',
    })
  })

  it('rejects missing, malformed, and oversized video output', async () => {
    expect(() =>
      pdfToVideoOutput({
        pageIndexes: [0],
        secondsPerPage: 2,
        resolution: '720p',
        transitionSeconds: 0.4,
        includeAnnotations: true,
        videoBytes: new Uint8Array(16),
      }),
    ).toThrow('valid WebM')
    await expect(
      runPdfToolBytes(await pdfWithWidths([612]), {
        kind: 'pdfToVideo',
        pageIndexes: [0],
        secondsPerPage: 11,
        resolution: '720p',
        transitionSeconds: 0.4,
        includeAnnotations: true,
        videoBytes,
      }),
    ).rejects.toThrow('1 to 10')
    await expect(
      runPdfToolBytes(await pdfWithWidths([612]), {
        kind: 'pdfToVideo',
        pageIndexes: [0],
        secondsPerPage: 2,
        resolution: '720p',
        transitionSeconds: 1.1,
        includeAnnotations: true,
        videoBytes,
      }),
    ).rejects.toThrow('0 to 1')
  })
})

describe('PDF to EPUB', () => {
  const semanticPage = {
    pageNumber: 1,
    width: 612,
    height: 792,
    text: 'Launch <Plan>\n\nPrivate generation',
    blocks: [
      { kind: 'heading' as const, text: 'Launch <Plan>', level: 1 as const },
      { kind: 'paragraph' as const, text: 'Private generation' },
      { kind: 'listItem' as const, text: '• Editable pages' },
    ],
    links: [
      { url: 'https://example.com/docs?a=1&b=2', label: 'Docs <home>' },
      { url: 'javascript:alert(1)', label: 'Unsafe' },
    ],
  }

  it('packages a safe reflowable EPUB 3 document', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612]), {
      kind: 'pdfToEpub',
      pageCount: 1,
      pageIndexes: [0],
      mode: 'reflowable',
      renderDpi: 150,
      includeAnnotations: true,
      baseName: '../Launch Plan.pdf',
      pages: [semanticPage],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.epub',
        fileName: '_Launch Plan_converted.epub',
        mimeType: 'application/epub+zip',
        extension: '.epub',
      }),
    )
    expect(decodeUtf8(output!.bytes.slice(0, 4))).toBe('PK\u0003\u0004')
    const archive = await JSZip.loadAsync(output!.bytes)
    expect(Object.keys(archive.files)[0]).toBe('mimetype')
    expect(await archive.file('mimetype')!.async('string')).toBe('application/epub+zip')
    const chapter = await archive.file('EPUB/text/page-1.xhtml')!.async('string')
    expect(chapter).toContain('<h1>Launch &lt;Plan&gt;</h1>')
    expect(chapter).toContain('<li>Editable pages</li>')
    expect(chapter).toContain('https://example.com/docs?a=1&amp;b=2')
    expect(chapter).not.toContain('javascript:')
    const packageDocument = await archive.file('EPUB/package.opf')!.async('string')
    expect(packageDocument).toContain(
      '<dc:identifier id="book-id">urn:genoffice:_Launch%20Plan</dc:identifier>',
    )
    expect(packageDocument).toContain('<meta property="rendition:layout">reflowable</meta>')
    expect(packageDocument).toContain('<dc:language>en</dc:language>')
  })

  it('embeds fixed-layout page images and validates prepared pages', async () => {
    const source = await pdfWithWidths([612])
    const output = await pdfToEpubBytes(source, {
      pageCount: 1,
      pageIndexes: [0],
      mode: 'fixed',
      renderDpi: 144,
      includeAnnotations: false,
      pages: [{ ...semanticPage, imageBytes: new Uint8Array([137, 80, 78, 71]) }],
    })
    const archive = await JSZip.loadAsync(output.bytes)
    expect(Array.from(await archive.file('EPUB/images/page-1.png')!.async('uint8array'))).toEqual([
      137, 80, 78, 71,
    ])
    expect(await archive.file('EPUB/package.opf')!.async('string')).toContain(
      '<meta property="rendition:layout">pre-paginated</meta>',
    )
    expect(await archive.file('EPUB/text/page-1.xhtml')!.async('string')).toContain(
      'content="width=612,height=792"',
    )
    await expect(
      pdfToEpubBytes(source, {
        pageCount: 1,
        pageIndexes: [0],
        mode: 'fixed',
        renderDpi: 144,
        includeAnnotations: false,
        pages: [semanticPage],
      }),
    ).rejects.toThrow('image is empty')
  })

  it('directs textless documents to fixed layout or OCR', async () => {
    await expect(
      pdfToEpubBytes(await pdfWithWidths([612]), {
        pageCount: 1,
        pageIndexes: [0],
        mode: 'reflowable',
        renderDpi: 150,
        includeAnnotations: true,
        pages: [{ ...semanticPage, text: '', blocks: [], links: [] }],
      }),
    ).rejects.toThrow('use fixed layout or run OCR first')
  })
})

describe('PDF to PowerPoint', () => {
  const page = {
    pageNumber: 1,
    width: 612,
    height: 792,
    textRuns: [
      {
        text: 'Editable launch',
        x: 48,
        y: 60,
        width: 180,
        height: 28,
        fontSize: 24,
        angle: 0,
        fontFamily: 'Arial',
        bold: true,
        italic: false,
      },
    ],
  }

  it('creates editable PowerPoint text boxes without page images', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612]), {
      kind: 'pdfToPptx',
      pageCount: 1,
      pageIndexes: [0],
      mode: 'editableText',
      renderDpi: 150,
      includeAnnotations: true,
      baseName: '../Launch:Deck.pdf',
      pages: [page],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.pptx',
        fileName: '_Launch_Deck_converted.pptx',
        mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        extension: '.pptx',
      }),
    )
    const archive = await JSZip.loadAsync(output!.bytes)
    const slide = await archive.file('ppt/slides/slide1.xml')!.async('string')
    expect(slide).toContain('Editable launch')
    expect(slide).toContain('PDF Text 1.1')
    expect(
      Object.entries(archive.files).some(
        ([name, entry]) => /^ppt\/media\//.test(name) && !entry.dir,
      ),
    ).toBe(false)
  })

  it('creates page-fidelity slides backed by PNG media', async () => {
    const output = await pdfToPptxBytes({
      pageCount: 1,
      pageIndexes: [0],
      mode: 'fidelity',
      renderDpi: 144,
      includeAnnotations: false,
      pages: [{ ...page, imageBytes: new Uint8Array([137, 80, 78, 71]) }],
    })
    const archive = await JSZip.loadAsync(output.bytes)
    const media = Object.keys(archive.files).find((name) => /^ppt\/media\/.*\.png$/.test(name))
    expect(media).toBeDefined()
    expect(Array.from(await archive.file(media!)!.async('uint8array'))).toEqual([137, 80, 78, 71])
    expect(await archive.file('ppt/slides/slide1.xml')!.async('string')).toContain('PDF Page 1')
  })

  it('requires text or a fidelity image for the selected mode', async () => {
    const valid = {
      pageCount: 1,
      pageIndexes: [0],
      renderDpi: 150,
      includeAnnotations: true,
    }
    await expect(
      pdfToPptxBytes({
        ...valid,
        mode: 'editableText',
        pages: [{ ...page, textRuns: [] }],
      }),
    ).rejects.toThrow('use page fidelity or run OCR first')
    await expect(pdfToPptxBytes({ ...valid, mode: 'fidelity', pages: [page] })).rejects.toThrow(
      'image is empty',
    )
  })
})

describe('PDF to Word', () => {
  const firstPage = {
    pageNumber: 1,
    width: 612,
    height: 792,
    textRuns: [
      {
        text: 'Editable launch',
        x: 48,
        y: 60,
        width: 180,
        height: 28,
        fontSize: 24,
        angle: 0,
        fontFamily: 'Arial',
        bold: true,
        italic: false,
      },
    ],
  }
  const secondPage = {
    ...firstPage,
    pageNumber: 2,
    textRuns: [{ ...firstPage.textRuns[0]!, text: 'Second page', bold: false }],
  }

  it('creates editable Word paragraphs with styles and page breaks', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612, 612]), {
      kind: 'pdfToDocx',
      pageCount: 2,
      pageIndexes: [0, 1],
      mode: 'editableText',
      renderDpi: 150,
      includeAnnotations: true,
      baseName: '../Launch:Brief.pdf',
      pages: [firstPage, secondPage],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.docx',
        fileName: '_Launch_Brief_converted.docx',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        extension: '.docx',
      }),
    )
    const archive = await JSZip.loadAsync(output!.bytes)
    const documentXml = await archive.file('word/document.xml')!.async('string')
    expect(documentXml).toContain('Editable launch')
    expect(documentXml).toContain('Second page')
    expect(documentXml).toContain('<w:b')
    expect(documentXml).toContain('<w:pageBreakBefore')
    expect(Object.keys(archive.files).some((name) => /^word\/media\//.test(name))).toBe(false)
  })

  it('creates page-fidelity Word documents backed by PNG media', async () => {
    const imageBytes = new Uint8Array([137, 80, 78, 71])
    const output = await pdfToDocxBytes({
      pageCount: 2,
      pageIndexes: [0, 1],
      mode: 'fidelity',
      renderDpi: 144,
      includeAnnotations: false,
      pages: [
        { ...firstPage, imageBytes, imageWidth: 1224, imageHeight: 1584 },
        { ...secondPage, imageBytes, imageWidth: 1224, imageHeight: 1584 },
      ],
    })
    const archive = await JSZip.loadAsync(output.bytes)
    const media = Object.keys(archive.files).filter((name) => /^word\/media\/.*\.png$/.test(name))
    expect(media).toHaveLength(2)
    expect(Array.from(await archive.file(media[0]!)!.async('uint8array'))).toEqual(
      Array.from(imageBytes),
    )
    expect(await archive.file('word/document.xml')!.async('string')).toContain('<w:pageBreakBefore')
  })

  it('requires text or a fidelity image for the selected mode', async () => {
    const valid = {
      pageCount: 1,
      pageIndexes: [0],
      renderDpi: 150,
      includeAnnotations: true,
    }
    await expect(
      pdfToDocxBytes({
        ...valid,
        mode: 'editableText',
        pages: [{ ...firstPage, textRuns: [] }],
      }),
    ).rejects.toThrow('use page fidelity or run OCR first')
    await expect(
      pdfToDocxBytes({ ...valid, mode: 'fidelity', pages: [firstPage] }),
    ).rejects.toThrow('image is empty')
  })

  it('keeps rotated PDF text editable instead of producing an empty document', async () => {
    const output = await pdfToDocxBytes({
      pageCount: 1,
      pageIndexes: [0],
      mode: 'editableText',
      renderDpi: 150,
      includeAnnotations: true,
      pages: [
        {
          ...firstPage,
          textRuns: [{ ...firstPage.textRuns[0]!, text: 'Vertical label', angle: 90 }],
        },
      ],
    })
    const archive = await JSZip.loadAsync(output.bytes)
    expect(await archive.file('word/document.xml')!.async('string')).toContain('Vertical label')
  })
})

describe('PDF to OpenDocument text', () => {
  const firstPage = {
    pageNumber: 1,
    width: 612,
    height: 792,
    textRuns: [
      {
        text: 'Editable <launch> & plan',
        x: 48,
        y: 60,
        width: 220,
        height: 28,
        fontSize: 24,
        angle: 0,
        fontFamily: 'GenOffice:Sans & Serif',
        bold: true,
        italic: true,
      },
    ],
  }
  const secondPage = {
    ...firstPage,
    pageNumber: 2,
    textRuns: [{ ...firstPage.textRuns[0]!, text: 'Second page', bold: false, italic: false }],
  }

  it('creates a standards-based editable ODT with text styles and page breaks', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612, 612]), {
      kind: 'pdfToOdt',
      pageCount: 2,
      pageIndexes: [0, 1],
      mode: 'editableText',
      renderDpi: 150,
      includeAnnotations: true,
      baseName: '../Launch:Brief.pdf',
      pages: [firstPage, secondPage],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.odt',
        fileName: '_Launch_Brief_converted.odt',
        mimeType: 'application/vnd.oasis.opendocument.text',
        extension: '.odt',
      }),
    )
    const header = new DataView(
      output!.bytes.buffer,
      output!.bytes.byteOffset,
      output!.bytes.byteLength,
    )
    const firstNameLength = header.getUint16(26, true)
    expect(new TextDecoder().decode(output!.bytes.slice(30, 30 + firstNameLength))).toBe('mimetype')
    expect(header.getUint16(8, true)).toBe(0)

    const archive = await JSZip.loadAsync(output!.bytes)
    expect(await archive.file('mimetype')!.async('string')).toBe(
      'application/vnd.oasis.opendocument.text',
    )
    const content = await archive.file('content.xml')!.async('string')
    const styles = await archive.file('styles.xml')!.async('string')
    const manifest = await archive.file('META-INF/manifest.xml')!.async('string')
    expect(XMLValidator.validate(content)).toBe(true)
    expect(XMLValidator.validate(styles)).toBe(true)
    expect(XMLValidator.validate(manifest)).toBe(true)
    expect(content).toContain('Editable')
    expect(content).toContain('&lt;launch&gt;')
    expect(content).toContain('&amp;')
    expect(content).toContain('Second<text:s/>page')
    expect(content).toContain('fo:font-weight="bold"')
    expect(content).toContain('fo:font-style="italic"')
    expect(content).toContain('fo:break-before="page"')
    expect(Object.keys(archive.files).some((name) => /^Pictures\//.test(name))).toBe(false)
  })

  it('creates page-fidelity ODT documents backed by PNG media', async () => {
    const imageBytes = new Uint8Array([137, 80, 78, 71])
    const output = await pdfToOdtBytes({
      pageCount: 2,
      pageIndexes: [0, 1],
      mode: 'fidelity',
      renderDpi: 144,
      includeAnnotations: false,
      pages: [
        { ...firstPage, imageBytes, imageWidth: 1224, imageHeight: 1584 },
        { ...secondPage, imageBytes, imageWidth: 1224, imageHeight: 1584 },
      ],
    })
    const archive = await JSZip.loadAsync(output.bytes)
    const images = Object.keys(archive.files).filter((name) => /^Pictures\/.*\.png$/.test(name))
    expect(images).toHaveLength(2)
    expect(Array.from(await archive.file(images[0]!)!.async('uint8array'))).toEqual(
      Array.from(imageBytes),
    )
    expect(await archive.file('content.xml')!.async('string')).toContain(
      'xlink:href="Pictures/page-1.png"',
    )
    expect(await archive.file('META-INF/manifest.xml')!.async('string')).toContain(
      'manifest:full-path="Pictures/page-2.png"',
    )
  })

  it('requires text or a fidelity image for the selected mode', async () => {
    const valid = {
      pageCount: 1,
      pageIndexes: [0],
      renderDpi: 150,
      includeAnnotations: true,
    }
    await expect(
      pdfToOdtBytes({
        ...valid,
        mode: 'editableText',
        pages: [{ ...firstPage, textRuns: [] }],
      }),
    ).rejects.toThrow('use page fidelity or run OCR first')
    await expect(pdfToOdtBytes({ ...valid, mode: 'fidelity', pages: [firstPage] })).rejects.toThrow(
      'image is empty',
    )
  })
})

describe('PDF to RTF', () => {
  const pages = [
    {
      pageNumber: 1,
      width: 612,
      height: 792,
      textRuns: [
        {
          text: '发布 {A\\B}',
          x: 48,
          y: 60,
          width: 120,
          height: 24,
          fontSize: 20,
          angle: 0,
          fontFamily: 'Arial; Unsafe',
          bold: true,
          italic: true,
        },
      ],
    },
    {
      pageNumber: 2,
      width: 792,
      height: 612,
      textRuns: [
        {
          text: 'Second page',
          x: 48,
          y: 60,
          width: 100,
          height: 18,
          fontSize: 14,
          angle: 0,
          fontFamily: 'Times New Roman',
          bold: false,
          italic: false,
        },
      ],
    },
  ]

  it('creates editable rich text with Unicode, styles, escaping, and page breaks', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612, 792]), {
      kind: 'pdfToRtf',
      pageCount: 2,
      pageIndexes: [0, 1],
      baseName: '../Launch:Brief.pdf',
      pages,
    })
    const rtf = decodeUtf8(output!.bytes)
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.rtf',
        fileName: '_Launch_Brief_converted.rtf',
        mimeType: 'application/rtf',
        extension: '.rtf',
      }),
    )
    expect(rtf).toMatch(/^\{\\rtf1/)
    expect(rtf).toContain('Arial Unsafe;')
    expect(rtf).toContain('\\b\\i ')
    expect(rtf).toContain('\\u21457?')
    expect(rtf).toContain('\\{A\\\\B\\}')
    expect(rtf).toContain('\\page')
    expect(rtf).toContain('Second page')
    expect(rtf.endsWith('}')).toBe(true)
  })

  it('rejects missing editable text and mismatched prepared pages', () => {
    expect(() =>
      pdfToRtfBytes({
        pageCount: 1,
        pageIndexes: [0],
        pages: [{ ...pages[0]!, textRuns: [] }],
      }),
    ).toThrow('run OCR first')
    expect(() => pdfToRtfBytes({ pageCount: 2, pageIndexes: [0, 1], pages: [pages[0]!] })).toThrow(
      'do not match the requested pages',
    )
  })
})

describe('PDF table export', () => {
  const tables = [
    {
      pageNumber: 1,
      tableNumber: 1,
      rows: [
        ['Name', 'Qty', 'Notes'],
        ['A, one', '2', '"quoted"'],
      ],
    },
    {
      pageNumber: 1,
      tableNumber: 2,
      rows: [
        ['Region', 'Value'],
        ['North & South', '12'],
      ],
    },
    {
      pageNumber: 3,
      tableNumber: 1,
      rows: [
        ['Month', 'Total'],
        ['August', '30'],
      ],
    },
  ]

  it('writes a single BOM-prefixed CSV with quoted cells', async () => {
    const output = await pdfTablesCsvOutput([tables[0]!], 'Launch.pdf')
    expect(output).toMatchObject({
      suffix: '_extracted.csv',
      mimeType: 'text/csv;charset=utf-8',
      extension: '.csv',
    })
    expect(Array.from(output.bytes.subarray(0, 3))).toEqual([0xef, 0xbb, 0xbf])
    expect(decodeUtf8(output.bytes.subarray(3))).toBe(
      '"Name","Qty","Notes"\r\n"A, one","2","""quoted"""\r\n',
    )
  })

  it('packages multiple tables as page-numbered CSV files', async () => {
    const output = await pdfTablesCsvOutput(tables, 'Launch:Plan.pdf')
    expect(output).toMatchObject({
      suffix: '_extracted_csv.zip',
      mimeType: 'application/zip',
      extension: '.zip',
    })
    const archive = await JSZip.loadAsync(output.bytes)
    expect(Object.keys(archive.files)).toEqual([
      'Launch_Plan_p1_t1.csv',
      'Launch_Plan_p1_t2.csv',
      'Launch_Plan_p3_t1.csv',
    ])
    expect(decodeUtf8(await archive.file('Launch_Plan_p1_t2.csv')!.async('uint8array'))).toContain(
      '"North & South","12"',
    )
  })

  it('creates one valid XLSX worksheet per detected table', async () => {
    const archive = await JSZip.loadAsync(await pdfTablesXlsxBytes(tables))
    const workbook = await archive.file('xl/workbook.xml')!.async('text')
    expect(workbook).toContain('name="Page 1 Table 1"')
    expect(workbook).toContain('name="Page 1 Table 2"')
    expect(workbook).toContain('name="Page 3"')
    const secondSheet = await archive.file('xl/worksheets/sheet2.xml')!.async('text')
    expect(secondSheet).toContain('North &amp; South')
    expect(secondSheet).toContain('<dimension ref="A1:B2"/>')
  })

  it('returns CSV archive and XLSX outputs through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'extractTables',
      format: 'both',
      pageIndexes: [0, 2],
      includeTwoColumnTextTables: false,
      baseName: 'Launch.pdf',
      tables: [tables[0]!, tables[2]!],
    })
    expect(
      outputs.map(({ suffix, mimeType, extension }) => ({ suffix, mimeType, extension })),
    ).toEqual([
      {
        suffix: '_extracted_csv.zip',
        mimeType: 'application/zip',
        extension: '.zip',
      },
      {
        suffix: '_tables.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        extension: '.xlsx',
      },
    ])
  })

  it('converts detected PDF tables to a named Excel workbook', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'pdfToXlsx',
      pageIndexes: [0, 2],
      includeTwoColumnTextTables: false,
      baseName: '../Launch:Plan.pdf',
      tables: [tables[0]!, tables[2]!],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.xlsx',
        fileName: '_Launch_Plan_converted.xlsx',
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        extension: '.xlsx',
      }),
    )
    const archive = await JSZip.loadAsync(output!.bytes)
    const workbook = await archive.file('xl/workbook.xml')!.async('text')
    expect(workbook).toContain('name="Page 1"')
    expect(workbook).toContain('name="Page 3"')
  })

  it('rejects missing, empty, and mismatched prepared table data', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      runPdfToolBytes(source, {
        kind: 'extractTables',
        format: 'csv',
        pageIndexes: [0],
        includeTwoColumnTextTables: false,
      }),
    ).rejects.toThrow('Extracted PDF tables are required')
    await expect(
      runPdfToolBytes(source, {
        kind: 'extractTables',
        format: 'csv',
        pageIndexes: [0],
        includeTwoColumnTextTables: false,
        baseName: 'Empty.pdf',
        tables: [],
      }),
    ).rejects.toThrow('No tables were detected')
    await expect(
      runPdfToolBytes(source, {
        kind: 'extractTables',
        format: 'xlsx',
        pageIndexes: [0],
        includeTwoColumnTextTables: false,
        baseName: 'Wrong.pdf',
        tables: [{ ...tables[0]!, pageNumber: 2 }],
      }),
    ).rejects.toThrow('do not match')
    await expect(pdfTablesCsvOutput([], 'Empty.pdf')).rejects.toThrow('No extracted PDF tables')
    await expect(pdfTablesXlsxBytes([])).rejects.toThrow('No extracted PDF tables')
  })
})

describe('PDF QR auto split', () => {
  it('calculates document ranges while removing divider pages', () => {
    expect(autoSplitPdfPageRanges(8, [2, 6], false)).toEqual([
      { firstPage: 0, lastPage: 1 },
      { firstPage: 3, lastPage: 5 },
      { firstPage: 7, lastPage: 7 },
    ])
    expect(autoSplitPdfPageRanges(8, [2, 6], true)).toEqual([
      { firstPage: 0, lastPage: 1 },
      { firstPage: 4, lastPage: 5 },
    ])
  })

  it('handles leading, trailing, and consecutive dividers without empty documents', () => {
    expect(autoSplitPdfPageRanges(7, [0, 1, 4, 6], false)).toEqual([
      { firstPage: 2, lastPage: 3 },
      { firstPage: 5, lastPage: 5 },
    ])
    expect(() => autoSplitPdfPageRanges(2, [2], false)).toThrow('invalid page')
  })

  it('generates a printable A4 divider with a square QR matrix', async () => {
    const modules = pdfAutoSplitDividerQrModules()
    expect(modules.length).toBeGreaterThan(20)
    expect(modules.every((row) => row.length === modules.length)).toBe(true)
    expect(modules.flat().some(Boolean)).toBe(true)
    const divider = await PDFDocument.load(await createPdfAutoSplitDividerBytes())
    expect(divider.getPageCount()).toBe(1)
    expect(divider.getPage(0).getSize()).toMatchObject({ width: 595.28, height: 841.89 })
  })

  it('creates a ZIP containing compact page-range PDFs', async () => {
    const source = await pdfWithWidths([100, 200, 300, 400, 500, 600])
    const archive = await JSZip.loadAsync(
      await autoSplitPdfZipBytes(source, [2, 5], false, 'Scan.pdf'),
    )
    expect(Object.keys(archive.files)).toEqual(['Scan_1.pdf', 'Scan_2.pdf'])
    expect(await pageWidths(await archive.file('Scan_1.pdf')!.async('uint8array'))).toEqual([
      100, 200,
    ])
    expect(await pageWidths(await archive.file('Scan_2.pdf')!.async('uint8array'))).toEqual([
      400, 500,
    ])
  })

  it('supports duplex divider removal and shared adapter metadata', async () => {
    const source = await pdfWithWidths([100, 200, 300, 400, 500])
    const outputs = await runPdfToolBytes(source, {
      kind: 'autoSplit',
      action: 'split',
      duplexMode: true,
      baseName: 'Duplex.pdf',
      dividerPageIndexes: [2],
    })
    expect(outputs[0]).toMatchObject({
      suffix: '_auto_split.zip',
      mimeType: 'application/zip',
      extension: '.zip',
    })
    const archive = await JSZip.loadAsync(outputs[0]!.bytes)
    expect(Object.keys(archive.files)).toEqual(['Duplex_1.pdf', 'Duplex_2.pdf'])
    expect(await pageWidths(await archive.file('Duplex_2.pdf')!.async('uint8array'))).toEqual([500])
  })

  it('returns the divider through the shared adapter and rejects missing detection', async () => {
    const source = await pdfWithWidths([100])
    const divider = await runPdfToolBytes(source, { kind: 'autoSplit', action: 'divider' })
    expect(divider[0]).toMatchObject({ suffix: '_auto_split_divider.pdf' })
    await expect(
      runPdfToolBytes(source, {
        kind: 'autoSplit',
        action: 'split',
        duplexMode: false,
      }),
    ).rejects.toThrow('Detected QR divider pages are required')
    await expect(
      runPdfToolBytes(source, {
        kind: 'autoSplit',
        action: 'split',
        duplexMode: false,
        baseName: 'No divider.pdf',
        dividerPageIndexes: [],
      }),
    ).rejects.toThrow('No QR divider pages')
  })
})

describe('insertPdfBytes', () => {
  it('inserts all source pages at the requested position', async () => {
    const result = await insertPdfBytes(
      await pdfWithWidths([100, 400]),
      await pdfWithWidths([200, 300]),
      0,
    )
    expect(result.count).toBe(2)
    expect(await pageWidths(result.merged)).toEqual([100, 200, 300, 400])
  })
})

describe('mergePdfBytes', () => {
  it('merges documents in input order', async () => {
    const result = await mergePdfBytes([
      await pdfWithWidths([100]),
      await pdfWithWidths([200, 300]),
    ])
    expect(await pageWidths(result)).toEqual([100, 200, 300])
  })

  it('rejects an empty input list', async () => {
    await expect(mergePdfBytes([])).rejects.toThrow('At least one PDF is required')
  })
})

describe('splitPdfBytes', () => {
  it('splits after zero-based page indexes and always includes the tail', async () => {
    const outputs = await splitPdfBytes(await pdfWithWidths([100, 200, 300, 400]), [1])
    expect(await Promise.all(outputs.map(pageWidths))).toEqual([
      [100, 200],
      [300, 400],
    ])
  })

  it('splits into fixed page-count chunks', async () => {
    const outputs = await splitPdfByPageCountBytes(
      await pdfWithWidths([100, 200, 300, 400, 500]),
      2,
    )
    expect(await Promise.all(outputs.map(pageWidths))).toEqual([[100, 200], [300, 400], [500]])
  })

  it('distributes remainder pages across a fixed document count', async () => {
    const outputs = await splitPdfByDocumentCountBytes(
      await pdfWithWidths([100, 200, 300, 400, 500, 600, 700]),
      3,
    )
    expect(await Promise.all(outputs.map(pageWidths))).toEqual([
      [100, 200, 300],
      [400, 500],
      [600, 700],
    ])
  })

  it('does not create empty files when document count exceeds page count', async () => {
    const outputs = await splitPdfByDocumentCountBytes(await pdfWithWidths([100, 200]), 5)
    expect(await Promise.all(outputs.map(pageWidths))).toEqual([[100], [200]])
  })

  it('packs the largest contiguous page ranges that fit a byte target', async () => {
    const source = await PDFDocument.create()
    for (let pageIndex = 0; pageIndex < 6; pageIndex++) {
      const page = source.addPage([200 + pageIndex * 10, 200])
      const content = new TextEncoder().encode(
        Array.from({ length: 1200 }, (_, index) =>
          String.fromCharCode(33 + ((index * 47 + pageIndex * 31) % 90)),
        ).join(''),
      )
      page.node.set(PDFName.of('Contents'), source.context.register(source.context.stream(content)))
    }
    const bytes = await source.save({ useObjectStreams: false })
    const singlePages = await splitPdfBySizeBytes(bytes, 1)
    const maxSinglePageSize = Math.max(...singlePages.map((output) => output.length))
    const targetSize = maxSinglePageSize * 2 + 256
    const outputs = await splitPdfBySizeBytes(bytes, targetSize)

    expect(outputs.length).toBeGreaterThan(1)
    expect(outputs.length).toBeLessThan(singlePages.length)
    expect(outputs.every((output) => output.length <= targetSize)).toBe(true)
    expect((await Promise.all(outputs.map(pageWidths))).flat()).toEqual([
      200, 210, 220, 230, 240, 250,
    ])
  })

  it('keeps range-local form fields when splitting by size', async () => {
    const source = await PDFDocument.create()
    const form = source.getForm()
    for (let pageIndex = 0; pageIndex < 3; pageIndex++) {
      const page = source.addPage([300, 200])
      form.createTextField(`size_field_${pageIndex + 1}`).addToPage(page)
    }
    const outputs = await splitPdfBySizeBytes(await source.save(), 1)
    const fieldNames = await Promise.all(
      outputs.map(async (bytes) =>
        (await PDFDocument.load(bytes))
          .getForm()
          .getFields()
          .map((field) => field.getName()),
      ),
    )
    expect(fieldNames).toEqual([['size_field_1'], ['size_field_2'], ['size_field_3']])
  })

  it('splits nested bookmarks through the requested chapter level', async () => {
    const source = await setPdfBookmarksBytes(await pdfWithWidths([100, 200, 300, 400, 500, 600]), [
      {
        title: 'Chapter 1',
        pageNumber: 1,
        children: [{ title: 'Section 1.1', pageNumber: 3, children: [] }],
      },
      { title: 'Chapter 2', pageNumber: 4, children: [] },
    ])

    const topLevel = await splitPdfByChaptersBytes(source, 0)
    expect(topLevel.map((output) => output.title)).toEqual(['Chapter 1', 'Chapter 2'])
    expect(await Promise.all(topLevel.map((output) => pageWidths(output.bytes)))).toEqual([
      [100, 200, 300],
      [400, 500, 600],
    ])

    const withSections = await splitPdfByChaptersBytes(source, 1)
    expect(withSections.map((output) => output.title)).toEqual([
      'Chapter 1',
      'Section 1.1',
      'Chapter 2',
    ])
    expect(await Promise.all(withSections.map((output) => pageWidths(output.bytes)))).toEqual([
      [100, 200],
      [300],
      [400, 500, 600],
    ])
  })

  it('merges same-page chapters or preserves their overlapping output', async () => {
    const document = await PDFDocument.create()
    const form = document.getForm()
    for (let pageIndex = 0; pageIndex < 4; pageIndex++) {
      const page = document.addPage([(pageIndex + 1) * 100, 200])
      form.createTextField(`field_${pageIndex + 1}`).addToPage(page)
    }
    const source = await setPdfBookmarksBytes(await document.save(), [
      { title: 'Intro A', pageNumber: 1, children: [] },
      { title: 'Intro B', pageNumber: 1, children: [] },
      { title: 'Body', pageNumber: 3, children: [] },
    ])

    const merged = await splitPdfByChaptersBytes(source, 0, false)
    expect(merged.map((output) => output.title)).toEqual(['Intro A Intro B', 'Body'])
    expect(await Promise.all(merged.map((output) => pageWidths(output.bytes)))).toEqual([
      [100, 200],
      [300, 400],
    ])
    expect(
      await Promise.all(
        merged.map(async (output) =>
          (await PDFDocument.load(output.bytes))
            .getForm()
            .getFields()
            .map((field) => field.getName()),
        ),
      ),
    ).toEqual([
      ['field_1', 'field_2'],
      ['field_3', 'field_4'],
    ])

    const duplicates = await splitPdfByChaptersBytes(source, 0, true)
    expect(duplicates.map((output) => output.title)).toEqual(['Intro A', 'Intro B', 'Body'])
    expect(await Promise.all(duplicates.map((output) => pageWidths(output.bytes)))).toEqual([
      [100],
      [100, 200],
      [300, 400],
    ])
  })

  it('rejects chapter splitting when the PDF has no bookmarks', async () => {
    await expect(splitPdfByChaptersBytes(await pdfWithWidths([100]), 0)).rejects.toThrow(
      'No PDF bookmarks',
    )
  })

  it('ignores external bookmarks when finding chapter boundaries', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 200])
    const secondPage = document.addPage([200, 200])
    const root = document.context.obj({})
    const rootRef = document.context.register(root)
    const external = document.context.obj({
      Title: PDFHexString.fromText('Website'),
      Parent: rootRef,
      A: { S: PDFName.of('URI'), URI: PDFHexString.fromText('https://example.com') },
    })
    const internal = document.context.obj({
      Title: PDFHexString.fromText('Internal'),
      Parent: rootRef,
      Dest: [secondPage.ref, PDFName.of('Fit')],
    })
    const externalRef = document.context.register(external)
    const internalRef = document.context.register(internal)
    external.set(PDFName.of('Next'), internalRef)
    internal.set(PDFName.of('Prev'), externalRef)
    root.set(PDFName.of('First'), externalRef)
    root.set(PDFName.of('Last'), internalRef)
    document.catalog.set(PDFName.of('Outlines'), rootRef)

    const outputs = await splitPdfByChaptersBytes(await document.save(), 0)
    expect(outputs.map((output) => output.title)).toEqual(['Internal'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([200])
  })

  it('keeps only form fields attached to each output range', async () => {
    const source = await PDFDocument.create()
    const form = source.getForm()
    for (let pageIndex = 0; pageIndex < 4; pageIndex++) {
      const page = source.addPage([300, 200])
      form.createTextField(`field_${pageIndex + 1}`).addToPage(page)
    }

    const outputs = await splitPdfByPageCountBytes(await source.save(), 2)
    const fieldNames = await Promise.all(
      outputs.map(async (bytes) =>
        (await PDFDocument.load(bytes))
          .getForm()
          .getFields()
          .map((field) => field.getName()),
      ),
    )
    expect(fieldNames).toEqual([
      ['field_1', 'field_2'],
      ['field_3', 'field_4'],
    ])
  })

  it('rejects non-positive split counts', async () => {
    const source = await pdfWithWidths([100, 200])
    await expect(splitPdfByPageCountBytes(source, 0)).rejects.toThrow('positive integer')
    await expect(splitPdfByDocumentCountBytes(source, 0)).rejects.toThrow('positive integer')
    await expect(splitPdfBySizeBytes(source, 0)).rejects.toThrow('positive safe integer')
  })
})

describe('PDF bookmarks', () => {
  it('writes and reads nested bookmarks with clamped page targets', async () => {
    const result = await setPdfBookmarksBytes(await pdfWithWidths([100, 200, 300]), [
      {
        title: 'Chapter 1',
        pageNumber: -5,
        children: [{ title: 'Section 1.1', pageNumber: 2, children: [] }],
      },
      { title: 'Chapter 2', pageNumber: 99, children: [] },
    ])
    expect(await listPdfBookmarksBytes(result)).toEqual([
      {
        title: 'Chapter 1',
        pageNumber: 1,
        children: [{ title: 'Section 1.1', pageNumber: 2, children: [] }],
      },
      { title: 'Chapter 2', pageNumber: 3, children: [] },
    ])
    const document = await PDFDocument.load(result)
    const outlines = document.catalog.lookup(PDFName.of('Outlines'), PDFDict)
    expect(outlines.lookup(PDFName.of('Count'), PDFNumber).asNumber()).toBe(3)
    expect(document.catalog.lookup(PDFName.of('PageMode'), PDFName).toString()).toBe('/UseOutlines')
  })

  it('reads GoTo action destinations', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 200])
    const secondPage = document.addPage([200, 200])
    const root = document.context.obj({})
    const rootRef = document.context.register(root)
    const item = document.context.obj({})
    item.set(PDFName.of('Title'), PDFHexString.fromText('Action target'))
    item.set(PDFName.of('Parent'), rootRef)
    item.set(
      PDFName.of('A'),
      document.context.obj({
        S: PDFName.of('GoTo'),
        D: document.context.obj([secondPage.ref, PDFName.of('Fit')]),
      }),
    )
    const itemRef = document.context.register(item)
    root.set(PDFName.of('First'), itemRef)
    root.set(PDFName.of('Last'), itemRef)
    document.catalog.set(PDFName.of('Outlines'), rootRef)

    expect(await listPdfBookmarksBytes(await document.save())).toEqual([
      { title: 'Action target', pageNumber: 2, children: [] },
    ])
  })

  it('clears the outline and rejects empty titles', async () => {
    const source = await setPdfBookmarksBytes(await pdfWithWidths([100]), [
      { title: 'Only chapter', pageNumber: 1, children: [] },
    ])
    const cleared = await setPdfBookmarksBytes(source, [])
    expect(await listPdfBookmarksBytes(cleared)).toEqual([])
    const document = await PDFDocument.load(cleared)
    expect(document.catalog.has(PDFName.of('Outlines'))).toBe(false)
    expect(document.catalog.has(PDFName.of('PageMode'))).toBe(false)
    await expect(
      setPdfBookmarksBytes(source, [{ title: ' ', pageNumber: 1, children: [] }]),
    ).rejects.toThrow('must not be empty')
  })
})

describe('analyzePdfBytes', () => {
  it('aggregates document, page, form, annotation, and resource information', async () => {
    const document = await PDFDocument.create()
    document.setTitle('Analysis fixture')
    document.setAuthor('GenOffice')
    document.setSubject('Local PDF analysis')
    document.setKeywords(['analysis', 'pdf'])
    const font = await document.embedFont('Helvetica')
    const image = await document.embedPng(tinyPngBytes())
    const firstPage = document.addPage([300, 200])
    firstPage.drawText('Analyzed text', { font })
    firstPage.drawImage(image, { x: 20, y: 20, width: 20, height: 20 })
    document.getForm().createTextField('name').addToPage(firstPage)
    const secondPage = document.addPage([400, 250])
    secondPage.setRotation(degrees(90))
    const note = document.context.obj({
      Type: PDFName.of('Annot'),
      Subtype: PDFName.of('Text'),
      Rect: [10, 10, 30, 30],
    })
    secondPage.node.set(
      PDFName.of('Annots'),
      document.context.obj([document.context.register(note)]),
    )
    await document.attach(new TextEncoder().encode('notes'), 'notes.txt', {
      mimeType: 'text/plain',
    })
    const acroForm = document.catalog.lookup(PDFName.of('AcroForm'), PDFDict)
    const signature = document.context.obj({
      FT: PDFName.of('Sig'),
      T: PDFHexString.fromText('approval'),
    })
    acroForm.lookup(PDFName.of('Fields'), PDFArray).push(document.context.register(signature))
    const bookmarked = await setPdfBookmarksBytes(
      await document.save({ useObjectStreams: false }),
      [
        {
          title: 'Chapter',
          pageNumber: 1,
          children: [{ title: 'Section', pageNumber: 2, children: [] }],
        },
      ],
    )
    const withXfa = await PDFDocument.load(bookmarked, { updateMetadata: false })
    withXfa.catalog
      .lookup(PDFName.of('AcroForm'), PDFDict)
      .set(PDFName.of('XFA'), PDFHexString.fromText('fixture'))
    const source = await withXfa.save({
      useObjectStreams: false,
      updateFieldAppearances: false,
    })

    const analysis = await analyzePdfBytes(source)
    expect(analysis.pageCount).toBe(2)
    expect(analysis.pdfVersion).toMatch(/^1\./)
    expect(analysis.fileSize).toBe(source.byteLength)
    expect(analysis.isEncrypted).toBe(false)
    expect(analysis.properties).toMatchObject({
      title: 'Analysis fixture',
      author: 'GenOffice',
      subject: 'Local PDF analysis',
    })
    expect(analysis.pages).toEqual([
      { pageNumber: 1, width: 300, height: 200, rotation: 0 },
      { pageNumber: 2, width: 400, height: 250, rotation: 90 },
    ])
    expect(analysis.form).toEqual({ fieldCount: 2, hasXfa: true, signatureCount: 1 })
    expect(analysis.annotations.totalCount).toBe(2)
    expect(analysis.annotations.typeBreakdown).toMatchObject({ Widget: 1, Text: 1 })
    expect(analysis.fonts).toContain('Helvetica')
    expect(analysis.imageCount).toBe(1)
    expect(analysis.attachmentCount).toBe(1)
    expect(analysis.bookmarkCount).toBe(2)
  })
})

describe('preflightPdfBytes', () => {
  it('passes a structurally sound PDF without declared standards', async () => {
    const report = await preflightPdfBytes(await pdfWithWidths([200, 300]))
    expect(report).toMatchObject({
      schema: 'genoffice.pdf.preflight',
      version: 1,
      status: 'pass',
      parseable: true,
      strictParsing: true,
      pageCount: 2,
      standards: [],
      disclaimer: 'local-structural-preflight',
    })
    expect(report.structure.startXrefInRange).toBe(true)
    expect(report.structure.startXrefTargetValid).toBe(true)
    expect(report.findings).toEqual([])
  })

  it('detects PDF/A and PDF/UA declarations without claiming formal compliance', async () => {
    const document = await PDFDocument.create()
    document.addPage([300, 200])
    const xmp = `<?xpacket begin=""?>
      <x:xmpmeta xmlns:x="adobe:ns:meta/">
        <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
          <rdf:Description xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"
            pdfaid:part="2" pdfaid:conformance="B" />
          <rdf:Description xmlns:pdfuaid="http://www.aiim.org/pdfua/ns/id/"
            pdfuaid:part="1" pdfuaid:rev="2024" />
        </rdf:RDF>
      </x:xmpmeta><?xpacket end="w"?>`
    document.catalog.set(
      PDFName.of('Metadata'),
      document.context.register(
        document.context.flateStream(new TextEncoder().encode(xmp), {
          Type: 'Metadata',
          Subtype: 'XML',
        }),
      ),
    )
    document.catalog.set(PDFName.of('MarkInfo'), document.context.obj({ Marked: true }))
    document.catalog.set(
      PDFName.of('StructTreeRoot'),
      document.context.obj({ Type: 'StructTreeRoot', K: [] }),
    )
    document.catalog.set(PDFName.of('Lang'), PDFString.of('en-US'))
    document.catalog.set(
      PDFName.of('OutputIntents'),
      document.context.obj([
        document.context.obj({
          Type: 'OutputIntent',
          S: 'GTS_PDFA1',
          OutputConditionIdentifier: PDFString.of('sRGB'),
        }),
      ]),
    )

    const report = await preflightPdfBytes(await document.save({ useObjectStreams: false }))
    expect(report.status).toBe('pass')
    expect(report.standards).toEqual([
      { family: 'PDF/A', part: '2', conformance: 'B', label: 'PDF/A-2b' },
      { family: 'PDF/UA', part: '1', revision: '2024', label: 'PDF/UA-1:2024' },
    ])
    expect(report.features).toMatchObject({
      hasXmpMetadata: true,
      xmpValid: true,
      tagged: true,
      marked: true,
      language: 'en-US',
      outputIntentCount: 1,
    })
    expect(report.findings).toEqual([
      { code: 'standardDeclaredOnly', severity: 'info', detail: 'PDF/A-2b' },
      { code: 'standardDeclaredOnly', severity: 'info', detail: 'PDF/UA-1:2024' },
    ])
  })

  it('reports an out-of-range startxref pointer even when the parser can recover', async () => {
    const source = await PDFDocument.create()
    source.addPage([320, 240])
    const bytes = new Uint8Array(await source.save({ useObjectStreams: false }))
    const text = new TextDecoder('latin1').decode(bytes)
    const match = /startxref\s+(\d+)\s+%%EOF\s*$/.exec(text)
    expect(match).not.toBeNull()
    const pointer = match![1]!
    const offset = match!.index + match![0].indexOf(pointer)
    bytes.fill(0x39, offset, offset + pointer.length)

    const report = await preflightPdfBytes(bytes)
    expect(report.status).toBe('error')
    expect(report.structure.startXrefInRange).toBe(false)
    expect(report.findings).toContainEqual({
      code: 'startXrefOutOfRange',
      severity: 'error',
      detail: '9'.repeat(pointer.length),
    })
  })

  it('rejects an in-range startxref pointer that does not target cross-reference data', async () => {
    const source = await PDFDocument.create()
    source.addPage([320, 240])
    const bytes = new Uint8Array(await source.save({ useObjectStreams: false }))
    const text = new TextDecoder('latin1').decode(bytes)
    const match = /startxref\s+(\d+)\s+%%EOF\s*$/.exec(text)
    expect(match).not.toBeNull()
    const pointer = match![1]!
    const offset = match!.index + match![0].indexOf(pointer)
    bytes.fill(0x30, offset, offset + pointer.length)

    const report = await preflightPdfBytes(bytes)
    expect(report.status).toBe('error')
    expect(report.structure.startXrefInRange).toBe(true)
    expect(report.structure.startXrefTargetValid).toBe(false)
    expect(report.findings).toContainEqual({
      code: 'startXrefTargetInvalid',
      severity: 'error',
      detail: '0',
    })
  })
})

describe('splitPdfSectionsBytes', () => {
  it('splits a page into merged row-major sections with matching dimensions', async () => {
    const outputs = await splitPdfSectionsBytes(await pdfWithWidths([400]), {
      pageIndexes: [0],
      rows: 2,
      columns: 2,
      merge: true,
      arrangement: 'rows',
    })
    expect(outputs).toHaveLength(1)
    expect(await pageSizes(outputs[0]!.bytes)).toEqual([
      [200, 100],
      [200, 100],
      [200, 100],
      [200, 100],
    ])
  })

  it('keeps unselected pages whole in merged output', async () => {
    const outputs = await splitPdfSectionsBytes(await pdfWithWidths([300, 400]), {
      pageIndexes: [1],
      rows: 1,
      columns: 2,
      merge: true,
    })
    expect(await pageSizes(outputs[0]!.bytes)).toEqual([
      [300, 200],
      [200, 200],
      [200, 200],
    ])
  })

  it('returns separately named files for every section', async () => {
    const source = await pdfWithWidths([400])
    const outputs = await splitPdfSectionsBytes(source, {
      pageIndexes: [0],
      rows: 2,
      columns: 2,
      merge: false,
      arrangement: 'columns',
    })
    expect(
      outputs.map(({ sourcePageNumber, sectionNumber }) => [sourcePageNumber, sectionNumber]),
    ).toEqual([
      [1, 1],
      [1, 2],
      [1, 3],
      [1, 4],
    ])
    expect(await Promise.all(outputs.map((output) => pageSizes(output.bytes)))).toEqual([
      [[200, 100]],
      [[200, 100]],
      [[200, 100]],
      [[200, 100]],
    ])

    const adapterOutputs = await runPdfToolBytes(source, {
      kind: 'splitSections',
      pageIndexes: [0],
      rows: 1,
      columns: 2,
      merge: false,
    })
    expect(adapterOutputs.map((output) => output.suffix)).toEqual([
      '_page_1_section_1.pdf',
      '_page_1_section_2.pdf',
    ])
  })

  it('rejects invalid grids and empty page selections', async () => {
    const source = await pdfWithWidths([400])
    await expect(
      splitPdfSectionsBytes(source, { pageIndexes: [0], rows: 0, columns: 2, merge: true }),
    ).rejects.toThrow('positive integer')
    await expect(
      splitPdfSectionsBytes(source, { pageIndexes: [], rows: 1, columns: 2, merge: true }),
    ).rejects.toThrow('At least one page')
  })
})

describe('removePdfImagesBytes', () => {
  it('removes images only from selected pages while preserving forms', async () => {
    const source = await PDFDocument.create()
    const image = await source.embedPng(tinyPngBytes())
    const firstPage = source.addPage([300, 200])
    firstPage.drawText('Keep this text')
    firstPage.drawImage(image, { x: 20, y: 20, width: 40, height: 40 })
    source.getForm().createTextField('name').addToPage(firstPage)
    const secondPage = source.addPage([300, 200])
    secondPage.drawImage(image, { x: 20, y: 20, width: 40, height: 40 })

    const result = await removePdfImagesBytes(await source.save(), [0])
    const document = await PDFDocument.load(result)
    expect(document.getPageCount()).toBe(2)
    expect(countImagesInResources(document, document.getPage(0).node.Resources())).toBe(0)
    expect(countImagesInResources(document, document.getPage(1).node.Resources())).toBe(1)
    expect(document.getForm().getTextField('name')).toBeDefined()
  })

  it('recursively removes images inside form XObjects', async () => {
    const source = await PDFDocument.create()
    const page = source.addPage([300, 200])
    const image = await source.embedPng(tinyPngBytes())
    const form = source.context.stream('q /Im0 Do Q', {
      Type: 'XObject',
      Subtype: 'Form',
      BBox: [0, 0, 100, 100],
      Resources: { XObject: { Im0: image.ref } },
    })
    const resources = page.node.Resources() ?? source.context.obj({})
    page.node.set(PDFName.of('Resources'), resources)
    const xObjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict) ?? source.context.obj({})
    resources.set(PDFName.of('XObject'), xObjects)
    xObjects.set(PDFName.of('Fm0'), source.context.register(form))

    const result = await removePdfImagesBytes(await source.save(), [0])
    const document = await PDFDocument.load(result)
    expect(countImagesInResources(document, document.getPage(0).node.Resources())).toBe(0)
  })
})

describe('invertPdfColorsBytes', () => {
  it('adds a white Difference layer only to selected pages', async () => {
    const source = await PDFDocument.create()
    const firstPage = source.addPage([300, 200])
    firstPage.drawText('Invert me')
    const secondPage = source.addPage([400, 250])
    secondPage.drawText('Leave me')
    const sourceBytes = await source.save({ useObjectStreams: false })
    const original = await PDFDocument.load(sourceBytes)
    const originalSecondContents = original.getPage(1).node.Contents()
    const originalSecondContentCount =
      originalSecondContents instanceof PDFArray ? originalSecondContents.size() : 1

    const result = await invertPdfColorsBytes(sourceBytes, [0])
    const document = await PDFDocument.load(result)
    const firstResources = document.getPage(0).node.Resources()
    const firstExtGState = firstResources?.lookupMaybe(PDFName.of('ExtGState'), PDFDict)
    expect(firstExtGState).toBeDefined()
    const invertState = document.context.lookup(firstExtGState?.get(PDFName.of('GOInvert')))
    expect(invertState).toBeInstanceOf(PDFDict)
    expect((invertState as PDFDict).lookupMaybe(PDFName.of('BM'), PDFName)?.decodeText()).toBe(
      'Difference',
    )
    const secondExtGState = document
      .getPage(1)
      .node.Resources()
      ?.lookupMaybe(PDFName.of('ExtGState'), PDFDict)
    expect(secondExtGState?.has(PDFName.of('GOInvert')) ?? false).toBe(false)

    const firstContents = document.getPage(0).node.Contents()
    expect(firstContents).toBeInstanceOf(PDFArray)
    const invertedStream = document.context.lookup(
      (firstContents as PDFArray).get((firstContents as PDFArray).size() - 1),
    )
    expect(invertedStream).toBeInstanceOf(PDFRawStream)
    const decoded = new TextDecoder().decode(
      decodePDFRawStream(invertedStream as PDFRawStream).decode(),
    )
    expect(decoded).toContain('/GOInvert gs')
    expect(decoded).toContain('1 1 1 rg')
    expect(decoded).toContain('0 0 300 200 re')

    const secondContents = document.getPage(1).node.Contents()
    const secondContentCount = secondContents instanceof PDFArray ? secondContents.size() : 1
    expect(secondContentCount).toBe(originalSecondContentCount)
  })

  it('preserves forms, bookmarks, attachments, and page geometry', async () => {
    const source = await PDFDocument.create()
    const firstPage = source.addPage([300, 200])
    source.getForm().createTextField('name').addToPage(firstPage)
    source.addPage([400, 250])
    let sourceBytes = await source.save({ useObjectStreams: false })
    sourceBytes = await setPdfBookmarksBytes(sourceBytes, [
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    sourceBytes = await addPdfAttachmentsBytes(sourceBytes, [
      { name: 'notes.txt', bytes: new TextEncoder().encode('hello') },
    ])

    const result = await invertPdfColorsBytes(sourceBytes, [0, 1])
    const document = await PDFDocument.load(result)
    expect(document.getPageCount()).toBe(2)
    expect(document.getPages().map((page) => page.getSize())).toEqual([
      { width: 300, height: 200 },
      { width: 400, height: 250 },
    ])
    expect(document.getForm().getTextField('name')).toBeDefined()
    expect(await listPdfBookmarksBytes(result)).toEqual([
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    expect((await listPdfAttachmentsBytes(result)).map((attachment) => attachment.name)).toEqual([
      'notes.txt',
    ])
  })

  it('rejects empty and invalid page selections', async () => {
    const source = await pdfWithWidths([100])
    await expect(invertPdfColorsBytes(source, [])).rejects.toThrow('At least one page')
    await expect(invertPdfColorsBytes(source, [1])).rejects.toThrow('invalid page')
  })
})

describe('replacePdfColorsBytes', () => {
  it('maps page luminosity through a soft mask only on selected pages', async () => {
    const source = await PDFDocument.create()
    const firstPage = source.addPage([300, 200])
    firstPage.drawText('Recolor me')
    const secondPage = source.addPage([400, 250])
    secondPage.drawText('Leave me')

    const result = await replacePdfColorsBytes(
      await source.save({ useObjectStreams: false }),
      [0],
      '#00ff00',
      '#000000',
    )
    const document = await PDFDocument.load(result)
    const extGState = document
      .getPage(0)
      .node.Resources()
      ?.lookupMaybe(PDFName.of('ExtGState'), PDFDict)
    const recolorState = document.context.lookup(extGState?.get(PDFName.of('GORecolor')))
    expect(recolorState).toBeInstanceOf(PDFDict)
    const softMask = (recolorState as PDFDict).lookupMaybe(PDFName.of('SMask'), PDFDict)
    expect(softMask?.lookupMaybe(PDFName.of('S'), PDFName)?.decodeText()).toBe('Luminosity')
    const maskStream = document.context.lookup(softMask?.get(PDFName.of('G')))
    expect(maskStream).toBeInstanceOf(PDFRawStream)
    const group = (maskStream as PDFRawStream).dict.lookupMaybe(PDFName.of('Group'), PDFDict)
    expect(group?.lookupMaybe(PDFName.of('S'), PDFName)?.decodeText()).toBe('Transparency')
    const secondExtGState = document
      .getPage(1)
      .node.Resources()
      ?.lookupMaybe(PDFName.of('ExtGState'), PDFDict)
    expect(secondExtGState?.has(PDFName.of('GORecolor')) ?? false).toBe(false)

    const contents = document.getPage(0).node.Contents()
    expect(contents).toBeInstanceOf(PDFArray)
    const overlay = document.context.lookup(
      (contents as PDFArray).get((contents as PDFArray).size() - 1),
    )
    expect(overlay).toBeInstanceOf(PDFRawStream)
    const decoded = new TextDecoder().decode(decodePDFRawStream(overlay as PDFRawStream).decode())
    expect(decoded).toContain('0 1 0 rg')
    expect(decoded).toContain('/GORecolor gs')
    expect(decoded).toContain('0 0 0 rg')
    expect(decoded).toContain('0 0 300 200 re')
  })

  it('preserves forms, bookmarks, attachments, and page geometry', async () => {
    const source = await PDFDocument.create()
    const firstPage = source.addPage([300, 200])
    firstPage.drawText('Searchable source text')
    source.getForm().createTextField('name').addToPage(firstPage)
    source.addPage([400, 250])
    let sourceBytes = await source.save({ useObjectStreams: false })
    sourceBytes = await setPdfBookmarksBytes(sourceBytes, [
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    sourceBytes = await addPdfAttachmentsBytes(sourceBytes, [
      { name: 'notes.txt', bytes: new TextEncoder().encode('hello') },
    ])

    const result = await replacePdfColorsBytes(sourceBytes, [0, 1], '#ffffff', '#000000')
    const document = await PDFDocument.load(result)
    expect(document.getPages().map((page) => page.getSize())).toEqual([
      { width: 300, height: 200 },
      { width: 400, height: 250 },
    ])
    expect(document.getForm().getTextField('name')).toBeDefined()
    expect(await listPdfBookmarksBytes(result)).toEqual([
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    expect((await listPdfAttachmentsBytes(result)).map((attachment) => attachment.name)).toEqual([
      'notes.txt',
    ])
  })

  it('rejects invalid colors and page selections', async () => {
    const source = await pdfWithWidths([100])
    await expect(replacePdfColorsBytes(source, [0], 'green', '#000000')).rejects.toThrow(
      'six-digit hex color',
    )
    await expect(replacePdfColorsBytes(source, [], '#ffffff', '#000000')).rejects.toThrow(
      'At least one page',
    )
  })
})

describe('overlayAdjustedPdfPagesBytes', () => {
  it('overlays selected pages while preserving document structure and geometry', async () => {
    const source = await PDFDocument.create()
    const firstPage = source.addPage([300, 200])
    firstPage.setCropBox(10, 20, 250, 150)
    firstPage.drawText('Searchable source text')
    source.getForm().createTextField('name').addToPage(firstPage)
    source.addPage([400, 250]).drawText('Unchanged page')
    let sourceBytes = await source.save({ useObjectStreams: false })
    sourceBytes = await setPdfBookmarksBytes(sourceBytes, [
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    sourceBytes = await addPdfAttachmentsBytes(sourceBytes, [
      { name: 'notes.txt', bytes: new TextEncoder().encode('hello') },
    ])

    const result = await overlayAdjustedPdfPagesBytes(sourceBytes, [0], [tinyPngBytes()])
    const document = await PDFDocument.load(result)
    expect(document.getPage(0).getCropBox()).toEqual({ x: 10, y: 20, width: 250, height: 150 })
    expect(document.getPage(1).getSize()).toEqual({ width: 400, height: 250 })
    expect(countImagesInResources(document, document.getPage(0).node.Resources())).toBe(1)
    expect(countImagesInResources(document, document.getPage(1).node.Resources())).toBe(0)
    expect(document.getForm().getTextField('name')).toBeDefined()
    expect(await listPdfBookmarksBytes(result)).toEqual([
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    expect((await listPdfAttachmentsBytes(result)).map((attachment) => attachment.name)).toEqual([
      'notes.txt',
    ])
  })

  it('rejects missing, mismatched, and invalid adjusted page images', async () => {
    const source = await pdfWithWidths([100])
    await expect(overlayAdjustedPdfPagesBytes(source, [], [])).rejects.toThrow('At least one page')
    await expect(overlayAdjustedPdfPagesBytes(source, [0], [])).rejects.toThrow('must match')
    await expect(overlayAdjustedPdfPagesBytes(source, [0], [new Uint8Array()])).rejects.toThrow(
      'image is empty',
    )
    await expect(overlayAdjustedPdfPagesBytes(source, [1], [tinyPngBytes()])).rejects.toThrow(
      'invalid page',
    )
  })
})

describe('cropPdfMarginsBytes', () => {
  it('applies margins to every page crop and media box', async () => {
    const bytes = await cropPdfMarginsBytes(await pdfWithWidths([200]), {
      top: 20,
      right: 30,
      bottom: 40,
      left: 50,
    })
    const page = (await PDFDocument.load(bytes)).getPage(0)
    expect(page.getMediaBox()).toEqual({ x: 50, y: 40, width: 120, height: 140 })
    expect(page.getCropBox()).toEqual({ x: 50, y: 40, width: 120, height: 140 })
  })

  it('applies distinct detected crop boxes while preserving page content', async () => {
    const source = await PDFDocument.create()
    source.addPage([200, 300]).drawText('First', { x: 40, y: 50 })
    source.addPage([400, 500]).drawText('Second', { x: 100, y: 120 })
    const bytes = await cropPdfPageBoxesBytes(await source.save(), [
      { x: 20, y: 30, width: 160, height: 240 },
      { x: 80, y: 100, width: 250, height: 320 },
    ])
    const pages = (await PDFDocument.load(bytes)).getPages()
    expect(pages[0]!.getCropBox()).toEqual({ x: 20, y: 30, width: 160, height: 240 })
    expect(pages[1]!.getCropBox()).toEqual({ x: 80, y: 100, width: 250, height: 320 })
    expect(pages[0]!.node.Contents()).toBeDefined()
    expect(pages[1]!.node.Contents()).toBeDefined()
  })

  it('rejects missing or out-of-page auto crop boxes', async () => {
    const source = await pdfWithWidths([200])
    await expect(cropPdfPageBoxesBytes(source, [])).rejects.toThrow('page count')
    await expect(
      cropPdfPageBoxesBytes(source, [{ x: -1, y: 0, width: 200, height: 200 }]),
    ).rejects.toThrow('outside page 1')
  })
})

describe('scalePdfPagesBytes', () => {
  it('centers pages on the selected target size', async () => {
    const bytes = await scalePdfPagesBytes(await pdfWithWidths([300]), {
      kind: 'scale',
      pageSize: 'LETTER',
      orientation: 'landscape',
      scaleFactor: 0.9,
    })
    const page = (await PDFDocument.load(bytes)).getPage(0)
    expect(page.getWidth()).toBe(792)
    expect(page.getHeight()).toBe(612)
  })

  it('keeps each original page size for mixed-size documents', async () => {
    const bytes = await scalePdfPagesBytes(await pdfWithWidths([200, 400]), {
      kind: 'scale',
      pageSize: 'KEEP',
      orientation: 'portrait',
      scaleFactor: 0.75,
    })
    expect(await pageWidths(bytes)).toEqual([200, 400])
  })
})

describe('nUpPdfBytes', () => {
  it('lays out multiple source pages on A4 sheets', async () => {
    const bytes = await nUpPdfBytes(await pdfWithWidths([100, 200, 300, 400, 500]), {
      rows: 2,
      columns: 2,
      orientation: 'portrait',
      innerMargin: 8,
      borderWidth: 1,
    })
    const document = await PDFDocument.load(bytes)
    expect(document.getPageCount()).toBe(2)
    expect(document.getPage(0).getWidth()).toBeCloseTo(595.28, 2)
    expect(document.getPage(0).getHeight()).toBeCloseTo(841.89, 2)
  })

  it('supports custom reading order, margins, and borders around fitted pages', async () => {
    const bytes = await nUpPdfBytes(await pdfWithWidths([100, 200, 300]), {
      rows: 1,
      columns: 3,
      orientation: 'landscape',
      arrangement: 'columns',
      readingDirection: 'rtl',
      innerMargin: 7,
      topMargin: 11,
      rightMargin: 13,
      bottomMargin: 17,
      leftMargin: 19,
      borderWidth: 2.5,
    })
    const document = await PDFDocument.load(bytes)
    expect(document.getPageCount()).toBe(1)
    expect(document.getPage(0).getWidth()).toBeCloseTo(841.89, 2)
    expect(document.getPage(0).getHeight()).toBeCloseTo(595.28, 2)

    const streams = await decodedPageStreams(bytes)
    expect(streams.join('\n')).toContain('2.5 w')
  })

  it('rejects excessive borders and margins that leave no content area', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      nUpPdfBytes(source, {
        rows: 1,
        columns: 1,
        orientation: 'portrait',
        borderWidth: 73,
      }),
    ).rejects.toThrow('borderWidth')
    await expect(
      nUpPdfBytes(source, {
        rows: 1,
        columns: 1,
        orientation: 'portrait',
        leftMargin: 300,
        rightMargin: 300,
      }),
    ).rejects.toThrow('Margins leave no room')
  })
})

describe('bookletPdfBytes', () => {
  it('creates saddle-stitch page pairs with blank padding', () => {
    expect(bookletPagePairs(5, 'both', false)).toEqual([
      [-1, 0],
      [1, -1],
      [-1, 2],
      [3, 4],
    ])
    expect(bookletPagePairs(4, 'back', true)).toEqual([[2, 1]])
  })

  it('places two booklet pages on each landscape sheet side', async () => {
    const bytes = await bookletPdfBytes(await pdfWithWidths([100, 100, 100, 100]), {
      spine: 'left',
      gutter: 12,
      border: true,
      duplexPass: 'both',
      flipOnShortEdge: false,
    })
    const document = await PDFDocument.load(bytes)
    expect(document.getPageCount()).toBe(2)
    expect(document.getPage(0).getWidth()).toBe(200)
    expect(document.getPage(0).getHeight()).toBe(100)
  })
})

describe('posterPdfBytes', () => {
  it('splits each source page into the requested printable grid', async () => {
    const bytes = await posterPdfBytes(await pdfWithWidths([600]), {
      pageSize: 'A5',
      rows: 2,
      columns: 3,
      readingDirection: 'rtl',
    })
    const document = await PDFDocument.load(bytes)
    expect(document.getPageCount()).toBe(6)
    expect(document.getPage(0).getWidth()).toBeCloseTo(419.53, 2)
    expect(document.getPage(0).getHeight()).toBeCloseTo(595.28, 2)
  })
})

describe('singlePagePdfBytes', () => {
  it('joins pages vertically or horizontally', async () => {
    const source = await pdfWithWidths([100, 200, 300])
    const vertical = await PDFDocument.load(
      await singlePagePdfBytes(source, { direction: 'vertical' }),
    )
    expect(vertical.getPage(0).getSize()).toEqual({ width: 300, height: 600 })

    const horizontal = await PDFDocument.load(
      await singlePagePdfBytes(source, { direction: 'horizontal' }),
    )
    expect(horizontal.getPage(0).getSize()).toEqual({ width: 600, height: 200 })
  })
})

describe('overlayPageAssignments', () => {
  it('cycles all overlay pages sequentially', () => {
    expect(overlayPageAssignments(7, [2, 1], 'sequential')).toEqual([
      { documentIndex: 0, pageIndex: 0 },
      { documentIndex: 0, pageIndex: 1 },
      { documentIndex: 1, pageIndex: 0 },
      { documentIndex: 0, pageIndex: 0 },
      { documentIndex: 0, pageIndex: 1 },
      { documentIndex: 1, pageIndex: 0 },
      { documentIndex: 0, pageIndex: 0 },
    ])
  })

  it('round-robins overlay documents and cycles their pages', () => {
    expect(overlayPageAssignments(6, [2, 3], 'interleaved')).toEqual([
      { documentIndex: 0, pageIndex: 0 },
      { documentIndex: 1, pageIndex: 0 },
      { documentIndex: 0, pageIndex: 1 },
      { documentIndex: 1, pageIndex: 1 },
      { documentIndex: 0, pageIndex: 0 },
      { documentIndex: 1, pageIndex: 2 },
    ])
  })

  it('applies fixed repeat counts and leaves remaining base pages unchanged', () => {
    expect(overlayPageAssignments(7, [2, 1], 'fixedRepeat', [1, 2])).toEqual([
      { documentIndex: 0, pageIndex: 0 },
      { documentIndex: 0, pageIndex: 1 },
      { documentIndex: 1, pageIndex: 0 },
      { documentIndex: 1, pageIndex: 0 },
      null,
      null,
      null,
    ])
  })
})

async function decodedPageStreams(bytes: Uint8Array): Promise<string[]> {
  const document = await PDFDocument.load(bytes)
  const contents = document.getPage(0).node.Contents()
  if (!(contents instanceof PDFArray)) return []
  return contents.asArray().map((streamRef) => {
    const stream = document.context.lookup(streamRef)
    if (!(stream instanceof PDFRawStream)) throw new Error('Expected a raw page content stream')
    return new TextDecoder().decode(decodePDFRawStream(stream).decode())
  })
}

describe('addPdfPageNumbersBytes', () => {
  const options = {
    pageIndexes: [2, 0],
    position: 'topRight' as const,
    margin: 'small' as const,
    fontSize: 14,
    font: 'courier' as const,
    fontColor: '#336699',
    startingNumber: 9,
    zeroPad: 3,
    textPattern: 'Page {n} of {total} - {filename}',
    baseName: 'Launch.pdf',
  }

  it('formats selected pages in document order with placeholders and zero padding', () => {
    expect(pdfPageNumberLabels(4, options)).toEqual([
      { pageIndex: 0, text: 'Page 009 of 4 - Launch' },
      { pageIndex: 2, text: 'Page 010 of 4 - Launch' },
    ])
  })

  it('draws colored page labels at the selected position without touching other pages', async () => {
    const source = await pdfWithWidths([500, 500, 500])
    const result = await addPdfPageNumbersBytes(source, options)
    const document = await PDFDocument.load(result)
    expect(document.getPageCount()).toBe(3)
    const streams = document.getPages().map((page) => {
      const contents = page.node.Contents()
      if (!(contents instanceof PDFArray)) return ''
      return contents
        .asArray()
        .map((reference) => {
          const stream = document.context.lookup(reference)
          return stream instanceof PDFRawStream
            ? decodeUtf8(decodePDFRawStream(stream).decode())
            : ''
        })
        .join('\n')
    })
    expect(streams[0]).toContain('0.2 0.4 0.6 rg')
    expect(streams[0]).toContain('<5061676520303039206F662033202D204C61756E6368> Tj')
    const transform = /1 0 0 1 ([\d.]+) ([\d.]+) Tm/.exec(streams[0]!)
    expect(Number(transform?.[1])).toBeGreaterThan(100)
    expect(Number(transform?.[2])).toBeGreaterThan(150)
    expect(streams[1]).toBe('')
    expect(streams[2]).toContain('<5061676520303130206F662033202D204C61756E6368> Tj')
  })

  it('returns the shared page-number tool output and rejects invalid settings', async () => {
    const source = await pdfWithWidths([200])
    const outputs = await runPdfToolBytes(source, {
      kind: 'pageNumbers',
      ...options,
      pageIndexes: [0],
    })
    expect(outputs[0]).toMatchObject({ suffix: '_page_numbers_added.pdf' })
    await expect(
      addPdfPageNumbersBytes(source, { ...options, pageIndexes: [0], zeroPad: 13 }),
    ).rejects.toThrow('padding must be between 0 and 12')
  })
})

describe('overlayPdfBytes', () => {
  it('preserves base pages and form fields', async () => {
    const base = await PDFDocument.create()
    const page = base.addPage([300, 200])
    const field = base.getForm().createTextField('name')
    field.addToPage(page)
    const result = await overlayPdfBytes(await base.save(), {
      overlayDocuments: [await pdfWithWidths([100])],
      mode: 'sequential',
      position: 'foreground',
      opacity: 0.6,
    })
    const overlaid = await PDFDocument.load(result)
    expect(overlaid.getPageCount()).toBe(1)
    expect(overlaid.getPage(0).getSize()).toEqual({ width: 300, height: 200 })
    expect(overlaid.getForm().getTextField('name')).toBeDefined()
  })

  it('places overlay content after or before the existing page stream', async () => {
    const base = await PDFDocument.create()
    base.addPage([300, 200]).drawText('base')
    const overlay = await PDFDocument.create()
    overlay.addPage([100, 100]).drawText('overlay')
    const source = await base.save()
    const overlayBytes = await overlay.save()
    const foreground = await decodedPageStreams(
      await overlayPdfBytes(source, {
        overlayDocuments: [overlayBytes],
        mode: 'sequential',
        position: 'foreground',
        opacity: 1,
      }),
    )
    const background = await decodedPageStreams(
      await overlayPdfBytes(source, {
        overlayDocuments: [overlayBytes],
        mode: 'sequential',
        position: 'background',
        opacity: 1,
      }),
    )
    expect(foreground.at(-1)).toContain('EmbeddedPdfPage')
    expect(background[0]).toContain('EmbeddedPdfPage')
  })
})

describe('overlayImagePdfBytes', () => {
  it('maps visual top-left placement across page rotations', () => {
    const options = { position: 'topLeft' as const, widthPercent: 20, margin: 10 }
    expect(pdfImageOverlayPlacement(200, 100, 0, 1, 1, options)).toEqual({
      x: 10,
      y: 50,
      width: 40,
      height: 40,
      rotation: 0,
    })
    expect(pdfImageOverlayPlacement(200, 100, 90, 1, 1, options)).toEqual({
      x: 30,
      y: 10,
      width: 20,
      height: 20,
      rotation: 90,
    })
    expect(pdfImageOverlayPlacement(200, 100, 180, 1, 1, options)).toEqual({
      x: 190,
      y: 50,
      width: 40,
      height: 40,
      rotation: 180,
    })
    expect(pdfImageOverlayPlacement(200, 100, 270, 1, 1, options)).toEqual({
      x: 170,
      y: 90,
      width: 20,
      height: 20,
      rotation: 270,
    })
  })

  it('supports custom visual coordinates and crop-box offsets', () => {
    expect(
      pdfImageOverlayPlacement(
        300,
        200,
        0,
        2,
        1,
        { position: 'custom', widthPercent: 20, margin: 0, x: 25, y: 30 },
        12,
        18,
      ),
    ).toEqual({ x: 37, y: 158, width: 60, height: 30, rotation: 0 })
  })

  it('embeds the image only on selected pages and preserves page geometry', async () => {
    const source = await pdfWithWidths([200, 300])
    const result = await overlayImagePdfBytes(source, {
      image: tinyPngBytes(),
      pageIndexes: [1],
      position: 'bottomRight',
      widthPercent: 25,
      margin: 12,
      opacity: 0.65,
      layer: 'foreground',
    })
    const document = await PDFDocument.load(result)
    expect(document.getPages().map((page) => page.getSize())).toEqual([
      { width: 200, height: 200 },
      { width: 300, height: 200 },
    ])
    expect(countImagesInResources(document, document.getPage(0).node.Resources())).toBe(0)
    expect(countImagesInResources(document, document.getPage(1).node.Resources())).toBe(1)
  })

  it('places the image before or after existing page content', async () => {
    const base = await PDFDocument.create()
    base.addPage([300, 200]).drawText('base')
    const source = await base.save()
    const operation = {
      image: tinyPngBytes(),
      pageIndexes: [0],
      position: 'center' as const,
      widthPercent: 20,
      margin: 0,
      opacity: 1,
    }
    const foreground = await decodedPageStreams(
      await overlayImagePdfBytes(source, { ...operation, layer: 'foreground' }),
    )
    const background = await decodedPageStreams(
      await overlayImagePdfBytes(source, { ...operation, layer: 'background' }),
    )
    expect(foreground.at(-1)).toContain('/Image-')
    expect(background[0]).toContain('/Image-')
  })

  it('validates images, page indexes, placement, opacity, and rotation', async () => {
    const source = await pdfWithWidths([200])
    const valid = {
      image: tinyPngBytes(),
      pageIndexes: [0],
      position: 'center' as const,
      widthPercent: 20,
      margin: 0,
      opacity: 1,
      layer: 'foreground' as const,
    }
    await expect(
      overlayImagePdfBytes(source, { ...valid, image: new Uint8Array([1]) }),
    ).rejects.toThrow('Only PNG and JPEG')
    await expect(overlayImagePdfBytes(source, { ...valid, pageIndexes: [1] })).rejects.toThrow(
      'invalid page',
    )
    await expect(overlayImagePdfBytes(source, { ...valid, widthPercent: 0 })).rejects.toThrow(
      'Image width',
    )
    await expect(overlayImagePdfBytes(source, { ...valid, opacity: 2 })).rejects.toThrow('opacity')
    expect(() => pdfImageOverlayPlacement(200, 100, 45, 1, 1, valid)).toThrow('multiple of 90')
  })

  it('returns the image-overlay adapter suffix', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([200]), {
      kind: 'overlayImage',
      image: tinyPngBytes(),
      pageIndexes: [0],
      position: 'center',
      widthPercent: 20,
      margin: 0,
      opacity: 1,
      layer: 'foreground',
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_image_overlayed.pdf'])
  })
})

describe('imagesToPdfBytes', () => {
  it('creates ordered A4 pages and auto-rotates landscape images', async () => {
    const result = await PDFDocument.load(
      await imagesToPdfBytes(
        [
          { image: tinyPngBytes(), width: 100, height: 200 },
          { image: tinyPngBytes(), width: 300, height: 100 },
        ],
        'maintainAspectRatio',
        true,
      ),
    )

    expect(result.getPageCount()).toBe(2)
    expect(result.getPage(0).getSize()).toEqual({ width: 595.28, height: 841.89 })
    expect(result.getPage(1).getSize()).toEqual({ width: 841.89, height: 595.28 })
  })

  it('uses image dimensions as the PDF page size', async () => {
    const result = await PDFDocument.load(
      await imagesToPdfBytes(
        [{ image: tinyPngBytes(), width: 320, height: 180 }],
        'fitDocumentToImage',
        false,
      ),
    )

    expect(result.getPage(0).getSize()).toEqual({ width: 320, height: 180 })
  })

  it('preserves vector PDF pages without rasterizing them', async () => {
    const vector = await PDFDocument.create()
    const vectorPage = vector.addPage([640, 360])
    vectorPage.drawRectangle({ x: 40, y: 50, width: 220, height: 120 })
    const result = await PDFDocument.load(
      await imagesToPdfBytes(
        [{ kind: 'vectorPdf', pdf: await vector.save({ useObjectStreams: false }) }],
        'fitDocumentToImage',
        false,
      ),
    )

    expect(result.getPage(0).getSize()).toEqual({ width: 640, height: 360 })
    const resources = result.getPage(0).node.lookupMaybe(PDFName.of('Resources'), PDFDict)
    const xObjects = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)
    expect(countImagesInResources(result, resources)).toBe(0)
    expect(
      xObjects?.keys().map((name) => {
        const xObject = result.context.lookup(xObjects.get(name))
        return xObject instanceof PDFRawStream
          ? xObject.dict.lookup(PDFName.of('Subtype'), PDFName).decodeText()
          : undefined
      }),
    ).toContain('Form')
  })

  it('fits and appends vector pages while validating prepared input', async () => {
    const vector = await PDFDocument.create()
    vector.addPage([300, 100]).drawRectangle({ x: 10, y: 10, width: 80, height: 30 })
    const vectorBytes = await vector.save({ useObjectStreams: false })
    const appended = await PDFDocument.load(
      await appendImagesToPdfBytes(
        await pdfWithWidths([200]),
        [{ kind: 'vectorPdf', pdf: vectorBytes }],
        'maintainAspectRatio',
        true,
      ),
    )
    expect(appended.getPages().map((page) => page.getWidth())).toEqual([200, 841.89])

    await expect(
      imagesToPdfBytes([{ kind: 'vectorPdf', pdf: Uint8Array.of(1, 2, 3) }], 'fillPage', false),
    ).rejects.toThrow('valid PDF')
    await expect(
      imagesToPdfBytes(
        [{ kind: 'vectorPdf', pdf: await pdfWithWidths([100, 200]) }],
        'fillPage',
        false,
      ),
    ).rejects.toThrow('exactly one')
  })

  it('appends image pages after the current PDF', async () => {
    const result = await PDFDocument.load(
      await appendImagesToPdfBytes(
        await pdfWithWidths([100, 200]),
        [{ image: tinyPngBytes(), width: 300, height: 100 }],
        'fillPage',
        true,
      ),
    )

    expect(result.getPages().map((page) => page.getWidth())).toEqual([100, 200, 841.89])
  })

  it('validates images and exposes operation suffixes', async () => {
    await expect(imagesToPdfBytes([], 'fillPage', false)).rejects.toThrow('at least one image')
    await expect(
      imagesToPdfBytes(
        [{ image: new Uint8Array([1]), width: 100, height: 100 }],
        'fillPage',
        false,
      ),
    ).rejects.toThrow('PNG or JPEG')

    const source = await pdfWithWidths([100])
    const operation = {
      kind: 'imagesToPdf' as const,
      images: [{ image: tinyPngBytes(), width: 10, height: 10 }],
      fitOption: 'fillPage' as const,
      autoRotate: false,
      appendToCurrent: false,
    }
    const created = await runPdfToolBytes(source, operation)
    const appended = await runPdfToolBytes(source, { ...operation, appendToCurrent: true })
    expect(created[0]?.suffix).toBe('_from_images.pdf')
    expect(appended[0]?.suffix).toBe('_with_images.pdf')
  })
})

describe('CBZ to PDF', () => {
  it('extracts supported images in natural order and ignores metadata', async () => {
    const archive = new JSZip()
    archive.file('page10.png', Uint8Array.of(10))
    archive.file('page2.jpg', Uint8Array.of(2))
    archive.file('chapter/page1.webp', Uint8Array.of(1))
    archive.file('__MACOSX/page0.png', Uint8Array.of(0))
    archive.file('.cover.png', Uint8Array.of(0))
    archive.file('ComicInfo.xml', '<ComicInfo/>')

    const images = await extractCbzImageEntries(await archive.generateAsync({ type: 'uint8array' }))
    expect(images.map((image) => image.name)).toEqual([
      'chapter/page1.webp',
      'page2.jpg',
      'page10.png',
    ])
    expect(images.map((image) => Array.from(image.bytes))).toEqual([[1], [2], [10]])
  })

  it('rejects invalid, empty, and unsafe comic archives', async () => {
    await expect(extractCbzImageEntries(new Uint8Array([1, 2, 3]))).rejects.toThrow(
      'invalid or damaged',
    )

    const noImages = new JSZip()
    noImages.file('ComicInfo.xml', '<ComicInfo/>')
    await expect(
      extractCbzImageEntries(await noImages.generateAsync({ type: 'uint8array' })),
    ).rejects.toThrow('no supported images')

    const unsafe = new JSZip()
    unsafe.file('../page1.png', tinyPngBytes())
    await expect(
      extractCbzImageEntries(await unsafe.generateAsync({ type: 'uint8array' })),
    ).rejects.toThrow('unsafe file path')
  })

  it('creates an independently named PDF from prepared comic pages', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'cbzToPdf',
      images: [
        { image: tinyPngBytes(), width: 120, height: 180 },
        { image: tinyPngBytes(), width: 180, height: 120 },
      ],
      fitOption: 'fitDocumentToImage',
      autoRotate: false,
      baseName: '../My Comic.cbz',
    })

    expect(output?.fileName).toBe('My Comic_converted.pdf')
    expect(await pageSizes(output!.bytes)).toEqual([
      [120, 180],
      [180, 120],
    ])
    expect(cbzPdfOutputFileName('..cbz')).toBe('comic_converted.pdf')
    expect(cbzPdfOutputFileName('../My Comic.cbr')).toBe('My Comic_converted.pdf')
  })
})

describe('PDF content page filtering', () => {
  it('keeps matched pages or removes them while preserving document order', () => {
    expect(contentFilterOutputPageIndexes(5, [1, 2, 3], [1, 3], 'keep')).toEqual([1, 3])
    expect(contentFilterOutputPageIndexes(5, [1, 2, 3], [1, 3], 'remove')).toEqual([0, 2, 4])
  })

  it('rejects mismatched analysis and empty output', () => {
    expect(() => contentFilterOutputPageIndexes(3, [0], [1], 'keep')).toThrow('scanned page')
    expect(() => contentFilterOutputPageIndexes(3, [0], [], 'keep')).toThrow('No pages matched')
    expect(() => contentFilterOutputPageIndexes(2, [0, 1], [0, 1], 'remove')).toThrow(
      'remove every page',
    )
  })

  it('produces filtered PDFs with stable suffixes', async () => {
    const source = await pdfWithWidths([100, 200, 300])
    const kept = await runPdfToolBytes(source, {
      kind: 'filterPages',
      criterion: 'text',
      action: 'keep',
      pageIndexes: [0, 1, 2],
      text: 'needle',
      caseSensitive: false,
      wholeWord: false,
      matchedPageIndexes: [1],
    })
    const removed = await runPdfToolBytes(source, {
      kind: 'filterPages',
      criterion: 'image',
      action: 'remove',
      pageIndexes: [0, 1, 2],
      caseSensitive: false,
      wholeWord: false,
      matchedPageIndexes: [0, 2],
    })
    expect(kept[0]!.suffix).toBe('_filtered_pages.pdf')
    expect(await pageWidths(kept[0]!.bytes)).toEqual([200])
    expect(removed[0]!.suffix).toBe('_content_removed.pdf')
    expect(await pageWidths(removed[0]!.bytes)).toEqual([200])
  })

  it('requires renderer analysis and a text query', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      runPdfToolBytes(source, {
        kind: 'filterPages',
        criterion: 'image',
        action: 'keep',
        pageIndexes: [0],
        caseSensitive: false,
        wholeWord: false,
      }),
    ).rejects.toThrow('analysis is required')
    await expect(
      runPdfToolBytes(source, {
        kind: 'filterPages',
        criterion: 'text',
        action: 'keep',
        pageIndexes: [0],
        text: ' ',
        caseSensitive: false,
        wholeWord: false,
        matchedPageIndexes: [0],
      }),
    ).rejects.toThrow('Enter text')
  })

  it('matches page size, visual orientation, and normalized rotation locally', async () => {
    const document = await PDFDocument.create()
    document.addPage([595.28, 841.89])
    document.addPage([792, 612])
    document.addPage([595.28, 841.89]).setRotation(degrees(90))
    document.addPage([300, 300])

    expect(
      geometricFilterMatchedPageIndexes(document, {
        criterion: 'pageSize',
        pageIndexes: [0, 1, 2, 3],
        pageSize: 'A4',
        comparator: 'equal',
      }),
    ).toEqual([0, 2])
    expect(
      geometricFilterMatchedPageIndexes(document, {
        criterion: 'orientation',
        pageIndexes: [0, 1, 2, 3],
        orientation: 'landscape',
      }),
    ).toEqual([1, 2])
    expect(
      geometricFilterMatchedPageIndexes(document, {
        criterion: 'rotation',
        pageIndexes: [0, 1, 2, 3],
        rotation: 90,
      }),
    ).toEqual([2])
  })

  it('runs geometric page filters without renderer analysis', async () => {
    const document = await PDFDocument.create()
    document.addPage([595.28, 841.89])
    document.addPage([792, 612])
    document.addPage([595.28, 841.89]).setRotation(degrees(90))
    const source = await document.save()
    const output = await runPdfToolBytes(source, {
      kind: 'filterPages',
      criterion: 'orientation',
      action: 'keep',
      pageIndexes: [0, 1, 2],
      orientation: 'landscape',
      caseSensitive: false,
      wholeWord: false,
    })
    expect(await pageSizes(output[0]!.bytes)).toEqual([
      [792, 612],
      [595.28, 841.89],
    ])
  })

  it('validates geometric filter settings', async () => {
    const source = await pdfWithWidths([200])
    await expect(
      runPdfToolBytes(source, {
        kind: 'filterPages',
        criterion: 'pageSize',
        action: 'keep',
        pageIndexes: [0],
        caseSensitive: false,
        wholeWord: false,
      }),
    ).rejects.toThrow('page size')
  })
})

describe('PDF document filtering', () => {
  async function filteredPdf(
    pageSizes: Array<readonly [number, number]>,
    rotation = 0,
  ): Promise<Uint8Array> {
    const document = await PDFDocument.create()
    for (const size of pageSizes) document.addPage([size[0], size[1]])
    document.getPage(0).setRotation(degrees(rotation))
    return document.save({ useObjectStreams: false })
  }

  it('matches page count, file size, first-page size, and normalized rotation', async () => {
    const source = await filteredPdf(
      [
        [595.28, 841.89],
        [300, 400],
      ],
      450,
    )
    expect(
      await documentMatchesFilterBytes(source, {
        criterion: 'pageCount',
        comparator: 'equal',
        pageCount: 2,
      }),
    ).toBe(true)
    expect(
      await documentMatchesFilterBytes(source, {
        criterion: 'fileSize',
        comparator: 'greater',
        fileSizeBytes: source.length - 1,
      }),
    ).toBe(true)
    expect(
      await documentMatchesFilterBytes(source, {
        criterion: 'pageSize',
        comparator: 'equal',
        pageSize: 'A4',
      }),
    ).toBe(true)
    expect(
      await documentMatchesFilterBytes(source, {
        criterion: 'rotation',
        comparator: 'equal',
        rotation: 90,
      }),
    ).toBe(true)
  })

  it('exports matching documents unchanged and assigns unique safe names', async () => {
    const current = await filteredPdf([[595.28, 841.89]])
    const matching = await filteredPdf([
      [300, 400],
      [300, 400],
    ])
    const rejected = await filteredPdf([
      [200, 200],
      [200, 200],
      [200, 200],
    ])
    const output = await runPdfToolBytes(current, {
      kind: 'filterDocuments',
      currentFileName: '../Report.pdf',
      documents: [
        { fileName: 'Report.pdf', bytes: matching },
        { fileName: 'Rejected.pdf', bytes: rejected },
      ],
      criterion: 'pageCount',
      comparator: 'less',
      caseSensitive: false,
      wholeWord: false,
      pageCount: 3,
    })
    expect(output.map(({ fileName }) => fileName)).toEqual(['Report.pdf', 'Report (2).pdf'])
    expect(output[0]!.bytes).toEqual(current)
    expect(output[1]!.bytes).toEqual(matching)
  })

  it('rejects invalid rules and an empty match set', async () => {
    const source = await filteredPdf([[300, 400]])
    await expect(
      documentMatchesFilterBytes(source, {
        criterion: 'pageCount',
        comparator: 'equal',
        pageCount: 0,
      }),
    ).rejects.toThrow('page count')
    await expect(
      runPdfToolBytes(source, {
        kind: 'filterDocuments',
        currentFileName: 'One.pdf',
        documents: [],
        criterion: 'pageCount',
        comparator: 'greater',
        caseSensitive: false,
        wholeWord: false,
        pageCount: 2,
      }),
    ).rejects.toThrow('No PDF documents matched')
  })

  it('filters analyzed document text and image matches without rewriting bytes', async () => {
    const current = await filteredPdf([[300, 400]])
    const matching = await filteredPdf([[400, 500]])
    const output = await runPdfToolBytes(current, {
      kind: 'filterDocuments',
      currentFileName: 'Current.pdf',
      currentContentMatched: false,
      documents: [{ fileName: 'Matching.pdf', bytes: matching, contentMatched: true }],
      criterion: 'text',
      comparator: 'equal',
      text: 'launch',
      caseSensitive: false,
      wholeWord: true,
    })
    expect(output).toHaveLength(1)
    expect(output[0]!.fileName).toBe('Matching.pdf')
    expect(output[0]!.bytes).toEqual(matching)
  })

  it('requires a text query and complete renderer analysis for content filters', async () => {
    const source = await filteredPdf([[300, 400]])
    await expect(
      runPdfToolBytes(source, {
        kind: 'filterDocuments',
        currentFileName: 'Current.pdf',
        currentContentMatched: true,
        documents: [],
        criterion: 'text',
        comparator: 'equal',
        text: ' ',
        caseSensitive: false,
        wholeWord: false,
      }),
    ).rejects.toThrow('Enter text')
    await expect(
      runPdfToolBytes(source, {
        kind: 'filterDocuments',
        currentFileName: 'Current.pdf',
        documents: [],
        criterion: 'image',
        comparator: 'equal',
        caseSensitive: false,
        wholeWord: false,
      }),
    ).rejects.toThrow('analysis is required')
  })
})

describe('PDF attachments', () => {
  it('adds, lists, replaces, renames, and deletes embedded files', async () => {
    const source = await pdfWithWidths([300])
    const added = await addPdfAttachmentsBytes(source, [
      {
        name: '../notes.txt',
        bytes: new TextEncoder().encode('first'),
        mimeType: 'text/plain',
      },
      {
        name: 'data.json',
        bytes: new TextEncoder().encode('{"ok":true}'),
        mimeType: 'application/json',
      },
    ])
    expect(await listPdfAttachmentsBytes(added)).toEqual([
      expect.objectContaining({ name: 'notes.txt', size: 5, mimeType: 'text/plain' }),
      expect.objectContaining({ name: 'data.json', size: 11, mimeType: 'application/json' }),
    ])

    const replaced = await addPdfAttachmentsBytes(added, [
      {
        name: 'notes.txt',
        bytes: new TextEncoder().encode('replacement'),
        mimeType: 'text/plain',
      },
    ])
    expect(
      (await listPdfAttachmentsBytes(replaced)).filter((item) => item.name === 'notes.txt'),
    ).toEqual([expect.objectContaining({ size: 11 })])

    const renamed = await renamePdfAttachmentBytes(replaced, 'data.json', 'result.json')
    expect((await listPdfAttachmentsBytes(renamed)).map((item) => item.name).sort()).toEqual([
      'notes.txt',
      'result.json',
    ])

    const deleted = await deletePdfAttachmentBytes(renamed, 'notes.txt')
    expect((await listPdfAttachmentsBytes(deleted)).map((item) => item.name)).toEqual([
      'result.json',
    ])
  })

  it('extracts embedded files into a ZIP archive', async () => {
    const source = await addPdfAttachmentsBytes(await pdfWithWidths([300]), [
      { name: 'alpha.txt', bytes: new TextEncoder().encode('alpha'), mimeType: 'text/plain' },
      { name: 'beta.txt', bytes: new TextEncoder().encode('beta'), mimeType: 'text/plain' },
    ])
    const archive = await JSZip.loadAsync(await extractPdfAttachmentsZipBytes(source))
    expect(await archive.file('alpha.txt')!.async('text')).toBe('alpha')
    expect(await archive.file('beta.txt')!.async('text')).toBe('beta')
  })
})

describe('rotatePdfPagesBytes', () => {
  it('rotates only the selected pages', async () => {
    const result = await rotatePdfPagesBytes(await pdfWithWidths([100, 200, 300]), [0, 2], 90)
    expect(await pageRotations(result)).toEqual([90, 0, 90])
  })
})

describe('autoRotatePdfPagesBytes', () => {
  it('applies a different detected correction to each page', async () => {
    const source = await PDFDocument.create()
    source.addPage([100, 200]).setRotation(degrees(90))
    source.addPage([200, 300]).setRotation(degrees(180))
    const result = await autoRotatePdfPagesBytes(await source.save(), [
      { pageIndex: 0, angle: 270 },
      { pageIndex: 1, angle: 90 },
    ])
    expect(await pageRotations(result)).toEqual([0, 270])
  })

  it('rejects duplicate and out-of-range page corrections', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      autoRotatePdfPagesBytes(source, [
        { pageIndex: 0, angle: 90 },
        { pageIndex: 0, angle: 180 },
      ]),
    ).rejects.toThrow('unique')
    await expect(autoRotatePdfPagesBytes(source, [{ pageIndex: 1, angle: 90 }])).rejects.toThrow(
      'out of range',
    )
  })
})

describe('removePdfPagesBytes', () => {
  it('removes selected pages while preserving the remaining order', async () => {
    const result = await removePdfPagesBytes(await pdfWithWidths([100, 200, 300]), [1])
    expect(await pageWidths(result)).toEqual([100, 300])
  })

  it('does not allow every page to be removed', async () => {
    await expect(removePdfPagesBytes(await pdfWithWidths([100]), [0])).rejects.toThrow(
      'At least one page must remain',
    )
  })
})

describe('rearrangePageIndexes', () => {
  it('matches the Stirling reverse, odd-even, and duplex presets', () => {
    expect(rearrangePageIndexes(5, { mode: 'reverse' })).toEqual([4, 3, 2, 1, 0])
    expect(rearrangePageIndexes(5, { mode: 'oddEven' })).toEqual([0, 2, 4, 1, 3])
    expect(rearrangePageIndexes(5, { mode: 'duplex' })).toEqual([0, 4, 1, 3, 2])
  })

  it('supports edge removal and page duplication presets', () => {
    expect(rearrangePageIndexes(4, { mode: 'removeFirst' })).toEqual([1, 2, 3])
    expect(rearrangePageIndexes(4, { mode: 'removeLast' })).toEqual([0, 1, 2])
    expect(rearrangePageIndexes(4, { mode: 'removeFirstAndLast' })).toEqual([1, 2])
    expect(rearrangePageIndexes(3, { mode: 'duplicate', duplicateCount: 2 })).toEqual([
      0, 0, 1, 1, 2, 2,
    ])
  })
})

describe('rearrangePdfPagesBytes', () => {
  it('supports custom page order with duplicates', async () => {
    const result = await rearrangePdfPagesBytes(await pdfWithWidths([100, 200, 300]), {
      mode: 'custom',
      customOrder: [2, 0, 2],
    })
    expect(await pageWidths(result)).toEqual([300, 100, 300])
  })
})

describe('processPdfFormBytes', () => {
  it('removes the read-only flag from form fields', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const field = document.getForm().createTextField('name')
    field.addToPage(page)
    field.enableReadOnly()

    const result = await processPdfFormBytes(await document.save(), 'unlock')
    const unlocked = await PDFDocument.load(result)
    expect(unlocked.getForm().getTextField('name').isReadOnly()).toBe(false)
  })

  it('removes field locks and unlocks a single XFA stream without dropping it', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const field = document.getForm().createTextField('name')
    field.addToPage(page)
    field.enableReadOnly()
    field.acroField.dict.set(PDFName.of('Lock'), document.context.obj({ Action: 'All' }))
    const acroForm = document.catalog.getAcroForm()!
    const xfa = document.context.register(
      document.context.flateStream(
        '<?xml version="1.0" encoding="UTF-8"?><field access = "readOnly"><child access=\'readOnly\'/></field>',
      ),
    )
    acroForm.dict.set(PDFName.of('XFA'), xfa)

    const result = await processPdfFormBytes(
      await document.save({ useObjectStreams: false, updateFieldAppearances: false }),
      'unlock',
    )
    const unlocked = await PDFDocument.load(result, { updateMetadata: false })
    const unlockedAcroForm = unlocked.catalog.getAcroForm()!
    const unlockedXfa = unlocked.context.lookup(unlockedAcroForm.dict.get(PDFName.of('XFA')))
    expect(unlockedXfa).toBeInstanceOf(PDFRawStream)
    if (!(unlockedXfa instanceof PDFRawStream)) throw new Error('Expected an XFA stream')
    const unlockedXml = decodeUtf8(decodePDFRawStream(unlockedXfa).decode())
    expect(unlockedXml).toContain('access="open"')
    expect(unlockedXml).toContain("access='open'")
    expect(unlockedXml).not.toContain('readOnly')

    const unlockedField = unlocked.getForm().getTextField('name')
    expect(unlockedField.isReadOnly()).toBe(false)
    expect(unlockedField.acroField.dict.has(PDFName.of('Lock'))).toBe(false)
  })

  it('unlocks UTF-16 XFA packet arrays and preserves packet names', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    document.getForm().createTextField('name').addToPage(page)
    const acroForm = document.catalog.getAcroForm()!
    const utf16Xml = '<?xml version="1.0" encoding="UTF-16"?><field access="readOnly" />'
    const utf16Bytes = new Uint8Array(utf16Xml.length * 2 + 2)
    utf16Bytes.set([0xff, 0xfe])
    for (let index = 0; index < utf16Xml.length; index++) {
      const code = utf16Xml.charCodeAt(index)
      utf16Bytes[index * 2 + 2] = code & 0xff
      utf16Bytes[index * 2 + 3] = code >> 8
    }
    const template = document.context.register(document.context.flateStream(utf16Bytes))
    const datasets = document.context.register(
      document.context.flateStream('<data access="readOnly" />'),
    )
    acroForm.dict.set(
      PDFName.of('XFA'),
      document.context.obj([
        PDFString.of('template'),
        template,
        PDFString.of('datasets'),
        datasets,
      ]),
    )

    const result = await processPdfFormBytes(
      await document.save({ useObjectStreams: false, updateFieldAppearances: false }),
      'unlock',
    )
    const unlocked = await PDFDocument.load(result, { updateMetadata: false })
    const packets = unlocked.catalog.getAcroForm()!.dict.lookup(PDFName.of('XFA'), PDFArray)
    expect(packets.lookup(0, PDFString).decodeText()).toBe('template')
    expect(packets.lookup(2, PDFString).decodeText()).toBe('datasets')
    const templateBytes = decodePDFRawStream(packets.lookup(1, PDFRawStream)).decode()
    expect(templateBytes.slice(0, 2)).toEqual(new Uint8Array([0xff, 0xfe]))
    expect(new TextDecoder('utf-16le').decode(templateBytes)).toContain('access="open"')
    expect(decodeUtf8(decodePDFRawStream(packets.lookup(3, PDFRawStream)).decode())).toContain(
      'access="open"',
    )
  })

  it('flattens form fields into page content', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const field = document.getForm().createTextField('name')
    field.addToPage(page)
    field.setText('GenOffice')
    field.updateAppearances(await document.embedFont('Helvetica'))

    const result = await processPdfFormBytes(await document.save(), 'flatten')
    const flattened = await PDFDocument.load(result)
    expect(flattened.getForm().getFields()).toHaveLength(0)
    expect(flattened.getPage(0).node.has(PDFName.of('Annots'))).toBe(false)
  })

  it('deletes selected form fields and their widgets while preserving other fields and XFA', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const form = document.getForm()
    form.createTextField('remove.me').addToPage(page, { x: 20, y: 120, width: 120, height: 20 })
    form.createTextField('keep.me').addToPage(page, { x: 20, y: 80, width: 120, height: 20 })
    const acroForm = document.catalog.getAcroForm()!
    const xfa = document.context.register(document.context.flateStream('<template />'))
    acroForm.dict.set(PDFName.of('XFA'), xfa)
    const source = await document.save({ useObjectStreams: false, updateFieldAppearances: false })

    const result = await processPdfFormBytes(source, 'delete', [], ['remove.me', 'missing'])
    const updated = await PDFDocument.load(result, { updateMetadata: false })
    expect(updated.catalog.getAcroForm()!.dict.has(PDFName.of('XFA'))).toBe(true)
    expect(
      updated
        .getForm()
        .getFields()
        .map((field) => field.getName()),
    ).toEqual(['keep.me'])
    const annotations = updated.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    expect(annotations.size()).toBe(1)
  })

  it('rejects deleting no fields or only unknown fields', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    document.getForm().createTextField('name').addToPage(page)
    const source = await document.save({ updateFieldAppearances: false })
    await expect(processPdfFormBytes(source, 'delete')).rejects.toThrow('At least one')
    await expect(processPdfFormBytes(source, 'delete', [], ['missing'])).rejects.toThrow(
      'No matching',
    )
  })

  it('renames fields in place and updates common properties while preserving XFA', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const form = document.getForm()
    const name = form.createTextField('profile.name')
    name.addToPage(page, { x: 20, y: 120, width: 120, height: 20 })
    const code = form.createTextField('profile.code')
    code.addToPage(page, { x: 20, y: 80, width: 120, height: 20 })
    code.enableReadOnly()
    code.enableRequired()
    const acroForm = document.catalog.getAcroForm()!
    const xfa = document.context.register(document.context.flateStream('<template />'))
    acroForm.dict.set(PDFName.of('XFA'), xfa)
    const source = await document.save({ useObjectStreams: false, updateFieldAppearances: false })

    const result = await processPdfFormBytes(
      source,
      'modify',
      [],
      [],
      [
        { name: 'profile.name', newName: 'displayName', readOnly: true, required: true },
        { name: 'profile.code', newName: 'profile.reference', readOnly: false, required: false },
      ],
    )
    const updated = await PDFDocument.load(result, { updateMetadata: false })
    expect(updated.catalog.getAcroForm()!.dict.has(PDFName.of('XFA'))).toBe(true)
    const updatedForm = updated.getForm()
    expect(updatedForm.getFields().map((field) => field.getName())).toEqual([
      'profile.displayName',
      'profile.reference',
    ])
    expect(updatedForm.getTextField('profile.displayName').isReadOnly()).toBe(true)
    expect(updatedForm.getTextField('profile.displayName').isRequired()).toBe(true)
    expect(updatedForm.getTextField('profile.reference').isReadOnly()).toBe(false)
    expect(updatedForm.getTextField('profile.reference').isRequired()).toBe(false)
    expect(updated.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).size()).toBe(2)
  })

  it('rejects invalid, duplicate, unknown, cross-group, and conflicting field modifications', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const form = document.getForm()
    form.createTextField('profile.name').addToPage(page)
    form.createTextField('profile.code').addToPage(page)
    const source = await document.save({ updateFieldAppearances: false })

    await expect(processPdfFormBytes(source, 'modify')).rejects.toThrow('At least one')
    await expect(
      processPdfFormBytes(
        source,
        'modify',
        [],
        [],
        [
          { name: 'profile.name', newName: 'alias' },
          { name: 'profile.name', required: true },
        ],
      ),
    ).rejects.toThrow('Duplicate')
    await expect(
      processPdfFormBytes(source, 'modify', [], [], [{ name: 'profile.missing', required: true }]),
    ).rejects.toThrow('No matching')
    await expect(
      processPdfFormBytes(
        source,
        'modify',
        [],
        [],
        [{ name: 'profile.name', newName: 'other.name' }],
      ),
    ).rejects.toThrow('another group')
    await expect(
      processPdfFormBytes(
        source,
        'modify',
        [],
        [],
        [{ name: 'profile.name', newName: 'profile.code' }],
      ),
    ).rejects.toThrow('already uses')
    await expect(
      processPdfFormBytes(
        source,
        'modify',
        [],
        [],
        [{ name: 'profile.name', newName: 'profile..name' }],
      ),
    ).rejects.toThrow('empty path')
  })

  it('updates labels, choice options, and multiselect flags without recreating fields', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([400, 260])
    const form = document.getForm()
    const region = form.createDropdown('profile.region')
    region.setOptions(['North', 'South', 'West'])
    region.select('South')
    region.acroField.dict.set(PDFName.of('TU'), PDFHexString.fromText('Current region'))
    region.addToPage(page, { x: 20, y: 180, width: 120, height: 24 })
    const interests = form.createOptionList('profile.interests')
    interests.setOptions(['PDF', 'Docs', 'Slides'])
    interests.enableMultiselect()
    interests.select(['PDF', 'Docs'])
    interests.addToPage(page, { x: 20, y: 80, width: 120, height: 70 })
    const source = await document.save({ useObjectStreams: false, updateFieldAppearances: false })

    const result = await processPdfFormBytes(
      source,
      'modify',
      [],
      [],
      [
        {
          name: 'profile.region',
          label: 'Office region',
          options: ['South', 'East', 'South', '  '],
          multiselect: false,
        },
        {
          name: 'profile.interests',
          label: '',
          options: ['Docs', 'Slides'],
          multiselect: false,
        },
      ],
    )
    const updated = await PDFDocument.load(result, { updateMetadata: false })
    const updatedForm = updated.getForm()
    const updatedRegion = updatedForm.getDropdown('profile.region')
    const updatedInterests = updatedForm.getOptionList('profile.interests')
    expect(updatedRegion.getOptions()).toEqual(['South', 'East'])
    expect(updatedRegion.getSelected()).toEqual(['South'])
    expect(updatedRegion.isMultiselect()).toBe(false)
    expect(
      updatedRegion.acroField.dict.lookup(PDFName.of('TU'), PDFString, PDFHexString).decodeText(),
    ).toBe('Office region')
    expect(updatedInterests.getOptions()).toEqual(['Docs', 'Slides'])
    expect(updatedInterests.getSelected()).toEqual(['Docs'])
    expect(updatedInterests.isMultiselect()).toBe(false)
    expect(updatedInterests.acroField.dict.has(PDFName.of('TU'))).toBe(false)
    expect(updated.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).size()).toBe(2)
  })

  it('rejects empty choice options and choice properties on other field types', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const form = document.getForm()
    form.createDropdown('region').addToPage(page)
    form.createTextField('name').addToPage(page)
    const source = await document.save({ updateFieldAppearances: false })

    await expect(
      processPdfFormBytes(source, 'modify', [], [], [{ name: 'region', options: [' ', ''] }]),
    ).rejects.toThrow('at least one option')
    await expect(
      processPdfFormBytes(source, 'modify', [], [], [{ name: 'name', options: ['Alpha'] }]),
    ).rejects.toThrow('choice fields')
  })

  it('creates supported fields on a plain PDF with values, flags, labels, and widgets', async () => {
    const document = await PDFDocument.create()
    document.addPage([500, 500])
    document.addPage([500, 500])
    const source = await document.save({ useObjectStreams: false })

    const result = await processPdfFormBytes(
      source,
      'create',
      [],
      [],
      [],
      [
        {
          name: 'profile.name',
          label: 'Display name',
          type: 'text',
          pageIndex: 0,
          x: 20,
          y: 440,
          width: 180,
          height: 36,
          required: true,
          multiline: true,
          defaultValue: 'GenOffice',
        },
        {
          name: 'profile.accepted',
          type: 'checkbox',
          pageIndex: 0,
          x: 20,
          y: 400,
          width: 20,
          height: 20,
          readOnly: true,
          defaultValue: true,
        },
        {
          name: 'profile.plan',
          type: 'radio',
          pageIndex: 0,
          x: 20,
          y: 350,
          width: 20,
          height: 20,
          options: ['Basic', 'Pro'],
          optionSpacing: 28,
          defaultValue: 'Pro',
        },
        {
          name: 'profile.region',
          type: 'dropdown',
          pageIndex: 1,
          x: 30,
          y: 420,
          width: 140,
          height: 24,
          options: ['North', 'South'],
          defaultValue: 'South',
        },
        {
          name: 'profile.interests',
          type: 'optionList',
          pageIndex: 1,
          x: 30,
          y: 300,
          width: 140,
          height: 90,
          options: ['PDF', 'Docs', 'Slides'],
          multiselect: true,
          defaultValue: ['PDF', 'Slides'],
        },
      ],
    )

    const updated = await PDFDocument.load(result, { updateMetadata: false })
    const form = updated.getForm()
    expect(form.getFields().map((field) => field.getName())).toEqual([
      'profile.name',
      'profile.accepted',
      'profile.plan',
      'profile.region',
      'profile.interests',
    ])
    const name = form.getTextField('profile.name')
    expect(name.getText()).toBe('GenOffice')
    expect(name.isMultiline()).toBe(true)
    expect(name.isRequired()).toBe(true)
    expect(name.acroField.dict.lookup(PDFName.of('TU'), PDFString, PDFHexString).decodeText()).toBe(
      'Display name',
    )
    expect(form.getCheckBox('profile.accepted').isChecked()).toBe(true)
    expect(form.getCheckBox('profile.accepted').isReadOnly()).toBe(true)
    const plan = form.getRadioGroup('profile.plan')
    expect(plan.getOptions()).toEqual(['Basic', 'Pro'])
    expect(plan.getSelected()).toBe('Pro')
    expect(plan.acroField.getWidgets()).toHaveLength(2)
    const radioRectangles = plan.acroField.getWidgets().map((widget) => widget.getRectangle())
    expect(radioRectangles[0]!.y - radioRectangles[1]!.y).toBe(28)
    expect(form.getDropdown('profile.region').getSelected()).toEqual(['South'])
    const interests = form.getOptionList('profile.interests')
    expect(interests.getSelected()).toEqual(['PDF', 'Slides'])
    expect(interests.isMultiselect()).toBe(true)
    expect(updated.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray).size()).toBe(4)
    expect(updated.getPage(1).node.lookup(PDFName.of('Annots'), PDFArray).size()).toBe(2)
  })

  it('preserves XFA while adding a field to an existing form', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    document.getForm().createTextField('existing').addToPage(page)
    const acroForm = document.catalog.getAcroForm()!
    const xfa = document.context.register(document.context.flateStream('<template />'))
    acroForm.dict.set(PDFName.of('XFA'), xfa)
    const source = await document.save({ useObjectStreams: false, updateFieldAppearances: false })

    const result = await processPdfFormBytes(
      source,
      'create',
      [],
      [],
      [],
      [
        {
          name: 'created',
          type: 'text',
          pageIndex: 0,
          x: 20,
          y: 40,
          width: 120,
          height: 24,
        },
      ],
    )
    const updated = await PDFDocument.load(result, { updateMetadata: false })
    expect(updated.catalog.getAcroForm()!.dict.has(PDFName.of('XFA'))).toBe(true)
    expect(
      updated
        .getForm()
        .getFields()
        .map((field) => field.getName()),
    ).toEqual(['existing', 'created'])
  })

  it('creates fields with non-Latin defaults without requiring a custom appearance font', async () => {
    const document = await PDFDocument.create()
    document.addPage([300, 200])

    const result = await processPdfFormBytes(
      await document.save(),
      'create',
      [],
      [],
      [],
      [
        {
          name: '姓名',
          type: 'text',
          pageIndex: 0,
          x: 20,
          y: 120,
          width: 120,
          height: 24,
          defaultValue: '艾达',
        },
        {
          name: '城市',
          type: 'dropdown',
          pageIndex: 0,
          x: 20,
          y: 80,
          width: 120,
          height: 24,
          options: ['北京', '上海'],
          defaultValue: '上海',
        },
      ],
    )
    const updated = await PDFDocument.load(result, { updateMetadata: false })
    expect(updated.getForm().getTextField('姓名').getText()).toBe('艾达')
    expect(updated.getForm().getDropdown('城市').getOptions()).toEqual(['北京', '上海'])
    expect(updated.getForm().getDropdown('城市').getSelected()).toEqual(['上海'])
  })

  it('rejects invalid form field creations before changing the document', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    document.getForm().createTextField('existing').addToPage(page)
    const source = await document.save({ updateFieldAppearances: false })
    const valid = {
      name: 'created',
      type: 'text' as const,
      pageIndex: 0,
      x: 20,
      y: 40,
      width: 120,
      height: 24,
    }

    await expect(processPdfFormBytes(source, 'create')).rejects.toThrow('At least one')
    await expect(
      processPdfFormBytes(source, 'create', [], [], [], [{ ...valid, name: 'existing.child' }]),
    ).rejects.toThrow('conflicts')
    await expect(
      processPdfFormBytes(
        source,
        'create',
        [],
        [],
        [],
        [{ ...valid, type: 'signature' as 'text' }],
      ),
    ).rejects.toThrow('Unsupported')
    await expect(
      processPdfFormBytes(source, 'create', [], [], [], [valid, { ...valid, name: 'created' }]),
    ).rejects.toThrow('conflicts')
    await expect(
      processPdfFormBytes(source, 'create', [], [], [], [{ ...valid, pageIndex: 1 }]),
    ).rejects.toThrow('page')
    await expect(
      processPdfFormBytes(source, 'create', [], [], [], [{ ...valid, x: 250 }]),
    ).rejects.toThrow('outside')
    await expect(
      processPdfFormBytes(
        source,
        'create',
        [],
        [],
        [],
        [{ ...valid, type: 'dropdown', options: [] }],
      ),
    ).rejects.toThrow('at least one option')
    await expect(
      processPdfFormBytes(
        source,
        'create',
        [],
        [],
        [],
        [{ ...valid, type: 'radio', options: ['One', 'Two'], defaultValue: 'Three' }],
      ),
    ).rejects.toThrow('must match')
  })

  it('lists and fills supported AcroForm field types while keeping the form editable', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([500, 500])
    const form = document.getForm()
    const name = form.createTextField('profile.name')
    name.acroField.dict.set(PDFName.of('TU'), PDFHexString.fromText('Display name'))
    name.enableMultiline()
    name.enableRequired()
    name.addToPage(page, { x: 20, y: 430, width: 200, height: 40 })
    const accepted = form.createCheckBox('profile.accepted')
    accepted.addToPage(page, { x: 20, y: 390, width: 20, height: 20 })
    const plan = form.createRadioGroup('profile.plan')
    plan.addOptionToPage('Basic', page, { x: 20, y: 350, width: 20, height: 20 })
    plan.addOptionToPage('Pro', page, { x: 60, y: 350, width: 20, height: 20 })
    const region = form.createDropdown('profile.region')
    region.setOptions(['North', 'South'])
    region.addToPage(page, { x: 20, y: 300, width: 100, height: 20 })
    const interests = form.createOptionList('profile.interests')
    interests.addOptions(['PDF', 'Docs', 'Slides'])
    interests.enableMultiselect()
    interests.addToPage(page, { x: 20, y: 220, width: 100, height: 60 })
    const submit = form.createButton('profile.submit')
    submit.addToPage('Submit', page, { x: 20, y: 170, width: 80, height: 24 })
    region.enableReadOnly()

    const source = await document.save({ updateFieldAppearances: false })
    expect(await listPdfFormFieldsBytes(source)).toEqual([
      expect.objectContaining({
        name: 'profile.name',
        label: 'Display name',
        type: 'text',
        required: true,
        multiline: true,
        value: '',
      }),
      expect.objectContaining({ name: 'profile.accepted', type: 'checkbox', value: false }),
      expect.objectContaining({
        name: 'profile.plan',
        type: 'radio',
        options: ['Basic', 'Pro'],
      }),
      expect.objectContaining({
        name: 'profile.region',
        type: 'dropdown',
        options: ['North', 'South'],
        readOnly: true,
      }),
      expect.objectContaining({
        name: 'profile.interests',
        type: 'optionList',
        options: ['PDF', 'Docs', 'Slides'],
        multiselect: true,
      }),
      expect.objectContaining({ name: 'profile.submit', type: 'button' }),
    ])

    region.disableReadOnly()
    const fillableSource = await document.save({ updateFieldAppearances: false })
    const filledBytes = await processPdfFormBytes(fillableSource, 'fill', [
      { name: 'profile.name', value: 'Alice 王' },
      { name: 'profile.accepted', value: true },
      { name: 'profile.plan', value: 'Pro' },
      { name: 'profile.region', value: ['South'] },
      { name: 'profile.interests', value: ['PDF', 'Slides'] },
    ])
    const filled = await PDFDocument.load(filledBytes)
    const filledForm = filled.getForm()
    expect(filledForm.getTextField('profile.name').getText()).toBe('Alice 王')
    expect(filledForm.getCheckBox('profile.accepted').isChecked()).toBe(true)
    expect(filledForm.getRadioGroup('profile.plan').getSelected()).toBe('Pro')
    expect(filledForm.getDropdown('profile.region').getSelected()).toEqual(['South'])
    expect(filledForm.getOptionList('profile.interests').getSelected()).toEqual(['PDF', 'Slides'])
    expect(filledForm.getFields()).toHaveLength(6)
    expect(filledForm.acroForm.dict.get(PDFName.of('NeedAppearances'))).toBe(PDFBool.True)
  })

  it('rejects read-only fields, unsupported buttons, duplicate names, and invalid options', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const form = document.getForm()
    const field = form.createTextField('name')
    field.addToPage(page)
    field.enableReadOnly()
    const button = form.createButton('submit')
    button.addToPage('Submit', page)
    const radio = form.createRadioGroup('plan')
    radio.addOptionToPage('Free', page)
    const source = await document.save({ updateFieldAppearances: false })

    await expect(
      processPdfFormBytes(source, 'fill', [{ name: 'name', value: 'A' }]),
    ).rejects.toThrow('name is read-only')
    await expect(
      processPdfFormBytes(source, 'fill', [{ name: 'submit', value: 'click' }]),
    ).rejects.toThrow('submit cannot be filled')
    await expect(
      processPdfFormBytes(source, 'fill', [
        { name: 'plan', value: 'Free' },
        { name: 'plan', value: 'Free' },
      ]),
    ).rejects.toThrow('unique')
    await expect(
      processPdfFormBytes(source, 'fill', [{ name: 'plan', value: 'Enterprise' }]),
    ).rejects.toThrow('unknown option')
  })
})

describe('PDF form data export', () => {
  it('preserves JSON value types and safely escapes CSV and XLSX values', async () => {
    const fields = [
      {
        name: 'customer,name',
        type: 'text' as const,
        readOnly: false,
        required: false,
        value: 'Alice "A"\n王',
      },
      {
        name: 'accepted',
        type: 'checkbox' as const,
        readOnly: false,
        required: false,
        value: true,
      },
      {
        name: 'interests',
        type: 'optionList' as const,
        readOnly: false,
        required: false,
        value: ['PDF', 'Slides'],
      },
      {
        name: 'submit',
        type: 'button' as const,
        readOnly: false,
        required: false,
      },
    ]
    const json = new TextDecoder().decode(pdfFormFieldsJsonBytes(fields))
    expect(JSON.parse(json)).toEqual({
      'customer,name': 'Alice "A"\n王',
      accepted: true,
      interests: ['PDF', 'Slides'],
      submit: null,
    })
    const csvBytes = pdfFormFieldsCsvBytes(fields)
    expect([...csvBytes.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf])
    const csv = new TextDecoder().decode(csvBytes)
    expect(csv.startsWith('"Field Name","Value"\r\n')).toBe(true)
    expect(csv).toContain('"customer,name","Alice ""A""\n王"')
    expect(csv).toContain('"interests","PDF; Slides"')
    expect(csv).toContain('"submit",""')

    const archive = await JSZip.loadAsync(await pdfFormFieldsXlsxBytes(fields))
    const workbook = await archive.file('xl/workbook.xml')!.async('text')
    expect(workbook).toContain('name="Form Fields"')
    const worksheet = await archive.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(worksheet).toContain('Field Name')
    expect(worksheet).toContain('customer,name')
    expect(worksheet).toContain('Alice &quot;A&quot;\n王')
    expect(worksheet).toContain('PDF; Slides')
    expect(worksheet).toContain('<dimension ref="A1:B5"/>')
  })

  it('exports JSON, CSV, and XLSX tool outputs from the current PDF values', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const form = document.getForm()
    const name = form.createTextField('name')
    name.addToPage(page)
    name.setText('GenOffice')
    const accepted = form.createCheckBox('accepted')
    accepted.addToPage(page)
    accepted.check()
    const outputs = await runPdfToolBytes(await document.save({ updateFieldAppearances: false }), {
      kind: 'forms',
      action: 'export',
    })
    expect(outputs.map((output) => [output.suffix, output.mimeType, output.extension])).toEqual([
      ['_form_data.json', 'application/json', '.json'],
      ['_form_data.csv', 'text/csv;charset=utf-8', '.csv'],
      [
        '_form_data.xlsx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        '.xlsx',
      ],
    ])
    expect(JSON.parse(new TextDecoder().decode(outputs[0]!.bytes))).toEqual({
      name: 'GenOffice',
      accepted: true,
    })
    expect(new TextDecoder().decode(outputs[1]!.bytes)).toContain('"accepted","true"')
    const archive = await JSZip.loadAsync(outputs[2]!.bytes)
    const worksheet = await archive.file('xl/worksheets/sheet1.xml')!.async('text')
    expect(worksheet).toContain('GenOffice')
    expect(worksheet).toContain('true')
  })

  it('rejects exporting PDFs without form fields', async () => {
    await expect(
      runPdfToolBytes(await pdfWithWidths([200]), { kind: 'forms', action: 'export' }),
    ).rejects.toThrow('No PDF form fields were found')
  })
})

describe('flattenPdfPagesBytes', () => {
  it('rebuilds every crop-sized page as a single image without interactive content', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Flatten me')
    const firstPage = source.addPage([400, 300])
    firstPage.setCropBox(10, 20, 250, 150)
    firstPage.setRotation(degrees(90))
    firstPage.drawText('Searchable source text')
    source.getForm().createTextField('name').addToPage(firstPage)
    source.addPage([320, 240]).drawText('Second page')

    const result = await flattenPdfPagesBytes(await source.save(), [tinyPngBytes(), tinyPngBytes()])
    const flattened = await PDFDocument.load(result)
    expect(flattened.getTitle()).toBe('Flatten me')
    expect(flattened.getPageCount()).toBe(2)
    expect(flattened.getPage(0).getSize()).toEqual({ width: 150, height: 250 })
    expect(flattened.getPage(0).getRotation().angle).toBe(0)
    expect(flattened.getPage(1).getSize()).toEqual({ width: 320, height: 240 })
    expect(flattened.catalog.has(PDFName.of('AcroForm'))).toBe(false)
    for (const page of flattened.getPages()) {
      expect(page.node.has(PDFName.of('Annots'))).toBe(false)
      expect(countImagesInResources(flattened, page.node.Resources())).toBe(1)
    }
  })

  it('rejects missing and empty rendered pages', async () => {
    const source = await pdfWithWidths([100, 200])
    await expect(flattenPdfPagesBytes(source, [tinyPngBytes()])).rejects.toThrow('every PDF page')
    await expect(flattenPdfPagesBytes(source, [tinyPngBytes(), new Uint8Array()])).rejects.toThrow(
      'image is empty',
    )
  })
})

describe('pdfToPdfaBytes', () => {
  it('creates a rasterized PDF/A-2b archive with XMP and an sRGB output intent', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Archive & Review')
    source.setAuthor('GenOffice')
    source.addPage([240, 320])
    const field = source.getForm().createTextField('approval')
    field.addToPage(source.getPage(0), { x: 20, y: 20, width: 100, height: 24 })
    source.catalog.set(PDFName.of('OpenAction'), source.context.obj({ S: 'JavaScript' }))

    const result = await pdfToPdfaBytes(await source.save(), {
      format: 'PDF/A-2b',
      archiveMode: 'raster',
      renderDpi: 150,
      imageQuality: 92,
      pageImages: [tinyPngBytes()],
    })
    const archive = await PDFDocument.load(result, { updateMetadata: false })
    expect(archive.getPageCount()).toBe(1)
    expect(archive.getTitle()).toBe('Archive & Review')
    expect(archive.catalog.has(PDFName.of('AcroForm'))).toBe(false)
    expect(archive.catalog.has(PDFName.of('OpenAction'))).toBe(false)
    expect(archive.catalog.lookup(PDFName.of('OutputIntents'), PDFArray).size()).toBe(1)
    const metadata = archive.context.lookup(archive.catalog.get(PDFName.of('Metadata')))
    if (!(metadata instanceof PDFRawStream)) throw new Error('Expected PDF/A XMP metadata stream')
    const xmp = decodeUtf8(decodePDFRawStream(metadata).decode())
    expect(xmp).toContain('pdfaid:part="2"')
    expect(xmp).toContain('pdfaid:conformance="B"')
    expect(xmp).toContain('Archive &amp; Review')

    const preflight = await preflightPdfBytes(result)
    expect(preflight.standards).toEqual([
      { family: 'PDF/A', part: '2', conformance: 'B', label: 'PDF/A-2b' },
    ])
    expect(preflight.features.outputIntentCount).toBe(1)
    expect(preflight.findings.map((finding) => finding.code)).not.toContain(
      'pdfaMissingOutputIntent',
    )
  })

  it('preserves selectable text and vector content when every used font is embedded', async () => {
    const source = await PDFDocument.create()
    source.registerFontkit(pdfLibFontkit)
    const fonts = await pdfJsonImportFonts()
    const font = await source.embedFont(fonts.regular, { subset: true })
    const page = source.addPage([320, 240])
    page.drawText('Archival vector text', { x: 28, y: 170, size: 20, font })
    page.drawRectangle({ x: 28, y: 110, width: 140, height: 32 })
    source.getForm().createTextField('approval').addToPage(page, {
      x: 28,
      y: 40,
      width: 120,
      height: 24,
    })
    source.catalog.set(PDFName.of('OpenAction'), source.context.obj({ S: 'JavaScript' }))
    const sourceBytes = await source.save()

    await expect(pdfaPreservationReportBytes(sourceBytes)).resolves.toEqual({
      eligible: true,
      fontCount: 1,
      embeddedFontCount: 1,
      unembeddedFonts: [],
    })
    const result = await pdfToPdfaBytes(sourceBytes, {
      format: 'PDF/A-2b',
      archiveMode: 'auto',
      renderDpi: 150,
      imageQuality: 92,
    })
    const archive = await PDFDocument.load(result, { updateMetadata: false })
    expect(countImagesInResources(archive, archive.getPage(0).node.Resources())).toBe(0)
    expect(archive.catalog.has(PDFName.of('AcroForm'))).toBe(false)
    expect(archive.catalog.has(PDFName.of('OpenAction'))).toBe(false)
    expect(archive.getPage(0).node.has(PDFName.of('Annots'))).toBe(false)
    await expect(analyzePdfFontsBytes(result)).resolves.toEqual(
      expect.objectContaining({ fontCount: 1, embeddedCount: 1 }),
    )
  })

  it('reports unembedded fonts and automatically requires the image fallback', async () => {
    const source = await PDFDocument.create()
    const font = await source.embedFont('Helvetica')
    source.addPage([240, 320]).drawText('Standard font text', { font })
    const sourceBytes = await source.save()

    await expect(pdfaPreservationReportBytes(sourceBytes)).resolves.toEqual({
      eligible: false,
      fontCount: 1,
      embeddedFontCount: 0,
      unembeddedFonts: ['Helvetica'],
    })
    await expect(
      runPdfToolBytes(sourceBytes, {
        kind: 'pdfToPdfa',
        format: 'PDF/A-2b',
        archiveMode: 'auto',
        renderDpi: 150,
        imageQuality: 92,
      }),
    ).rejects.toThrow(/image fallback.*Helvetica/i)

    const [output] = await runPdfToolBytes(sourceBytes, {
      kind: 'pdfToPdfa',
      format: 'PDF/A-2b',
      archiveMode: 'auto',
      renderDpi: 150,
      imageQuality: 92,
      pageImages: [tinyPngBytes()],
    })
    const archive = await PDFDocument.load(output!.bytes)
    expect(countImagesInResources(archive, archive.getPage(0).node.Resources())).toBe(1)
    await expect(analyzePdfFontsBytes(output!.bytes)).resolves.toEqual(
      expect.objectContaining({ fontCount: 0, embeddedCount: 0 }),
    )
  })

  it('validates archive settings and prepared page images', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      pdfToPdfaBytes(source, {
        format: 'PDF/A-2b',
        archiveMode: 'raster',
        renderDpi: 71,
        imageQuality: 92,
        pageImages: [tinyPngBytes()],
      }),
    ).rejects.toThrow(/72 to 600/i)
  })

  it('uses a stable PDF/A output suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'pdfToPdfa',
      format: 'PDF/A-2b',
      archiveMode: 'auto',
      renderDpi: 150,
      imageQuality: 92,
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_PDFA-2b.pdf'])
  })
})

describe('scannerEffectPdfPagesBytes', () => {
  it('rebuilds rotated crop-sized pages from scanner images', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Scan me')
    const firstPage = source.addPage([400, 300])
    firstPage.setCropBox(10, 20, 250, 150)
    firstPage.setRotation(degrees(90))
    source.addPage([320, 240])

    const result = await scannerEffectPdfPagesBytes(await source.save(), [
      tinyPngBytes(),
      tinyPngBytes(),
    ])
    const scanned = await PDFDocument.load(result)
    expect(scanned.getTitle()).toBe('Scan me')
    expect(scanned.getPageCount()).toBe(2)
    expect(scanned.getPage(0).getSize()).toEqual({ width: 150, height: 250 })
    expect(scanned.getPage(0).getRotation().angle).toBe(0)
    expect(scanned.getPage(1).getSize()).toEqual({ width: 320, height: 240 })
    for (const page of scanned.getPages()) {
      expect(countImagesInResources(scanned, page.node.Resources())).toBe(1)
    }
  })

  it('rejects missing and empty scanner page images', async () => {
    const source = await pdfWithWidths([100, 200])
    await expect(scannerEffectPdfPagesBytes(source, [tinyPngBytes()])).rejects.toThrow(
      'every PDF page',
    )
    await expect(
      scannerEffectPdfPagesBytes(source, [tinyPngBytes(), new Uint8Array()]),
    ).rejects.toThrow('image is empty')
  })
})

describe('deskewPdfPagesBytes', () => {
  it('rasterizes corrected pages while preserving unchanged pages as PDF content', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Deskew me')
    source.addPage([200, 300]).drawText('First vector page')
    source.addPage([320, 240]).drawText('Second vector page')
    const result = await deskewPdfPagesBytes(await source.save(), [1], 8, [
      { pageIndex: 1, angle: -4.2, image: tinyPngBytes() },
    ])
    const deskewed = await PDFDocument.load(result)
    expect(deskewed.getTitle()).toBe('Deskew me')
    expect(deskewed.getPageCount()).toBe(2)
    expect(deskewed.getPage(0).getSize()).toEqual({ width: 200, height: 300 })
    expect(countImagesInResources(deskewed, deskewed.getPage(0).node.Resources())).toBe(0)
    expect(countImagesInResources(deskewed, deskewed.getPage(1).node.Resources())).toBe(1)
  })

  it('returns the original bytes when no selected page needs correction', async () => {
    const source = await pdfWithWidths([100, 200])
    expect(Array.from(await deskewPdfPagesBytes(source, [0, 1], 8, []))).toEqual(Array.from(source))
  })

  it('validates prepared page indexes, angles, and images', async () => {
    const source = await pdfWithWidths([100, 200])
    await expect(
      deskewPdfPagesBytes(source, [0], 8, [{ pageIndex: 1, angle: 2, image: tinyPngBytes() }]),
    ).rejects.toThrow('outside the selected pages')
    await expect(
      deskewPdfPagesBytes(source, [0], 8, [{ pageIndex: 0, angle: 9, image: tinyPngBytes() }]),
    ).rejects.toThrow('angle is invalid')
    await expect(
      deskewPdfPagesBytes(source, [0], 8, [{ pageIndex: 0, angle: 2, image: new Uint8Array() }]),
    ).rejects.toThrow('image is empty')
  })
})

describe('comparisonPdfPagesBytes', () => {
  it('builds image-only comparison pages with source metadata and requested sizes', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Compare me')
    source.addPage([100, 200])

    const result = await comparisonPdfPagesBytes(await source.save(), [
      { image: tinyPngBytes(), width: 320, height: 240 },
      { image: tinyPngBytes(), width: 500, height: 700 },
    ])
    const compared = await PDFDocument.load(result)
    expect(compared.getTitle()).toBe('Compare me')
    expect(await pageSizes(result)).toEqual([
      [320, 240],
      [500, 700],
    ])
    for (const page of compared.getPages()) {
      expect(countImagesInResources(compared, page.node.Resources())).toBe(1)
    }
  })

  it('rejects missing, empty, and invalid comparison pages', async () => {
    const source = await pdfWithWidths([100])
    await expect(comparisonPdfPagesBytes(source, [])).rejects.toThrow('At least one')
    await expect(
      comparisonPdfPagesBytes(source, [{ image: new Uint8Array(), width: 100, height: 100 }]),
    ).rejects.toThrow('image is empty')
    await expect(
      comparisonPdfPagesBytes(source, [{ image: tinyPngBytes(), width: 0, height: 100 }]),
    ).rejects.toThrow('size is invalid')
  })
})

describe('scannerImageSplitPdfBytes', () => {
  it('combines detected photos into image-only PDF pages with source metadata', async () => {
    const source = await PDFDocument.create()
    source.setTitle('Photo scan')
    source.addPage([600, 800])
    const result = await scannerImageSplitPdfBytes(await source.save(), [
      { image: tinyPngBytes(), width: 300, height: 200 },
      { image: tinyPngBytes(), width: 180, height: 240 },
    ])
    const split = await PDFDocument.load(result)
    expect(split.getTitle()).toBe('Photo scan')
    expect(await pageSizes(result)).toEqual([
      [300, 200],
      [180, 240],
    ])
    expect(
      split.getPages().map((page) => countImagesInResources(split, page.node.Resources())),
    ).toEqual([1, 1])
  })

  it('rejects empty and invalid detected photo pages', async () => {
    const source = await pdfWithWidths([100])
    await expect(scannerImageSplitPdfBytes(source, [])).rejects.toThrow('At least one')
    await expect(
      scannerImageSplitPdfBytes(source, [{ image: new Uint8Array(), width: 100, height: 100 }]),
    ).rejects.toThrow('image is empty')
    await expect(
      scannerImageSplitPdfBytes(source, [{ image: tinyPngBytes(), width: 100, height: -1 }]),
    ).rejects.toThrow('size is invalid')
  })
})

describe('compressPdfPagesBytes', () => {
  it('rebuilds large documents from compact page images while keeping page sizes', async () => {
    const source = await PDFDocument.create()
    for (const width of [320, 480]) {
      const page = source.addPage([width, 240])
      const content = new Uint8Array(50_000)
      for (let index = 0; index < content.length; index++) content[index] = (index * 73) % 251
      page.node.set(PDFName.of('Contents'), source.context.register(source.context.stream(content)))
    }
    const sourceBytes = await source.save({ useObjectStreams: false })
    const result = await compressPdfPagesBytes(sourceBytes, [tinyPngBytes(), tinyPngBytes()])

    expect(result.length).toBeLessThan(sourceBytes.length)
    expect(await pageSizes(result)).toEqual([
      [320, 240],
      [480, 240],
    ])
  })

  it('keeps the source bytes when page rasterization would make the PDF larger', async () => {
    const source = await pdfWithWidths([100])
    const result = await compressPdfPagesBytes(source, [tinyPngBytes()])
    expect(result).toEqual(source)
  })

  it('keeps an intentional raster transformation even when it increases file size', async () => {
    const source = await pdfWithWidths([100])
    const result = await compressPdfPagesBytes(source, [tinyPngBytes()], {
      forceRasterized: true,
    })
    expect(result).not.toEqual(source)
    expect(await pageSizes(result)).toEqual([[100, 200]])
  })
})

describe('redactPdfPagesBytes', () => {
  it('rebuilds every page without searchable source objects or interactive content', async () => {
    const source = await PDFDocument.create()
    const page = source.addPage([400, 300])
    page.drawText('Top secret customer record')
    source.getForm().createTextField('secret_field').addToPage(page)

    const result = await redactPdfPagesBytes(await source.save(), [tinyPngBytes()])
    const redacted = await PDFDocument.load(result)
    expect(redacted.getPageCount()).toBe(1)
    expect(redacted.getPage(0).getSize()).toEqual({ width: 400, height: 300 })
    expect(redacted.catalog.has(PDFName.of('AcroForm'))).toBe(false)
    expect(redacted.getPage(0).node.has(PDFName.of('Annots'))).toBe(false)
    expect(countImagesInResources(redacted, redacted.getPage(0).node.Resources())).toBe(1)
    expect(new TextDecoder('latin1').decode(result)).not.toContain('Top secret customer record')
  })

  it('securely rasterizes only affected pages and preserves other vector pages', async () => {
    const source = await PDFDocument.create()
    source.addPage([400, 300]).drawText('Remove this private value')
    source.addPage([500, 320]).drawText('Keep this public vector text')

    const result = await redactSelectedPdfPagesBytes(await source.save(), [
      { pageIndex: 0, image: tinyPngBytes() },
    ])
    const redacted = await PDFDocument.load(result)
    expect(redacted.getPageCount()).toBe(2)
    expect(redacted.getPage(0).getSize()).toEqual({ width: 400, height: 300 })
    expect(redacted.getPage(1).getSize()).toEqual({ width: 500, height: 320 })
    expect(countImagesInResources(redacted, redacted.getPage(0).node.Resources())).toBe(1)
    expect(countImagesInResources(redacted, redacted.getPage(1).node.Resources())).toBe(0)
    const raw = new TextDecoder('latin1').decode(result)
    expect(raw).not.toContain('Remove this private value')
  })

  it('validates selected redaction pages', async () => {
    const source = await pdfWithWidths([100])
    await expect(redactSelectedPdfPagesBytes(source, [])).rejects.toThrow(/at least one/i)
    await expect(
      redactSelectedPdfPagesBytes(source, [{ pageIndex: 1, image: tinyPngBytes() }]),
    ).rejects.toThrow(/outside/i)
    await expect(
      redactSelectedPdfPagesBytes(source, [
        { pageIndex: 0, image: tinyPngBytes() },
        { pageIndex: 0, image: tinyPngBytes() },
      ]),
    ).rejects.toThrow(/unique/i)
  })
})

describe('addPdfCommentsBytes', () => {
  it('adds standard text annotations while preserving existing annotations', async () => {
    const source = await PDFDocument.create()
    const firstPage = source.addPage([400, 300])
    const existingLink = source.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [0, 0, 100, 20],
    })
    firstPage.node.set(
      PDFName.of('Annots'),
      source.context.obj([source.context.register(existingLink)]),
    )
    source.addPage([500, 400])

    const result = await addPdfCommentsBytes(await source.save(), [
      {
        pageIndex: 1,
        x: 42,
        y: 88,
        width: 20,
        height: 20,
        text: '需要核对这里的数据',
        author: '审核员',
        subject: '数据复核',
      },
      {
        pageIndex: 99,
        x: 0,
        y: 0,
        width: 20,
        height: 20,
        text: 'invalid page',
      },
    ])
    const document = await PDFDocument.load(result)
    const firstAnnotations = document.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const secondAnnotations = document.getPage(1).node.lookup(PDFName.of('Annots'), PDFArray)
    expect(firstAnnotations.size()).toBe(1)
    expect(secondAnnotations.size()).toBe(1)
    const annotation = secondAnnotations.lookup(0, PDFDict)
    expect(annotation.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()).toBe('Text')
    expect(annotation.lookupMaybe(PDFName.of('Name'), PDFName)?.decodeText()).toBe('Comment')
    expect(annotation.lookupMaybe(PDFName.of('Contents'), PDFHexString)?.decodeText()).toBe(
      '需要核对这里的数据',
    )
    expect(annotation.lookupMaybe(PDFName.of('T'), PDFHexString)?.decodeText()).toBe('审核员')
    expect(annotation.lookupMaybe(PDFName.of('Subj'), PDFHexString)?.decodeText()).toBe('数据复核')
    const rect = annotation.lookup(PDFName.of('Rect'), PDFArray)
    expect(
      Array.from({ length: rect.size() }, (_, index) => rect.lookup(index, PDFNumber).asNumber()),
    ).toEqual([42, 88, 62, 108])
  })

  it('rejects requests with no valid comments', async () => {
    await expect(
      addPdfCommentsBytes(await pdfWithWidths([100]), [
        { pageIndex: 0, x: 0, y: 0, width: 20, height: 20, text: ' ' },
      ]),
    ).rejects.toThrow('No valid comments were provided')
  })
})

describe('updatePdfMetadataBytes', () => {
  const metadata = {
    title: 'Updated title',
    author: 'Ada Lovelace',
    subject: 'Metadata migration',
    keywords: 'local, pdf',
    creator: 'GenOffice',
    producer: 'GenOffice PDF',
    creationDate: '2026-08-01T10:20:30.000Z',
    modificationDate: '2026-08-13T04:00:00.000Z',
    trapped: 'False' as const,
    custom: [
      { key: 'Department', value: 'Engineering' },
      { key: 'ReviewStage', value: 'Approved' },
    ],
  }

  it('updates standard and custom properties without changing pages', async () => {
    const source = await pdfWithWidths([100, 200])
    const result = await updatePdfMetadataBytes(source, { deleteAll: false, metadata })
    const analysis = await analyzePdfBytes(result)

    expect(analysis.pages.map((page) => page.width)).toEqual([100, 200])
    expect(analysis.properties).toMatchObject({
      title: 'Updated title',
      author: 'Ada Lovelace',
      subject: 'Metadata migration',
      creator: 'GenOffice',
      producer: 'GenOffice PDF',
      creationDate: '2026-08-01T10:20:30.000Z',
      modificationDate: '2026-08-13T04:00:00.000Z',
      trapped: 'False',
      custom: [
        { key: 'Department', value: 'Engineering' },
        { key: 'ReviewStage', value: 'Approved' },
      ],
    })
  })

  it('replaces old custom fields and clears empty standard fields', async () => {
    const first = await updatePdfMetadataBytes(await pdfWithWidths([100]), {
      deleteAll: false,
      metadata,
    })
    const second = await updatePdfMetadataBytes(first, {
      deleteAll: false,
      metadata: {
        ...metadata,
        title: '',
        creationDate: '',
        custom: [{ key: 'Owner', value: 'Office team' }],
      },
    })
    const analysis = await analyzePdfBytes(second)
    expect(analysis.properties.title).toBeUndefined()
    expect(analysis.properties.creationDate).toBeUndefined()
    expect(analysis.properties.custom).toEqual([{ key: 'Owner', value: 'Office team' }])
  })

  it('clears Info, XMP metadata, and PieceInfo without automatic replacements', async () => {
    const document = await PDFDocument.load(
      await updatePdfMetadataBytes(await pdfWithWidths([100]), { deleteAll: false, metadata }),
      { updateMetadata: false },
    )
    document.catalog.set(
      PDFName.of('Metadata'),
      document.context.register(document.context.stream(new TextEncoder().encode('<xmp />'))),
    )
    document.catalog.set(PDFName.of('PieceInfo'), document.context.obj({ GenOffice: {} }))
    const result = await updatePdfMetadataBytes(await document.save({ useObjectStreams: false }), {
      deleteAll: true,
      metadata,
    })
    const cleared = await PDFDocument.load(result, { updateMetadata: false })
    const info = cleared.context.lookup(cleared.context.trailerInfo.Info, PDFDict)

    expect(info.keys()).toHaveLength(0)
    expect(cleared.catalog.has(PDFName.of('Metadata'))).toBe(false)
    expect(cleared.catalog.has(PDFName.of('PieceInfo'))).toBe(false)
    expect((await analyzePdfBytes(result)).properties.custom).toEqual([])
  })

  it('validates dates and custom keys and exposes the operation suffix', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      updatePdfMetadataBytes(source, {
        deleteAll: false,
        metadata: { ...metadata, creationDate: 'not-a-date' },
      }),
    ).rejects.toThrow('Creation date')
    await expect(
      updatePdfMetadataBytes(source, {
        deleteAll: false,
        metadata: { ...metadata, custom: [{ key: 'Title', value: 'reserved' }] },
      }),
    ).rejects.toThrow('reserved')
    const outputs = await runPdfToolBytes(source, {
      kind: 'metadata',
      deleteAll: false,
      metadata,
    })
    expect(outputs[0]?.suffix).toBe('_metadata.pdf')
  })
})

describe('pdfAutoRenameFileName', () => {
  it('creates a cross-platform safe PDF filename', () => {
    expect(pdfAutoRenameFileName('Quarter / Review: 2026?')).toBe('Quarter Review 2026.pdf')
    expect(pdfAutoRenameFileName('山东文旅年度报告.pdf')).toBe('山东文旅年度报告.pdf')
    expect(pdfAutoRenameFileName('CON')).toBe('_CON.pdf')
  })

  it('limits the UTF-8 filename length and rejects empty titles', () => {
    const fileName = pdfAutoRenameFileName('山东文旅'.repeat(100))
    expect(new TextEncoder().encode(fileName).length).toBeLessThanOrEqual(240)
    expect(fileName.endsWith('.pdf')).toBe(true)
    expect(() => pdfAutoRenameFileName(' / : ? ')).toThrow('No usable PDF title')
  })

  it('exposes a complete output filename and preserves the PDF bytes', async () => {
    const source = await pdfWithWidths([100, 200])
    const outputs = await runPdfToolBytes(source, {
      kind: 'autoRename',
      strategy: 'largestHeading',
      suggestedName: 'Board / Report',
    })
    expect(outputs).toHaveLength(1)
    expect(outputs[0]?.fileName).toBe('Board Report.pdf')
    expect(outputs[0]?.suffix).toBe('')
    expect(outputs[0]?.bytes).toEqual(source)
  })
})

describe('removePdfAnnotationsBytes', () => {
  it('removes every annotation and its form root while preserving other document structures', async () => {
    const document = await PDFDocument.create()
    const firstPage = document.addPage([300, 200])
    firstPage.drawText('Searchable page content')
    document.getForm().createTextField('name').addToPage(firstPage)
    const note = document.context.obj({
      Type: PDFName.of('Annot'),
      Subtype: PDFName.of('Text'),
      Rect: [10, 10, 30, 30],
    })
    const firstAnnotations = firstPage.node.lookup(PDFName.of('Annots'), PDFArray)
    firstAnnotations.push(document.context.register(note))
    const secondPage = document.addPage([400, 250])
    const link = document.context.obj({
      Type: PDFName.of('Annot'),
      Subtype: PDFName.of('Link'),
      Rect: [0, 0, 100, 20],
      A: { S: PDFName.of('URI') },
    })
    secondPage.node.set(
      PDFName.of('Annots'),
      document.context.obj([document.context.register(link)]),
    )
    let sourceBytes = await document.save({ useObjectStreams: false })
    sourceBytes = await setPdfBookmarksBytes(sourceBytes, [
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    sourceBytes = await addPdfAttachmentsBytes(sourceBytes, [
      { name: 'notes.txt', bytes: new TextEncoder().encode('hello') },
    ])

    const result = await removePdfAnnotationsBytes(sourceBytes)
    const cleaned = await PDFDocument.load(result)
    expect(cleaned.getPages().map((page) => page.node.has(PDFName.of('Annots')))).toEqual([
      false,
      false,
    ])
    expect(cleaned.catalog.has(PDFName.of('AcroForm'))).toBe(false)
    expect(cleaned.getPages().map((page) => page.getSize())).toEqual([
      { width: 300, height: 200 },
      { width: 400, height: 250 },
    ])
    expect(await listPdfBookmarksBytes(result)).toEqual([
      { title: 'Start', pageNumber: 1, children: [] },
    ])
    expect((await listPdfAttachmentsBytes(result)).map((attachment) => attachment.name)).toEqual([
      'notes.txt',
    ])
  })

  it('is a valid no-op for PDFs without annotations', async () => {
    const result = await removePdfAnnotationsBytes(await pdfWithWidths([100, 200]))
    expect(await pageWidths(result)).toEqual([100, 200])
  })
})

describe('sanitizePdfBytes', () => {
  it('removes active content, attachments, metadata, XMP, and external link actions', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    document.setTitle('Sensitive title')
    document.addJavaScript('open', 'app.alert("hello")')
    await document.attach(new Uint8Array([1, 2, 3]), 'secret.bin')
    const metadata = document.context.register(
      document.context.stream(new TextEncoder().encode('<x:xmpmeta />')),
    )
    document.catalog.set(PDFName.of('Metadata'), metadata)
    const link = document.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [0, 0, 100, 20],
      A: { S: 'URI' },
    })
    const linkRef = document.context.register(link)
    page.node.set(PDFName.of('Annots'), document.context.obj([linkRef]))
    const chainedJavaScript = document.context.register(
      document.context.obj({ S: 'JavaScript', JS: PDFString.of('nested();') }),
    )
    document.catalog.set(
      PDFName.of('OpenAction'),
      document.context.register(
        document.context.obj({ S: 'GoTo', D: [page.ref, 'Fit'], Next: [chainedJavaScript] }),
      ),
    )

    expect((await auditPdfJavaScriptBytes(await document.save())).actions).not.toHaveLength(0)

    const result = await sanitizePdfBytes(await document.save(), {
      removeJavaScript: true,
      removeEmbeddedFiles: true,
      removeXmpMetadata: true,
      removeMetadata: true,
      removeLinks: true,
    })
    const sanitized = await PDFDocument.load(result, { updateMetadata: false })
    const names = sanitized.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
    expect(names?.has(PDFName.of('JavaScript')) ?? false).toBe(false)
    expect(names?.has(PDFName.of('EmbeddedFiles')) ?? false).toBe(false)
    expect(sanitized.catalog.has(PDFName.of('Metadata'))).toBe(false)
    expect(sanitized.getTitle()).toBeUndefined()
    const annotations = sanitized.getPage(0).node.lookup(PDFName.of('Annots'), PDFArray)
    const sanitizedLink = annotations.lookup(0, PDFDict)
    expect(sanitizedLink.has(PDFName.of('A'))).toBe(false)
    expect((await auditPdfJavaScriptBytes(result)).actions).toHaveLength(0)
    expect(sanitized.catalog.has(PDFName.of('OpenAction'))).toBe(true)
  })
})

describe('auditPdfJavaScriptBytes', () => {
  it('lists named, document, page, annotation, form, and chained JavaScript actions', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const action = (code: string) =>
      document.context.obj({ S: 'JavaScript', JS: PDFString.of(code) })

    const childAction = action('child();')
    const childActionRef = document.context.register(childAction)
    const openAction = action('open();')
    openAction.set(PDFName.of('Next'), document.context.obj([childActionRef]))
    document.catalog.set(PDFName.of('OpenAction'), document.context.register(openAction))
    document.catalog.set(PDFName.of('AA'), document.context.obj({ WC: action('close();') }))

    const namedAction = action('named();')
    const namedKid = document.context.obj({
      Names: [PDFString.of('Startup'), document.context.register(namedAction)],
    })
    document.catalog.set(
      PDFName.of('Names'),
      document.context.obj({ JavaScript: { Kids: [document.context.register(namedKid)] } }),
    )

    page.node.set(PDFName.of('AA'), document.context.obj({ O: action('pageOpen();') }))
    const annotation = document.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [20, 20, 40, 40],
      A: action('annotate();'),
      AA: { E: action('hover();') },
    })
    page.node.set(
      PDFName.of('Annots'),
      document.context.obj([document.context.register(annotation)]),
    )

    const field = document.context.obj({
      FT: 'Tx',
      T: PDFString.of('Customer'),
      A: action('focus();'),
      AA: { K: action('validate();') },
    })
    document.catalog.set(
      PDFName.of('AcroForm'),
      document.context.obj({ Fields: [document.context.register(field)] }),
    )

    const audit = await auditPdfJavaScriptBytes(await document.save({ useObjectStreams: false }))
    expect(audit.actions).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ source: 'named', name: 'Startup', code: 'named();' }),
        expect.objectContaining({ source: 'document', trigger: 'OpenAction', code: 'open();' }),
        expect.objectContaining({
          source: 'document',
          trigger: 'OpenAction / Next 1',
          code: 'child();',
        }),
        expect.objectContaining({ source: 'document', trigger: 'WC', code: 'close();' }),
        expect.objectContaining({
          source: 'page',
          pageNumber: 1,
          trigger: 'O',
          code: 'pageOpen();',
        }),
        expect.objectContaining({
          source: 'annotation',
          pageNumber: 1,
          annotationNumber: 1,
          trigger: 'A',
          code: 'annotate();',
        }),
        expect.objectContaining({ source: 'annotation', trigger: 'E', code: 'hover();' }),
        expect.objectContaining({
          source: 'form',
          fieldName: 'Customer',
          trigger: 'A',
          code: 'focus();',
        }),
        expect.objectContaining({
          source: 'form',
          fieldName: 'Customer',
          trigger: 'K',
          code: 'validate();',
        }),
      ]),
    )
    expect(audit.actions).toHaveLength(9)
    expect(audit.uniqueScriptCount).toBe(9)
    expect(audit.totalCodeBytes).toBeGreaterThan(0)
  })

  it('decodes stream scripts without executing them and reports duplicate code once', async () => {
    const document = await PDFDocument.create()
    document.addPage([100, 100])
    const stream = document.context.register(
      document.context.flateStream(new TextEncoder().encode('shared();')),
    )
    const streamAction = document.context.obj({ S: 'JavaScript', JS: stream })
    document.catalog.set(PDFName.of('OpenAction'), document.context.register(streamAction))
    document.catalog.set(PDFName.of('AA'), document.context.obj({ WC: streamAction }))

    const audit = await auditPdfJavaScriptBytes(await document.save({ useObjectStreams: false }))
    expect(audit.actions.map((entry) => entry.code)).toEqual(['shared();', 'shared();'])
    expect(audit.uniqueScriptCount).toBe(1)
    expect(audit.totalCodeBytes).toBe(18)
  })

  it('returns an empty audit for PDFs without JavaScript actions', async () => {
    const audit = await auditPdfJavaScriptBytes(await pdfWithWidths([200]))
    expect(audit).toEqual({ actions: [], uniqueScriptCount: 0, totalCodeBytes: 0 })
  })
})

describe('repairPdfBytes', () => {
  it('rebuilds a valid cross-reference table from a damaged startxref pointer', async () => {
    const document = await PDFDocument.create()
    document.addPage([320, 240])
    document.addPage([640, 480])
    document.setTitle('Repair fixture')
    const source = await document.save({ useObjectStreams: false })
    const damaged = new Uint8Array(source)
    const sourceText = new TextDecoder('latin1').decode(damaged)
    const startXrefMatch = /startxref\s+(\d+)\s+%%EOF\s*$/.exec(sourceText)
    expect(startXrefMatch).not.toBeNull()
    const pointer = startXrefMatch![1]!
    const pointerOffset = startXrefMatch!.index + startXrefMatch![0].indexOf(pointer)
    damaged.fill(0x30, pointerOffset, pointerOffset + pointer.length)

    const repaired = await repairPdfBytes(damaged)
    const repairedText = new TextDecoder('latin1').decode(repaired)
    const repairedStartXref = /startxref\s+(\d+)\s+%%EOF\s*$/.exec(repairedText)
    expect(repairedStartXref).not.toBeNull()
    const xrefOffset = Number(repairedStartXref![1])
    expect(new TextDecoder().decode(repaired.slice(xrefOffset, xrefOffset + 4))).toBe('xref')

    const reloaded = await PDFDocument.load(repaired, {
      throwOnInvalidObject: true,
      updateMetadata: false,
    })
    expect(reloaded.getPages().map((page) => [page.getWidth(), page.getHeight()])).toEqual([
      [320, 240],
      [640, 480],
    ])
    expect(reloaded.getTitle()).toBe('Repair fixture')
  })
})

describe('decompressPdfBytes', () => {
  it('decodes filtered streams and writes them without compression metadata', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([300, 200])
    const content = new TextEncoder().encode(`q\n${'0 0 0 rg\n'.repeat(200)}Q\n`)
    const stream = document.context.flateStream(content, {
      DecodeParms: document.context.obj({ Predictor: 1 }),
      D: document.context.obj({ Predictor: 1 }),
    })
    const streamReference = document.context.register(stream)
    page.node.set(PDFName.of('Contents'), streamReference)
    const source = await document.save({ useObjectStreams: false })

    const decompressed = await decompressPdfBytes(source)
    const result = await PDFDocument.load(decompressed, { updateMetadata: false })
    const resultStream = result.context.lookup(streamReference)
    expect(resultStream).toBeInstanceOf(PDFRawStream)
    const rawStream = resultStream as PDFRawStream
    expect(rawStream.dict.has(PDFName.of('Filter'))).toBe(false)
    expect(rawStream.dict.has(PDFName.of('DecodeParms'))).toBe(false)
    expect(rawStream.dict.has(PDFName.of('D'))).toBe(false)
    expect(rawStream.getContents()).toEqual(content)
    expect(decompressed.length).toBeGreaterThan(source.length)
  })
})

describe('removePdfSignaturesBytes', () => {
  it('keeps the signature appearance while preserving ordinary form fields', async () => {
    const document = await PDFDocument.create()
    const page = document.addPage([400, 240])
    const form = document.getForm()
    form.createTextField('customer_name').addToPage(page, {
      x: 30,
      y: 160,
      width: 180,
      height: 28,
    })

    const signatureDictionary = document.context.obj({
      FT: 'Sig',
      T: PDFHexString.fromText('approval_signature'),
      V: { Type: 'Sig', Filter: 'Adobe.PPKLite', SubFilter: 'adbe.pkcs7.detached' },
    })
    const signatureReference = document.context.register(signatureDictionary)
    const widget = PDFWidgetAnnotation.create(document.context, signatureReference)
    widget.setRectangle({ x: 40, y: 50, width: 160, height: 48 })
    widget.setP(page.ref)
    const appearance = document.context.formXObject([], {
      BBox: [0, 0, 160, 48],
      Resources: {},
    })
    const appearanceReference = document.context.register(appearance)
    widget.setNormalAppearance(appearanceReference)
    const widgetReference = document.context.register(widget.dict)
    signatureDictionary.set(PDFName.of('Kids'), document.context.obj([widgetReference]))
    form.acroForm.addField(signatureReference)
    page.node.addAnnot(widgetReference)

    const source = await document.save({ useObjectStreams: false, updateFieldAppearances: false })
    const loadedSource = await PDFDocument.load(source, { updateMetadata: false })
    expect(
      loadedSource
        .getForm()
        .getFields()
        .some((field) => field instanceof PDFSignature),
    ).toBe(true)

    const unsigned = await removePdfSignaturesBytes(source)
    const result = await PDFDocument.load(unsigned, { updateMetadata: false })
    expect(
      result
        .getForm()
        .getFields()
        .map((field) => field.getName()),
    ).toEqual(['customer_name'])
    expect(
      result
        .getForm()
        .getFields()
        .some((field) => field instanceof PDFSignature),
    ).toBe(false)
    const annotations = result.getPage(0).node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    expect(annotations?.size()).toBe(1)
    const contentStreams = result.getPage(0).node.Contents()
    const contentText = contentStreams
      ? Array.from(
          { length: contentStreams instanceof PDFArray ? contentStreams.size() : 1 },
          (_, i) => {
            const stream =
              contentStreams instanceof PDFArray
                ? contentStreams.lookup(i, PDFRawStream)
                : contentStreams
            return new TextDecoder().decode(decodePDFRawStream(stream as PDFRawStream).decode())
          },
        ).join('\n')
      : ''
    expect(contentText).toContain('/UnsignedWidget')
    expect(result.context.lookup(signatureReference)).toBeUndefined()
  })
})

describe('runPdfToolBytes', () => {
  it('returns stable output suffixes for adapters', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200]), {
      kind: 'split',
      mode: 'afterPages',
      splitAfterPages: [0],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_split_1.pdf', '_split_2.pdf'])
  })

  it('extracts pages in requested order with the Stirling-compatible suffix', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'extractPages',
      pageIndexes: [2, 0, 2],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_extracted_pages.pdf'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([300, 100, 300])
  })

  it('merges external PDFs around the current document in the requested order', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([200, 300]), {
      kind: 'merge',
      documents: [await pdfWithWidths([100]), await pdfWithWidths([400, 500])],
      currentDocumentIndex: 1,
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_merged.pdf'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([100, 200, 300, 400, 500])
  })

  it('exports prepared comparison pages with a stable suffix', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'compare',
      comparisonDocument: await pdfWithWidths([200]),
      renderDpi: 150,
      threshold: 0.1,
      pages: [{ image: tinyPngBytes(), width: 200, height: 300 }],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_comparison.pdf'])
    expect(await pageSizes(outputs[0]!.bytes)).toEqual([[200, 300]])
  })

  it('requires prepared comparison pages', async () => {
    await expect(
      runPdfToolBytes(await pdfWithWidths([100]), {
        kind: 'compare',
        comparisonDocument: await pdfWithWidths([200]),
        renderDpi: 150,
        threshold: 0.1,
      }),
    ).rejects.toThrow('comparison pages')
  })

  it('rejects invalid comparison settings and empty documents', async () => {
    const source = await pdfWithWidths([100])
    const comparisonDocument = await pdfWithWidths([200])
    const pages = [{ image: tinyPngBytes(), width: 200, height: 300 }]
    await expect(
      runPdfToolBytes(source, {
        kind: 'compare',
        comparisonDocument,
        renderDpi: 301,
        threshold: 0.1,
        pages,
      }),
    ).rejects.toThrow('72 to 300')
    await expect(
      runPdfToolBytes(source, {
        kind: 'compare',
        comparisonDocument,
        renderDpi: 150,
        threshold: -0.1,
        pages,
      }),
    ).rejects.toThrow('from 0 to 1')
    await expect(
      runPdfToolBytes(source, {
        kind: 'compare',
        comparisonDocument: new Uint8Array(),
        renderDpi: 150,
        threshold: 0.1,
        pages,
      }),
    ).rejects.toThrow('PDF is empty')
  })

  it('exports detected scanned images with a stable suffix', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'scannerImageSplit',
      angleThreshold: 10,
      tolerance: 30,
      minArea: 10000,
      minContourArea: 500,
      borderSize: 1,
      renderDpi: 150,
      pages: [{ image: tinyPngBytes(), width: 200, height: 300 }],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_scanned_images.pdf'])
    expect(await pageSizes(outputs[0]!.bytes)).toEqual([[200, 300]])
  })

  it('validates scanned-image detection settings and prepared pages', async () => {
    const source = await pdfWithWidths([100])
    const valid = {
      kind: 'scannerImageSplit' as const,
      angleThreshold: 10,
      tolerance: 30,
      minArea: 10000,
      minContourArea: 500,
      borderSize: 1,
      renderDpi: 150,
    }
    await expect(runPdfToolBytes(source, valid)).rejects.toThrow('Detected scanned image')
    await expect(runPdfToolBytes(source, { ...valid, tolerance: 256, pages: [] })).rejects.toThrow(
      'tolerance',
    )
    await expect(runPdfToolBytes(source, { ...valid, minArea: 0, pages: [] })).rejects.toThrow(
      'minimum area',
    )
    await expect(runPdfToolBytes(source, { ...valid, renderDpi: 301, pages: [] })).rejects.toThrow(
      'rendering DPI',
    )
  })

  it('rejects invalid merge inputs', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      runPdfToolBytes(source, { kind: 'merge', documents: [], currentDocumentIndex: 0 }),
    ).rejects.toThrow('at least one PDF')
    await expect(
      runPdfToolBytes(source, {
        kind: 'merge',
        documents: [await pdfWithWidths([200])],
        currentDocumentIndex: 2,
      }),
    ).rejects.toThrow('position is invalid')
    await expect(
      runPdfToolBytes(source, {
        kind: 'merge',
        documents: [new Uint8Array()],
        currentDocumentIndex: 0,
      }),
    ).rejects.toThrow('PDF is empty')
  })

  it('rejects invalid or empty page extraction selections', async () => {
    const source = await pdfWithWidths([100, 200])
    await expect(
      runPdfToolBytes(source, { kind: 'extractPages', pageIndexes: [0, 2] }),
    ).rejects.toThrow('invalid page')
    await expect(
      runPdfToolBytes(source, { kind: 'extractPages', pageIndexes: [] }),
    ).rejects.toThrow('At least one page')
  })

  it('returns the repair suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), { kind: 'repair' })
    expect(outputs.map((output) => output.suffix)).toEqual(['_repaired.pdf'])
  })

  it('returns the decompress suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), { kind: 'decompress' })
    expect(outputs.map((output) => output.suffix)).toEqual(['_decompressed.pdf'])
  })

  it('runs a validated single-output PDF pipeline in order', async () => {
    const source = await PDFDocument.create()
    const page = source.addPage([400, 300])
    page.drawText('Pipeline source text')
    source.getForm().createTextField('pipeline_field').addToPage(page)
    const link = source.context.obj({
      Type: 'Annot',
      Subtype: 'Link',
      Rect: [10, 10, 100, 30],
      A: { S: 'URI', URI: 'https://example.com' },
    })
    page.node.set(PDFName.of('Annots'), source.context.obj([source.context.register(link)]))

    const outputs = await runPdfToolBytes(await source.save(), {
      kind: 'pipeline',
      steps: [
        {
          kind: 'sanitize',
          removeJavaScript: true,
          removeEmbeddedFiles: true,
          removeXmpMetadata: true,
          removeMetadata: false,
          removeLinks: true,
        },
        { kind: 'flattenForms' },
        { kind: 'removeAnnotations' },
        { kind: 'repair' },
      ],
    })

    expect(outputs.map((output) => output.suffix)).toEqual(['_processed.pdf'])
    const processed = await PDFDocument.load(outputs[0]!.bytes)
    expect(processed.getPageCount()).toBe(1)
    expect(processed.getForm().getFields()).toHaveLength(0)
    expect(processed.getPage(0).node.has(PDFName.of('Annots'))).toBe(false)
  })

  it('rejects empty, oversized, and no-op cleanup pipelines', async () => {
    const source = await pdfWithWidths([100])
    await expect(runPdfToolBytes(source, { kind: 'pipeline', steps: [] })).rejects.toThrow(
      /1 to 12/i,
    )
    await expect(
      runPdfToolBytes(source, {
        kind: 'pipeline',
        steps: Array.from({ length: 13 }, () => ({ kind: 'repair' as const })),
      }),
    ).rejects.toThrow(/1 to 12/i)
    await expect(
      runPdfToolBytes(source, {
        kind: 'pipeline',
        steps: [
          {
            kind: 'sanitize',
            removeJavaScript: false,
            removeEmbeddedFiles: false,
            removeXmpMetadata: false,
            removeMetadata: false,
            removeLinks: false,
          },
        ],
      }),
    ).rejects.toThrow(/cleanup action/i)
  })

  it('returns the compressed suffix through the shared adapter', async () => {
    const source = await pdfWithWidths([100])
    const outputs = await runPdfToolBytes(source, {
      kind: 'compress',
      renderDpi: 120,
      imageQuality: 72,
      pageImages: [tinyPngBytes()],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_compressed.pdf'])
  })

  it('keeps line-art output through the shared adapter when it is larger', async () => {
    const source = await pdfWithWidths([100])
    const outputs = await runPdfToolBytes(source, {
      kind: 'compress',
      renderDpi: 120,
      imageQuality: 72,
      lineArt: true,
      lineArtThreshold: 55,
      lineArtEdgeLevel: 1,
      pageImages: [tinyPngBytes()],
    })
    expect(outputs[0]?.bytes).not.toEqual(source)
  })

  it('validates line-art compression settings in the shared adapter', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      runPdfToolBytes(source, {
        kind: 'compress',
        renderDpi: 120,
        imageQuality: 72,
        lineArt: true,
        lineArtThreshold: 101,
        lineArtEdgeLevel: 1,
        pageImages: [tinyPngBytes()],
      }),
    ).rejects.toThrow('0 to 100')
  })

  it('returns the redacted suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'redact',
      patterns: ['secret'],
      useRegex: false,
      wholeWord: false,
      color: '#000000',
      padding: 2,
      renderDpi: 144,
      pageImages: [tinyPngBytes()],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_redacted.pdf'])
  })

  it('returns locally selected redaction pages through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 120]), {
      kind: 'redact',
      mode: 'areas',
      patterns: [],
      useRegex: false,
      wholeWord: false,
      color: '#000000',
      padding: 0,
      renderDpi: 144,
      areas: [{ pageIndex: 0, x: 0.1, y: 0.1, width: 0.4, height: 0.2 }],
      pages: [{ pageIndex: 0, image: tinyPngBytes() }],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_redacted.pdf'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([100, 120])
  })

  it('returns the commented suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'comments',
      comments: [{ pageIndex: 0, x: 10, y: 20, width: 20, height: 20, text: 'Review this' }],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_commented.pdf'])
  })

  it('returns the auto-rotated suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'autoRotate',
      inferUndetected: true,
      pageRotations: [{ pageIndex: 0, angle: 90 }],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_auto_rotated.pdf'])
    expect(await pageRotations(outputs[0]!.bytes)).toEqual([90])
  })

  it('returns the unsigned suffix through the shared adapter', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'removeSignatures',
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_unsigned.pdf'])
  })

  it('uses sanitized chapter titles in split output suffixes', async () => {
    const source = await setPdfBookmarksBytes(await pdfWithWidths([100, 200]), [
      { title: 'Intro / Overview', pageNumber: 1, children: [] },
      { title: '第二章', pageNumber: 2, children: [] },
    ])
    const outputs = await runPdfToolBytes(source, {
      kind: 'split',
      mode: 'chapters',
      bookmarkLevel: 0,
      allowDuplicates: false,
    })
    expect(outputs.map((output) => output.suffix)).toEqual([
      '_chapter_1_Intro_Overview.pdf',
      '_chapter_2_第二章.pdf',
    ])
  })

  it('returns stable suffixes for page composition tools', async () => {
    const source = await pdfWithWidths([100, 100, 100, 100])
    const booklet = await runPdfToolBytes(source, {
      kind: 'booklet',
      spine: 'left',
      gutter: 0,
      border: false,
      duplexPass: 'both',
      flipOnShortEdge: false,
    })
    const poster = await runPdfToolBytes(source, {
      kind: 'poster',
      pageSize: 'A4',
      rows: 2,
      columns: 2,
      readingDirection: 'ltr',
    })
    const singlePage = await runPdfToolBytes(source, {
      kind: 'singlePage',
      direction: 'vertical',
    })
    const overlay = await runPdfToolBytes(source, {
      kind: 'overlay',
      overlayDocuments: [await pdfWithWidths([100])],
      mode: 'sequential',
      position: 'foreground',
      opacity: 1,
    })
    expect([
      booklet[0]!.suffix,
      poster[0]!.suffix,
      singlePage[0]!.suffix,
      overlay[0]!.suffix,
    ]).toEqual(['_booklet.pdf', '_poster.pdf', '_single_page.pdf', '_overlayed.pdf'])
  })

  it('dispatches prepared auto crop boxes with a stable output suffix', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([200]), {
      kind: 'crop',
      mode: 'auto',
      whiteThreshold: 250,
      padding: 6,
      pageBoxes: [{ x: 20, y: 30, width: 160, height: 140 }],
    })
    expect(output!.suffix).toBe('_auto_cropped.pdf')
    expect((await PDFDocument.load(output!.bytes)).getPage(0).getCropBox()).toEqual({
      x: 20,
      y: 30,
      width: 160,
      height: 140,
    })
  })

  it('requires prepared page boxes for auto crop dispatch', async () => {
    await expect(
      runPdfToolBytes(await pdfWithWidths([200]), {
        kind: 'crop',
        mode: 'auto',
        whiteThreshold: 250,
        padding: 6,
      }),
    ).rejects.toThrow('Auto crop boxes are required')
  })

  it('describes ZIP attachment exports for platform adapters', async () => {
    const source = await addPdfAttachmentsBytes(await pdfWithWidths([100]), [
      { name: 'notes.txt', bytes: new TextEncoder().encode('hello') },
    ])
    const [output] = await runPdfToolBytes(source, {
      kind: 'attachments',
      action: 'extract',
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_attachments.zip',
        mimeType: 'application/zip',
        extension: '.zip',
      }),
    )
  })

  it('returns a stable suffix for bookmark edits', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'bookmarks',
      bookmarks: [{ title: 'Chapter', pageNumber: 1, children: [] }],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_with_toc.pdf'])
  })

  it('returns the Stirling-compatible suffix for image removal', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'removeImages',
      pageIndexes: [0],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_images_removed.pdf'])
  })

  it('packages extracted images with Stirling-compatible names', async () => {
    const archiveBytes = await extractPdfImagesZipBytes(
      [
        { pageNumber: 2, imageNumber: 1, bytes: new Uint8Array([1, 2, 3]) },
        { pageNumber: 4, imageNumber: 2, bytes: new Uint8Array([4, 5]) },
      ],
      '../My:PDF.pdf',
      'png',
    )
    const archive = await JSZip.loadAsync(archiveBytes)
    expect(Object.keys(archive.files)).toEqual(['_My_PDF_page_2_1.png', '_My_PDF_page_4_2.png'])
    expect(Array.from(await archive.file('_My_PDF_page_2_1.png')!.async('uint8array'))).toEqual([
      1, 2, 3,
    ])
  })

  it('creates an empty extracted-images ZIP for image-free PDFs', async () => {
    const archive = await JSZip.loadAsync(await extractPdfImagesZipBytes([], 'empty.pdf', 'jpg'))
    expect(Object.keys(archive.files)).toEqual([])
  })

  it('returns a stable adapter output for embedded image extraction', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'extractImages',
      format: 'gif',
      baseName: 'document.pdf',
      images: [{ pageNumber: 1, imageNumber: 1, bytes: new Uint8Array([71, 73, 70]) }],
    })
    expect(outputs[0]).toEqual(
      expect.objectContaining({
        suffix: '_extracted-images.zip',
        mimeType: 'application/zip',
        extension: '.zip',
      }),
    )
    const archive = await JSZip.loadAsync(outputs[0]!.bytes)
    expect(Object.keys(archive.files)).toEqual(['document_page_1_1.gif'])
  })

  it('exports one rendered PDF page as a directly downloadable image', async () => {
    const output = await pdfPageImagesOutput({
      pageCount: 12,
      pageIndexes: [1],
      format: 'jpg',
      outputMode: 'multiple',
      renderDpi: 150,
      imageQuality: 85,
      colorMode: 'color',
      includeAnnotations: true,
      baseName: '../Quarter:Review.pdf',
      images: [{ pageNumber: 2, bytes: new Uint8Array([255, 216, 255]) }],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_page_2.jpg',
        fileName: '_Quarter_Review_page_02.jpg',
        mimeType: 'image/jpeg',
        extension: '.jpg',
      }),
    )
    expect(Array.from(output.bytes)).toEqual([255, 216, 255])
  })

  it('packages multiple rendered PDF pages in page order', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'pdfToImages',
      pageCount: 3,
      pageIndexes: [2, 0],
      format: 'webp',
      outputMode: 'multiple',
      renderDpi: 144,
      imageQuality: 90,
      colorMode: 'greyscale',
      includeAnnotations: false,
      baseName: 'slides.pdf',
      images: [
        { pageNumber: 3, bytes: new Uint8Array([3]) },
        { pageNumber: 1, bytes: new Uint8Array([1]) },
      ],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_images.zip',
        mimeType: 'application/zip',
        extension: '.zip',
      }),
    )
    const archive = await JSZip.loadAsync(output!.bytes)
    expect(Object.keys(archive.files)).toEqual(['slides_page_3.webp', 'slides_page_1.webp'])
    expect(Array.from(await archive.file('slides_page_3.webp')!.async('uint8array'))).toEqual([3])
  })

  it('validates prepared PDF page images and export settings', async () => {
    const valid = {
      pageCount: 2,
      pageIndexes: [0],
      format: 'png' as const,
      outputMode: 'multiple' as const,
      renderDpi: 150,
      imageQuality: 92,
      colorMode: 'color' as const,
      includeAnnotations: true,
      images: [{ pageNumber: 1, bytes: new Uint8Array([1]) }],
    }
    await expect(pdfPageImagesOutput({ ...valid, renderDpi: 301 })).rejects.toThrow('72 to 300')
    await expect(pdfPageImagesOutput({ ...valid, imageQuality: 9 })).rejects.toThrow('10 to 100')
    await expect(pdfPageImagesOutput({ ...valid, images: undefined })).rejects.toThrow(
      'Rendered PDF page images are required',
    )
    await expect(
      pdfPageImagesOutput({ ...valid, images: [{ pageNumber: 2, bytes: new Uint8Array([1]) }] }),
    ).rejects.toThrow('order is invalid')
  })

  it('packages rendered PDF pages as a naturally ordered CBZ archive', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'pdfToCbz',
      pageCount: 3,
      pageIndexes: [2, 0],
      format: 'webp',
      renderDpi: 150,
      imageQuality: 88,
      colorMode: 'greyscale',
      includeAnnotations: false,
      baseName: '../Comic:Issue.pdf',
      images: [
        { pageNumber: 3, bytes: new Uint8Array([3]) },
        { pageNumber: 1, bytes: new Uint8Array([1]) },
      ],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_converted.cbz',
        fileName: '_Comic_Issue_converted.cbz',
        mimeType: 'application/vnd.comicbook+zip',
        extension: '.cbz',
      }),
    )
    const archive = await JSZip.loadAsync(output!.bytes)
    expect(Object.keys(archive.files)).toEqual(['page_003.webp', 'page_001.webp'])
    expect(Array.from(await archive.file('page_003.webp')!.async('uint8array'))).toEqual([3])
  })

  it('validates CBZ image format and prepared page order', async () => {
    const valid = {
      pageCount: 1,
      pageIndexes: [0],
      format: 'png' as const,
      renderDpi: 150,
      imageQuality: 90,
      colorMode: 'color' as const,
      includeAnnotations: true,
      images: [{ pageNumber: 1, bytes: new Uint8Array([1]) }],
    }
    await expect(pdfToCbzBytes({ ...valid, format: 'gif' as 'png' })).rejects.toThrow(
      'format is invalid',
    )
    await expect(
      pdfToCbzBytes({ ...valid, images: [{ pageNumber: 2, bytes: new Uint8Array([1]) }] }),
    ).rejects.toThrow('order is invalid')
  })

  it('returns a single long image without wrapping it in a ZIP file', async () => {
    const output = await pdfPageImagesOutput({
      pageCount: 2,
      pageIndexes: [0, 1],
      format: 'gif',
      outputMode: 'single',
      renderDpi: 96,
      imageQuality: 90,
      colorMode: 'blackwhite',
      includeAnnotations: false,
      baseName: 'handout.pdf',
      images: [{ pageNumber: 1, bytes: new Uint8Array([71, 73, 70]) }],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: '_long.gif',
        fileName: 'handout_long.gif',
        mimeType: 'image/gif',
        extension: '.gif',
      }),
    )
  })

  it('packages searchable, escaped PDF pages as an offline HTML archive', async () => {
    const [output] = await runPdfToolBytes(await pdfWithWidths([612, 612]), {
      kind: 'pdfToHtml',
      pageCount: 2,
      pageIndexes: [1],
      renderDpi: 144,
      includeAnnotations: false,
      baseName: '../Quarter <Review>.pdf',
      pages: [
        {
          pageNumber: 2,
          width: 612,
          height: 792,
          imageBytes: new Uint8Array([137, 80, 78, 71]),
          text: 'Revenue <script>alert(1)</script>',
          textRuns: [
            {
              text: 'Revenue <script>',
              x: 40,
              y: 72,
              fontSize: 18,
              angle: 0,
              fontFamily: 'Inter" onclick="alert(1)',
              bold: true,
            },
          ],
        },
      ],
    })
    expect(output).toEqual(
      expect.objectContaining({
        suffix: 'ToHtml.zip',
        fileName: '_Quarter _Review_ToHtml.zip',
        mimeType: 'application/zip',
        extension: '.zip',
      }),
    )
    const archive = await JSZip.loadAsync(output!.bytes)
    expect(Object.keys(archive.files)).toEqual(['index.html', 'pages/', 'pages/page-2.png'])
    const html = await archive.file('index.html')!.async('string')
    expect(html).toContain('Content-Security-Policy')
    expect(html).toContain('Revenue &lt;script&gt;alert(1)&lt;/script&gt;')
    expect(html).toContain('font-family="Inter&quot; onclick=&quot;alert(1)"')
    expect(html).not.toContain('<script>')
  })

  it('requires prepared PDF HTML pages in the requested order', async () => {
    const valid = {
      pageCount: 2,
      pageIndexes: [0],
      renderDpi: 150,
      includeAnnotations: true,
      pages: [
        {
          pageNumber: 1,
          width: 100,
          height: 200,
          imageBytes: new Uint8Array([1]),
          text: '',
          textRuns: [],
        },
      ],
    }
    await expect(pdfToHtmlZipBytes({ ...valid, renderDpi: 301 })).rejects.toThrow('72 to 300')
    await expect(pdfToHtmlZipBytes({ ...valid, pages: undefined })).rejects.toThrow(
      'Rendered PDF HTML pages are required',
    )
    await expect(
      pdfToHtmlZipBytes({ ...valid, pages: [{ ...valid.pages[0]!, pageNumber: 2 }] }),
    ).rejects.toThrow('order is invalid')
  })

  it('rejects image extraction before renderer data preparation', async () => {
    await expect(
      runPdfToolBytes(await pdfWithWidths([100]), { kind: 'extractImages', format: 'png' }),
    ).rejects.toThrow('Extracted image data is required')
  })

  it('returns a stable suffix for annotation removal', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'removeAnnotations',
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_annotations_removed.pdf'])
  })

  it('removes detected blank pages and optionally exports them', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'removeBlanks',
      threshold: 10,
      whitePercent: 99.9,
      includeBlankPages: true,
      blankPageIndexes: [1],
    })
    expect(outputs.map((output) => output.suffix)).toEqual([
      '_nonBlankPages.pdf',
      '_blankPages.pdf',
    ])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([100, 300])
    expect(await pageWidths(outputs[1]!.bytes)).toEqual([200])
  })

  it('returns only the cleaned PDF when blank-page export is disabled', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200, 300]), {
      kind: 'removeBlanks',
      threshold: 10,
      whitePercent: 99.9,
      includeBlankPages: false,
      blankPageIndexes: [1],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_nonBlankPages.pdf'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([100, 300])
  })

  it('keeps an all-blank document as a valid non-empty PDF', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200]), {
      kind: 'removeBlanks',
      threshold: 10,
      whitePercent: 99.9,
      includeBlankPages: true,
      blankPageIndexes: [0, 1],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_allBlankPages.pdf'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([100, 200])
  })

  it('returns a full non-blank copy when no blank pages are detected', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100, 200]), {
      kind: 'removeBlanks',
      threshold: 10,
      whitePercent: 99.9,
      includeBlankPages: true,
      blankPageIndexes: [],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_nonBlankPages.pdf'])
    expect(await pageWidths(outputs[0]!.bytes)).toEqual([100, 200])
  })

  it('validates blank-page scan input from platform adapters', async () => {
    const source = await pdfWithWidths([100])
    await expect(
      runPdfToolBytes(source, {
        kind: 'removeBlanks',
        threshold: 256,
        whitePercent: 99.9,
        includeBlankPages: false,
        blankPageIndexes: [],
      }),
    ).rejects.toThrow('threshold must be from 0 to 255')
    await expect(
      runPdfToolBytes(source, {
        kind: 'removeBlanks',
        threshold: 10,
        whitePercent: 0,
        includeBlankPages: false,
        blankPageIndexes: [],
      }),
    ).rejects.toThrow('whitePercent must be greater than 0 and at most 100')
    await expect(
      runPdfToolBytes(source, {
        kind: 'removeBlanks',
        threshold: 10,
        whitePercent: 99.9,
        includeBlankPages: false,
      }),
    ).rejects.toThrow('Detected blank page indexes are required')
  })

  it('returns the Stirling-compatible suffix for full color inversion', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'invertColors',
      pageIndexes: [0],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_inverted.pdf'])
  })

  it('returns a stable suffix for high-contrast and custom recoloring', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'replaceColors',
      pageIndexes: [0],
      textColor: '#ffffff',
      backgroundColor: '#000000',
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_recolored.pdf'])
  })

  it('returns a stable suffix for raster color adjustments', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'adjustColors',
      pageIndexes: [0],
      pageImages: [tinyPngBytes()],
      contrast: 100,
      brightness: 100,
      saturation: 100,
      red: 100,
      green: 100,
      blue: 100,
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_adjusted.pdf'])
  })

  it('returns the Stirling-compatible suffix for the scanner effect', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'scannerEffect',
      quality: 'high',
      rotation: 'slight',
      colorspace: 'grayscale',
      border: 20,
      rotate: 0,
      rotateVariance: 2,
      brightness: 1.03,
      contrast: 1.06,
      blur: 0.1,
      noise: 1,
      yellowish: false,
      renderDpi: 150,
      seed: 2026,
      pageImages: [tinyPngBytes()],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_scanner_effect.pdf'])
  })

  it('returns a stable suffix for local deskewing', async () => {
    const outputs = await runPdfToolBytes(await pdfWithWidths([100]), {
      kind: 'deskew',
      pageIndexes: [0],
      maxAngle: 8,
      renderDpi: 150,
      includeAnnotations: true,
      pages: [],
    })
    expect(outputs.map((output) => output.suffix)).toEqual(['_deskewed.pdf'])
  })
})
