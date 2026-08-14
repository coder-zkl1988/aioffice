import {
  PDFArray,
  PDFDocument,
  PDFDropdown,
  PDFHexString,
  PDFName,
  PDFOptionList,
  degrees,
  rgb,
} from 'pdf-lib'
import type { PDFRef } from 'pdf-lib'
import {
  applyPdfClassificationMetadata,
  extractPagesBytes,
  insertBlankPageBytes,
  insertPdfBytes,
  removeEmptyPdfSignatureFields,
} from '@genoffice/pdf-tools'
import type {
  DrawingInput,
  ImageEditFailure,
  ImageEditInput,
  MarkupInput,
  PageImageRef,
  PagePreviewRequest,
  SavePdfRequest,
  StampInput,
  TextEditFailure,
  TextEditInput,
  TextEditValidation,
} from '../../../pdf/src/shared/ipc'
import { stampDrawPlacement } from '../../../pdf/src/shared/ipc'

const color = ([red, green, blue]: [number, number, number]) => rgb(red, green, blue)

function base64FromBytes(data: ArrayBuffer): string {
  const bytes = new Uint8Array(data)
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function bytesFromBase64(value: string): ArrayBuffer {
  const binary = atob(value)
  return Uint8Array.from(binary, (character) => character.charCodeAt(0)).buffer
}

async function pdfApi<T>(action: string, body?: unknown): Promise<T> {
  const response = await fetch(new URL(`./api/pdf/${action}`, document.baseURI), {
    method: body === undefined ? 'GET' : 'POST',
    headers: body === undefined ? undefined : { 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body),
  })
  if (!response.ok) {
    const result = (await response.json().catch(() => ({}))) as { error?: string }
    throw new Error(result.error || `PDF 服务返回 HTTP ${response.status}`)
  }
  return (await response.json()) as T
}

export async function validateWebPdfTextEdits(
  source: ArrayBuffer,
  edits: TextEditInput[],
): Promise<TextEditValidation[]> {
  const result = await pdfApi<{ validations: TextEditValidation[] }>('text/validate', {
    pdfBase64: base64FromBytes(source),
    edits,
  })
  return result.validations
}

export async function applyWebPdfTextEdits(
  source: ArrayBuffer,
  edits: TextEditInput[],
): Promise<{ data: ArrayBuffer; skipped: TextEditFailure[] }> {
  const result = await pdfApi<{ pdfBase64: string; skipped: TextEditFailure[] }>('text/apply', {
    pdfBase64: base64FromBytes(source),
    edits,
  })
  return { data: bytesFromBase64(result.pdfBase64), skipped: result.skipped }
}

export async function listWebPdfEditFonts(): Promise<string[]> {
  return (await pdfApi<{ fonts: string[] }>('fonts')).fonts
}

export async function listWebPdfPageImages(source: ArrayBuffer): Promise<PageImageRef[]> {
  return (
    await pdfApi<{ images: PageImageRef[] }>('images/list', {
      pdfBase64: base64FromBytes(source),
    })
  ).images
}

export async function renderWebPdfPageImage(
  source: ArrayBuffer,
  pageIndex: number,
  rect: [number, number, number, number],
): Promise<string | null> {
  return (
    await pdfApi<{ pngBase64: string | null }>('images/render', {
      pdfBase64: base64FromBytes(source),
      pageIndex,
      rect,
    })
  ).pngBase64
}

export async function renderWebPdfPagePreview(
  source: ArrayBuffer,
  request: Omit<PagePreviewRequest, 'path'>,
): Promise<string | null> {
  return (
    await pdfApi<{ pngBase64: string | null }>('images/preview', {
      pdfBase64: base64FromBytes(source),
      ...request,
    })
  ).pngBase64
}

export async function applyWebPdfImageEdits(
  source: ArrayBuffer,
  edits: ImageEditInput[],
): Promise<{ data: ArrayBuffer; skipped: ImageEditFailure[] }> {
  const result = await pdfApi<{ pdfBase64: string; skipped: ImageEditFailure[] }>('images/apply', {
    pdfBase64: base64FromBytes(source),
    edits,
  })
  return { data: bytesFromBase64(result.pdfBase64), skipped: result.skipped }
}

function bounds(values: number[]): [number, number, number, number] {
  const xs = values.filter((_, index) => index % 2 === 0)
  const ys = values.filter((_, index) => index % 2 === 1)
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
}

function appendAnnotation(
  document: PDFDocument,
  page: ReturnType<PDFDocument['getPage']>,
  reference: PDFRef,
): void {
  const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
  if (annotations) annotations.push(reference)
  else page.node.set(PDFName.of('Annots'), document.context.obj([reference]))
}

function drawMarkup(page: ReturnType<PDFDocument['getPage']>, markup: MarkupInput): void {
  for (const quad of markup.quads) {
    const [x1, y1, x2, y2] = bounds(quad)
    if (markup.type === 'highlight') {
      page.drawRectangle({
        x: x1,
        y: y1,
        width: x2 - x1,
        height: y2 - y1,
        color: color(markup.color),
        opacity: 0.35,
      })
    } else {
      const y = markup.type === 'underline' ? y1 + 1 : (y1 + y2) / 2
      page.drawLine({
        start: { x: x1, y },
        end: { x: x2, y },
        color: color(markup.color),
        thickness: 1,
      })
    }
  }
}

async function drawDrawing(
  document: PDFDocument,
  page: ReturnType<PDFDocument['getPage']>,
  drawing: DrawingInput,
): Promise<void> {
  if (drawing.kind === 'image') {
    const image = await document.embedPng(drawing.image)
    const [x1, y1, x2, y2] = drawing.rect
    page.drawImage(image, { x: x1, y: y1, width: x2 - x1, height: y2 - y1 })
    return
  }
  if (drawing.kind === 'note') {
    const [x, y] = drawing.at
    const annotation = document.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [x, y - 18, x + 20, y],
      Name: 'Comment',
      C: drawing.color,
      F: 4,
      P: page.ref,
    })
    annotation.set(PDFName.of('Contents'), PDFHexString.fromText(drawing.contents))
    annotation.set(PDFName.of('T'), PDFHexString.fromText(drawing.author?.trim() || 'GenOffice'))
    if (drawing.subject?.trim()) {
      annotation.set(PDFName.of('Subj'), PDFHexString.fromText(drawing.subject.trim()))
    }
    appendAnnotation(document, page, document.context.register(annotation))
    return
  }
  if (drawing.kind === 'ink') {
    for (const path of drawing.paths) {
      for (let index = 0; index + 3 < path.length; index += 2) {
        page.drawLine({
          start: { x: path[index]!, y: path[index + 1]! },
          end: { x: path[index + 2]!, y: path[index + 3]! },
          color: color(drawing.color),
          thickness: drawing.width,
        })
      }
    }
    return
  }
  if (drawing.kind === 'rect' || drawing.kind === 'ellipse') {
    const [x1, y1, x2, y2] = drawing.rect
    const left = Math.min(x1, x2)
    const bottom = Math.min(y1, y2)
    const width = Math.abs(x2 - x1)
    const height = Math.abs(y2 - y1)
    const options = {
      x: left,
      y: bottom,
      width,
      height,
      borderColor: color(drawing.color),
      borderWidth: drawing.width,
    }
    if (drawing.kind === 'rect') page.drawRectangle(options)
    else {
      page.drawEllipse({
        x: left + width / 2,
        y: bottom + height / 2,
        xScale: width / 2,
        yScale: height / 2,
        borderColor: options.borderColor,
        borderWidth: options.borderWidth,
      })
    }
    return
  }
  page.drawLine({
    start: { x: drawing.from[0], y: drawing.from[1] },
    end: { x: drawing.to[0], y: drawing.to[1] },
    color: color(drawing.color),
    thickness: drawing.width,
  })
}

async function drawStamp(
  document: PDFDocument,
  page: ReturnType<PDFDocument['getPage']>,
  stamp: StampInput,
  imageCache: Map<string, Awaited<ReturnType<PDFDocument['embedPng']>>>,
  stampImages: string[],
): Promise<void> {
  const imageData = stamp.image || stampImages[stamp.imageIndex ?? -1]
  if (!imageData) return
  let image = imageCache.get(imageData)
  if (!image) {
    image = await document.embedPng(imageData)
    imageCache.set(imageData, image)
  }
  const placement = stampDrawPlacement(stamp.rect, stamp.rotation)
  page.drawImage(image, {
    x: placement.x,
    y: placement.y,
    width: placement.width,
    height: placement.height,
    opacity: stamp.opacity ?? 1,
    rotate: degrees(placement.rotation),
  })
}

function applyForms(document: PDFDocument, request: SavePdfRequest): void {
  const form = document.getForm()
  for (const value of request.formValues) {
    try {
      const scalarValue = Array.isArray(value.value) ? (value.value[0] ?? '') : (value.value ?? '')
      if (value.kind === 'text') form.getTextField(value.name).setText(scalarValue)
      else if (value.kind === 'checkbox') {
        const field = form.getCheckBox(value.name)
        if (value.checked) field.check()
        else field.uncheck()
      } else if (value.kind === 'radio') {
        const field = form.getRadioGroup(value.name)
        if (scalarValue) field.select(scalarValue)
        else field.clear()
      } else {
        const field = form.getField(value.name)
        if (!(field instanceof PDFDropdown || field instanceof PDFOptionList)) continue
        const selection = Array.isArray(value.value)
          ? value.value.filter((entry) => entry.length > 0)
          : scalarValue
        if (Array.isArray(selection) ? selection.length > 0 : selection) field.select(selection)
        else field.clear()
      }
    } catch {
      // A stale or unsupported form field is skipped without blocking other edits.
    }
  }
}

export async function applyWebPdfSave(
  source: ArrayBuffer,
  request: SavePdfRequest,
): Promise<{
  bytes: Uint8Array
  skippedTextEdits: TextEditFailure[]
  skippedImageEdits: Array<{ editIndex: number; pageIndex: number; reason: string }>
}> {
  const document = await PDFDocument.load(source, { updateMetadata: false })
  const pages = document.getPages()
  applyForms(document, request)
  removeEmptyPdfSignatureFields(document, request.removeSignatureFields ?? [])

  for (const rotation of request.rotations || []) {
    const page = pages[rotation.pageIndex]
    if (page) page.setRotation(degrees((page.getRotation().angle + rotation.delta) % 360))
  }
  for (const markup of request.markups) {
    const page = pages[markup.pageIndex]
    if (page) drawMarkup(page, markup)
  }
  for (const drawing of request.drawings) {
    const page = pages[drawing.pageIndex]
    if (page) await drawDrawing(document, page, drawing)
  }
  const stampImageCache = new Map<string, Awaited<ReturnType<PDFDocument['embedPng']>>>()
  for (const stamp of request.stamps) {
    const page = pages[stamp.pageIndex]
    if (page) await drawStamp(document, page, stamp, stampImageCache, request.stampImages ?? [])
  }

  if (request.metadata) {
    if (request.metadata.title !== undefined) document.setTitle(request.metadata.title)
    if (request.metadata.author !== undefined) document.setAuthor(request.metadata.author)
    if (request.metadata.subject !== undefined) document.setSubject(request.metadata.subject)
    if (request.metadata.keywords !== undefined) {
      document.setKeywords(request.metadata.keywords.split(/[,，;；]/).map((item) => item.trim()))
    }
  }
  if (request.classification) {
    applyPdfClassificationMetadata(document, request.classification)
  }

  for (const pageIndex of [...(request.deletedPages || [])].sort((left, right) => right - left)) {
    if (pageIndex >= 0 && pageIndex < document.getPageCount() && document.getPageCount() > 1) {
      document.removePage(pageIndex)
    }
  }
  if (request.pageOrder?.length) {
    const deleted = new Set(request.deletedPages || [])
    const target = request.pageOrder
      .filter((pageIndex) => !deleted.has(pageIndex))
      .map((pageIndex) => pages[pageIndex])
      .filter((page) => page !== undefined)
    if (target.length === document.getPageCount()) {
      while (document.getPageCount()) document.removePage(0)
      for (const page of target) document.addPage(page)
    }
  }

  return {
    bytes: await document.save({ useObjectStreams: false }),
    skippedTextEdits: (request.textEdits || []).map((edit) => ({
      pageIndex: edit.pageIndex,
      oldText: edit.oldText,
      reason: 'Web 版暂不支持修改 PDF 内容流文本',
    })),
    skippedImageEdits: (request.imageEdits || []).map((edit, editIndex) => ({
      editIndex,
      pageIndex: edit.pageIndex,
      reason: 'Web 版暂不支持修改 PDF 内容流图片',
    })),
  }
}

export async function extractWebPdf(
  source: ArrayBuffer,
  pageIndexes: number[],
): Promise<Uint8Array> {
  return extractPagesBytes(source, pageIndexes)
}

export async function insertWebPdf(
  source: ArrayBuffer,
  addition: ArrayBuffer,
  afterPageIndex: number,
): Promise<{ bytes: Uint8Array; count: number }> {
  const result = await insertPdfBytes(source, addition, afterPageIndex)
  return { bytes: result.merged, count: result.count }
}

export async function insertWebPdfBlankPage(
  source: ArrayBuffer,
  afterPageIndex: number,
  options?: Parameters<typeof insertBlankPageBytes>[2],
): Promise<{ bytes: Uint8Array; count: number }> {
  const result = await insertBlankPageBytes(source, afterPageIndex, options)
  return { bytes: result.merged, count: result.count }
}
