import { PDFDocument, degrees, rgb } from 'pdf-lib'
import type {
  DrawingInput,
  MarkupInput,
  SavePdfRequest,
  StampInput,
  TextEditFailure,
} from '../../../pdf/src/shared/ipc'

const color = ([red, green, blue]: [number, number, number]) => rgb(red, green, blue)

function bounds(values: number[]): [number, number, number, number] {
  const xs = values.filter((_, index) => index % 2 === 0)
  const ys = values.filter((_, index) => index % 2 === 1)
  return [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]
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
    page.drawText(drawing.contents, {
      x: drawing.at[0],
      y: drawing.at[1],
      size: 9,
      color: color(drawing.color),
      maxWidth: 180,
    })
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
): Promise<void> {
  const image = await document.embedPng(stamp.image)
  const [x1, y1, x2, y2] = stamp.rect
  page.drawImage(image, {
    x: x1,
    y: y1,
    width: x2 - x1,
    height: y2 - y1,
    opacity: stamp.opacity ?? 1,
  })
}

function applyForms(document: PDFDocument, request: SavePdfRequest): void {
  const form = document.getForm()
  for (const value of request.formValues) {
    try {
      if (value.kind === 'text') form.getTextField(value.name).setText(value.value || '')
      else if (value.kind === 'checkbox') {
        const field = form.getCheckBox(value.name)
        if (value.checked) field.check()
        else field.uncheck()
      } else if (value.kind === 'radio') form.getRadioGroup(value.name).select(value.value || '')
      else form.getDropdown(value.name).select(value.value || '')
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
  for (const stamp of request.stamps) {
    const page = pages[stamp.pageIndex]
    if (page) await drawStamp(document, page, stamp)
  }

  if (request.metadata) {
    if (request.metadata.title !== undefined) document.setTitle(request.metadata.title)
    if (request.metadata.author !== undefined) document.setAuthor(request.metadata.author)
    if (request.metadata.subject !== undefined) document.setSubject(request.metadata.subject)
    if (request.metadata.keywords !== undefined) {
      document.setKeywords(request.metadata.keywords.split(/[,，;；]/).map((item) => item.trim()))
    }
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
  const input = await PDFDocument.load(source)
  const output = await PDFDocument.create()
  const pages = await output.copyPages(input, pageIndexes)
  for (const page of pages) output.addPage(page)
  return output.save({ useObjectStreams: false })
}

export async function insertWebPdf(
  source: ArrayBuffer,
  addition: ArrayBuffer,
  afterPageIndex: number,
): Promise<{ bytes: Uint8Array; count: number }> {
  const output = await PDFDocument.load(source)
  const input = await PDFDocument.load(addition)
  const pages = await output.copyPages(input, input.getPageIndices())
  let insertionIndex = Math.min(Math.max(afterPageIndex + 1, 0), output.getPageCount())
  for (const page of pages) output.insertPage(insertionIndex++, page)
  return { bytes: await output.save({ useObjectStreams: false }), count: pages.length }
}
