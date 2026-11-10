/**
 * PDF ↔ Sheets table bridge: export one detected table as an .xlsx workbook,
 * then render the edited rows back into the page as a clean table image placed
 * over the original table's rectangle.
 *
 * The write-back goes through the ordinary 'addImageEdit' op (kind insertImage,
 * layer aboveText), so undo, the save pipeline and the image bake path all work
 * without new state buckets: the single opaque image carries the white cover
 * and the edited table together.
 */
import JSZip from 'jszip'
import type { PdfExtractedTable } from '@genoffice/pdf-tools'
import type { RunPdfToolRequest, RunPdfToolResult } from '../shared/ipc'
import { detectedTableBounds, type PdfDetectedTable } from './extract-tables'

/** Where one exported table went: kept in memory for the editing session only */
export interface TableBridgeRecord {
  sourcePath: string
  xlsxPath: string
  pageNumber: number
  tableNumber: number
  /** The original table's footprint in PDF user space (y up) */
  bounds: { x0: number; y0: number; x1: number; y1: number }
  /** Original column widths in PDF points, in column order ([] = even distribution) */
  columnWidths: number[]
  /** Original median row height in PDF points, used to size the edited table's text */
  rowHeight: number
  exportedAt: string
  /** The pending image edit that currently paints this table, when one is staged */
  bridgeEditId?: string
}

export interface TableBridgeRunner {
  runTool(request: RunPdfToolRequest): Promise<RunPdfToolResult>
}

export type BridgeExportResult =
  | { ok: true; savedPath: string }
  | { ok: false; canceled: true }
  | { ok: false; canceled?: false; error: string }

export interface BridgeExportRequest {
  filePath: string
  /** Name of the source PDF; main derives the workbook name from it */
  baseName: string
  table: PdfDetectedTable
  includeTwoColumnTextTables: boolean
}

/** One table → one sheet workbook; the table rows travel with the operation so the
    main process writes exactly this table, not every table on the page. */
export async function exportTableToXlsx(
  runner: TableBridgeRunner,
  request: BridgeExportRequest,
): Promise<BridgeExportResult> {
  const table: PdfExtractedTable = {
    pageNumber: request.table.pageNumber,
    tableNumber: request.table.tableNumber,
    rows: request.table.rows.map((row) => row.map((cell) => cell.text)),
  }
  const result = await runner.runTool({
    path: request.filePath,
    baseName: request.baseName,
    operation: {
      kind: 'extractTables',
      format: 'xlsx',
      pageIndexes: [request.table.pageNumber - 1],
      includeTwoColumnTextTables: request.includeTwoColumnTextTables,
      baseName: request.baseName,
      tables: [table],
    },
  })
  if (!result.ok) return { ok: false, error: result.error }
  if ('canceled' in result) return { ok: false, canceled: true }
  return { ok: true, savedPath: result.savedPath }
}

/** Safety margin around the detected table box when the edited table is painted:
    the detected box ends at the baseline, so descenders and old rules need cover */
export const TABLE_BRIDGE_PADDING_PT = 2

/** The image footprint for a bridge record: the table's box plus the padding */
export function padTableBounds(
  bounds: { x0: number; y0: number; x1: number; y1: number },
  pad = TABLE_BRIDGE_PADDING_PT,
): [number, number, number, number] {
  return [bounds.x0 - pad, bounds.y0 - pad, bounds.x1 + pad, bounds.y1 + pad]
}

/** The bridge record for an exported table, with the geometry the write-back needs */
export function tableBridgeRecord(
  sourcePath: string,
  xlsxPath: string,
  table: PdfDetectedTable,
  exportedAt = new Date().toISOString(),
): TableBridgeRecord {
  return {
    sourcePath,
    xlsxPath,
    pageNumber: table.pageNumber,
    tableNumber: table.tableNumber,
    bounds: detectedTableBounds(table),
    columnWidths: tableColumnWidths(table),
    rowHeight: tableMedianRowHeight(table),
    exportedAt,
  }
}

/**
 * Column widths in PDF points from the detected cells' run geometry: the gap
 * between two columns splits at the midpoint between the previous column's
 * rightmost ink and the next column's leftmost ink. [] when any column has no
 * runs (merged/empty cells), so the renderer falls back to even widths.
 */
export function tableColumnWidths(table: PdfDetectedTable): number[] {
  const columnCount = Math.max(0, ...table.rows.map((row) => row.length))
  if (columnCount < 2) return []
  const lefts = Array.from({ length: columnCount }, () => Number.POSITIVE_INFINITY)
  const rights = Array.from({ length: columnCount }, () => Number.NEGATIVE_INFINITY)
  for (const row of table.rows) {
    row.forEach((cell, column) => {
      for (const run of cell.runs) {
        lefts[column] = Math.min(lefts[column]!, run.x)
        rights[column] = Math.max(rights[column]!, run.x + run.width)
      }
    })
  }
  if (
    lefts.some((left) => !Number.isFinite(left)) ||
    rights.some((right) => !Number.isFinite(right))
  )
    return []
  const edges = [lefts[0]!]
  for (let column = 1; column < columnCount; column++)
    edges.push((rights[column - 1]! + lefts[column]!) / 2)
  edges.push(rights[columnCount - 1]!)
  const widths = edges.slice(1).map((edge, index) => edge - edges[index]!)
  return widths.every((width) => width > 0) ? widths : []
}

/** Median height of the original table's rows in PDF points (0 when no runs) */
export function tableMedianRowHeight(table: PdfDetectedTable): number {
  const heights = table.rows
    .map((row) => Math.max(0, ...row.flatMap((cell) => cell.runs.map((run) => run.height))))
    .filter((height) => height > 0)
  if (heights.length === 0) return 0
  const sorted = [...heights].sort((left, right) => left - right)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0 ? (sorted[middle - 1]! + sorted[middle]!) / 2 : sorted[middle]!
}

// ── workbook reading ──────────────────────────────────────────────────

const XLSX_MAX_ROWS = 10_000
const XLSX_MAX_COLUMNS = 1_000
const XLSX_MAX_PART_CHARS = 32 * 1024 * 1024

function decodeXmlText(value: string): string {
  return value.replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos);/gi, (match, entity: string) => {
    const lower = entity.toLowerCase()
    if (lower === 'amp') return '&'
    if (lower === 'lt') return '<'
    if (lower === 'gt') return '>'
    if (lower === 'quot') return '"'
    if (lower === 'apos') return "'"
    const code = lower.startsWith('#x')
      ? Number.parseInt(lower.slice(2), 16)
      : Number(lower.slice(1))
    return Number.isFinite(code) ? String.fromCodePoint(code) : match
  })
}

function attribute(tag: string, name: string): string | null {
  const pattern = new RegExp(`(?:^|\\s)${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i')
  const match = pattern.exec(tag)
  return match ? (match[1] ?? match[2] ?? '') : null
}

function zipEntry(zip: JSZip, path: string): JSZip.JSZipObject | null {
  const wanted = path.replace(/^\/+/, '').toLowerCase()
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue
    if (name.replace(/^\/+/, '').toLowerCase() === wanted) return entry
  }
  return null
}

async function readZipText(zip: JSZip, path: string): Promise<string | null> {
  const entry = zipEntry(zip, path)
  if (!entry) return null
  const text = await entry.async('string')
  return text.length > XLSX_MAX_PART_CHARS ? null : text
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function firstSheetPath(workbookXml: string, relsXml: string | null): string | null {
  const sheetTag = /<sheet\b[^>]*>/i.exec(workbookXml)?.[0] ?? ''
  const relationId = attribute(sheetTag, 'r:id')
  if (relationId && relsXml) {
    const relationshipPattern = new RegExp(
      `<Relationship\\b[^>]*\\bId\\s*=\\s*["']${escapeRegExp(relationId)}["'][^>]*>`,
      'i',
    )
    const relationshipTag = relationshipPattern.exec(relsXml)?.[0]
    const target = relationshipTag ? attribute(relationshipTag, 'Target') : null
    if (target) {
      const clean = target.replace(/^\/+/, '')
      return clean.startsWith('xl/') ? clean : `xl/${clean}`
    }
  }
  return null
}

function worksheetPaths(zip: JSZip): string[] {
  return Object.keys(zip.files)
    .filter((name) => /^\/?xl\/worksheets\/[^/]+\.xml$/i.test(name) && !zip.files[name]!.dir)
    .map((name) => name.replace(/^\/+/, ''))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }))
}

/** All <t> runs of a rich-text string, minus phonetic (rPh) readings */
function textRunsOf(xml: string): string {
  const body = xml.replace(/<rPh\b[\s\S]*?<\/rPh>/gi, '')
  let text = ''
  const textPattern = /<t\b[^>]*>([\s\S]*?)<\/t>/gi
  let run: RegExpExecArray | null
  while ((run = textPattern.exec(body)) !== null) text += decodeXmlText(run[1] ?? '')
  return text
}

function parseSharedStrings(xml: string): string[] {
  const strings: string[] = []
  const itemPattern = /<si\b([^>]*?)(?:\/>|>([\s\S]*?)<\/si>)/gi
  let match: RegExpExecArray | null
  while ((match = itemPattern.exec(xml)) !== null) strings.push(textRunsOf(match[2] ?? ''))
  return strings
}

function columnIndexFromReference(reference: string | null): number | null {
  const letters = reference ? /^([A-Z]+)\d+$/i.exec(reference)?.[1] : null
  if (!letters) return null
  let index = 0
  for (const character of letters.toUpperCase()) index = index * 26 + character.charCodeAt(0) - 64
  return index - 1
}

function cellValue(body: string, type: string | null, sharedStrings: string[]): string {
  const value = /<v\b[^>]*>([\s\S]*?)<\/v>/i.exec(body)?.[1]
  if (type === 'inlineStr') {
    const inline = /<is\b[^>]*>([\s\S]*?)<\/is>/i.exec(body)?.[1]
    return inline ? textRunsOf(inline) : ''
  }
  if (type === 's') {
    const index = Number(value ?? '')
    return Number.isInteger(index) ? (sharedStrings[index] ?? '') : ''
  }
  if (value === undefined) return ''
  if (type === 'b') return value === '1' ? 'TRUE' : 'FALSE'
  return decodeXmlText(value)
}

function worksheetRows(worksheetXml: string, sharedStrings: string[]): string[][] {
  const byRow = new Map<number, string[]>()
  const rowPattern = /<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/gi
  let rowMatch: RegExpExecArray | null
  let sequentialRow = 0
  while ((rowMatch = rowPattern.exec(worksheetXml)) !== null) {
    const rowAttributes = rowMatch[1] ?? ''
    const rowNumber = Number(attribute(rowAttributes, 'r') ?? '')
    const rowIndex = Number.isInteger(rowNumber) && rowNumber > 0 ? rowNumber - 1 : sequentialRow
    sequentialRow = rowIndex + 1
    if (rowIndex >= XLSX_MAX_ROWS) throw new Error('The workbook has too many rows')
    const cells: string[] = byRow.get(rowIndex) ?? []
    const body = rowMatch[2] ?? ''
    const cellPattern = /<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/gi
    let cellMatch: RegExpExecArray | null
    let sequentialColumn = 0
    while ((cellMatch = cellPattern.exec(body)) !== null) {
      const cellAttributes = cellMatch[1] ?? ''
      const column = columnIndexFromReference(attribute(cellAttributes, 'r')) ?? sequentialColumn
      sequentialColumn = column + 1
      if (column >= XLSX_MAX_COLUMNS) throw new Error('The workbook has too many columns')
      cells[column] = cellValue(cellMatch[2] ?? '', attribute(cellAttributes, 't'), sharedStrings)
    }
    byRow.set(rowIndex, cells)
  }
  const lastRow = Math.max(-1, ...byRow.keys())
  if (lastRow < 0) return []
  return Array.from({ length: lastRow + 1 }, (_, row) => {
    const cells = byRow.get(row) ?? []
    return Array.from({ length: cells.length }, (_, column) => cells[column] ?? '')
  })
}

/**
 * Read the first worksheet of an .xlsx as plain text rows: cell text as it
 * displays, with shared strings, inline strings and plain values all resolved.
 * Number/date formatting is not applied — the bridge only needs the text a
 * table cell carries (see risks on date cells).
 */
export async function xlsxRowsFromBytes(data: ArrayBuffer | Uint8Array): Promise<string[][]> {
  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(data)
  } catch {
    throw new Error('The workbook is not a readable .xlsx file')
  }
  const workbookXml = await readZipText(zip, 'xl/workbook.xml')
  if (!workbookXml) throw new Error('The workbook has no workbook part')
  const relsXml = await readZipText(zip, 'xl/_rels/workbook.xml.rels')
  const sharedStringsXml = await readZipText(zip, 'xl/sharedStrings.xml')
  const sharedStrings = sharedStringsXml ? parseSharedStrings(sharedStringsXml) : []
  const sheetPath = firstSheetPath(workbookXml, relsXml) ?? worksheetPaths(zip)[0]
  const worksheetXml = sheetPath ? await readZipText(zip, sheetPath) : null
  if (!worksheetXml) throw new Error('The workbook has no readable worksheet')
  return worksheetRows(worksheetXml, sharedStrings)
}

// ── table image ───────────────────────────────────────────────────────

/** Supersampling factor relative to PDF points, like stamps.ts */
const IMAGE_SS = 4
const MAX_IMAGE_PIXELS = 8192
// Document content: the rendered table is page data, so black ink on white paper
// stays hardcoded and never follows the UI theme (CLAUDE.md rule 4).
const TABLE_INK = '#000000'
const TABLE_PAPER = '#ffffff'
const TABLE_FONT =
  '"Helvetica Neue", Helvetica, Arial, "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif'

export interface TableImageOptions {
  /** Target footprint in PDF points (the original table's bounds) */
  widthPt: number
  heightPt: number
  /** Original column widths in PDF points; empty/omitted = even distribution */
  columnWidths?: readonly number[]
  /** Original median row height in PDF points; 0/omitted = the footprint's row height */
  rowHeightPt?: number
  supersample?: number
  maxPixels?: number
}

export interface TableImageLayout {
  /** Canvas size in pixels */
  width: number
  height: number
  /** Pixels per PDF point */
  scale: number
  columns: { x: number; width: number }[]
  /** Row band height in pixels (the edited rows share the footprint evenly) */
  rowHeight: number
  /** Base font size in pixels before per-cell shrink-to-fit */
  fontSize: number
  padding: number
  rowCount: number
}

/** Trim trailing empty rows/columns, pad every row to the same width, and reject a blank grid */
export function normalizeTableRows(rows: readonly (readonly string[])[]): string[][] {
  const grid = rows.map((row) => row.map((cell) => (cell == null ? '' : String(cell))))
  while (grid.length > 0 && grid[grid.length - 1]!.every((cell) => !cell.trim())) grid.pop()
  if (grid.length === 0) throw new Error('The workbook has no data rows')
  const columnCount = Math.max(...grid.map((row) => row.length))
  let width = columnCount
  while (width > 0 && grid.every((row) => !(row[width - 1] ?? '').trim())) width--
  if (width === 0) throw new Error('The workbook has no data rows')
  if (width > XLSX_MAX_COLUMNS) throw new Error('The workbook has too many columns')
  return grid.map((row) => Array.from({ length: width }, (_, column) => row[column] ?? ''))
}

function columnFractions(count: number, original?: readonly number[]): number[] {
  const widths = (original ?? []).filter((width) => Number.isFinite(width) && width > 0)
  if (widths.length === 0 || count === 0) return Array.from({ length: count }, () => 1 / count)
  let values: number[]
  if (count <= widths.length) values = widths.slice(0, count)
  else {
    const average = widths.reduce((sum, width) => sum + width, 0) / widths.length
    values = [...widths, ...Array.from({ length: count - widths.length }, () => average)]
  }
  const total = values.reduce((sum, width) => sum + width, 0)
  return values.map((width) => width / total)
}

/**
 * Pure layout for the write-back image: the edited rows are laid out inside the
 * original table's rectangle, column widths follow the original geometry when
 * supplied, and the font size is fitted to the original median row height (and
 * capped by the row band actually available).
 */
export function layoutTableImage(
  rows: readonly (readonly string[])[],
  options: TableImageOptions,
): TableImageLayout {
  const grid = normalizeTableRows(rows)
  const columnCount = grid[0]!.length
  const rowCount = grid.length
  const fractions = columnFractions(columnCount, options.columnWidths)

  const widthPt = Math.max(1, options.widthPt)
  const heightPt = Math.max(1, options.heightPt)
  const maxPixels = Math.max(16, options.maxPixels ?? MAX_IMAGE_PIXELS)
  let scale = Math.max(0.01, options.supersample ?? IMAGE_SS)
  scale = Math.min(scale, maxPixels / widthPt, maxPixels / heightPt)
  const width = Math.max(1, Math.round(widthPt * scale))
  const height = Math.max(1, Math.round(heightPt * scale))
  const scaleX = width / widthPt
  const scaleY = height / heightPt

  let cumulative = 0
  const columns = fractions.map((fraction, index) => {
    const x = Math.round(cumulative * width)
    cumulative += fraction
    const end = index === fractions.length - 1 ? width : Math.round(cumulative * width)
    return { x, width: Math.max(0, end - x) }
  })

  const rowHeight = height / rowCount
  const originalRowPt =
    options.rowHeightPt && options.rowHeightPt > 0 ? options.rowHeightPt : heightPt / rowCount
  const fontSize = Math.max(1, Math.min(originalRowPt * 0.62 * scaleY, rowHeight * 0.72))
  const narrowest = Math.min(...columns.map((column) => column.width))
  const padding = Math.max(
    0,
    Math.min(fontSize * 0.3, rowHeight * 0.2, narrowest > 0 ? narrowest * 0.25 : 0),
  )
  return {
    width,
    height,
    scale: Math.min(scaleX, scaleY),
    columns,
    rowHeight,
    fontSize,
    padding,
    rowCount,
  }
}

/** Break a cell's text into lines that fit `maxWidth`, splitting long words by character */
function wrappedLines(context: CanvasRenderingContext2D, text: string, maxWidth: number): string[] {
  if (maxWidth <= 0 || !text) return []
  if (context.measureText(text).width <= maxWidth) return [text]
  const words = text.split(/\s+/u).filter(Boolean)
  const lines: string[] = []
  let current = ''
  const push = (value: string) => {
    if (value) lines.push(value)
  }
  const appendWord = (word: string) => {
    if (!current) {
      let piece = ''
      for (const character of word) {
        if (piece && context.measureText(piece + character).width > maxWidth) {
          push(piece)
          piece = character
        } else piece += character
      }
      current = piece
      return
    }
    const candidate = `${current} ${word}`
    if (context.measureText(candidate).width <= maxWidth) current = candidate
    else {
      push(current)
      current = ''
      appendWord(word)
    }
  }
  for (const word of words.length > 0 ? words : [text]) appendWord(word)
  push(current)
  return lines
}

/** Draw the edited rows as a clean bordered table on a white canvas; returns base64 PNG */
export function renderTableImage(
  rows: readonly (readonly string[])[],
  layout: TableImageLayout,
): string {
  const grid = normalizeTableRows(rows)
  const canvas = document.createElement('canvas')
  canvas.width = layout.width
  canvas.height = layout.height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas is unavailable in this view')

  context.fillStyle = TABLE_PAPER
  context.fillRect(0, 0, layout.width, layout.height)
  context.textBaseline = 'middle'
  context.textAlign = 'left'

  for (let row = 0; row < grid.length; row++) {
    const top = Math.round((row * layout.height) / grid.length)
    const bottom = Math.round(((row + 1) * layout.height) / grid.length)
    const bandHeight = Math.max(1, bottom - top)
    const bold = row === 0
    for (let column = 0; column < layout.columns.length; column++) {
      const box = layout.columns[column]!
      const text = (grid[row]![column] ?? '').trim()
      if (!text || box.width <= 0) continue
      let fontSize = layout.fontSize
      let lines: string[] = []
      for (let attempt = 0; attempt < 8; attempt++) {
        context.font = `${bold ? '600 ' : ''}${fontSize}px ${TABLE_FONT}`
        lines = wrappedLines(context, text, Math.max(1, box.width - layout.padding * 2))
        const lineHeight = fontSize * 1.2
        const maxLines = Math.max(1, Math.floor((bandHeight - layout.padding * 2) / lineHeight))
        if (lines.length <= maxLines || fontSize <= 1.5) {
          lines = lines.slice(0, maxLines)
          break
        }
        fontSize *= 0.85
      }
      const lineHeight = fontSize * 1.2
      const totalHeight = lineHeight * lines.length
      const startY = top + (bandHeight - totalHeight) / 2 + lineHeight / 2
      context.fillStyle = TABLE_INK
      lines.forEach((line, index) => {
        context.fillText(line, box.x + layout.padding, startY + index * lineHeight)
      })
    }
  }

  // Grid last so the rules stay crisp over any text
  context.strokeStyle = TABLE_INK
  context.lineWidth = Math.max(1, Math.round(layout.scale * 0.75))
  for (let row = 0; row <= grid.length; row++) {
    const y = Math.round((row * layout.height) / grid.length)
    context.beginPath()
    context.moveTo(0, y)
    context.lineTo(layout.width, y)
    context.stroke()
  }
  for (const box of layout.columns) {
    context.beginPath()
    context.moveTo(box.x, 0)
    context.lineTo(box.x, layout.height)
    context.stroke()
  }
  context.beginPath()
  context.moveTo(layout.width - 0.5, 0)
  context.lineTo(layout.width - 0.5, layout.height)
  context.stroke()

  return canvas.toDataURL('image/png').split(',')[1] ?? ''
}
