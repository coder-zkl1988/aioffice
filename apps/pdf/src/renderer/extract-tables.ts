import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { PdfExtractedTable } from '@genoffice/pdf-tools'

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
}

export interface PdfTableTextRun {
  text: string
  x: number
  y: number
  width: number
  height: number
  sourceStart?: number
  sourceEnd?: number
}

export interface PdfDetectedTableCell {
  text: string
  runs: PdfTableTextRun[]
}

export interface PdfDetectedTable {
  pageNumber: number
  tableNumber: number
  rows: PdfDetectedTableCell[][]
}

interface PdfTableTextRow {
  runs: PdfTableTextRun[]
  y: number
  height: number
}

interface CandidateRow extends PdfTableTextRow {
  boundaries: number[]
}

export interface PdfTableDetectionOptions {
  includeTwoColumnTextTables: boolean
}

function median(values: number[]): number {
  if (values.length === 0) return 0
  const sorted = [...values].sort((left, right) => left - right)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0 ? (sorted[middle - 1]! + sorted[middle]!) / 2 : sorted[middle]!
}

function averageCharacterWidth(runs: PdfTableTextRun[]): number {
  const width = runs.reduce((sum, run) => sum + Math.max(0, run.width), 0)
  const characters = runs.reduce((sum, run) => sum + Math.max(1, run.text.trim().length), 0)
  return characters === 0 ? 6 : width / characters
}

function joinRuns(runs: PdfTableTextRun[]): string {
  const ordered = [...runs].sort((left, right) => left.x - right.x)
  let result = ''
  let previous: PdfTableTextRun | undefined
  for (const run of ordered) {
    const text = run.text.replace(/\s+/gu, ' ').trim()
    if (!text) continue
    const gap = previous ? run.x - (previous.x + previous.width) : 0
    if (result && gap > Math.max(1, averageCharacterWidth([previous!, run]) * 0.35)) result += ' '
    result += text
    previous = run
  }
  return result
}

export function groupPdfTableRows(runs: PdfTableTextRun[]): PdfTableTextRow[] {
  const rows: PdfTableTextRun[][] = []
  for (const run of [...runs].sort((left, right) => right.y - left.y || left.x - right.x)) {
    if (!run.text.trim() || run.height <= 0 || run.width < 0) continue
    const row = rows.find((candidate) => {
      const anchor = candidate[0]!
      return Math.abs(anchor.y - run.y) <= Math.max(2, Math.min(anchor.height, run.height) * 0.35)
    })
    if (row) row.push(run)
    else rows.push([run])
  }
  return rows.map((row) => ({
    runs: row.sort((left, right) => left.x - right.x),
    y: row.reduce((sum, run) => sum + run.y, 0) / row.length,
    height: Math.max(...row.map((run) => run.height)),
  }))
}

function candidateRow(row: PdfTableTextRow): CandidateRow | null {
  if (row.runs.length < 2) return null
  const minimumGap = Math.max(10, averageCharacterWidth(row.runs) * 2.5)
  const boundaries: number[] = []
  for (let index = 1; index < row.runs.length; index++) {
    const left = row.runs[index - 1]!
    const right = row.runs[index]!
    const gap = right.x - (left.x + left.width)
    if (gap >= minimumGap) boundaries.push(left.x + left.width + gap / 2)
  }
  return boundaries.length > 0 ? { ...row, boundaries } : null
}

function splitCandidateBlocks(rows: PdfTableTextRow[]): CandidateRow[][] {
  const candidates = rows.map(candidateRow)
  const bodyHeight = median(rows.map((row) => row.height).filter((height) => height > 0)) || 12
  const blocks: CandidateRow[][] = []
  let current: CandidateRow[] = []
  let previousIndex = -2
  for (let index = 0; index < candidates.length; index++) {
    const candidate = candidates[index]
    if (!candidate) {
      if (current.length >= 2) blocks.push(current)
      current = []
      previousIndex = -2
      continue
    }
    const previous = current.at(-1)
    const verticalGap = previous ? previous.y - candidate.y : 0
    if (
      current.length > 0 &&
      (index !== previousIndex + 1 || verticalGap > Math.max(36, bodyHeight * 3))
    ) {
      if (current.length >= 2) blocks.push(current)
      current = []
    }
    current.push(candidate)
    previousIndex = index
  }
  if (current.length >= 2) blocks.push(current)
  return blocks
}

function supportedBoundaries(rows: CandidateRow[]): number[] {
  const tolerance = Math.max(8, median(rows.map((row) => row.height)) * 1.2)
  const clusters: number[][] = []
  for (const boundary of rows.flatMap((row) => row.boundaries)) {
    const cluster = clusters.find((values) => Math.abs(median(values) - boundary) <= tolerance)
    if (cluster) cluster.push(boundary)
    else clusters.push([boundary])
  }
  const support = Math.max(2, Math.ceil(rows.length * 0.5))
  return clusters
    .filter((cluster) => cluster.length >= support)
    .map(median)
    .sort((left, right) => left - right)
}

function dataLike(value: string): boolean {
  const normalized = value.trim()
  return (
    /^[-+]?[$€£¥￥]?\s*\d[\d,.]*(?:\s*%|\s*[A-Z]{3})?$/u.test(normalized) ||
    /^\d{1,4}[-/.]\d{1,2}(?:[-/.]\d{1,4})?$/u.test(normalized) ||
    /^(?:yes|no|true|false|是|否)$/iu.test(normalized)
  )
}

function detailedRowsForBoundaries(
  rows: CandidateRow[],
  boundaries: number[],
): PdfDetectedTableCell[][] {
  return rows.map((row) => {
    const cells = Array.from({ length: boundaries.length + 1 }, () => [] as PdfTableTextRun[])
    for (const run of row.runs) {
      const center = run.x + run.width / 2
      const column = boundaries.findIndex((boundary) => center < boundary)
      cells[column < 0 ? cells.length - 1 : column]!.push(run)
    }
    return cells.map((runs) => ({ text: joinRuns(runs), runs }))
  })
}

function validTableRows(rows: string[][], includeTwoColumnTextTables: boolean): boolean {
  if (rows.length < 2 || rows[0]!.length < 2) return false
  const columnCount = rows[0]!.length
  const populatedRows = rows.filter((row) => row.filter((cell) => cell.trim()).length >= 2)
  if (populatedRows.length < Math.max(2, Math.ceil(rows.length * 0.75))) return false
  if (columnCount > 2 || includeTwoColumnTextTables) return true
  const bodyCells = rows.slice(1).flat()
  return bodyCells.length > 0 && bodyCells.filter(dataLike).length / bodyCells.length >= 0.25
}

export function detectPdfTablesFromRuns(
  pageNumber: number,
  runs: PdfTableTextRun[],
  options: PdfTableDetectionOptions,
): PdfExtractedTable[] {
  return detectPdfTableDetailsFromRuns(pageNumber, runs, options).map((table) => ({
    pageNumber: table.pageNumber,
    tableNumber: table.tableNumber,
    rows: table.rows.map((row) => row.map((cell) => cell.text)),
  }))
}

export function detectPdfTableDetailsFromRuns(
  pageNumber: number,
  runs: PdfTableTextRun[],
  options: PdfTableDetectionOptions,
): PdfDetectedTable[] {
  const rows = groupPdfTableRows(runs)
  const tables: PdfDetectedTable[] = []
  for (const block of splitCandidateBlocks(rows)) {
    const boundaries = supportedBoundaries(block)
    if (boundaries.length === 0) continue
    const tableRows = detailedRowsForBoundaries(block, boundaries)
    if (
      !validTableRows(
        tableRows.map((row) => row.map((cell) => cell.text)),
        options.includeTwoColumnTextTables,
      )
    )
      continue
    tables.push({
      pageNumber,
      tableNumber: tables.length + 1,
      rows: tableRows,
    })
  }
  return tables
}

function textRuns(items: RawTextItem[]): PdfTableTextRun[] {
  return items.flatMap((item) => {
    if (!item.str?.trim() || !item.transform || item.transform.length < 6) return []
    const height = item.height || Math.hypot(item.transform[2] ?? 0, item.transform[3] ?? 0) || 1
    return [
      {
        text: item.str,
        x: item.transform[4] ?? 0,
        y: item.transform[5] ?? 0,
        width: Math.max(0, item.width ?? 0),
        height,
      },
    ]
  })
}

export async function extractPdfTables(
  document: PDFDocumentProxy,
  pageIndexes: number[],
  options: PdfTableDetectionOptions,
): Promise<PdfExtractedTable[]> {
  const tables: PdfExtractedTable[] = []
  for (const pageIndex of pageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= document.numPages) {
      throw new Error('Table export page indexes contain an invalid page')
    }
    const page = await document.getPage(pageIndex + 1)
    try {
      const content = await page.getTextContent()
      tables.push(
        ...detectPdfTablesFromRuns(
          pageIndex + 1,
          textRuns(content.items as RawTextItem[]),
          options,
        ),
      )
    } finally {
      page.cleanup()
    }
  }
  return tables
}
