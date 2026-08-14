import type { SearchIndex } from './search'
import {
  detectPdfTableDetailsFromRuns,
  type PdfDetectedTable,
  type PdfDetectedTableCell,
  type PdfTableTextRun,
} from './extract-tables'

export type NumberAuditFindingKind = 'arithmetic' | 'consistency' | 'tableFormula' | 'tableTotal'

export interface NumberAuditFinding {
  kind: NumberAuditFindingKind
  pageIndex: number
  anchorText: string
  occurrence: number
  stated: string
  expected: string
  label?: string
  canonicalPageIndex?: number
  tableNumber?: number
  rowNumber?: number
  columnNumber?: number
}

export interface NumberAuditReport {
  pagesExamined: number
  expressionsChecked: number
  namedFigures: number
  tablesExamined: number
  tableChecks: number
  findings: NumberAuditFinding[]
  truncated: boolean
}

interface ParsedNumber {
  value: number
  percent: boolean
}

interface NamedFigure {
  label: string
  value: number
  rawValue: string
  pageIndex: number
  anchorText: string
  occurrence: number
}

const CURRENCY = '£$€¥￥'
const NUMBER_SOURCE = `[${CURRENCY}]?[ \\t]*-?(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?%?`
const EXPRESSION_SOURCE = `${NUMBER_SOURCE}(?:[ \\t]*[+\\-×xX*/÷][ \\t]*${NUMBER_SOURCE})+`
const EQUALS_EXPRESSION = new RegExp(
  `(${EXPRESSION_SOURCE})[ \\t]*[=＝][ \\t]*(${NUMBER_SOURCE})`,
  'giu',
)
const TOTAL_EXPRESSION = new RegExp(
  `(?:grand[ \\t]+total|subtotal|total|sum|合计|總計|总计|小计|小計|总和|總和)[ \\t]*[:：\\-]?[ \\t]*(${NUMBER_SOURCE})[ \\t]*[（(](${EXPRESSION_SOURCE})[）)]`,
  'giu',
)
const NAMED_FIGURE = new RegExp(`^\\s*([^:：\\n]{2,80}?)\\s*[:：]\\s*(${NUMBER_SOURCE})`, 'u')
const NUMBER_AT_START = new RegExp(
  `^(?:[${CURRENCY}][ \\t]*[+\\-]?|[+\\-][ \\t]*[${CURRENCY}]?|[+\\-]?)(?:\\d{1,3}(?:,\\d{3})+|\\d+)(?:\\.\\d+)?%?`,
  'u',
)
const GENERIC_LABELS = new Set([
  'total',
  'grand total',
  'subtotal',
  'sum',
  '合计',
  '總計',
  '总计',
  '小计',
  '小計',
  '总和',
  '總和',
])
const TOTAL_ROW_LABELS = new Set([...GENERIC_LABELS, '合計', 'grandtotal', 'sub total', '合 計'])
const QUANTITY_HEADERS = [
  'qty',
  'quantity',
  'units',
  'unit count',
  'count',
  '数量',
  '數量',
  '件数',
  '件數',
]
const PRICE_HEADERS = [
  'unit price',
  'price per unit',
  'unit cost',
  'price',
  'rate',
  '单价',
  '單價',
  '单位成本',
  '單位成本',
]
const AMOUNT_HEADERS = [
  'line total',
  'line amount',
  'extended price',
  'amount',
  'total price',
  'total cost',
  '金额',
  '金額',
  '价税合计',
  '價稅合計',
]

function parseNumber(raw: string): ParsedNumber | null {
  const trimmed = raw.trim()
  const parenthesized = trimmed.startsWith('(') && trimmed.endsWith(')')
  const percent = trimmed.includes('%')
  const unsigned = parenthesized ? trimmed.slice(1, -1) : trimmed
  const cleaned = unsigned.replace(new RegExp(`[${CURRENCY},\\s]`, 'gu'), '').replace('%', '')
  if (!cleaned || cleaned === '-' || cleaned === '+') return null
  const value = Number(cleaned) * (parenthesized ? -1 : 1)
  if (!Number.isFinite(value)) return null
  return { value: percent ? value / 100 : value, percent }
}

function evaluateExpression(expression: string): number | null {
  const values: number[] = []
  const operators: string[] = []
  let cursor = 0
  let expectsNumber = true

  while (cursor < expression.length) {
    while (/\s/u.test(expression[cursor] ?? '')) cursor += 1
    if (cursor >= expression.length) break

    if (expectsNumber) {
      const match = NUMBER_AT_START.exec(expression.slice(cursor))
      if (!match) return null
      const parsed = parseNumber(match[0])
      if (!parsed) return null
      values.push(parsed.value)
      cursor += match[0].length
      expectsNumber = false
      continue
    }

    const operator = expression[cursor]
    if (!operator || !'+-×xX*/÷'.includes(operator)) return null
    operators.push(operator)
    cursor += 1
    expectsNumber = true
  }

  if (expectsNumber || values.length !== operators.length + 1) return null

  const collapsedValues = [values[0]!]
  const collapsedOperators: string[] = []
  for (let index = 0; index < operators.length; index += 1) {
    const operator = operators[index]!
    const right = values[index + 1]!
    if ('×xX*÷/'.includes(operator)) {
      const left = collapsedValues.pop()!
      if ('÷/'.includes(operator) && right === 0) return null
      collapsedValues.push('×xX*'.includes(operator) ? left * right : left / right)
    } else {
      collapsedOperators.push(operator)
      collapsedValues.push(right)
    }
  }

  let result = collapsedValues[0]!
  for (let index = 0; index < collapsedOperators.length; index += 1) {
    result =
      collapsedOperators[index] === '+'
        ? result + collapsedValues[index + 1]!
        : result - collapsedValues[index + 1]!
  }
  return Number.isFinite(result) ? result : null
}

function formatNumber(value: number, percent: boolean): string {
  const scaled = percent ? value * 100 : value
  const normalized = Math.abs(scaled) < 1e-12 ? 0 : scaled
  const formatted = normalized
    .toFixed(8)
    .replace(/\.0+$/u, '')
    .replace(/(\.\d*?[1-9])0+$/u, '$1')
  return `${formatted}${percent ? '%' : ''}`
}

function occurrenceAt(text: string, anchor: string, targetIndex: number): number {
  const lower = text.toLocaleLowerCase()
  const query = anchor.toLocaleLowerCase()
  let occurrence = 0
  let from = 0
  while (from <= targetIndex) {
    const index = lower.indexOf(query, from)
    if (index < 0 || index > targetIndex) break
    occurrence += 1
    from = index + query.length
  }
  return Math.max(1, occurrence)
}

function normalizeLabel(label: string): string {
  return label
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/^[\s•·▪◦*-]+/u, '')
    .replace(/[:：\-—–_\s]+/gu, ' ')
    .trim()
}

function isSpecificLabel(label: string): boolean {
  if (!label || GENERIC_LABELS.has(label)) return false
  const cjkCount = (label.match(/[\p{Script=Han}]/gu) ?? []).length
  const letterCount = (label.match(/[\p{L}]/gu) ?? []).length
  return cjkCount >= 2 || letterCount >= 3
}

function differs(left: number, right: number, tolerance: number): boolean {
  return Math.abs(left - right) > tolerance
}

function normalizedHeader(value: string): string {
  return value
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/[_:：\-—–]+/gu, ' ')
    .replace(/\s+/gu, ' ')
    .trim()
}

function headerMatches(value: string, candidates: string[]): boolean {
  const header = normalizedHeader(value)
  return candidates.some((candidate) => header === candidate || header.includes(candidate))
}

function totalRowLabel(row: PdfDetectedTableCell[]): string | null {
  const firstTextCell = row.find((cell) => cell.text.trim() && !parseNumber(cell.text))
  if (!firstTextCell) return null
  const label = normalizedHeader(firstTextCell.text)
  return TOTAL_ROW_LABELS.has(label) ? label : null
}

function tableRuns(page: SearchIndex[number]): PdfTableTextRun[] {
  return page.items.flatMap((item) => {
    if (item.rot || item.end <= item.start) return []
    const text = page.text.slice(item.start, item.end)
    if (!text.trim()) return []
    return [
      {
        text,
        x: item.x,
        y: item.y,
        width: item.w,
        height: item.h,
        sourceStart: item.start,
        sourceEnd: item.end,
      },
    ]
  })
}

function cellAnchor(
  pageText: string,
  cell: PdfDetectedTableCell,
): { anchorText: string; occurrence: number } | null {
  const run = [...cell.runs]
    .filter(
      (candidate) =>
        candidate.sourceStart !== undefined &&
        candidate.sourceEnd !== undefined &&
        candidate.sourceEnd > candidate.sourceStart,
    )
    .sort((left, right) => right.text.trim().length - left.text.trim().length)[0]
  if (!run || run.sourceStart === undefined || run.sourceEnd === undefined) return null
  const raw = pageText.slice(run.sourceStart, run.sourceEnd)
  const leading = raw.length - raw.trimStart().length
  const anchorText = raw.trim()
  if (!anchorText) return null
  return {
    anchorText,
    occurrence: occurrenceAt(pageText, anchorText, run.sourceStart + leading),
  }
}

function findHeaderColumn(
  header: PdfDetectedTableCell[],
  candidates: string[],
  excluded: Set<number> = new Set(),
): number {
  return header.findIndex(
    (cell, columnIndex) => !excluded.has(columnIndex) && headerMatches(cell.text, candidates),
  )
}

function findAmountColumn(header: PdfDetectedTableCell[]): number {
  return header.findIndex((cell) => {
    const normalized = normalizedHeader(cell.text)
    return normalized === 'total' || headerMatches(cell.text, AMOUNT_HEADERS)
  })
}

function auditTable(
  table: PdfDetectedTable,
  pageText: string,
  tolerance: number,
  addFinding: (finding: NumberAuditFinding) => void,
): number {
  const header = table.rows[0]
  if (!header) return 0
  let checks = 0
  const formulaValues = new Map<string, number>()

  const amountColumn = findAmountColumn(header)
  const quantityColumn = findHeaderColumn(
    header,
    QUANTITY_HEADERS,
    new Set(amountColumn >= 0 ? [amountColumn] : []),
  )
  const priceColumn = findHeaderColumn(
    header,
    PRICE_HEADERS,
    new Set([amountColumn, quantityColumn].filter((column) => column >= 0)),
  )

  if (quantityColumn >= 0 && priceColumn >= 0 && amountColumn >= 0) {
    table.rows.slice(1).forEach((row, bodyIndex) => {
      if (totalRowLabel(row)) return
      const quantity = parseNumber(row[quantityColumn]?.text ?? '')
      const price = parseNumber(row[priceColumn]?.text ?? '')
      const amountCell = row[amountColumn]
      const amount = parseNumber(amountCell?.text ?? '')
      if (!quantity || !price || !amount || !amountCell) return
      checks += 1
      const expected = quantity.value * price.value
      formulaValues.set(`${bodyIndex + 1}:${amountColumn}`, expected)
      if (!differs(expected, amount.value, tolerance)) return
      const anchor = cellAnchor(pageText, amountCell)
      if (!anchor) return
      addFinding({
        kind: 'tableFormula',
        pageIndex: table.pageNumber - 1,
        ...anchor,
        stated: amountCell.text.trim(),
        expected: formatNumber(expected, amount.percent),
        label: header[amountColumn]?.text.trim(),
        tableNumber: table.tableNumber,
        rowNumber: bodyIndex + 2,
        columnNumber: amountColumn + 1,
      })
    })
  }

  let segmentStart = 1
  for (let rowIndex = 1; rowIndex < table.rows.length; rowIndex += 1) {
    const row = table.rows[rowIndex]!
    if (!totalRowLabel(row)) continue
    for (let columnIndex = 1; columnIndex < row.length; columnIndex += 1) {
      const totalCell = row[columnIndex]
      const stated = parseNumber(totalCell?.text ?? '')
      if (!totalCell || !stated) continue
      const addends = table.rows
        .slice(segmentStart, rowIndex)
        .map((candidate, offset) => ({ candidate, sourceRow: segmentStart + offset }))
        .filter(({ candidate }) => !totalRowLabel(candidate))
        .map(({ candidate, sourceRow }) => {
          const formulaValue = formulaValues.get(`${sourceRow}:${columnIndex}`)
          return formulaValue === undefined
            ? parseNumber(candidate[columnIndex]?.text ?? '')
            : ({ value: formulaValue, percent: false } satisfies ParsedNumber)
        })
        .filter((value): value is ParsedNumber => value !== null)
      if (addends.length < 2) continue
      checks += 1
      const expected = addends.reduce((sum, value) => sum + value.value, 0)
      if (!differs(expected, stated.value, stated.percent ? tolerance / 100 : tolerance)) continue
      const anchor = cellAnchor(pageText, totalCell)
      if (!anchor) continue
      addFinding({
        kind: 'tableTotal',
        pageIndex: table.pageNumber - 1,
        ...anchor,
        stated: totalCell.text.trim(),
        expected: formatNumber(expected, stated.percent),
        label: header[columnIndex]?.text.trim(),
        tableNumber: table.tableNumber,
        rowNumber: rowIndex + 1,
        columnNumber: columnIndex + 1,
      })
    }
    segmentStart = rowIndex + 1
  }

  return checks
}

export function auditNumbersInSearchIndex(
  index: SearchIndex,
  options: { tolerance?: number; maxFindings?: number } = {},
): NumberAuditReport {
  const tolerance = options.tolerance ?? 0.01
  const maxFindings = options.maxFindings ?? 100
  if (!Number.isFinite(tolerance) || tolerance < 0)
    throw new Error('tolerance must be non-negative')
  if (!Number.isInteger(maxFindings) || maxFindings < 1) {
    throw new Error('maxFindings must be a positive integer')
  }

  const findings: NumberAuditFinding[] = []
  const figures = new Map<string, NamedFigure[]>()
  let expressionsChecked = 0
  let namedFigures = 0
  let tablesExamined = 0
  let tableChecks = 0
  let truncated = false

  const addFinding = (finding: NumberAuditFinding): void => {
    if (findings.length < maxFindings) findings.push(finding)
    else truncated = true
  }

  index.forEach((page, pageIndex) => {
    for (const match of page.text.matchAll(EQUALS_EXPRESSION)) {
      expressionsChecked += 1
      const expression = match[1] ?? ''
      const statedRaw = (match[2] ?? '').trim()
      const computed = evaluateExpression(expression)
      const stated = parseNumber(statedRaw)
      if (
        computed === null ||
        !stated ||
        !differs(computed, stated.value, stated.percent ? tolerance / 100 : tolerance)
      )
        continue
      const anchorText = match[0].trim()
      addFinding({
        kind: 'arithmetic',
        pageIndex,
        anchorText,
        occurrence: occurrenceAt(page.text, anchorText, match.index ?? 0),
        stated: statedRaw,
        expected: formatNumber(computed, stated.percent),
      })
    }

    for (const match of page.text.matchAll(TOTAL_EXPRESSION)) {
      expressionsChecked += 1
      const statedRaw = (match[1] ?? '').trim()
      const expression = match[2] ?? ''
      const computed = evaluateExpression(expression)
      const stated = parseNumber(statedRaw)
      if (
        computed === null ||
        !stated ||
        !differs(computed, stated.value, stated.percent ? tolerance / 100 : tolerance)
      )
        continue
      const anchorText = match[0].trim()
      addFinding({
        kind: 'arithmetic',
        pageIndex,
        anchorText,
        occurrence: occurrenceAt(page.text, anchorText, match.index ?? 0),
        stated: statedRaw,
        expected: formatNumber(computed, stated.percent),
      })
    }

    for (const lineMatch of page.text.matchAll(/[^\n]+/gu)) {
      const line = lineMatch[0]
      const match = NAMED_FIGURE.exec(line)
      if (!match) continue
      const label = normalizeLabel(match[1] ?? '')
      if (!isSpecificLabel(label)) continue
      const rawValue = (match[2] ?? '').trim()
      const parsed = parseNumber(rawValue)
      if (!parsed) continue
      const anchorText = match[0].trim()
      const sourceIndex = (lineMatch.index ?? 0) + (match.index ?? 0) + line.indexOf(anchorText)
      figures.set(label, [
        ...(figures.get(label) ?? []),
        {
          label,
          value: parsed.value,
          rawValue,
          pageIndex,
          anchorText,
          occurrence: occurrenceAt(page.text, anchorText, sourceIndex),
        },
      ])
      namedFigures += 1
    }

    const tables = detectPdfTableDetailsFromRuns(pageIndex + 1, tableRuns(page), {
      includeTwoColumnTextTables: false,
    })
    tablesExamined += tables.length
    for (const table of tables) {
      tableChecks += auditTable(table, page.text, tolerance, addFinding)
    }
  })

  for (const records of figures.values()) {
    const canonical = records[0]
    if (!canonical) continue
    for (const record of records.slice(1)) {
      if (record.pageIndex === canonical.pageIndex) continue
      if (!differs(record.value, canonical.value, tolerance)) continue
      addFinding({
        kind: 'consistency',
        pageIndex: record.pageIndex,
        anchorText: record.anchorText,
        occurrence: record.occurrence,
        stated: record.rawValue,
        expected: canonical.rawValue,
        label: record.label,
        canonicalPageIndex: canonical.pageIndex,
      })
    }
  }

  return {
    pagesExamined: index.length,
    expressionsChecked,
    namedFigures,
    tablesExamined,
    tableChecks,
    findings,
    truncated,
  }
}
