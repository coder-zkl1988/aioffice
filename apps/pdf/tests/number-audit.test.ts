import { describe, expect, it } from 'vitest'
import { auditNumbersInSearchIndex } from '../src/renderer/number-audit'
import type { SearchIndex } from '../src/renderer/search'

function indexFor(...pages: string[]): SearchIndex {
  return pages.map((text) => ({
    text,
    lower: text.toLocaleLowerCase(),
    items: text ? [{ start: 0, end: text.length, x: 0, y: 700, w: text.length * 6, h: 12 }] : [],
  }))
}

function tableIndex(rows: string[][], positions = [20, 170, 300, 430]): SearchIndex {
  let text = ''
  const items: SearchIndex[number]['items'] = []
  rows.forEach((row, rowIndex) => {
    row.forEach((cell, columnIndex) => {
      const start = text.length
      text += cell
      items.push({
        start,
        end: text.length,
        x: positions[columnIndex]!,
        y: 700 - rowIndex * 24,
        w: Math.max(12, cell.length * 7),
        h: 12,
      })
      text += columnIndex === row.length - 1 ? '\n' : ' '
    })
  })
  return [{ text, lower: text.toLocaleLowerCase(), items }]
}

describe('auditNumbersInSearchIndex', () => {
  it('checks addition, subtraction, currency, multiplication, division, and percentages', () => {
    const report = auditNumbersInSearchIndex(
      indexFor(
        [
          'Correct: £1,000 + £500 = £1,500',
          'Wrong: 500 + 300 = 900',
          'Net: 1000 - 250 = 750',
          'Margin: 100 × 20% = 25',
          'Average: 100 / 4 = 25',
          'Rate: 10% + 5% = 15.02%',
        ].join('\n'),
      ),
    )
    expect(report.expressionsChecked).toBe(6)
    expect(report.findings).toHaveLength(3)
    expect(report.findings[0]).toMatchObject({
      kind: 'arithmetic',
      stated: '900',
      expected: '800',
      pageIndex: 0,
    })
    expect(report.findings[1]).toMatchObject({ stated: '25', expected: '20' })
    expect(report.findings[2]).toMatchObject({ stated: '15.02%', expected: '15%' })
  })

  it('checks stated totals followed by parenthesized addends', () => {
    const report = auditNumbersInSearchIndex(
      indexFor('Grand Total: 900 (300 + 250 + 200)\n总计：600（200 + 200 + 200）'),
    )
    expect(report.expressionsChecked).toBe(2)
    expect(report.findings).toHaveLength(1)
    expect(report.findings[0]).toMatchObject({ stated: '900', expected: '750' })
  })

  it('flags cross-page conflicts for normalized named figures', () => {
    const report = auditNumbersInSearchIndex(
      indexFor('Net Profit: £1,200\nTotal: 1,200', 'net-profit —: £1,250\nTotal: 1,300'),
    )
    expect(report.namedFigures).toBe(2)
    expect(report.findings).toEqual([
      expect.objectContaining({
        kind: 'consistency',
        pageIndex: 1,
        canonicalPageIndex: 0,
        label: 'net profit',
        stated: '£1,250',
        expected: '£1,200',
      }),
    ])
  })

  it('ignores same-page repetitions, generic totals, and differences within tolerance', () => {
    const report = auditNumbersInSearchIndex(
      indexFor('VAT: 100\nVAT: 105\nTotal: 300', 'VAT: 100.005\nTotal: 400'),
    )
    expect(report.findings).toEqual([])
  })

  it('tracks the exact occurrence used for comment anchoring', () => {
    const text = '500 + 300 = 900\n500 + 300 = 900'
    const report = auditNumbersInSearchIndex(indexFor(text))
    expect(report.findings.map((finding) => finding.occurrence)).toEqual([1, 2])
  })

  it('caps findings and validates options', () => {
    const report = auditNumbersInSearchIndex(indexFor('1 + 1 = 3\n2 + 2 = 5'), {
      maxFindings: 1,
    })
    expect(report.findings).toHaveLength(1)
    expect(report.truncated).toBe(true)
    expect(() => auditNumbersInSearchIndex([], { tolerance: -1 })).toThrow('non-negative')
    expect(() => auditNumbersInSearchIndex([], { maxFindings: 0 })).toThrow('positive integer')
  })

  it('checks quantity multiplied by unit price when headers make the formula explicit', () => {
    const report = auditNumbersInSearchIndex(
      tableIndex([
        ['Item', 'Qty', 'Unit Price', 'Amount'],
        ['AURA One', '2', '$10', '$20'],
        ['AURA Mini', '3', '$5', '$99'],
      ]),
    )
    expect(report).toMatchObject({ tablesExamined: 1, tableChecks: 2 })
    expect(report.findings).toEqual([
      expect.objectContaining({
        kind: 'tableFormula',
        pageIndex: 0,
        tableNumber: 1,
        rowNumber: 3,
        columnNumber: 4,
        anchorText: '$99',
        stated: '$99',
        expected: '15',
      }),
    ])
  })

  it('checks numeric columns in explicit total rows', () => {
    const report = auditNumbersInSearchIndex(
      tableIndex([
        ['Item', 'Qty', 'Unit Price', 'Amount'],
        ['AURA One', '2', '$10', '$20'],
        ['AURA Mini', '3', '$5', '$15'],
        ['Total', '5', '', '$40'],
      ]),
    )
    expect(report.tableChecks).toBe(4)
    expect(report.findings).toEqual([
      expect.objectContaining({
        kind: 'tableTotal',
        rowNumber: 4,
        columnNumber: 4,
        anchorText: '$40',
        stated: '$40',
        expected: '35',
      }),
    ])
  })

  it('uses corrected row formulas when checking a dependent total', () => {
    const report = auditNumbersInSearchIndex(
      tableIndex([
        ['Item', 'Qty', 'Unit Price', 'Amount'],
        ['AURA One', '2', '$10', '$20'],
        ['AURA Mini', '3', '$5', '$99'],
        ['Total', '5', '', '$40'],
      ]),
    )
    expect(report.findings).toEqual([
      expect.objectContaining({ kind: 'tableFormula', stated: '$99', expected: '15' }),
      expect.objectContaining({ kind: 'tableTotal', stated: '$40', expected: '35' }),
    ])
  })

  it('skips inferred formulas when semantic headers or explicit total rows are absent', () => {
    const report = auditNumbersInSearchIndex(
      tableIndex([
        ['Product', 'Value A', 'Value B', 'Value C'],
        ['AURA One', '2', '10', '99'],
        ['AURA Mini', '3', '5', '15'],
      ]),
    )
    expect(report.tablesExamined).toBe(1)
    expect(report.tableChecks).toBe(0)
    expect(report.findings).toEqual([])
  })

  it('does not confuse total units with a standalone total amount column', () => {
    const report = auditNumbersInSearchIndex(
      tableIndex([
        ['Product', 'Total Units', 'Unit Price', 'Amount'],
        ['AURA One', '2', '10', '20'],
        ['AURA Mini', '3', '5', '15'],
      ]),
    )
    expect(report.tableChecks).toBe(2)
    expect(report.findings).toEqual([])
  })

  it('uses the exact repeated table-cell occurrence for comment anchoring', () => {
    const report = auditNumbersInSearchIndex(
      tableIndex([
        ['Item', 'Qty', 'Unit Price', 'Amount'],
        ['A', '2', '10', '99'],
        ['B', '3', '5', '99'],
      ]),
    )
    expect(
      report.findings
        .filter((finding) => finding.kind === 'tableFormula')
        .map((finding) => finding.occurrence),
    ).toEqual([1, 2])
  })
})
