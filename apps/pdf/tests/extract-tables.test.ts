import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  detectPdfTableDetailsFromRuns,
  detectPdfTablesFromRuns,
  extractPdfTables,
  type PdfTableTextRun,
} from '../src/renderer/extract-tables'

function run(text: string, x: number, y: number, width = Math.max(10, text.length * 7)) {
  return { text, x, y, width, height: 12 }
}

function row(y: number, values: string[], positions = [20, 180, 320]): PdfTableTextRun[] {
  return values.map((value, index) => run(value, positions[index]!, y))
}

function item(text: string, x: number, y: number, width = Math.max(10, text.length * 7)) {
  return { str: text, transform: [12, 0, 0, 12, x, y], width, height: 12 }
}

describe('PDF table extraction', () => {
  it('detects aligned three-column rows as a table', () => {
    const tables = detectPdfTablesFromRuns(
      2,
      [
        ...row(700, ['Product', 'Qty', 'Price']),
        ...row(680, ['AURA One', '12', '499']),
        ...row(660, ['AURA Mini', '8', '299']),
      ],
      { includeTwoColumnTextTables: false },
    )
    expect(tables).toEqual([
      {
        pageNumber: 2,
        tableNumber: 1,
        rows: [
          ['Product', 'Qty', 'Price'],
          ['AURA One', '12', '499'],
          ['AURA Mini', '8', '299'],
        ],
      },
    ])
  })

  it('retains source runs in detailed table detection without changing export rows', () => {
    const runs = [...row(700, ['Product', 'Qty', 'Amount']), ...row(680, ['AURA', '2', '20'])].map(
      (entry, index) => ({ ...entry, sourceStart: index * 10, sourceEnd: index * 10 + 4 }),
    )
    const [table] = detectPdfTableDetailsFromRuns(1, runs, {
      includeTwoColumnTextTables: false,
    })
    expect(table?.rows[1]?.[2]).toMatchObject({
      text: '20',
      runs: [expect.objectContaining({ text: '20', sourceStart: 50, sourceEnd: 54 })],
    })
    expect(
      detectPdfTablesFromRuns(1, runs, { includeTwoColumnTextTables: false })[0]?.rows,
    ).toEqual([
      ['Product', 'Qty', 'Amount'],
      ['AURA', '2', '20'],
    ])
  })

  it('keeps separated aligned blocks as separate tables', () => {
    const tables = detectPdfTablesFromRuns(
      1,
      [
        ...row(700, ['Name', 'Qty', 'Price']),
        ...row(680, ['One', '1', '10']),
        ...row(560, ['Region', 'Units', 'Rate']),
        ...row(540, ['East', '2', '20']),
      ],
      { includeTwoColumnTextTables: false },
    )
    expect(tables.map((table) => table.tableNumber)).toEqual([1, 2])
    expect(tables.map((table) => table.rows[0])).toEqual([
      ['Name', 'Qty', 'Price'],
      ['Region', 'Units', 'Rate'],
    ])
  })

  it('avoids prose-like two-column false positives unless explicitly enabled', () => {
    const runs = [
      ...row(700, ['Feature', 'Description'], [20, 220]),
      ...row(680, ['Privacy', 'Runs locally'], [20, 220]),
      ...row(660, ['Editing', 'Keeps elements editable'], [20, 220]),
    ]
    expect(detectPdfTablesFromRuns(1, runs, { includeTwoColumnTextTables: false })).toEqual([])
    expect(detectPdfTablesFromRuns(1, runs, { includeTwoColumnTextTables: true })).toHaveLength(1)
  })

  it('accepts numeric two-column tables by default', () => {
    const tables = detectPdfTablesFromRuns(
      1,
      [
        ...row(700, ['Month', 'Total'], [20, 220]),
        ...row(680, ['July', '18'], [20, 220]),
        ...row(660, ['August', '30'], [20, 220]),
      ],
      { includeTwoColumnTextTables: false },
    )
    expect(tables).toHaveLength(1)
  })

  it('rejects rows without a consistently supported column boundary', () => {
    const tables = detectPdfTablesFromRuns(
      1,
      [
        ...row(700, ['One', 'Alpha'], [20, 150]),
        ...row(680, ['Two', 'Beta'], [20, 270]),
        ...row(660, ['Three', 'Gamma'], [20, 390]),
      ],
      { includeTwoColumnTextTables: true },
    )
    expect(tables).toEqual([])
  })

  it('extracts requested pages and always cleans PDF.js resources', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 2,
      getPage: vi.fn(async (_pageNumber: number) => ({
        getTextContent: async () => ({
          items: [
            item('Name', 20, 700),
            item('Qty', 220, 700),
            item('AURA', 20, 680),
            item('12', 220, 680),
          ],
        }),
        cleanup,
      })),
    } as unknown as PDFDocumentProxy
    const tables = await extractPdfTables(document, [1], {
      includeTwoColumnTextTables: false,
    })
    expect(document.getPage).toHaveBeenCalledWith(2)
    expect(tables).toMatchObject([
      {
        pageNumber: 2,
        rows: [
          ['Name', 'Qty'],
          ['AURA', '12'],
        ],
      },
    ])
    expect(cleanup).toHaveBeenCalledOnce()
    await expect(
      extractPdfTables(document, [2], { includeTwoColumnTextTables: false }),
    ).rejects.toThrow('invalid page')
  })

  it('cleans page resources when text extraction fails', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 1,
      getPage: async () => ({
        getTextContent: async () => {
          throw new Error('text failure')
        },
        cleanup,
      }),
    } as unknown as PDFDocumentProxy
    await expect(
      extractPdfTables(document, [0], { includeTwoColumnTextTables: false }),
    ).rejects.toThrow('text failure')
    expect(cleanup).toHaveBeenCalledOnce()
  })
})
