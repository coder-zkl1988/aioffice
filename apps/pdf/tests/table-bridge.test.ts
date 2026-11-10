import JSZip from 'jszip'
import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { pdfTablesXlsxBytes } from '@genoffice/pdf-tools'
import {
  detectedTableBounds,
  detectPdfTableDetailsFromRuns,
  extractPdfTableDetails,
  type PdfDetectedTable,
  type PdfTableTextRun,
} from '../src/renderer/extract-tables'
import {
  exportTableToXlsx,
  layoutTableImage,
  normalizeTableRows,
  padTableBounds,
  TABLE_BRIDGE_PADDING_PT,
  tableBridgeRecord,
  tableColumnWidths,
  tableMedianRowHeight,
  xlsxRowsFromBytes,
} from '../src/renderer/table-bridge'
import { runEditOps, type OpContext } from '../src/renderer/edit-ops'
import type { EditSnapshot } from '../src/renderer/edit-state'

function run(text: string, x: number, y: number, width = Math.max(10, text.length * 7)) {
  return { text, x, y, width, height: 12 }
}

function item(text: string, x: number, y: number, width = Math.max(10, text.length * 7)) {
  return { str: text, transform: [12, 0, 0, 12, x, y], width, height: 12 }
}

function tableOf(runs: PdfTableTextRun[]): PdfDetectedTable {
  const [table] = detectPdfTableDetailsFromRuns(2, runs, {
    includeTwoColumnTextTables: true,
  })
  if (!table) throw new Error('fixture did not detect a table')
  return table
}

const emptySnapshot = (): EditSnapshot => ({
  markups: [],
  annotDeletes: [],
  noteEdits: [],
  drawings: [],
  textEdits: [],
  textInserts: [],
  imageEdits: [],
  stampCfg: null,
  formEdits: new Map(),
  formFields: [],
  formWidgetEdits: new Map(),
  blankPages: [],
  duplicates: [],
  rotations: new Map(),
  deleted: new Set(),
  order: null,
  metadata: null,
  classification: null,
})

const editContext = (over: Partial<OpContext> = {}): OpContext => ({
  readOnly: false,
  pageCount: 5,
  deleted: new Set(),
  claimedImages: new Set(),
  ...over,
})

describe('detectedTableBounds', () => {
  it('takes min/max over runs with the PDF baseline+height box (y up)', () => {
    const table = tableOf([
      run('Name', 20, 700, 40),
      run('Qty', 180, 700, 20),
      run('Alpha', 20, 680, 35),
      run('12', 180, 680, 14),
    ])
    expect(detectedTableBounds(table)).toEqual({ x0: 20, y0: 680, x1: 200, y1: 712 })
  })

  it('is zero for a table with no runs', () => {
    expect(detectedTableBounds({ pageNumber: 1, tableNumber: 1, rows: [] })).toEqual({
      x0: 0,
      y0: 0,
      x1: 0,
      y1: 0,
    })
  })

  it('pads the box on every side for the write-back image footprint', () => {
    const bounds = { x0: 20, y0: 680, x1: 200, y1: 712 }
    expect(padTableBounds(bounds)).toEqual([
      20 - TABLE_BRIDGE_PADDING_PT,
      680 - TABLE_BRIDGE_PADDING_PT,
      200 + TABLE_BRIDGE_PADDING_PT,
      712 + TABLE_BRIDGE_PADDING_PT,
    ])
    expect(padTableBounds(bounds, 0)).toEqual([20, 680, 200, 712])
  })
})

describe('table geometry', () => {
  it('splits column widths at the midpoint between adjacent ink', () => {
    const table = tableOf([
      run('Name', 20, 700, 40),
      run('Qty', 180, 700, 20),
      run('Alpha', 20, 680, 50),
      run('12', 180, 680, 10),
    ])
    // col 0 ink: 20..70; col 1 ink: 180..200; boundary = (70+180)/2 = 125
    expect(tableColumnWidths(table)).toEqual([105, 75])
  })

  it('falls back to [] when a column has no runs', () => {
    const table = tableOf([
      run('Name', 20, 700, 40),
      run('Qty', 180, 700, 20),
      run('Alpha', 20, 680, 40),
      run('12', 180, 680, 20),
    ])
    table.rows[0]![1]!.runs = []
    table.rows[1]![1]!.runs = []
    expect(tableColumnWidths(table)).toEqual([])
  })

  it('reports the median original row height', () => {
    const table = tableOf([
      run('Name', 20, 700, 40),
      run('Qty', 180, 700, 20),
      run('Alpha', 20, 680, 40),
      run('12', 180, 680, 20),
    ])
    table.rows[0]![0]!.runs[0]!.height = 10
    table.rows[0]![1]!.runs[0]!.height = 10
    table.rows[1]![0]!.runs[0]!.height = 14
    table.rows[1]![1]!.runs[0]!.height = 12
    expect(tableMedianRowHeight(table)).toBe(12)
  })
})

describe('extractPdfTableDetails', () => {
  it('keeps runs per cell and cleans page resources', async () => {
    const cleanup = vi.fn()
    const document = {
      numPages: 1,
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
    const tables = await extractPdfTableDetails(document, [0], {
      includeTwoColumnTextTables: false,
    })
    expect(document.getPage).toHaveBeenCalledWith(1)
    expect(tables).toHaveLength(1)
    expect(tables[0]!.rows[0]![0]!.runs[0]).toMatchObject({ text: 'Name', x: 20, y: 700 })
    expect(cleanup).toHaveBeenCalledOnce()
    await expect(
      extractPdfTableDetails(document, [1], { includeTwoColumnTextTables: false }),
    ).rejects.toThrow('invalid page')
  })
})

describe('xlsxRowsFromBytes', () => {
  it('round-trips the table writer output', async () => {
    const bytes = await pdfTablesXlsxBytes([
      {
        pageNumber: 2,
        tableNumber: 1,
        rows: [
          ['Product', 'Qty'],
          ['AURA One', '12'],
        ],
      },
    ])
    expect(await xlsxRowsFromBytes(bytes)).toEqual([
      ['Product', 'Qty'],
      ['AURA One', '12'],
    ])
  })

  it('reads shared strings, inline strings, numbers and cached formulas', async () => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="Sheet1" sheetId="3" r:id="rId7"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships><Relationship Id="rId7" Type="worksheet" Target="worksheets/sheet9.xml"/></Relationships>',
    )
    zip.file('xl/sharedStrings.xml', '<sst><si><t>Name</t></si><si><t>Beta</t></si><si/></sst>')
    zip.file(
      'xl/worksheets/sheet9.xml',
      '<worksheet><sheetData>' +
        '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>Qty &amp; more</t></is></c></row>' +
        '<row r="2"><c r="A2" t="inlineStr"><is><t>Alpha</t></is></c><c r="B2"><v>42</v></c></row>' +
        '<row r="3"><c r="A3"><f>1+1</f><v>2</v></c><c r="B3" t="s"><v>1</v></c></row>' +
        '</sheetData></worksheet>',
    )
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    expect(await xlsxRowsFromBytes(bytes)).toEqual([
      ['Name', 'Qty & more'],
      ['Alpha', '42'],
      ['2', 'Beta'],
    ])
  })

  it('rejects bytes that are not an xlsx package', async () => {
    await expect(xlsxRowsFromBytes(new Uint8Array([1, 2, 3]))).rejects.toThrow(
      'not a readable .xlsx',
    )
  })
})

describe('normalizeTableRows', () => {
  it('pads short rows, trims trailing empty rows/columns and keeps interior blanks', () => {
    expect(normalizeTableRows([['A', 'B', ''], ['C'], ['', '']])).toEqual([
      ['A', 'B'],
      ['C', ''],
    ])
  })

  it('rejects a blank grid', () => {
    expect(() => normalizeTableRows([[''], ['']])).toThrow('no data rows')
    expect(() => normalizeTableRows([])).toThrow('no data rows')
  })
})

describe('layoutTableImage', () => {
  it('uses the original column proportions and the original row height', () => {
    const layout = layoutTableImage(
      [
        ['Name', 'Qty', 'Price'],
        ['AURA', '2', '99'],
      ],
      { widthPt: 300, heightPt: 120, columnWidths: [100, 50, 150], rowHeightPt: 12 },
    )
    expect(layout.width).toBe(1200)
    expect(layout.height).toBe(480)
    expect(layout.columns.map((column) => column.width)).toEqual([400, 200, 600])
    expect(layout.columns.at(-1)!.x + layout.columns.at(-1)!.width).toBe(1200)
    expect(layout.rowHeight).toBe(240)
    // 12pt row * 0.62 * 4 px/pt
    expect(layout.fontSize).toBeCloseTo(29.76, 5)
    expect(layout.padding).toBeGreaterThan(0)
  })

  it('distributes extra columns evenly and rescales to the original width', () => {
    const layout = layoutTableImage(
      [
        ['A', 'B', 'C'],
        ['1', '2', '3'],
      ],
      { widthPt: 225, heightPt: 100, columnWidths: [100, 50], rowHeightPt: 10 },
    )
    // widths become [100, 50, 75] over 225pt → fractions 4/9, 2/9, 1/3
    expect(layout.columns.map((column) => column.x)).toEqual([0, 400, 600])
    expect(layout.columns.map((column) => column.width)).toEqual([400, 200, 300])
  })

  it('falls back to even columns when the edited grid has fewer columns', () => {
    const layout = layoutTableImage([['Only']], {
      widthPt: 90,
      heightPt: 40,
      columnWidths: [100, 100, 100],
      rowHeightPt: 0,
    })
    expect(layout.columns).toHaveLength(1)
    expect(layout.columns[0]).toEqual({ x: 0, width: layout.width })
    // No original row height: the footprint's single row band drives the font size
    expect(layout.fontSize).toBeCloseTo(layout.rowHeight * 0.62, 5)
  })

  it('caps a huge footprint at the pixel limit without losing the aspect', () => {
    const layout = layoutTableImage([['A', 'B']], {
      widthPt: 4000,
      heightPt: 2000,
      maxPixels: 4096,
    })
    expect(Math.max(layout.width, layout.height)).toBe(4096)
    expect(layout.width / layout.height).toBeCloseTo(2, 1)
  })
})

describe('exportTableToXlsx', () => {
  it('scopes the extractTables operation to the one bridged table', async () => {
    const table = tableOf([
      run('Name', 20, 700, 40),
      run('Qty', 180, 700, 20),
      run('Alpha', 20, 680, 35),
      run('12', 180, 680, 14),
    ])
    const runTool = vi.fn(async () => ({
      ok: true as const,
      savedPath: '/tmp/Doc_tables.xlsx',
      count: 1,
    }))
    const result = await exportTableToXlsx(
      { runTool },
      {
        filePath: '/tmp/Doc.pdf',
        baseName: 'Doc.pdf',
        table,
        includeTwoColumnTextTables: true,
      },
    )
    expect(result).toEqual({ ok: true, savedPath: '/tmp/Doc_tables.xlsx' })
    expect(runTool).toHaveBeenCalledWith({
      path: '/tmp/Doc.pdf',
      baseName: 'Doc.pdf',
      operation: {
        kind: 'extractTables',
        format: 'xlsx',
        pageIndexes: [1],
        includeTwoColumnTextTables: true,
        baseName: 'Doc.pdf',
        tables: [
          {
            pageNumber: 2,
            tableNumber: 1,
            rows: [
              ['Name', 'Qty'],
              ['Alpha', '12'],
            ],
          },
        ],
      },
    })
  })

  it('reports a canceled save dialog without an error', async () => {
    const table = tableOf([
      run('Name', 20, 700, 40),
      run('Qty', 180, 700, 20),
      run('Alpha', 20, 680, 35),
      run('12', 180, 680, 14),
    ])
    const result = await exportTableToXlsx(
      { runTool: async () => ({ ok: true, canceled: true }) },
      {
        filePath: '/tmp/Doc.pdf',
        baseName: 'Doc.pdf',
        table,
        includeTwoColumnTextTables: false,
      },
    )
    expect(result).toEqual({ ok: false, canceled: true })
  })
})

describe('write-back staging', () => {
  const table = tableOf([
    run('Name', 20, 700, 40),
    run('Qty', 180, 700, 20),
    run('Alpha', 20, 680, 35),
    run('12', 180, 680, 14),
  ])

  const insertOp = (image: string) => {
    const bridge = tableBridgeRecord('/tmp/Doc.pdf', '/tmp/Doc_tables.xlsx', table)
    return {
      op: 'addImageEdit' as const,
      input: {
        kind: 'insertImage' as const,
        pageIndex: bridge.pageNumber - 1,
        image,
        rect: [bridge.bounds.x0, bridge.bounds.y0, bridge.bounds.x1, bridge.bounds.y1] as [
          number,
          number,
          number,
          number,
        ],
        layer: 'aboveText' as const,
        rotate: 0,
      },
    }
  }

  let sequence = 0
  const newId = () => `bridge${++sequence}`

  it('lands in the imageEdits bucket with an undo step and replaces its own previous image', () => {
    const first = runEditOps(emptySnapshot(), [insertOp('AAA')], editContext(), newId)
    expect(first.plan.failures).toEqual([])
    expect([...first.plan.touched]).toEqual(['imageEdits'])
    expect(first.state.imageEdits).toHaveLength(1)
    expect(first.state.imageEdits[0]!.input).toMatchObject({
      kind: 'insertImage',
      pageIndex: 1,
      image: 'AAA',
      layer: 'aboveText',
      rect: [20, 680, 200, 712],
    })
    const firstId = first.state.imageEdits[0]!.id

    // A second update removes the previous bridge image and adds the new one in one batch
    const second = runEditOps(
      first.state,
      [
        { op: 'removeImageEdit', id: firstId },
        { ...insertOp('BBB'), id: 'bridge-next' },
      ],
      editContext(),
      newId,
    )
    expect(second.plan.failures).toEqual([])
    expect(second.state.imageEdits).toHaveLength(1)
    expect(second.state.imageEdits[0]!.id).toBe('bridge-next')
    expect(second.state.imageEdits[0]!.input).toMatchObject({ image: 'BBB' })

    // Undo restores the whole snapshot the second batch started from
    expect(first.state.imageEdits[0]!.input).toMatchObject({ image: 'AAA' })
  })

  it('refuses a rect outside the addressable page range', () => {
    const result = runEditOps(
      emptySnapshot(),
      [insertOp('AAA')],
      editContext({ pageCount: 1 }),
      newId,
    )
    expect(result.plan.failures[0]!.error).toContain('out of range')
    expect(result.state.imageEdits).toEqual([])
  })
})
