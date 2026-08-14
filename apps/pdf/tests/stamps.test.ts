import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  DEFAULT_HEADER_FOOTER,
  DEFAULT_WATERMARK,
  buildStamps,
  compactStampImages,
  createWatermarkUuid,
  formatWatermarkDatePattern,
  parseStampPageRange,
  resolveHeaderFooterText,
  resolveWatermarkText,
  type HeaderFooterConfig,
  type WatermarkConfig,
} from '../src/renderer/stamps'

interface FakeCtx {
  texts: string[]
  font: string
  fillStyle: string
  textAlign: string
  textBaseline: string
  translate: () => void
  rotate: () => void
  scale: () => void
  measureText: (text: string) => { width: number }
  fillText: (text: string, x: number, y: number) => void
}

let contexts: FakeCtx[]

function makeCtx(): FakeCtx {
  const ctx: FakeCtx = {
    texts: [],
    font: '',
    fillStyle: '',
    textAlign: '',
    textBaseline: '',
    translate: vi.fn(),
    rotate: vi.fn(),
    scale: vi.fn(),
    measureText: (text) => ({ width: text.length * 10 }),
    fillText: (text) => {
      ctx.texts.push(text)
    },
  }
  return ctx
}

beforeEach(() => {
  contexts = []
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(() => {
    const ctx = makeCtx()
    contexts.push(ctx)
    return ctx as unknown as CanvasRenderingContext2D
  })
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue(
    'data:image/png;base64,FAKEBASE64',
  )
})

afterEach(() => {
  vi.restoreAllMocks()
})

const wm = (over: Partial<WatermarkConfig> = {}): WatermarkConfig => ({
  ...DEFAULT_WATERMARK,
  text: 'CONFIDENTIAL',
  ...over,
})

const page = (origIdx: number, displayNo: number, pw = 600, ph = 800) => ({
  origIdx,
  pw,
  ph,
  displayNo,
})

describe('buildStamps watermark', () => {
  it('emits one centered stamp per page with the configured opacity and rotation', () => {
    const stamps = buildStamps([page(0, 1), page(3, 2)], wm({ opacity: 0.3 }), null)
    expect(stamps).toHaveLength(2)
    expect(stamps[0]).toMatchObject({
      pageIndex: 0,
      image: 'FAKEBASE64',
      opacity: 0.3,
      rotation: 35,
    })
    const [x1, y1, x2, y2] = stamps[0]!.rect
    expect((x1 + x2) / 2).toBe(300)
    expect((y1 + y2) / 2).toBe(400)
    expect(x2 - x1).toBeGreaterThan(0)
    expect(y2 - y1).toBeGreaterThan(0)
    expect(stamps[1]!.pageIndex).toBe(3)
  })

  it('reuses one rendered bitmap for same-size pages but not across sizes', () => {
    buildStamps([page(0, 1), page(1, 2)], wm(), null)
    expect(contexts).toHaveLength(2)

    contexts = []
    buildStamps([page(0, 1, 600, 800), page(1, 2, 595, 842)], wm(), null)
    expect(contexts).toHaveLength(4)
  })

  it('places a single stamp in the selected nine-grid position', () => {
    const stamps = buildStamps(
      [page(0, 1, 600, 800)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 2,
        sizeRatio: 0.1,
        position: 'topRight',
        angle: 0,
      }),
      null,
    )

    expect(stamps[0]!.rect).toEqual([455.5, 715.5, 575.5, 775.5])
  })

  it('uses the selected margin preset for nine-grid placement', () => {
    const small = buildStamps(
      [page(0, 1, 600, 800)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 2,
        sizeRatio: 0.1,
        position: 'bottomLeft',
        margin: 'small',
        angle: 0,
      }),
      null,
    )
    const extraLarge = buildStamps(
      [page(0, 1, 600, 800)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 2,
        sizeRatio: 0.1,
        position: 'bottomLeft',
        margin: 'x-large',
        angle: 0,
      }),
      null,
    )

    expect(small[0]!.rect).toEqual([14, 14, 134, 74])
    expect(extraLarge[0]!.rect).toEqual([52.5, 52.5, 172.5, 112.5])
  })

  it('places a single watermark at a normalized custom center', () => {
    const stamps = buildStamps(
      [page(0, 1, 600, 800), page(1, 2, 300, 400)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 2,
        sizeRatio: 0.1,
        customPosition: { xRatio: 0.25, yRatio: 0.75 },
        angle: 0,
      }),
      null,
    )

    expect(stamps[0]!.rect).toEqual([90, 570, 210, 630])
    expect(stamps[1]!.rect).toEqual([45, 285, 105, 315])
  })

  it('clamps a rotated custom watermark inside the page bounds', () => {
    const stamps = buildStamps(
      [page(0, 1, 200, 200)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 4,
        sizeRatio: 0.2,
        customPosition: { xRatio: -1, yRatio: 2 },
        angle: 45,
      }),
      null,
    )
    const [x1, y1, x2, y2] = stamps[0]!.rect
    expect(x1).toBeGreaterThanOrEqual(0)
    expect(y1).toBeGreaterThanOrEqual(0)
    expect(x2).toBeLessThanOrEqual(200)
    expect(y2).toBeLessThanOrEqual(200)
  })

  it('falls back to the page center for non-finite custom coordinates', () => {
    const stamps = buildStamps(
      [page(0, 1, 600, 800)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 2,
        sizeRatio: 0.1,
        customPosition: { xRatio: Number.NaN, yRatio: Number.POSITIVE_INFINITY },
        angle: 0,
      }),
      null,
    )

    expect(stamps[0]!.rect).toEqual([240, 370, 360, 430])
  })

  it('keeps rotated edge stamps inside the page', () => {
    const stamps = buildStamps(
      [page(0, 1, 200, 200)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 4,
        sizeRatio: 0.2,
        position: 'bottomLeft',
        angle: 45,
      }),
      null,
    )
    const [x1, y1, x2, y2] = stamps[0]!.rect
    expect(x1).toBeGreaterThanOrEqual(0)
    expect(y1).toBeGreaterThanOrEqual(0)
    expect(x2).toBeLessThanOrEqual(200)
    expect(y2).toBeLessThanOrEqual(200)
  })

  it('renders up to five trimmed text lines into one reusable tile', () => {
    const stamps = buildStamps(
      [page(0, 1)],
      wm({ text: ' First\nSecond\nThird\nFourth\nFifth\nSixth ' }),
      null,
    )
    expect(stamps).toHaveLength(1)
    expect(contexts.at(-1)?.texts).toEqual(['First', 'Second', 'Third', 'Fourth', 'Fifth'])
  })

  it('uses the prepared PNG directly for an image watermark', () => {
    const stamps = buildStamps(
      [page(0, 1)],
      wm({ type: 'image', image: 'IMAGEBASE64', imageAspectRatio: 2, sizeRatio: 0.1 }),
      null,
    )
    expect(stamps).toEqual([
      {
        pageIndex: 0,
        image: 'IMAGEBASE64',
        rect: [240, 370, 360, 430],
        opacity: DEFAULT_WATERMARK.opacity,
        rotation: DEFAULT_WATERMARK.angle,
      },
    ])
    expect(contexts).toHaveLength(0)
  })

  it('builds staggered tiled watermarks and caps pathological tile counts', () => {
    const tiled = buildStamps(
      [page(0, 1, 200, 200)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 1,
        layout: 'tiled',
        angle: 0,
        sizeRatio: 0.1,
        horizontalSpacing: 20,
        verticalSpacing: 20,
      }),
      null,
    )
    expect(tiled).toHaveLength(36)
    expect(tiled[0]!.rect).toEqual([-20, -20, 0, 0])
    expect(tiled[6]!.rect).toEqual([0, 20, 20, 40])

    const capped = buildStamps(
      [page(0, 1, 10_000, 10_000)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 1,
        layout: 'tiled',
        angle: 0,
        sizeRatio: 0.0001,
        horizontalSpacing: 0,
        verticalSpacing: 0,
      }),
      null,
    )
    expect(capped).toHaveLength(10_000)

    const documentCapped = buildStamps(
      Array.from({ length: 26 }, (_, index) => page(index, index + 1, 10_000, 10_000)),
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 1,
        layout: 'tiled',
        angle: 0,
        sizeRatio: 0.0001,
        horizontalSpacing: 0,
        verticalSpacing: 0,
      }),
      null,
    )
    expect(documentCapped).toHaveLength(10_000)
  })

  it('emits nothing for a blank watermark and no header/footer', () => {
    expect(buildStamps([page(0, 1)], wm({ text: '  ' }), null)).toEqual([])
    expect(buildStamps([page(0, 1)], wm({ type: 'image', image: '' }), null)).toEqual([])
    expect(buildStamps([page(0, 1)], null, null)).toEqual([])
  })
})

describe('resolveWatermarkText', () => {
  it('resolves date, page, and filename variables in a stable order', () => {
    expect(
      resolveWatermarkText('@date @time | @page_number/@page_count | @filename | @filename_full', {
        pageNumber: 3,
        totalPages: 12,
        fileName: 'Quarterly.Report.pdf',
        dateTime: new Date(2026, 7, 14, 9, 5, 7),
      }),
    ).toBe('2026-08-14 09:05:07 | 3/12 | Quarterly.Report | Quarterly.Report.pdf')
  })

  it('supports date parts and escaped at signs', () => {
    expect(
      resolveWatermarkText('@@review @year-@month-@day @datetime @total_pages', {
        pageNumber: 1,
        totalPages: 2,
        dateTime: new Date(2026, 0, 2, 3, 4, 5),
      }),
    ).toBe('@review 2026-01-02 2026-01-02 03:04:05 2')
  })

  it('resolves document metadata and a stable operation identifier', () => {
    expect(
      resolveWatermarkText('@title | @author | @subject | @keywords | @uuid', {
        pageNumber: 1,
        totalPages: 1,
        title: 'Launch Plan',
        author: 'GenOffice',
        subject: 'Review',
        keywords: 'alpha, beta',
        uuid: 'a1b2c3d4',
      }),
    ).toBe('Launch Plan | GenOffice | Review | alpha, beta | a1b2c3d4')
  })

  it('creates an eight-character hexadecimal identifier', () => {
    expect(createWatermarkUuid()).toMatch(/^[0-9a-f]{8}$/)
  })

  it('renders variable text independently for each page', () => {
    buildStamps(
      [page(0, 1), page(1, 2)],
      wm({ text: 'Page @page of @total_pages', appliedAt: '2026-08-14T00:00:00.000Z' }),
      null,
    )

    expect(contexts[1]!.texts).toEqual(['Page 1 of 2'])
    expect(contexts[3]!.texts).toEqual(['Page 2 of 2'])
  })

  it('supports safe custom date formats and rejects invalid patterns', () => {
    const dateTime = new Date(2026, 7, 14, 21, 5, 7, 42)
    expect(formatWatermarkDatePattern("yyyy/MM/dd 'at' hh:mm:ss a", dateTime)).toBe(
      '2026/08/14 at 09:05:07 PM',
    )
    expect(
      resolveWatermarkText('@date{dd-MM-yy_HH:mm:ss.SSS}', {
        pageNumber: 1,
        totalPages: 1,
        dateTime,
      }),
    ).toBe('14-08-26_21:05:07.042')
    expect(formatWatermarkDatePattern('yyyy<script>', dateTime)).toBeNull()
    expect(formatWatermarkDatePattern('MMMM dd', dateTime)).toBeNull()
    expect(
      resolveWatermarkText('@date{QQQQ}', {
        pageNumber: 1,
        totalPages: 1,
        dateTime,
      }),
    ).toBe('[invalid date format]')
    expect(
      resolveWatermarkText(`@date{${'y'.repeat(51)}}`, {
        pageNumber: 1,
        totalPages: 1,
        dateTime,
      }),
    ).toBe('[invalid date format]')
  })

  it('parses deduplicated page ranges and all-page selection', () => {
    expect(parseStampPageRange('3, 1-2,2', 4)).toEqual([1, 2, 3])
    expect(parseStampPageRange('1 - 2, 4', 4)).toEqual([1, 2, 4])
    expect(parseStampPageRange('all', 3)).toEqual([1, 2, 3])
    expect(parseStampPageRange('0,2', 3)).toBeNull()
    expect(parseStampPageRange('2-4', 3)).toBeNull()
    expect(parseStampPageRange('2-', 3)).toBeNull()
  })

  it('applies watermarks only to selected original pages after a reorder', () => {
    buildStamps([page(2, 1), page(0, 2), page(1, 3)], wm({ text: '@page/@total_pages' }), null, {
      watermarkPageIndexes: [0],
    })

    expect(contexts.at(-1)?.texts).toEqual(['2/3'])
  })

  it('uses the same metadata and identifier on every selected page', () => {
    buildStamps(
      [page(0, 1), page(1, 2)],
      wm({ text: '@title @uuid @page', uuid: 'deadbeef' }),
      null,
      { metadata: { title: 'Roadmap' } },
    )

    expect(contexts[1]!.texts).toEqual(['Roadmap deadbeef 1'])
    expect(contexts[3]!.texts).toEqual(['Roadmap deadbeef 2'])
  })

  it('keeps header and footer output on all pages when a watermark range is selected', () => {
    const stamps = buildStamps(
      [page(0, 1), page(1, 2)],
      wm(),
      { ...DEFAULT_HEADER_FOOTER, pageNumber: true },
      { watermarkPageIndexes: [1] },
    )

    expect(stamps.map((stamp) => stamp.pageIndex)).toEqual([0, 1, 1])
  })
})

describe('compactStampImages', () => {
  it('deduplicates bitmap payloads while preserving stamp placement', () => {
    const source = buildStamps(
      [page(0, 1, 200, 200)],
      wm({
        type: 'image',
        image: 'IMAGEBASE64',
        imageAspectRatio: 1,
        layout: 'tiled',
        angle: 0,
      }),
      null,
    )
    const compacted = compactStampImages(source)
    expect(compacted.stampImages).toEqual(['IMAGEBASE64'])
    expect(compacted.stamps).toHaveLength(source.length)
    expect(compacted.stamps[0]).toEqual({ ...source[0], image: '', imageIndex: 0 })
    expect(compacted.stamps.every((stamp) => stamp.image === '' && stamp.imageIndex === 0)).toBe(
      true,
    )
  })

  it('keeps distinct header and watermark bitmaps in separate slots', () => {
    const compacted = compactStampImages([
      { pageIndex: 0, image: 'A', rect: [0, 0, 10, 10] },
      { pageIndex: 0, image: 'B', rect: [0, 10, 10, 20] },
      { pageIndex: 1, image: 'A', rect: [0, 0, 10, 10] },
    ])
    expect(compacted.stampImages).toEqual(['A', 'B'])
    expect(compacted.stamps.map((stamp) => stamp.imageIndex)).toEqual([0, 1, 0])
  })
})

describe('buildStamps header/footer', () => {
  const hf = (over: Partial<HeaderFooterConfig> = {}): HeaderFooterConfig => ({
    ...DEFAULT_HEADER_FOOTER,
    pageNumber: true,
    ...over,
  })

  it('renders an auto page-number footer positioned at the bottom margin', () => {
    const stamps = buildStamps([page(5, 1), page(6, 2)], null, hf())
    expect(stamps).toHaveLength(2)
    const barH = DEFAULT_HEADER_FOOTER.fontSize * 2.2
    const margin = 800 * 0.035
    expect(stamps[0]).toEqual({
      pageIndex: 5,
      image: 'FAKEBASE64',
      rect: [0, margin, 600, margin + barH],
    })
    expect(contexts[0]!.texts).toEqual(['1 / 2'])
    expect(contexts[1]!.texts).toEqual(['2 / 2'])
  })

  it('honors startAt for page numbering', () => {
    buildStamps([page(0, 1)], null, hf({ startAt: 10 }))
    expect(contexts[0]!.texts).toEqual(['10 / 1'])
  })

  it('formats Bates numbers with filename, total pages, and zero padding', () => {
    buildStamps(
      [page(0, 1), page(1, 2)],
      null,
      hf({
        startAt: 7,
        pageNumberFormat: 'DOC-{filename}-{n}/{total}',
        pageNumberZeroPad: 4,
      }),
      { fileName: '/incoming/Launch.Deck.pdf' },
    )

    expect(contexts[0]!.texts).toEqual(['DOC-Launch.Deck-0007/2'])
    expect(contexts[1]!.texts).toEqual(['DOC-Launch.Deck-0008/2'])
  })

  it('places an automatic page number at any nine-grid position', () => {
    const stamps = buildStamps(
      [page(0, 1, 600, 800)],
      null,
      hf({
        pageNumberFormat: 'Page {page}',
        pageNumberPosition: 'middleRight',
        pageNumberMargin: 'large',
      }),
    )

    expect(stamps).toHaveLength(1)
    expect(stamps[0]!.rect[1]).toBeCloseTo(390.1)
    expect(stamps[0]!.rect[3]).toBeCloseTo(409.9)
    expect(contexts[0]!.texts).toEqual(['Page 1'])
    expect(contexts[0]!.textAlign).toBe('right')
  })

  it('replaces an occupied header slot with the automatic page number', () => {
    buildStamps(
      [page(0, 1)],
      null,
      hf({
        headerLeft: 'Document',
        headerRight: 'Old value',
        pageNumberPosition: 'topRight',
        pageNumberFormat: '{page}',
      }),
    )

    expect(contexts[0]!.texts).toEqual(['Document'])
    expect(contexts[1]!.texts).toEqual(['1'])
  })

  it('numbers only selected original pages in their current display order', () => {
    const stamps = buildStamps([page(2, 1), page(0, 2), page(1, 3)], null, hf({ startAt: 10 }), {
      headerFooterPageIndexes: [0, 1],
    })

    expect(stamps.map((stamp) => stamp.pageIndex)).toEqual([0, 1])
    expect(contexts[0]!.texts).toEqual(['10 / 3'])
    expect(contexts[1]!.texts).toEqual(['11 / 3'])
  })

  it('uses the configured header and footer font size and color', () => {
    buildStamps([page(0, 1)], null, hf({ headerCenter: 'Title', fontSize: 18, color: '#217346' }))

    expect(contexts).toHaveLength(2)
    expect(contexts[0]!.font).toContain('18px')
    expect(contexts[0]!.fillStyle).toBe('#217346')
    expect(contexts[1]!.font).toContain('18px')
    expect(contexts[1]!.fillStyle).toBe('#217346')
  })

  it('uses the selected standard font family', () => {
    buildStamps([page(0, 1)], null, hf({ fontFamily: 'mono' }))
    expect(contexts[0]!.font).toContain('ui-monospace')
  })

  it('places headers at the top and fills {page}/{total} placeholders', () => {
    const stamps = buildStamps(
      [page(0, 1)],
      null,
      hf({ headerLeft: 'p{page} of {total}', headerRight: 'ACME', pageNumber: false }),
    )
    expect(stamps).toHaveLength(1)
    expect(contexts[0]!.texts).toEqual(['p1 of 1', 'ACME'])
    const barH = DEFAULT_HEADER_FOOTER.fontSize * 2.2
    const margin = Math.min(800 * 0.035, 26)
    expect(stamps[0]!.rect).toEqual([0, 800 - margin - barH, 600, 800 - margin])
  })

  it('uses footerCenter text when pageNumber is disabled', () => {
    buildStamps([page(0, 1)], null, hf({ pageNumber: false, footerCenter: 'Page {page}' }))
    expect(contexts[0]!.texts).toEqual(['Page 1'])
  })

  it('emits no bars when all segments are blank and pageNumber is off', () => {
    expect(buildStamps([page(0, 1)], null, hf({ pageNumber: false }))).toEqual([])
  })

  it('combines watermark, header, and footer stamps for one page', () => {
    const stamps = buildStamps([page(0, 1)], wm(), hf({ headerCenter: 'Title' }))
    expect(stamps).toHaveLength(3)
    expect(stamps.map((s) => s.pageIndex)).toEqual([0, 0, 0])
  })
})

describe('resolveHeaderFooterText', () => {
  it('supports both page aliases and clamps Bates padding to twelve digits', () => {
    expect(resolveHeaderFooterText('{page}|{n}|{total}|{filename}', 12, 30, 'report.pdf', 99)).toBe(
      '000000000012|000000000012|30|report',
    )
  })

  it('keeps dotfiles and extensionless file names intact', () => {
    expect(resolveHeaderFooterText('{filename}', 1, 1, '.archive', 0)).toBe('.archive')
    expect(resolveHeaderFooterText('{filename}', 1, 1, 'README', 0)).toBe('README')
  })
})
