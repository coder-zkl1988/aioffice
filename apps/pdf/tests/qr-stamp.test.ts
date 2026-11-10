import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  DEFAULT_WATERMARK,
  MAX_QR_CONTENT_BYTES,
  QR_QUIET_ZONE_MODULES,
  buildStamps,
  renderQrPreviewBase64,
  validateQrContent,
  type WatermarkConfig,
} from '../src/renderer/stamps'

interface RecordedContext {
  fillRects: [number, number, number, number][]
  fillStyle: string
  fillRect: (x: number, y: number, width: number, height: number) => void
}

let contexts: RecordedContext[]
let canvases: HTMLCanvasElement[]
let drawCalls: number

beforeEach(() => {
  contexts = []
  canvases = []
  drawCalls = 0
  vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockImplementation(function (
    this: HTMLCanvasElement,
  ) {
    const context: RecordedContext = {
      fillRects: [],
      fillStyle: '',
      fillRect(x: number, y: number, width: number, height: number) {
        this.fillRects.push([x, y, width, height])
      },
    }
    contexts.push(context)
    canvases.push(this)
    return context as unknown as CanvasRenderingContext2D
  } as never)
  vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockImplementation(() => {
    const n = drawCalls
    drawCalls += 1
    // A real PNG data URL always starts with the PNG signature, so the payload is a base64 PNG.
    return `data:image/png;base64,iVBORw0KGgo${n}`
  })
})

afterEach(() => {
  vi.restoreAllMocks()
})

const qrWm = (over: Partial<WatermarkConfig> = {}): WatermarkConfig => ({
  ...DEFAULT_WATERMARK,
  type: 'qrcode',
  text: 'https://example.com/doc',
  ...over,
})

const page = (origIdx: number, displayNo: number, pw = 600, ph = 800) => ({
  origIdx,
  pw,
  ph,
  displayNo,
})

describe('validateQrContent', () => {
  it('accepts non-blank content up to the UTF-8 byte limit', () => {
    expect(validateQrContent('hello')).toBeNull()
    expect(validateQrContent('a'.repeat(MAX_QR_CONTENT_BYTES))).toBeNull()
    // 400 three-byte characters are 1200 bytes even though the string is shorter than 1000.
    expect(validateQrContent('汉'.repeat(400))).toBe('tooLong')
  })

  it('rejects blank content', () => {
    expect(validateQrContent('')).toBe('blank')
    expect(validateQrContent(' \n\t ')).toBe('blank')
  })
})

describe('buildStamps qrcode watermark', () => {
  it('emits one square QR stamp per page', () => {
    const stamps = buildStamps([page(0, 1), page(1, 2)], qrWm({ opacity: 0.3 }), null)

    expect(stamps).toHaveLength(2)
    expect(stamps[0]).toMatchObject({ pageIndex: 0, opacity: 0.3, rotation: 35 })
    expect(stamps[1]!.pageIndex).toBe(1)
    expect(stamps[0]!.image.startsWith('iVBORw0KGgo')).toBe(true)
    const [x1, y1, x2, y2] = stamps[0]!.rect
    expect((x1 + x2) / 2).toBe(300)
    expect((y1 + y2) / 2).toBe(400)
    expect(x2 - x1).toBeCloseTo(y2 - y1, 6)
  })

  it('draws the matrix at module size with a four-module quiet zone', () => {
    buildStamps([page(0, 1)], qrWm(), null)

    expect(contexts).toHaveLength(1)
    const [context] = contexts as [RecordedContext]
    const canvas = canvases[0]!
    const [firstX, firstY, moduleWidth] = context.fillRects[0]!
    // The top-left finder pattern is dark, so the first drawn cell starts past the quiet zone.
    expect(firstX).toBe(QR_QUIET_ZONE_MODULES * moduleWidth)
    expect(firstY).toBe(QR_QUIET_ZONE_MODULES * moduleWidth)
    expect(moduleWidth).toBeGreaterThan(0)
    expect(canvas.width).toBe(canvas.height)
    expect(canvas.width % moduleWidth).toBe(0)
    // Symmetric quiet zone: modules span width minus both borders.
    const modules = canvas.width / moduleWidth - QR_QUIET_ZONE_MODULES * 2
    expect(modules).toBeGreaterThanOrEqual(21)
    expect(context.fillStyle).toBe(DEFAULT_WATERMARK.color)
  })

  it('reuses one bitmap for same-size pages with the same resolved text', () => {
    const stamps = buildStamps([page(0, 1), page(1, 2)], qrWm(), null)

    expect(contexts).toHaveLength(1)
    expect(stamps[0]!.image).toBe(stamps[1]!.image)
  })

  it('resolves tokens per page so each page encodes its own text', () => {
    const stamps = buildStamps([page(0, 1), page(1, 2)], qrWm({ text: 'Page @page' }), null)

    expect(contexts).toHaveLength(2)
    expect(stamps).toHaveLength(2)
    expect(stamps[0]!.image).not.toBe(stamps[1]!.image)
  })

  it('skips pages when the resolved QR content is blank or too long', () => {
    expect(buildStamps([page(0, 1)], qrWm({ text: '   ' }), null)).toEqual([])
    expect(buildStamps([page(0, 1)], qrWm({ text: 'a'.repeat(1001) }), null)).toEqual([])
    expect(contexts).toHaveLength(0)
  })

  it('encodes content at exactly the byte limit', () => {
    const stamps = buildStamps([page(0, 1)], qrWm({ text: 'a'.repeat(1000) }), null)
    expect(stamps).toHaveLength(1)
  })
})

describe('renderQrPreviewBase64', () => {
  it('returns a base64 PNG for encodable text and null otherwise', () => {
    expect(renderQrPreviewBase64('https://example.com', '#d0342c')).toBe('iVBORw0KGgo0')
    expect(renderQrPreviewBase64('', '#d0342c')).toBeNull()
    expect(renderQrPreviewBase64('a'.repeat(1001), '#d0342c')).toBeNull()
  })
})
