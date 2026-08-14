import type { StampInput } from '../shared/ipc'

/** Bitmap supersampling factor relative to PDF pt — stays sharp even when enlarged for print */
const SS = 4
const MAX_DOCUMENT_WATERMARK_TILES = 10_000

export type WatermarkType = 'text' | 'image'
export type WatermarkLayout = 'single' | 'tiled'
export type WatermarkMargin = 'small' | 'medium' | 'large' | 'x-large'
export const WATERMARK_MARGIN_FACTORS: Record<WatermarkMargin, number> = {
  small: 0.02,
  medium: 0.035,
  large: 0.05,
  'x-large': 0.075,
}
export interface WatermarkCustomPosition {
  /** Normalized center position from the left edge of the page. */
  xRatio: number
  /** Normalized center position from the bottom edge of the page. */
  yRatio: number
}
export type WatermarkPosition =
  | 'topLeft'
  | 'topCenter'
  | 'topRight'
  | 'middleLeft'
  | 'center'
  | 'middleRight'
  | 'bottomLeft'
  | 'bottomCenter'
  | 'bottomRight'

export type HeaderFooterFont = 'sans' | 'serif' | 'mono'

export interface WatermarkConfig {
  type: WatermarkType
  layout: WatermarkLayout
  position: WatermarkPosition
  margin: WatermarkMargin
  customPosition: WatermarkCustomPosition | null
  text: string
  /** Normalized PNG, base64 without the data prefix */
  image: string
  imageAspectRatio: number
  /** Counterclockwise angle */
  angle: number
  opacity: number
  color: string
  /** Text size or image height as a ratio of page width */
  sizeRatio: number
  /** Tiled watermark spacing in PDF points */
  horizontalSpacing: number
  verticalSpacing: number
  /** Captured when the stamp is applied so preview and saved output use the same date/time. */
  appliedAt: string
  /** Stable short identifier shared by every page in one watermark operation. */
  uuid: string
}

export interface HeaderFooterConfig {
  headerLeft: string
  headerCenter: string
  headerRight: string
  footerLeft: string
  footerCenter: string
  footerRight: string
  /** Auto page number in the footer center (overrides footerCenter) */
  pageNumber: boolean
  startAt: number
  pageNumberFormat: string
  pageNumberZeroPad: number
  pageNumberPosition: WatermarkPosition
  pageNumberMargin: WatermarkMargin
  fontFamily: HeaderFooterFont
  fontSize: number
  color: string
}

export const DEFAULT_WATERMARK: WatermarkConfig = {
  type: 'text',
  layout: 'single',
  position: 'center',
  margin: 'medium',
  customPosition: null,
  text: '',
  image: '',
  imageAspectRatio: 1,
  angle: 35,
  opacity: 0.18,
  color: '#d0342c',
  sizeRatio: 0.11,
  horizontalSpacing: 54,
  verticalSpacing: 54,
  appliedAt: '',
  uuid: '',
}

export const DEFAULT_HEADER_FOOTER: HeaderFooterConfig = {
  headerLeft: '',
  headerCenter: '',
  headerRight: '',
  footerLeft: '',
  footerCenter: '',
  footerRight: '',
  pageNumber: false,
  startAt: 1,
  pageNumberFormat: '{page} / {total}',
  pageNumberZeroPad: 0,
  pageNumberPosition: 'bottomCenter',
  pageNumberMargin: 'medium',
  fontFamily: 'sans',
  fontSize: 9,
  color: '#666666',
}

const FONT_FAMILIES: Record<HeaderFooterFont, string> = {
  sans: '-apple-system, "PingFang SC", "Microsoft YaHei", "Segoe UI", sans-serif',
  serif: 'Georgia, "Times New Roman", "Noto Serif CJK SC", serif',
  mono: 'ui-monospace, "SFMono-Regular", Consolas, "Liberation Mono", monospace',
}

const FONT = (px: number, bold = false, family: HeaderFooterFont = 'sans') =>
  `${bold ? '600 ' : ''}${px}px ${FONT_FAMILIES[family]}`

function toBase64(canvas: HTMLCanvasElement): string {
  return canvas.toDataURL('image/png').split(',')[1] ?? ''
}

interface WatermarkTile {
  image: string
  width: number
  height: number
}

function renderTextWatermarkTile(
  cfg: WatermarkConfig,
  text: string,
  pw: number,
  ph: number,
): WatermarkTile | null {
  const lines = text
    .trim()
    .split(/\r?\n/)
    .map((line) => line.trim())
    .slice(0, 5)
  if (lines.length === 0 || lines.every((line) => !line.trim())) return null

  const measureCanvas = document.createElement('canvas')
  const measureContext = measureCanvas.getContext('2d')
  if (!measureContext) return null
  let fontSize = Math.max(8, pw * cfg.sizeRatio)
  measureContext.font = FONT(fontSize * SS, true)
  const measuredWidth = Math.max(
    ...lines.map((line) => measureContext.measureText(line).width / SS),
  )
  const maxWidth = Math.max(40, Math.min(pw * (cfg.layout === 'tiled' ? 0.62 : 0.82), ph * 0.82))
  if (measuredWidth > maxWidth) fontSize *= maxWidth / measuredWidth

  const padding = Math.max(2, fontSize * 0.18)
  const lineHeight = fontSize * 1.22
  measureContext.font = FONT(fontSize * SS, true)
  const width = Math.max(
    fontSize,
    Math.max(...lines.map((line) => measureContext.measureText(line).width / SS)) + padding * 2,
  )
  const height = lineHeight * lines.length + padding * 2
  const canvas = document.createElement('canvas')
  canvas.width = Math.max(1, Math.ceil(width * SS))
  canvas.height = Math.max(1, Math.ceil(height * SS))
  const context = canvas.getContext('2d')
  if (!context) return null
  context.scale(SS, SS)
  context.font = FONT(fontSize, true)
  context.fillStyle = cfg.color
  context.textAlign = 'center'
  context.textBaseline = 'middle'
  lines.forEach((line, index) => {
    context.fillText(line, width / 2, padding + lineHeight * (index + 0.5))
  })
  return { image: toBase64(canvas), width, height }
}

function watermarkTile(
  cfg: WatermarkConfig,
  text: string,
  pw: number,
  ph: number,
): WatermarkTile | null {
  if (cfg.type === 'text') return renderTextWatermarkTile(cfg, text, pw, ph)
  if (!cfg.image || !Number.isFinite(cfg.imageAspectRatio) || cfg.imageAspectRatio <= 0) return null
  let height = Math.max(8, pw * cfg.sizeRatio)
  let width = height * cfg.imageAspectRatio
  const maxWidth = pw * 0.82
  const maxHeight = ph * 0.82
  const scale = Math.min(1, maxWidth / width, maxHeight / height)
  width *= scale
  height *= scale
  return { image: cfg.image, width, height }
}

export interface WatermarkTokenContext {
  pageNumber: number
  totalPages: number
  fileName?: string
  dateTime?: Date
  title?: string
  author?: string
  subject?: string
  keywords?: string
  uuid?: string
}

function twoDigits(value: number): string {
  return String(value).padStart(2, '0')
}

const DATE_PATTERN_TOKENS = [
  'yyyy',
  'SSS',
  'yy',
  'SS',
  'MM',
  'dd',
  'HH',
  'hh',
  'mm',
  'ss',
  'S',
  'M',
  'd',
  'H',
  'h',
  'm',
  's',
  'a',
] as const
const DATE_PATTERN_LETTERS = /[A-Za-z]/

export function formatWatermarkDatePattern(pattern: string, dateTime: Date): string | null {
  if (!pattern || pattern.length > 50 || !/^[A-Za-z0-9/\-:. ,'_]+$/.test(pattern)) return null
  if (Number.isNaN(dateTime.getTime())) return null

  const hour12 = dateTime.getHours() % 12 || 12
  const values: Record<(typeof DATE_PATTERN_TOKENS)[number], string> = {
    yyyy: String(dateTime.getFullYear()).padStart(4, '0'),
    yy: twoDigits(dateTime.getFullYear() % 100),
    MM: twoDigits(dateTime.getMonth() + 1),
    M: String(dateTime.getMonth() + 1),
    dd: twoDigits(dateTime.getDate()),
    d: String(dateTime.getDate()),
    HH: twoDigits(dateTime.getHours()),
    H: String(dateTime.getHours()),
    hh: twoDigits(hour12),
    h: String(hour12),
    mm: twoDigits(dateTime.getMinutes()),
    m: String(dateTime.getMinutes()),
    ss: twoDigits(dateTime.getSeconds()),
    s: String(dateTime.getSeconds()),
    SSS: String(dateTime.getMilliseconds()).padStart(3, '0'),
    SS: String(dateTime.getMilliseconds()).padStart(3, '0').slice(0, 2),
    S: String(dateTime.getMilliseconds()).padStart(3, '0').slice(0, 1),
    a: dateTime.getHours() < 12 ? 'AM' : 'PM',
  }
  let output = ''
  let quoted = false
  for (let index = 0; index < pattern.length;) {
    if (pattern[index] === "'") {
      if (pattern[index + 1] === "'") {
        output += "'"
        index += 2
      } else {
        quoted = !quoted
        index += 1
      }
      continue
    }
    if (quoted) {
      output += pattern[index]
      index += 1
      continue
    }
    if (DATE_PATTERN_LETTERS.test(pattern[index]!)) {
      let end = index + 1
      while (pattern[end] === pattern[index]) end += 1
      const token = pattern.slice(index, end) as (typeof DATE_PATTERN_TOKENS)[number]
      if (!DATE_PATTERN_TOKENS.includes(token)) return null
      output += values[token]
      index = end
      continue
    }
    output += pattern[index]
    index += 1
  }
  return quoted ? null : output
}

export function parseStampPageRange(value: string, totalPages: number): number[] | null {
  const normalized = value
    .trim()
    .toLowerCase()
    .replace(/\s*[-–]\s*/g, '-')
  if (normalized === 'all') return Array.from({ length: totalPages }, (_, index) => index + 1)
  if (!normalized) return null
  const pages = new Set<number>()
  for (const token of normalized.split(/[,，;；\s]+/).filter(Boolean)) {
    const match = /^(\d+)(?:-(\d+))?$/.exec(token)
    if (!match) return null
    const start = Number(match[1])
    const end = Number(match[2] ?? match[1])
    if (start < 1 || end < start || end > totalPages) return null
    for (let page = start; page <= end; page += 1) pages.add(page)
  }
  return pages.size > 0 ? [...pages].sort((left, right) => left - right) : null
}

export function createWatermarkUuid(): string {
  const bytes = new Uint8Array(4)
  globalThis.crypto.getRandomValues(bytes)
  return Array.from(bytes, (value) => value.toString(16).padStart(2, '0')).join('')
}

export function resolveWatermarkText(
  template: string,
  {
    pageNumber,
    totalPages,
    fileName = '',
    dateTime = new Date(),
    title = '',
    author = '',
    subject = '',
    keywords = '',
    uuid = '',
  }: WatermarkTokenContext,
): string {
  const marker = '\u0000GENOFFICE_AT\u0000'
  const validDate = Number.isNaN(dateTime.getTime()) ? new Date() : dateTime
  const date = `${validDate.getFullYear()}-${twoDigits(validDate.getMonth() + 1)}-${twoDigits(validDate.getDate())}`
  const time = `${twoDigits(validDate.getHours())}:${twoDigits(validDate.getMinutes())}:${twoDigits(validDate.getSeconds())}`
  const extensionIndex = fileName.lastIndexOf('.')
  const fileStem = extensionIndex > 0 ? fileName.slice(0, extensionIndex) : fileName
  const customDatePattern = /@date\{([^}]*)\}/g

  return template
    .replaceAll('@@', marker)
    .replace(
      customDatePattern,
      (_token, pattern: string) =>
        formatWatermarkDatePattern(pattern, validDate) ?? '[invalid date format]',
    )
    .replaceAll('@filename_full', fileName)
    .replaceAll('@total_pages', String(totalPages))
    .replaceAll('@page_count', String(totalPages))
    .replaceAll('@page_number', String(pageNumber))
    .replaceAll('@datetime', `${date} ${time}`)
    .replaceAll('@filename', fileStem)
    .replaceAll('@keywords', keywords)
    .replaceAll('@subject', subject)
    .replaceAll('@author', author)
    .replaceAll('@title', title)
    .replaceAll('@uuid', uuid)
    .replaceAll('@date', date)
    .replaceAll('@time', time)
    .replaceAll('@year', String(validDate.getFullYear()))
    .replaceAll('@month', twoDigits(validDate.getMonth() + 1))
    .replaceAll('@day', twoDigits(validDate.getDate()))
    .replaceAll('@page', String(pageNumber))
    .replaceAll(marker, '@')
}

function singleWatermarkCenter(
  cfg: WatermarkConfig,
  page: { pw: number; ph: number },
  tile: WatermarkTile,
): readonly [number, number] {
  const radians = (cfg.angle * Math.PI) / 180
  const boundWidth =
    Math.abs(tile.width * Math.cos(radians)) + Math.abs(tile.height * Math.sin(radians))
  const boundHeight =
    Math.abs(tile.width * Math.sin(radians)) + Math.abs(tile.height * Math.cos(radians))
  const margin = ((page.pw + page.ph) / 2) * WATERMARK_MARGIN_FACTORS[cfg.margin]
  if (cfg.customPosition) {
    const xRatio = Number.isFinite(cfg.customPosition.xRatio)
      ? Math.max(0, Math.min(1, cfg.customPosition.xRatio))
      : 0.5
    const yRatio = Number.isFinite(cfg.customPosition.yRatio)
      ? Math.max(0, Math.min(1, cfg.customPosition.yRatio))
      : 0.5
    const targetX = xRatio * page.pw
    const targetY = yRatio * page.ph
    const halfX = Math.max(tile.width, boundWidth) / 2
    const halfY = Math.max(tile.height, boundHeight) / 2
    const minX = halfX
    const maxX = page.pw - halfX
    const minY = halfY
    const maxY = page.ph - halfY
    return [
      minX > maxX ? page.pw / 2 : Math.max(minX, Math.min(maxX, targetX)),
      minY > maxY ? page.ph / 2 : Math.max(minY, Math.min(maxY, targetY)),
    ]
  }
  if (cfg.position === 'center') return [page.pw / 2, page.ph / 2]
  const horizontal = cfg.position.endsWith('Left')
    ? 'left'
    : cfg.position.endsWith('Right')
      ? 'right'
      : 'center'
  const vertical = cfg.position.startsWith('top')
    ? 'top'
    : cfg.position.startsWith('bottom')
      ? 'bottom'
      : 'middle'
  const centerX =
    horizontal === 'left'
      ? margin + boundWidth / 2
      : horizontal === 'right'
        ? page.pw - margin - boundWidth / 2
        : page.pw / 2
  const centerY =
    vertical === 'top'
      ? page.ph - margin - boundHeight / 2
      : vertical === 'bottom'
        ? margin + boundHeight / 2
        : page.ph / 2
  return [
    Math.max(tile.width / 2, Math.min(page.pw - tile.width / 2, centerX)),
    Math.max(tile.height / 2, Math.min(page.ph - tile.height / 2, centerY)),
  ]
}

function addWatermarkStamps(
  out: StampInput[],
  cfg: WatermarkConfig,
  page: { origIdx: number; pw: number; ph: number },
  tile: WatermarkTile,
  limit = MAX_DOCUMENT_WATERMARK_TILES,
): number {
  let count = 0
  const addAt = (centerX: number, centerY: number) => {
    out.push({
      pageIndex: page.origIdx,
      image: tile.image,
      rect: [
        centerX - tile.width / 2,
        centerY - tile.height / 2,
        centerX + tile.width / 2,
        centerY + tile.height / 2,
      ],
      opacity: cfg.opacity,
      rotation: cfg.angle,
    })
    count += 1
  }

  if (cfg.layout === 'single') {
    addAt(...singleWatermarkCenter(cfg, page, tile))
    return count
  }

  const radians = (cfg.angle * Math.PI) / 180
  const boundWidth =
    Math.abs(tile.width * Math.cos(radians)) + Math.abs(tile.height * Math.sin(radians))
  const boundHeight =
    Math.abs(tile.width * Math.sin(radians)) + Math.abs(tile.height * Math.cos(radians))
  const stepX = Math.max(8, boundWidth + Math.max(0, cfg.horizontalSpacing))
  const stepY = Math.max(8, boundHeight + Math.max(0, cfg.verticalSpacing))
  let row = 0
  for (let centerY = -boundHeight / 2; centerY <= page.ph + boundHeight / 2; centerY += stepY) {
    const offset = row % 2 === 0 ? 0 : stepX / 2
    for (
      let centerX = -boundWidth / 2 + offset;
      centerX <= page.pw + boundWidth / 2;
      centerX += stepX
    ) {
      addAt(centerX, centerY)
      if (count >= limit) return count
    }
    row += 1
  }
  return count
}

/** Header or footer bar: transparent canvas of full page width × bar height, in left/center/right segments */
function renderBar(
  parts: [string, string, string],
  pw: number,
  fontSize: number,
  color: string,
  fontFamily: HeaderFooterFont = 'sans',
  horizontalMargin = pw * 0.06,
): string | null {
  if (parts.every((p) => !p.trim())) return null
  const h = Math.round(fontSize * 2.2)
  const canvas = document.createElement('canvas')
  canvas.width = Math.round(pw * SS)
  canvas.height = h * SS
  const ctx = canvas.getContext('2d')
  if (!ctx) return null
  ctx.scale(SS, SS)
  ctx.font = FONT(fontSize, false, fontFamily)
  ctx.fillStyle = color
  ctx.textBaseline = 'middle'
  const y = h / 2
  const [left, center, right] = parts
  if (left.trim()) {
    ctx.textAlign = 'left'
    ctx.fillText(left, horizontalMargin, y)
  }
  if (center.trim()) {
    ctx.textAlign = 'center'
    ctx.fillText(center, pw / 2, y)
  }
  if (right.trim()) {
    ctx.textAlign = 'right'
    ctx.fillText(right, pw - horizontalMargin, y)
  }
  return toBase64(canvas)
}

function fileNameWithoutExtension(fileName = ''): string {
  const simpleName = fileName.split(/[\\/]/).at(-1) ?? ''
  const extensionIndex = simpleName.lastIndexOf('.')
  return extensionIndex > 0 ? simpleName.slice(0, extensionIndex) : simpleName
}

/** Replace page-number placeholders shared by page numbers, headers, and footers. */
export function resolveHeaderFooterText(
  template: string,
  page: number,
  total: number,
  fileName = '',
  zeroPad = 0,
): string {
  const width = Math.max(0, Math.min(12, Math.trunc(zeroPad) || 0))
  const formattedPage = width > 0 ? String(page).padStart(width, '0') : String(page)
  return template
    .replaceAll('{page}', formattedPage)
    .replaceAll('{n}', formattedPage)
    .replaceAll('{total}', String(total))
    .replaceAll('{filename}', fileNameWithoutExtension(fileName))
}

function positionedBarRect(
  position: WatermarkPosition,
  pw: number,
  ph: number,
  barHeight: number,
  margin: WatermarkMargin,
): [number, number, number, number] {
  const marginY = ph * WATERMARK_MARGIN_FACTORS[margin]
  if (position.startsWith('top')) return [0, ph - marginY - barHeight, pw, ph - marginY]
  if (position.startsWith('middle') || position === 'center') {
    return [0, ph / 2 - barHeight / 2, pw, ph / 2 + barHeight / 2]
  }
  return [0, marginY, pw, marginY + barHeight]
}

function positionedBarParts(position: WatermarkPosition, text: string): [string, string, string] {
  if (position.endsWith('Left')) return [text, '', '']
  if (position.endsWith('Right')) return ['', '', text]
  return ['', text, '']
}

/**
 * Build stamps for each target page. pages provides (original page index, unrotated
 * page size, display number). Headers/footers are placed in unrotated coordinates and
 * follow rotation consistently with the page's /Rotate.
 */
export function buildStamps(
  pages: { origIdx: number; pw: number; ph: number; displayNo: number }[],
  watermark: WatermarkConfig | null,
  hf: HeaderFooterConfig | null,
  context: {
    fileName?: string
    watermarkPageIndexes?: readonly number[] | null
    headerFooterPageIndexes?: readonly number[] | null
    metadata?: { title?: string; author?: string; subject?: string; keywords?: string }
  } = {},
): StampInput[] {
  const out: StampInput[] = []
  const total = pages.length
  // Pages of the same size share one watermark bitmap, avoiding re-rendering in large docs.
  const wmCache = new Map<string, WatermarkTile | null>()
  let watermarkCount = 0
  const watermarkDateTime = new Date(watermark?.appliedAt || Date.now())
  const watermarkPageIndexes = context.watermarkPageIndexes
    ? new Set(context.watermarkPageIndexes)
    : null
  const headerFooterPageIndexes = context.headerFooterPageIndexes
    ? new Set(context.headerFooterPageIndexes)
    : null
  let headerFooterSequence = 0

  for (const p of pages) {
    if (watermark && (!watermarkPageIndexes || watermarkPageIndexes.has(p.origIdx))) {
      const text =
        watermark.type === 'text'
          ? resolveWatermarkText(watermark.text, {
              pageNumber: p.displayNo,
              totalPages: total,
              fileName: context.fileName,
              dateTime: watermarkDateTime,
              ...context.metadata,
              uuid: watermark.uuid,
            })
          : ''
      const key = `${Math.round(p.pw)}x${Math.round(p.ph)}:${text}`
      if (!wmCache.has(key)) wmCache.set(key, watermarkTile(watermark, text, p.pw, p.ph))
      const tile = wmCache.get(key)
      if (tile && watermarkCount < MAX_DOCUMENT_WATERMARK_TILES) {
        watermarkCount += addWatermarkStamps(
          out,
          watermark,
          p,
          tile,
          MAX_DOCUMENT_WATERMARK_TILES - watermarkCount,
        )
      }
    }

    if (!hf || (headerFooterPageIndexes && !headerFooterPageIndexes.has(p.origIdx))) continue
    headerFooterSequence += 1
    const no = hf.startAt + (headerFooterPageIndexes ? headerFooterSequence : p.displayNo) - 1
    const barH = hf.fontSize * 2.2
    const margin = Math.min(p.ph * 0.035, 26)
    const fontFamily = hf.fontFamily ?? 'sans'
    const pageNumberFormat = hf.pageNumberFormat || '{page} / {total}'
    const pageNumberPosition = hf.pageNumberPosition ?? 'bottomCenter'
    const pageNumberMargin = hf.pageNumberMargin ?? 'medium'
    const pageNumberZeroPad = hf.pageNumberZeroPad ?? 0
    const fill = (template: string, zeroPad = 0) =>
      resolveHeaderFooterText(template, no, total, context.fileName, zeroPad)
    const pageNumberText = fill(pageNumberFormat, pageNumberZeroPad)
    const headerParts = [hf.headerLeft, hf.headerCenter, hf.headerRight].map((s) => fill(s)) as [
      string,
      string,
      string,
    ]
    const footerParts = [hf.footerLeft, hf.footerCenter, hf.footerRight].map((s) => fill(s)) as [
      string,
      string,
      string,
    ]
    if (hf.pageNumber) {
      const index = pageNumberPosition.endsWith('Left')
        ? 0
        : pageNumberPosition.endsWith('Right')
          ? 2
          : 1
      if (pageNumberPosition.startsWith('top')) headerParts[index] = ''
      if (pageNumberPosition.startsWith('bottom')) footerParts[index] = ''
    }

    const header = renderBar(headerParts, p.pw, hf.fontSize, hf.color, fontFamily)
    if (header) {
      out.push({
        pageIndex: p.origIdx,
        image: header,
        rect: [0, p.ph - margin - barH, p.pw, p.ph - margin],
      })
    }

    const footer = renderBar(footerParts, p.pw, hf.fontSize, hf.color, fontFamily)
    if (footer) {
      out.push({ pageIndex: p.origIdx, image: footer, rect: [0, margin, p.pw, margin + barH] })
    }

    if (hf.pageNumber) {
      const pageNumber = renderBar(
        positionedBarParts(pageNumberPosition, pageNumberText),
        p.pw,
        hf.fontSize,
        hf.color,
        fontFamily,
        p.pw * WATERMARK_MARGIN_FACTORS[pageNumberMargin],
      )
      if (pageNumber) {
        out.push({
          pageIndex: p.origIdx,
          image: pageNumber,
          rect: positionedBarRect(pageNumberPosition, p.pw, p.ph, barH, pageNumberMargin),
        })
      }
    }
  }
  return out
}

/** Compact repeated watermark bitmaps before IPC / web save serialization. */
export function compactStampImages(stamps: StampInput[]): {
  stamps: StampInput[]
  stampImages: string[]
} {
  const stampImages: string[] = []
  const indices = new Map<string, number>()
  return {
    stamps: stamps.map((stamp) => {
      let imageIndex = indices.get(stamp.image)
      if (imageIndex === undefined) {
        imageIndex = stampImages.length
        indices.set(stamp.image, imageIndex)
        stampImages.push(stamp.image)
      }
      return { ...stamp, image: '', imageIndex }
    }),
    stampImages,
  }
}
