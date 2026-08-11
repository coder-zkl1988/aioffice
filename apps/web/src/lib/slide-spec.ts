import { GENOFFICE_NATIVE_VISUAL_SKILL } from './slide-visual-skill'

export interface WebSlideSpecRequest {
  brief: string
  title?: string
  styleSkill?: string
  deckContext?: Record<string, unknown>
  images?: { url: string; caption?: string }[]
  width?: number
  height?: number
}

export interface WebSlideSpecPrompts {
  system: string
  user: string
}

type TextAlign = 'left' | 'center' | 'right'
type VerticalAlign = 'top' | 'middle' | 'bottom'
type ShapeName =
  | 'rect'
  | 'roundRect'
  | 'ellipse'
  | 'triangle'
  | 'diamond'
  | 'hexagon'
  | 'chevron'
  | 'rightArrow'
  | 'leftArrow'
  | 'star5'
type DashName = 'solid' | 'dash' | 'dot'
type ArrowName = 'none' | 'arrow' | 'triangle' | 'stealth'
type ChartName = 'bar' | 'line' | 'pie' | 'doughnut'

interface SlideElementBase {
  x: number
  y: number
  w: number
  h: number
  z: number
  name?: string
}

export interface SlideTextElement extends SlideElementBase {
  kind: 'text'
  text: string
  fontSize: number
  fontFace?: string
  color: string
  bold: boolean
  italic: boolean
  align: TextAlign
  valign: VerticalAlign
  bullet: boolean
  margin: number
  fill?: string
  stroke?: string
  opacity: number
}

export interface SlideShapeElement extends SlideElementBase {
  kind: 'shape'
  shape: ShapeName
  fill?: string
  stroke?: string
  strokeWidth: number
  opacity: number
  rotate: number
}

export interface SlideLineElement extends SlideElementBase {
  kind: 'line'
  color: string
  width: number
  dash: DashName
  startArrow: ArrowName
  endArrow: ArrowName
  opacity: number
}

export interface SlideImageElement extends SlideElementBase {
  kind: 'image'
  url: string
  fit: 'cover' | 'contain'
  alt: string
  opacity: number
  rotate: number
}

export interface SlideTableElement extends SlideElementBase {
  kind: 'table'
  rows: string[][]
  headerRows: number
  fontSize: number
  fontFace?: string
  color: string
  headerColor: string
  headerFill: string
  bodyFill: string
  borderColor: string
  accentColor: string
}

export interface SlideChartElement extends SlideElementBase {
  kind: 'chart'
  chart: ChartName
  title?: string
  categories: string[]
  series: Array<{ name: string; values: number[] }>
  colors: string[]
  showLegend: boolean
  showValues: boolean
  direction: 'column' | 'bar'
  textColor: string
}

export type SlideSpecElement =
  | SlideTextElement
  | SlideShapeElement
  | SlideLineElement
  | SlideImageElement
  | SlideTableElement
  | SlideChartElement

export interface SlideSpec {
  version: 1
  title: string
  layout: string
  width: number
  height: number
  background: string
  elements: SlideSpecElement[]
  speakerNotes?: string
}

export const SLIDE_SPEC_MARKER_PREFIX = 'genoffice-slide-spec:v1:'

const SHAPES = new Set<ShapeName>([
  'rect',
  'roundRect',
  'ellipse',
  'triangle',
  'diamond',
  'hexagon',
  'chevron',
  'rightArrow',
  'leftArrow',
  'star5',
])
const ALIGNS = new Set<TextAlign>(['left', 'center', 'right'])
const VALIGNS = new Set<VerticalAlign>(['top', 'middle', 'bottom'])
const DASHES = new Set<DashName>(['solid', 'dash', 'dot'])
const ARROWS = new Set<ArrowName>(['none', 'arrow', 'triangle', 'stealth'])
const CHARTS = new Set<ChartName>(['bar', 'line', 'pie', 'doughnut'])

function boundedDimension(value: number | undefined, fallback: number): number {
  if (!Number.isFinite(value)) return fallback
  return Math.min(3840, Math.max(320, Math.round(value!)))
}

export function webSlideSpecDimensions(request: Pick<WebSlideSpecRequest, 'width' | 'height'>): {
  width: number
  height: number
} {
  return {
    width: boundedDimension(request.width, 1280),
    height: boundedDimension(request.height, 720),
  }
}

export function buildWebSlideSpecPrompts(request: WebSlideSpecRequest): WebSlideSpecPrompts {
  const { width, height } = webSlideSpecDimensions(request)
  const images = (request.images ?? [])
    .filter((image) => /^https?:\/\//i.test(image.url))
    .slice(0, 8)
  const approvedImages = images.length
    ? images
        .map(
          (image, index) =>
            `${index + 1}. ${image.url}${image.caption ? ` - ${image.caption}` : ''}`,
        )
        .join('\n')
    : '(none)'

  return {
    system: [
      'You are a senior presentation designer creating one editable PowerPoint slide.',
      `The fixed canvas is ${width}x${height} px. Return JSON only, without Markdown or explanation.`,
      GENOFFICE_NATIVE_VISUAL_SKILL,
      'Output this exact top-level structure:',
      '{"version":1,"title":"takeaway title","layout":"semantic layout id","background":"#RRGGBB","speakerNotes":"optional","elements":[...]}',
      'Every element needs x,y,w,h in pixels and optional z (0-100). Coordinates must stay inside the canvas.',
      'Allowed element schemas:',
      '- text: {"kind":"text","text":"...","x":0,"y":0,"w":100,"h":50,"z":10,"fontSize":32,"fontFace":"optional","color":"#RRGGBB","bold":false,"italic":false,"align":"left|center|right","valign":"top|middle|bottom","bullet":false,"margin":0,"fill":"optional #RRGGBB","stroke":"optional #RRGGBB","opacity":100}',
      '- shape: {"kind":"shape","shape":"rect|roundRect|ellipse|triangle|diamond|hexagon|chevron|rightArrow|leftArrow|star5","x":0,"y":0,"w":100,"h":50,"z":1,"fill":"optional #RRGGBB","stroke":"optional #RRGGBB","strokeWidth":1,"opacity":100,"rotate":0}',
      '- line: {"kind":"line","x":0,"y":0,"w":100,"h":0,"z":3,"color":"#RRGGBB","width":2,"dash":"solid|dash|dot","startArrow":"none|arrow|triangle|stealth","endArrow":"none|arrow|triangle|stealth","opacity":100}',
      '- image: {"kind":"image","url":"approved URL","x":0,"y":0,"w":100,"h":100,"z":2,"fit":"cover|contain","alt":"description","opacity":100,"rotate":0}',
      '- table: {"kind":"table","rows":[["A","B"]],"headerRows":1,"x":0,"y":0,"w":100,"h":100,"z":5,"fontSize":18,"fontFace":"optional","color":"#RRGGBB","headerColor":"#RRGGBB","headerFill":"#RRGGBB","bodyFill":"#RRGGBB","borderColor":"#RRGGBB","accentColor":"#RRGGBB"}',
      '- chart: {"kind":"chart","chart":"bar|line|pie|doughnut","title":"optional","categories":["A"],"series":[{"name":"Series","values":[1]}],"x":0,"y":0,"w":100,"h":100,"z":5,"colors":["#RRGGBB"],"showLegend":true,"showValues":false,"direction":"column|bar","textColor":"#RRGGBB"}',
      'Use only approved image URLs. If no image is approved, use native text, shapes, lines, tables, and charts.',
      'For Chinese text, prefer fontFace "Noto Sans CJK SC". For Latin text, prefer "Aptos" or "Arial".',
      'Use z to place backgrounds first and text last. Do not rasterize text into images.',
      'Keep copy concise enough to fit. Preserve the language of the supplied content.',
    ].join('\n'),
    user: [
      `Slide title: ${request.title?.trim() || '(derive a concise takeaway title)'}`,
      `Slide brief:\n${request.brief.trim()}`,
      `Deck design system:\n${request.styleSkill?.trim() || '(clean, modern, professional)'}`,
      `Deck context:\n${JSON.stringify(request.deckContext ?? {}, null, 2)}`,
      `Approved image URLs:\n${approvedImages}`,
    ].join('\n\n'),
  }
}

function record(value: unknown, field: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${field} 必须是对象`)
  }
  return value as Record<string, unknown>
}

function textValue(value: unknown, fallback = '', max = 4000): string {
  return (typeof value === 'string' ? value : fallback).trim().slice(0, max)
}

function numberValue(value: unknown, fallback: number, min: number, max: number): number {
  const number = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(number)) return fallback
  return Math.min(max, Math.max(min, number))
}

function booleanValue(value: unknown, fallback: boolean): boolean {
  return typeof value === 'boolean' ? value : fallback
}

function enumValue<T extends string>(value: unknown, allowed: Set<T>, fallback: T): T {
  return typeof value === 'string' && allowed.has(value as T) ? (value as T) : fallback
}

function colorValue(value: unknown, fallback: string): string {
  if (typeof value !== 'string') return fallback
  const match = /^#?([0-9a-f]{6})$/i.exec(value.trim())
  return match ? `#${match[1]!.toUpperCase()}` : fallback
}

function optionalColor(value: unknown): string | undefined {
  if (value === undefined || value === null || value === '') return undefined
  return colorValue(value, '#000000')
}

function position(raw: Record<string, unknown>, width: number, height: number): SlideElementBase {
  const x = numberValue(raw.x, 0, 0, width - 1)
  const y = numberValue(raw.y, 0, 0, height - 1)
  const w = numberValue(raw.w, 100, 1, width - x)
  const h = numberValue(raw.h, 50, 1, height - y)
  return {
    x,
    y,
    w,
    h,
    z: numberValue(raw.z, 10, 0, 100),
    ...(textValue(raw.name, '', 80) ? { name: textValue(raw.name, '', 80) } : {}),
  }
}

function fitFontSize(text: string, width: number, height: number, requested: number): number {
  const lineBreaks = Math.max(1, text.split('\n').length)
  for (let size = Math.round(requested); size >= 14; size -= 1) {
    const charWidthPx = size * (96 / 72) * 0.55
    const lineHeightPx = size * (96 / 72) * 1.25
    const charsPerLine = Math.max(1, Math.floor(width / charWidthPx))
    const estimatedLines = Math.max(lineBreaks, Math.ceil(text.length / charsPerLine))
    if (estimatedLines * lineHeightPx <= height) return size
  }
  return 14
}

function parseTextElement(
  raw: Record<string, unknown>,
  width: number,
  height: number,
): SlideTextElement {
  const base = position(raw, width, height)
  const text = textValue(raw.text)
  if (!text) throw new Error('文本元素不能为空')
  const requestedFontSize = numberValue(raw.fontSize, 24, 14, 72)
  return {
    ...base,
    kind: 'text',
    text,
    fontSize: fitFontSize(text, base.w, base.h, requestedFontSize),
    ...(textValue(raw.fontFace, '', 80) ? { fontFace: textValue(raw.fontFace, '', 80) } : {}),
    color: colorValue(raw.color, '#171717'),
    bold: booleanValue(raw.bold, false),
    italic: booleanValue(raw.italic, false),
    align: enumValue(raw.align, ALIGNS, 'left'),
    valign: enumValue(raw.valign, VALIGNS, 'top'),
    bullet: booleanValue(raw.bullet, false),
    margin: numberValue(raw.margin, 0, 0, 24),
    ...(optionalColor(raw.fill) ? { fill: optionalColor(raw.fill) } : {}),
    ...(optionalColor(raw.stroke) ? { stroke: optionalColor(raw.stroke) } : {}),
    opacity: numberValue(raw.opacity, 100, 0, 100),
  }
}

function parseShapeElement(
  raw: Record<string, unknown>,
  width: number,
  height: number,
): SlideShapeElement {
  return {
    ...position(raw, width, height),
    kind: 'shape',
    shape: enumValue(raw.shape, SHAPES, 'rect'),
    ...(optionalColor(raw.fill) ? { fill: optionalColor(raw.fill) } : {}),
    ...(optionalColor(raw.stroke) ? { stroke: optionalColor(raw.stroke) } : {}),
    strokeWidth: numberValue(raw.strokeWidth, 1, 0, 12),
    opacity: numberValue(raw.opacity, 100, 0, 100),
    rotate: numberValue(raw.rotate, 0, -360, 360),
  }
}

function parseLineElement(
  raw: Record<string, unknown>,
  width: number,
  height: number,
): SlideLineElement {
  return {
    ...position(raw, width, height),
    kind: 'line',
    color: colorValue(raw.color, '#171717'),
    width: numberValue(raw.width, 2, 0.5, 12),
    dash: enumValue(raw.dash, DASHES, 'solid'),
    startArrow: enumValue(raw.startArrow, ARROWS, 'none'),
    endArrow: enumValue(raw.endArrow, ARROWS, 'none'),
    opacity: numberValue(raw.opacity, 100, 0, 100),
  }
}

function parseImageElement(
  raw: Record<string, unknown>,
  width: number,
  height: number,
  approvedImages: Set<string>,
): SlideImageElement {
  const url = textValue(raw.url, '', 3000)
  if (!approvedImages.has(url)) throw new Error(`图片 URL 未获批准：${url || '(empty)'}`)
  return {
    ...position(raw, width, height),
    kind: 'image',
    url,
    fit: raw.fit === 'contain' ? 'contain' : 'cover',
    alt: textValue(raw.alt, 'Presentation image', 500),
    opacity: numberValue(raw.opacity, 100, 0, 100),
    rotate: numberValue(raw.rotate, 0, -360, 360),
  }
}

function parseTableElement(
  raw: Record<string, unknown>,
  width: number,
  height: number,
): SlideTableElement {
  const sourceRows = Array.isArray(raw.rows) ? raw.rows.slice(0, 10) : []
  const rows = sourceRows.map((row) =>
    (Array.isArray(row) ? row : []).slice(0, 8).map((cell) => textValue(cell, '', 300)),
  )
  const columns = Math.max(0, ...rows.map((row) => row.length))
  if (rows.length === 0 || columns === 0) throw new Error('表格元素必须包含数据')
  for (const row of rows) while (row.length < columns) row.push('')
  return {
    ...position(raw, width, height),
    kind: 'table',
    rows,
    headerRows: Math.round(numberValue(raw.headerRows, 1, 0, Math.min(3, rows.length))),
    fontSize: numberValue(raw.fontSize, 18, 12, 32),
    ...(textValue(raw.fontFace, '', 80) ? { fontFace: textValue(raw.fontFace, '', 80) } : {}),
    color: colorValue(raw.color, '#171717'),
    headerColor: colorValue(raw.headerColor, '#FFFFFF'),
    headerFill: colorValue(raw.headerFill, '#1F5EFF'),
    bodyFill: colorValue(raw.bodyFill, '#FFFFFF'),
    borderColor: colorValue(raw.borderColor, '#D9DDE7'),
    accentColor: colorValue(raw.accentColor, '#1F5EFF'),
  }
}

function parseChartElement(
  raw: Record<string, unknown>,
  width: number,
  height: number,
): SlideChartElement {
  const categories = (Array.isArray(raw.categories) ? raw.categories : [])
    .slice(0, 12)
    .map((value) => textValue(value, '', 100))
  const sourceSeries = (Array.isArray(raw.series) ? raw.series : []).slice(0, 6)
  const series = sourceSeries.map((item, index) => {
    const source = record(item, `series[${index}]`)
    const values = (Array.isArray(source.values) ? source.values : [])
      .slice(0, categories.length)
      .map((value) => numberValue(value, 0, -1_000_000_000, 1_000_000_000))
    while (values.length < categories.length) values.push(0)
    return { name: textValue(source.name, `Series ${index + 1}`, 100), values }
  })
  if (categories.length === 0 || series.length === 0) throw new Error('图表元素必须包含分类和序列')
  const colors = (Array.isArray(raw.colors) ? raw.colors : [])
    .slice(0, 8)
    .map((value) => colorValue(value, '#1F5EFF'))
  return {
    ...position(raw, width, height),
    kind: 'chart',
    chart: enumValue(raw.chart, CHARTS, 'bar'),
    ...(textValue(raw.title, '', 200) ? { title: textValue(raw.title, '', 200) } : {}),
    categories,
    series,
    colors: colors.length ? colors : ['#1F5EFF', '#16A085', '#F4B740', '#E05252'],
    showLegend: booleanValue(raw.showLegend, series.length > 1),
    showValues: booleanValue(raw.showValues, false),
    direction: raw.direction === 'bar' ? 'bar' : 'column',
    textColor: colorValue(raw.textColor, '#171717'),
  }
}

function extractJsonSource(value: string): string {
  const trimmed = value.trim()
  if (!trimmed) throw new Error('模型未返回幻灯片规范')
  if (trimmed.length > 500_000) throw new Error('模型返回的幻灯片规范过大')
  const fenced = /```(?:json)?\s*([\s\S]*?)```/i.exec(trimmed)
  const candidate = (fenced?.[1] ?? trimmed).trim()
  const start = candidate.indexOf('{')
  const end = candidate.lastIndexOf('}')
  if (start < 0 || end <= start) throw new Error('模型未返回有效 JSON')
  return candidate.slice(start, end + 1)
}

export function parseGeneratedSlideSpec(value: string, request: WebSlideSpecRequest): SlideSpec {
  let decoded: unknown
  try {
    decoded = JSON.parse(extractJsonSource(value))
  } catch (error) {
    throw new Error(
      `无法解析模型返回的幻灯片 JSON：${error instanceof Error ? error.message : String(error)}`,
      { cause: error },
    )
  }
  const raw = record(decoded, '幻灯片规范')
  const { width, height } = webSlideSpecDimensions(request)
  const approvedImages = new Set(
    (request.images ?? [])
      .map((image) => image.url.trim())
      .filter((url) => /^https?:\/\//i.test(url)),
  )
  const sourceElements = Array.isArray(raw.elements) ? raw.elements.slice(0, 40) : []
  const elements = sourceElements.map((value, index): SlideSpecElement => {
    const element = record(value, `elements[${index}]`)
    switch (element.kind) {
      case 'text':
        return parseTextElement(element, width, height)
      case 'shape':
        return parseShapeElement(element, width, height)
      case 'line':
        return parseLineElement(element, width, height)
      case 'image':
        return parseImageElement(element, width, height, approvedImages)
      case 'table':
        return parseTableElement(element, width, height)
      case 'chart':
        return parseChartElement(element, width, height)
      default:
        throw new Error(`不支持的元素类型：${String(element.kind)}`)
    }
  })
  const hasMeaningfulContent = elements.some(
    (element) => element.kind === 'text' || element.kind === 'table' || element.kind === 'chart',
  )
  if (!hasMeaningfulContent) throw new Error('幻灯片必须包含可编辑的文本、表格或图表')
  return {
    version: 1,
    title: textValue(raw.title, request.title || 'Untitled slide', 300),
    layout: textValue(raw.layout, 'title_body', 80),
    width,
    height,
    background: colorValue(raw.background, '#FFFFFF'),
    elements,
    ...(textValue(raw.speakerNotes, '', 8000)
      ? { speakerNotes: textValue(raw.speakerNotes, '', 8000) }
      : {}),
  }
}

export function createSlideSpecMarker(spec: SlideSpec): string {
  return `${SLIDE_SPEC_MARKER_PREFIX}${encodeURIComponent(JSON.stringify(spec))}`
}

export function parseSlideSpecMarker(value: string): SlideSpec | null {
  if (!value.startsWith(SLIDE_SPEC_MARKER_PREFIX)) return null
  try {
    const decoded = JSON.parse(decodeURIComponent(value.slice(SLIDE_SPEC_MARKER_PREFIX.length)))
    const raw = record(decoded, '幻灯片标记')
    if (raw.version !== 1 || !Array.isArray(raw.elements)) throw new Error('版本或元素无效')
    return decoded as SlideSpec
  } catch (error) {
    throw new Error(`幻灯片标记损坏：${error instanceof Error ? error.message : String(error)}`, {
      cause: error,
    })
  }
}
