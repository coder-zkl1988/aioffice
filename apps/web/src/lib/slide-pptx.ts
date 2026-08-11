import PptxGenJS from 'pptxgenjs'
import type {
  SlideChartElement,
  SlideImageElement,
  SlideLineElement,
  SlideShapeElement,
  SlideSpec,
  SlideTableElement,
  SlideTextElement,
} from './slide-spec'

export interface FetchedSlideImage {
  base64: string
  mime: string
  ext: string
}

export type SlideImageFetcher = (url: string) => Promise<FetchedSlideImage | null>

export interface NativePptxResult {
  bytes: ArrayBuffer
  imageFailures: Array<{ page: number; url: string }>
}

const PX_PER_INCH = 96

function inches(px: number): number {
  return px / PX_PER_INCH
}

function pptxColor(value: string): string {
  return value.replace(/^#/, '').toUpperCase()
}

function transparency(opacity: number): number {
  return Math.max(0, Math.min(100, 100 - opacity))
}

function defaultFont(text: string): string {
  return /[\u3000-\u9fff\uf900-\ufaff]/.test(text) ? 'Noto Sans CJK SC' : 'Aptos'
}

function objectName(kind: string, index: number, name?: string): string {
  return name?.trim() || `GenOffice ${kind} ${index + 1}`
}

function addText(slide: PptxGenJS.Slide, element: SlideTextElement, index: number): void {
  slide.addText(element.text, {
    x: inches(element.x),
    y: inches(element.y),
    w: inches(element.w),
    h: inches(element.h),
    objectName: objectName('Text', index, element.name),
    fontFace: element.fontFace || defaultFont(element.text),
    fontSize: element.fontSize,
    color: pptxColor(element.color),
    bold: element.bold,
    italic: element.italic,
    align: element.align,
    valign: element.valign,
    bullet: element.bullet,
    margin: element.margin,
    breakLine: false,
    fit: 'shrink',
    isTextBox: true,
    transparency: transparency(element.opacity),
    ...(element.fill
      ? { fill: { color: pptxColor(element.fill), transparency: transparency(element.opacity) } }
      : {}),
    line: element.stroke
      ? { color: pptxColor(element.stroke), width: 1, transparency: transparency(element.opacity) }
      : { type: 'none' },
  })
}

function addShape(
  presentation: PptxGenJS,
  slide: PptxGenJS.Slide,
  element: SlideShapeElement,
  index: number,
): void {
  const shapes = presentation.ShapeType
  const shapeMap = {
    rect: shapes.rect,
    roundRect: shapes.roundRect,
    ellipse: shapes.ellipse,
    triangle: shapes.triangle,
    diamond: shapes.diamond,
    hexagon: shapes.hexagon,
    chevron: shapes.chevron,
    rightArrow: shapes.rightArrow,
    leftArrow: shapes.leftArrow,
    star5: shapes.star5,
  } as const
  slide.addShape(shapeMap[element.shape], {
    x: inches(element.x),
    y: inches(element.y),
    w: inches(element.w),
    h: inches(element.h),
    objectName: objectName('Shape', index, element.name),
    rotate: element.rotate,
    fill: element.fill
      ? { color: pptxColor(element.fill), transparency: transparency(element.opacity) }
      : { color: 'FFFFFF', transparency: 100 },
    line:
      element.stroke && element.strokeWidth > 0
        ? {
            color: pptxColor(element.stroke),
            width: element.strokeWidth,
            transparency: transparency(element.opacity),
          }
        : { type: 'none' },
  })
}

function addLine(
  presentation: PptxGenJS,
  slide: PptxGenJS.Slide,
  element: SlideLineElement,
  index: number,
): void {
  slide.addShape(presentation.ShapeType.line, {
    x: inches(element.x),
    y: inches(element.y),
    w: inches(element.w),
    h: inches(element.h),
    objectName: objectName('Line', index, element.name),
    line: {
      color: pptxColor(element.color),
      width: element.width,
      dashType: element.dash === 'dot' ? 'sysDot' : element.dash,
      beginArrowType: element.startArrow,
      endArrowType: element.endArrow,
      transparency: transparency(element.opacity),
    },
  })
}

function addImage(
  slide: PptxGenJS.Slide,
  element: SlideImageElement,
  image: FetchedSlideImage,
  index: number,
): void {
  const w = inches(element.w)
  const h = inches(element.h)
  slide.addImage({
    data: `${image.mime};base64,${image.base64}`,
    x: inches(element.x),
    y: inches(element.y),
    w,
    h,
    objectName: objectName('Image', index, element.name),
    altText: element.alt,
    rotate: element.rotate,
    transparency: transparency(element.opacity),
    sizing: { type: element.fit, w, h },
  })
}

function addTable(slide: PptxGenJS.Slide, element: SlideTableElement, index: number): void {
  const rows: PptxGenJS.TableRow[] = element.rows.map((row, rowIndex) =>
    row.map((cell) => ({
      text: cell,
      options: {
        bold: rowIndex < element.headerRows,
        color: pptxColor(rowIndex < element.headerRows ? element.headerColor : element.color),
        fill: {
          color: pptxColor(rowIndex < element.headerRows ? element.headerFill : element.bodyFill),
        },
        border: { color: pptxColor(element.borderColor), pt: 1 },
        margin: 6,
        valign: 'middle' as const,
      },
    })),
  )
  slide.addTable(rows, {
    x: inches(element.x),
    y: inches(element.y),
    w: inches(element.w),
    h: inches(element.h),
    objectName: objectName('Table', index, element.name),
    fontFace: element.fontFace || defaultFont(element.rows.flat().join('')),
    fontSize: element.fontSize,
    color: pptxColor(element.color),
    border: { color: pptxColor(element.borderColor), pt: 1 },
    fill: { color: pptxColor(element.bodyFill) },
    margin: 6,
    valign: 'middle',
    autoPage: false,
  })
}

function addChart(
  presentation: PptxGenJS,
  slide: PptxGenJS.Slide,
  element: SlideChartElement,
  index: number,
): void {
  const chartMap = {
    bar: presentation.ChartType.bar,
    line: presentation.ChartType.line,
    pie: presentation.ChartType.pie,
    doughnut: presentation.ChartType.doughnut,
  } as const
  slide.addChart(
    chartMap[element.chart],
    element.series.map((series) => ({
      name: series.name,
      labels: element.categories,
      values: series.values,
    })),
    {
      x: inches(element.x),
      y: inches(element.y),
      w: inches(element.w),
      h: inches(element.h),
      objectName: objectName('Chart', index, element.name),
      chartColors: element.colors.map(pptxColor),
      showLegend: element.showLegend,
      showValue: element.showValues,
      showTitle: Boolean(element.title),
      title: element.title,
      titleColor: pptxColor(element.textColor),
      titleFontFace: defaultFont(element.title || ''),
      titleFontSize: 18,
      catAxisLabelColor: pptxColor(element.textColor),
      catAxisLabelFontFace: defaultFont(element.categories.join('')),
      catAxisLabelFontSize: 12,
      valAxisLabelColor: pptxColor(element.textColor),
      valAxisLabelFontSize: 11,
      showPercent: element.chart === 'pie' || element.chart === 'doughnut',
      showLabel: element.chart === 'pie' || element.chart === 'doughnut',
      showSerName: false,
      legendColor: pptxColor(element.textColor),
      legendFontFace: defaultFont(element.series.map((series) => series.name).join('')),
      legendFontSize: 11,
      showCatAxisTitle: false,
      showValAxisTitle: false,
      catAxisLineColor: 'D9DDE7',
      valAxisLineColor: 'D9DDE7',
      valGridLine: { color: 'E8EBF2', size: 1 },
      showLeaderLines: true,
      dataLabelColor: pptxColor(element.textColor),
      dataLabelFontSize: 11,
      ...(element.chart === 'bar'
        ? { barDir: element.direction === 'bar' ? 'bar' : 'col', barGrouping: 'clustered' }
        : {}),
    },
  )
}

function arrayBuffer(output: ArrayBuffer | Uint8Array | Blob): Promise<ArrayBuffer> | ArrayBuffer {
  if (output instanceof ArrayBuffer) return output
  if (output instanceof Uint8Array) {
    return output.buffer.slice(
      output.byteOffset,
      output.byteOffset + output.byteLength,
    ) as ArrayBuffer
  }
  return output.arrayBuffer()
}

export async function compileNativeSlides(
  specs: SlideSpec[],
  fetchImage: SlideImageFetcher,
): Promise<NativePptxResult> {
  if (specs.length === 0 || specs.length > 100) throw new Error('生成页数无效')
  const presentation = new PptxGenJS()
  presentation.layout = 'LAYOUT_WIDE'
  presentation.author = 'GenOffice Web'
  presentation.company = 'GenOffice'
  presentation.subject = 'Editable AI-generated presentation'
  presentation.title = specs[0]?.title || 'Generated presentation'

  const imageUrls = [
    ...new Set(
      specs.flatMap((spec) =>
        spec.elements
          .filter((element): element is SlideImageElement => element.kind === 'image')
          .map((element) => element.url),
      ),
    ),
  ]
  const imageEntries = await Promise.all(
    imageUrls.map(async (url) => [url, await fetchImage(url)] as const),
  )
  const images = new Map(imageEntries)
  const imageFailures: Array<{ page: number; url: string }> = []

  specs.forEach((spec, pageIndex) => {
    const slide = presentation.addSlide()
    slide.background = { color: pptxColor(spec.background) }
    const elements = spec.elements
      .map((element, index) => ({ element, index }))
      .sort((left, right) => left.element.z - right.element.z || left.index - right.index)
    for (const { element, index } of elements) {
      switch (element.kind) {
        case 'text':
          addText(slide, element, index)
          break
        case 'shape':
          addShape(presentation, slide, element, index)
          break
        case 'line':
          addLine(presentation, slide, element, index)
          break
        case 'image': {
          const image = images.get(element.url)
          if (image) addImage(slide, element, image, index)
          else imageFailures.push({ page: pageIndex + 1, url: element.url })
          break
        }
        case 'table':
          addTable(slide, element, index)
          break
        case 'chart':
          addChart(presentation, slide, element, index)
          break
      }
    }
    if (spec.speakerNotes) slide.addNotes(spec.speakerNotes)
  })

  const output = await presentation.write({ outputType: 'arraybuffer', compression: true })
  return { bytes: await arrayBuffer(output as ArrayBuffer | Uint8Array | Blob), imageFailures }
}
