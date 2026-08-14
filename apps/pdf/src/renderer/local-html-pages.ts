import type { PdfRasterPage } from '@genoffice/pdf-tools'

const DEFAULT_PAGE_WIDTH_CSS = 794
const DEFAULT_PAGE_HEIGHT_CSS = 1123
const DEFAULT_PAGE_WIDTH_POINTS = 595.28
const DEFAULT_PAGE_HEIGHT_POINTS = 841.89
const DEFAULT_RASTER_SCALE = 1.5

export interface LocalHtmlPageOptions {
  maxPages?: number
  includePageNumbers?: boolean
  optimizeForEbook?: boolean
  pageWidthCss?: number
  pageHeightCss?: number
  pageWidthPoints?: number
  pageHeightPoints?: number
  rasterScale?: number
}

function iframeLoaded(frame: HTMLIFrameElement): Promise<void> {
  return new Promise((resolve, reject) => {
    frame.addEventListener('load', () => resolve(), { once: true })
    frame.addEventListener('error', () => reject(new Error('Could not render local HTML')), {
      once: true,
    })
  })
}

function canvasBytes(canvas: HTMLCanvasElement, optimize: boolean): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) return reject(new Error('Could not encode local HTML page'))
        void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
      },
      optimize ? 'image/jpeg' : 'image/png',
      optimize ? 0.84 : undefined,
    )
  })
}

export function localHtmlPageBreaks(
  body: HTMLElement,
  totalHeight: number,
  pageHeight: number,
  maxPages: number,
): number[] {
  const bodyTop = body.getBoundingClientRect().top
  const candidates = [
    ...body.querySelectorAll(
      'article,h1,h2,h3,h4,h5,h6,p,li,tr,pre,blockquote,img,[data-pdf-page-break-before]',
    ),
  ]
    .flatMap((element) => {
      const bounds = element.getBoundingClientRect()
      const top = bounds.top - bodyTop
      const bottom = bounds.bottom - bodyTop
      return element.hasAttribute('data-pdf-page-break-before') ? [top, bottom] : [bottom]
    })
    .filter((value) => Number.isFinite(value) && value > 0)
    .sort((left, right) => left - right)
  const forced = [...body.querySelectorAll('[data-pdf-page-break-before]')]
    .map((element) => element.getBoundingClientRect().top - bodyTop)
    .filter((value) => Number.isFinite(value) && value > 0)
    .sort((left, right) => left - right)
  const breaks = [0]
  while (breaks.at(-1)! + pageHeight < totalHeight) {
    const start = breaks.at(-1)!
    const target = start + pageHeight
    const forcedBreak = forced.find((value) => value > start + 40 && value <= target)
    const minimum = start + pageHeight * 0.72
    const candidate = candidates.filter((value) => value >= minimum && value <= target - 24).at(-1)
    breaks.push(forcedBreak ?? (candidate && candidate > start + 40 ? candidate : target))
    if (breaks.length > maxPages) {
      throw new Error(`Document may contain no more than ${maxPages} pages`)
    }
  }
  return breaks
}

function grayscaleCanvas(context: CanvasRenderingContext2D, width: number, height: number): void {
  const image = context.getImageData(0, 0, width, height)
  for (let offset = 0; offset < image.data.length; offset += 4) {
    const gray = Math.round(
      image.data[offset]! * 0.299 +
        image.data[offset + 1]! * 0.587 +
        image.data[offset + 2]! * 0.114,
    )
    image.data[offset] = gray
    image.data[offset + 1] = gray
    image.data[offset + 2] = gray
  }
  context.putImageData(image, 0, 0)
}

export async function renderLocalHtmlPages(
  html: string,
  options: LocalHtmlPageOptions = {},
): Promise<PdfRasterPage[]> {
  const pageWidthCss = options.pageWidthCss ?? DEFAULT_PAGE_WIDTH_CSS
  const pageHeightCss = options.pageHeightCss ?? DEFAULT_PAGE_HEIGHT_CSS
  const pageWidthPoints = options.pageWidthPoints ?? DEFAULT_PAGE_WIDTH_POINTS
  const pageHeightPoints = options.pageHeightPoints ?? DEFAULT_PAGE_HEIGHT_POINTS
  const rasterScale = options.rasterScale ?? DEFAULT_RASTER_SCALE
  const maxPages = options.maxPages ?? 40
  if (
    !Number.isFinite(pageWidthCss) ||
    !Number.isFinite(pageHeightCss) ||
    pageWidthCss < 100 ||
    pageHeightCss < 100 ||
    !Number.isFinite(rasterScale) ||
    rasterScale < 1 ||
    rasterScale > 3 ||
    !Number.isInteger(maxPages) ||
    maxPages < 1 ||
    maxPages > 200
  ) {
    throw new Error('Local HTML page settings are invalid')
  }

  const frame = document.createElement('iframe')
  frame.setAttribute('sandbox', 'allow-same-origin')
  frame.setAttribute('aria-hidden', 'true')
  frame.style.position = 'fixed'
  frame.style.left = '-100000px'
  frame.style.top = '0'
  frame.style.width = `${pageWidthCss}px`
  frame.style.height = `${pageHeightCss}px`
  frame.style.border = '0'
  document.body.appendChild(frame)
  const loaded = iframeLoaded(frame)
  frame.srcdoc = html
  try {
    await loaded
    const frameDocument = frame.contentDocument
    if (!frameDocument) throw new Error('Could not render local HTML')
    await frameDocument.fonts?.ready
    await Promise.all(
      [...frameDocument.images].map((image) => image.decode().catch(() => undefined)),
    )
    const body = frameDocument.body
    const totalHeight = Math.max(body.scrollHeight, body.getBoundingClientRect().height, 1)
    const breaks = localHtmlPageBreaks(body, totalHeight, pageHeightCss, maxPages)
    const { default: html2canvas } = await import('html2canvas')
    const pages: PdfRasterPage[] = []
    for (let index = 0; index < breaks.length; index++) {
      const y = breaks[index]!
      const next = breaks[index + 1] ?? totalHeight
      const contentHeight = Math.min(pageHeightCss, Math.max(1, next - y))
      const rendered = await html2canvas(body, {
        allowTaint: false,
        backgroundColor: '#ffffff',
        height: contentHeight,
        logging: false,
        scale: rasterScale,
        scrollX: 0,
        scrollY: 0,
        useCORS: false,
        width: pageWidthCss,
        windowHeight: pageHeightCss,
        windowWidth: pageWidthCss,
        x: 0,
        y,
      })
      const page = document.createElement('canvas')
      page.width = Math.round(pageWidthCss * rasterScale)
      page.height = Math.round(pageHeightCss * rasterScale)
      const context = page.getContext('2d')
      if (!context) throw new Error('Could not create local HTML page canvas')
      context.fillStyle = '#ffffff'
      context.fillRect(0, 0, page.width, page.height)
      context.drawImage(rendered, 0, 0)
      if (options.includePageNumbers) {
        context.fillStyle = '#ffffff'
        context.fillRect(0, page.height - 34 * rasterScale, page.width, 34 * rasterScale)
        context.fillStyle = '#666666'
        context.font = `${11 * rasterScale}px -apple-system,BlinkMacSystemFont,"Segoe UI",Arial,sans-serif`
        context.textAlign = 'center'
        context.textBaseline = 'middle'
        context.fillText(String(index + 1), page.width / 2, page.height - 16 * rasterScale)
      }
      if (options.optimizeForEbook) grayscaleCanvas(context, page.width, page.height)
      pages.push({
        image: await canvasBytes(page, options.optimizeForEbook === true),
        width: pageWidthPoints,
        height: pageHeightPoints,
      })
    }
    return pages
  } finally {
    frame.remove()
  }
}
