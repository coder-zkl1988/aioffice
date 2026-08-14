import type { PdfImagesToPdfPage, PdfRasterPage, PdfVectorPage } from '@genoffice/pdf-tools'

export type PdfImageColorMode = 'color' | 'greyscale' | 'blackwhite'

const MAX_IMAGE_FILES = 200
const MAX_IMAGE_FILE_BYTES = 50 * 1024 * 1024
const MAX_SVG_FILE_BYTES = 10 * 1024 * 1024
const MAX_TOTAL_IMAGE_BYTES = 200 * 1024 * 1024
const MAX_IMAGE_PIXELS = 40_000_000
const MAX_IMAGE_DIMENSION = 16_384
const SVG_RASTER_SCALE = 2

interface DecodedPdfImage {
  source: CanvasImageSource
  width: number
  height: number
  rasterWidth: number
  rasterHeight: number
  close: () => void
}

export interface SanitizedSvg {
  markup: string
  width: number
  height: number
}

function luminance(red: number, green: number, blue: number): number {
  return Math.round(red * 0.299 + green * 0.587 + blue * 0.114)
}

export function convertImagePixelsForPdf(
  data: Uint8ClampedArray,
  colorMode: PdfImageColorMode,
): Uint8ClampedArray {
  if (colorMode === 'color') return data
  for (let offset = 0; offset < data.length; offset += 4) {
    const value = luminance(data[offset]!, data[offset + 1]!, data[offset + 2]!)
    const output = colorMode === 'blackwhite' ? (value < 128 ? 0 : 255) : value
    data[offset] = output
    data[offset + 1] = output
    data[offset + 2] = output
  }
  return data
}

function canvasPngBytes(canvas: HTMLCanvasElement): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob((blob) => {
      if (!blob) {
        reject(new Error('Could not encode image as PNG'))
        return
      }
      void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
    }, 'image/png')
  })
}

function svgLengthPixels(value: string | null): number | undefined {
  if (!value || value.trim().endsWith('%')) return undefined
  const match = /^\s*(\d+(?:\.\d+)?|\.\d+)\s*(px|pt|pc|in|cm|mm)?\s*$/iu.exec(value)
  if (!match) return undefined
  const number = Number(match[1])
  const scale =
    match[2]?.toLowerCase() === 'pt'
      ? 96 / 72
      : match[2]?.toLowerCase() === 'pc'
        ? 16
        : match[2]?.toLowerCase() === 'in'
          ? 96
          : match[2]?.toLowerCase() === 'cm'
            ? 96 / 2.54
            : match[2]?.toLowerCase() === 'mm'
              ? 96 / 25.4
              : 1
  const pixels = number * scale
  return Number.isFinite(pixels) && pixels > 0 ? pixels : undefined
}

function svgViewBox(root: Element): readonly [number, number] | undefined {
  const values = root
    .getAttribute('viewBox')
    ?.trim()
    .split(/[\s,]+/u)
    .map(Number)
  if (
    !values ||
    values.length !== 4 ||
    values.some((value) => !Number.isFinite(value)) ||
    values[2]! <= 0 ||
    values[3]! <= 0
  ) {
    return undefined
  }
  return [values[2]!, values[3]!]
}

function svgDimensions(root: Element): readonly [number, number] {
  const viewBox = svgViewBox(root)
  let width = svgLengthPixels(root.getAttribute('width'))
  let height = svgLengthPixels(root.getAttribute('height'))
  if (width && !height && viewBox) height = width * (viewBox[1] / viewBox[0])
  if (height && !width && viewBox) width = height * (viewBox[0] / viewBox[1])
  width ??= viewBox?.[0] ?? 300
  height ??= viewBox?.[1] ?? 150
  if (
    width > MAX_IMAGE_DIMENSION ||
    height > MAX_IMAGE_DIMENSION ||
    width * height > MAX_IMAGE_PIXELS
  ) {
    throw new Error('SVG dimensions are too large')
  }
  return [width, height]
}

function safeSvgReference(value: string): boolean {
  const reference = value.trim()
  return (
    reference.startsWith('#') ||
    /^data:image\/(?:png|jpe?g|webp);base64,[a-z\d+/=\s]+$/iu.test(reference)
  )
}

function hasUnsafeCss(value: string): boolean {
  if (/@import|expression\s*\(|javascript\s*:/iu.test(value)) return true
  return [...value.matchAll(/url\(\s*(['"]?)(.*?)\1\s*\)/giu)].some(
    (match) => !safeSvgReference(match[2] ?? ''),
  )
}

function svgStyleProperty(style: string | null, property: string): string | undefined {
  if (!style) return undefined
  const match = new RegExp(`(?:^|;)\\s*${property}\\s*:\\s*([^;]+)`, 'iu').exec(style)
  return match?.[1]?.trim()
}

function gradientStrokeFallback(root: Element, id: string): string | undefined {
  const gradient = [root, ...Array.from(root.querySelectorAll('[id]'))].find(
    (element) => element.getAttribute('id') === id,
  )
  if (
    !gradient ||
    !['lineargradient', 'radialgradient'].includes(gradient.localName.toLowerCase())
  ) {
    return undefined
  }
  const colors = Array.from(gradient.children)
    .filter((element) => element.localName.toLowerCase() === 'stop')
    .map(
      (stop) =>
        stop.getAttribute('stop-color')?.trim() ??
        svgStyleProperty(stop.getAttribute('style'), 'stop-color'),
    )
    .filter(
      (color): color is string =>
        typeof color === 'string' && color.length > 0 && !hasUnsafeCss(color),
    )
  return colors[Math.floor(colors.length / 2)]
}

function replaceGradientStrokeDeclarations(value: string, root: Element): string {
  return value.replace(
    /(\bstroke\s*:\s*)url\(\s*(['"]?)#([^)'"\s]+)\2\s*\)/giu,
    (match, prefix: string, _quote: string, id: string) => {
      const fallback = gradientStrokeFallback(root, id)
      return fallback ? `${prefix}${fallback}` : match
    },
  )
}

export function sanitizeSvgForPdf(svg: string): SanitizedSvg {
  const document = new DOMParser().parseFromString(svg, 'image/svg+xml')
  const root = document.documentElement
  if (
    document.doctype ||
    root.localName.toLowerCase() !== 'svg' ||
    root.querySelector('parsererror')
  ) {
    throw new Error('SVG markup is invalid')
  }

  root
    .querySelectorAll(
      'script, foreignObject, iframe, object, embed, audio, video, animate, animateMotion, animateTransform, set',
    )
    .forEach((element) => element.remove())
  root.querySelectorAll('style').forEach((element) => {
    if (hasUnsafeCss(element.textContent ?? '')) element.remove()
  })
  for (const element of [root, ...Array.from(root.querySelectorAll('*'))]) {
    for (const attribute of Array.from(element.attributes)) {
      const name = attribute.name.toLowerCase()
      if (
        name.startsWith('on') ||
        ((name === 'href' || name.endsWith(':href')) && !safeSvgReference(attribute.value)) ||
        ((name === 'style' || attribute.value.includes('url(')) && hasUnsafeCss(attribute.value))
      ) {
        element.removeAttribute(attribute.name)
      }
    }
    const stroke = element.getAttribute('stroke')
    const strokeReference = /^\s*url\(\s*(['"]?)#([^)'"\s]+)\1\s*\)\s*$/iu.exec(stroke ?? '')
    if (strokeReference) {
      const fallback = gradientStrokeFallback(root, strokeReference[2]!)
      if (fallback) element.setAttribute('stroke', fallback)
    }
    const style = element.getAttribute('style')
    if (style) element.setAttribute('style', replaceGradientStrokeDeclarations(style, root))
  }
  root.querySelectorAll('style').forEach((element) => {
    element.textContent = replaceGradientStrokeDeclarations(element.textContent ?? '', root)
  })

  const [width, height] = svgDimensions(root)
  root.setAttribute('width', `${width}px`)
  root.setAttribute('height', `${height}px`)
  if (!svgViewBox(root)) root.setAttribute('viewBox', `0 0 ${width} ${height}`)
  return { markup: new XMLSerializer().serializeToString(root), width, height }
}

function isSvgFile(file: File): boolean {
  return file.type === 'image/svg+xml' || /\.svg$/iu.test(file.name)
}

async function decodeSvg(file: File): Promise<DecodedPdfImage> {
  if (file.size > MAX_SVG_FILE_BYTES) throw new Error(`${file.name}: SVG must be 10 MB or smaller`)
  const sanitized = sanitizeSvgForPdf(await file.text())
  const scale = Math.min(
    SVG_RASTER_SCALE,
    MAX_IMAGE_DIMENSION / Math.max(sanitized.width, sanitized.height),
    Math.sqrt(MAX_IMAGE_PIXELS / (sanitized.width * sanitized.height)),
  )
  const rasterWidth = Math.max(1, Math.round(sanitized.width * scale))
  const rasterHeight = Math.max(1, Math.round(sanitized.height * scale))
  const objectUrl = URL.createObjectURL(
    new Blob([sanitized.markup], { type: 'image/svg+xml;charset=utf-8' }),
  )
  const image = new Image()
  try {
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve()
      image.onerror = () => reject(new Error(`${file.name}: unsupported or damaged SVG`))
      image.src = objectUrl
    })
  } catch (error) {
    URL.revokeObjectURL(objectUrl)
    throw error
  }
  return {
    source: image,
    width: sanitized.width,
    height: sanitized.height,
    rasterWidth,
    rasterHeight,
    close: () => URL.revokeObjectURL(objectUrl),
  }
}

async function prepareVectorSvgPage(file: File): Promise<PdfVectorPage> {
  if (file.size > MAX_SVG_FILE_BYTES) throw new Error(`${file.name}: SVG must be 10 MB or smaller`)
  const sanitized = sanitizeSvgForPdf(await file.text())
  const document = new DOMParser().parseFromString(sanitized.markup, 'image/svg+xml')
  const root = document.documentElement
  const [{ jsPDF }, { svg2pdf }] = await Promise.all([import('jspdf'), import('svg2pdf.js')])
  const pdf = new jsPDF({
    orientation: sanitized.width > sanitized.height ? 'landscape' : 'portrait',
    unit: 'pt',
    format: [sanitized.width, sanitized.height],
    compress: true,
    putOnlyUsedFonts: true,
    precision: 16,
  })
  await svg2pdf(root, pdf, {
    x: 0,
    y: 0,
    width: sanitized.width,
    height: sanitized.height,
    loadExternalStyleSheets: false,
  })
  return { kind: 'vectorPdf', pdf: new Uint8Array(pdf.output('arraybuffer')) }
}

async function decodeImage(file: File): Promise<DecodedPdfImage> {
  if (isSvgFile(file)) return decodeSvg(file)
  try {
    const image = await createImageBitmap(file, { imageOrientation: 'from-image' })
    return {
      source: image,
      width: image.width,
      height: image.height,
      rasterWidth: image.width,
      rasterHeight: image.height,
      close: () => image.close(),
    }
  } catch {
    throw new Error(`${file.name}: unsupported or damaged image`)
  }
}

export function prepareImagesForPdf(
  files: File[],
  colorMode: PdfImageColorMode,
): Promise<PdfImagesToPdfPage[]>
export function prepareImagesForPdf(
  files: File[],
  colorMode: PdfImageColorMode,
  preserveSvgVectors: false,
): Promise<PdfRasterPage[]>
export async function prepareImagesForPdf(
  files: File[],
  colorMode: PdfImageColorMode,
  preserveSvgVectors = true,
): Promise<PdfImagesToPdfPage[]> {
  if (files.length === 0) throw new Error('Choose at least one image')
  if (files.length > MAX_IMAGE_FILES)
    throw new Error(`Choose no more than ${MAX_IMAGE_FILES} images`)
  if (!['color', 'greyscale', 'blackwhite'].includes(colorMode)) {
    throw new Error('Image color mode is invalid')
  }
  if (files.some((file) => file.size > MAX_IMAGE_FILE_BYTES)) {
    throw new Error('Each image must be 50 MB or smaller')
  }
  if (files.reduce((total, file) => total + file.size, 0) > MAX_TOTAL_IMAGE_BYTES) {
    throw new Error('Selected images must total 200 MB or less')
  }

  const pages: PdfImagesToPdfPage[] = []
  for (const file of files) {
    if (preserveSvgVectors && isSvgFile(file) && colorMode === 'color') {
      pages.push(await prepareVectorSvgPage(file))
      continue
    }
    const image = await decodeImage(file)
    try {
      if (
        image.rasterWidth <= 0 ||
        image.rasterHeight <= 0 ||
        image.rasterWidth > MAX_IMAGE_DIMENSION ||
        image.rasterHeight > MAX_IMAGE_DIMENSION ||
        image.rasterWidth * image.rasterHeight > MAX_IMAGE_PIXELS
      ) {
        throw new Error(`${file.name}: image dimensions are too large`)
      }
      const canvas = document.createElement('canvas')
      canvas.width = image.rasterWidth
      canvas.height = image.rasterHeight
      const context = canvas.getContext('2d')
      if (!context) throw new Error('Canvas is unavailable')
      context.drawImage(image.source, 0, 0, canvas.width, canvas.height)
      if (colorMode !== 'color') {
        const pixels = context.getImageData(0, 0, canvas.width, canvas.height)
        convertImagePixelsForPdf(pixels.data, colorMode)
        context.putImageData(pixels, 0, 0)
      }
      pages.push({
        image: await canvasPngBytes(canvas),
        width: image.width,
        height: image.height,
      })
      canvas.width = 0
      canvas.height = 0
    } finally {
      image.close()
    }
  }
  return pages
}
