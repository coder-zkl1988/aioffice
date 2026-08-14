// eslint-disable-next-line @typescript-eslint/triple-slash-reference
/// <reference path="./gifenc.d.ts" />

import { GIFEncoder, applyPalette, quantize } from 'gifenc'
import { ImageKind, OPS } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy, PDFPageProxy } from 'pdfjs-dist'
import type { PdfExtractedImage, PdfExtractImageFormat } from '@genoffice/pdf-tools'

interface PdfJsImageData {
  width: number
  height: number
  kind?: number
  data?: Uint8Array | Uint8ClampedArray
  bitmap?: CanvasImageSource
}

const MAX_IMAGE_PIXELS = 32_000_000
const MAX_IMAGE_COUNT = 1_000

export function packedPdfImageToRgba(image: PdfJsImageData): Uint8ClampedArray {
  const { width, height, data } = image
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1) {
    throw new Error('Embedded image has invalid dimensions')
  }
  if (width * height > MAX_IMAGE_PIXELS) throw new Error('Embedded image is too large to extract')
  if (!data) throw new Error('Embedded image pixel data is unavailable')

  const pixels = width * height
  const rgba = new Uint8ClampedArray(pixels * 4)
  if (image.kind === ImageKind.RGBA_32BPP) {
    if (data.length < rgba.length) throw new Error('Embedded RGBA image data is incomplete')
    rgba.set(data.subarray(0, rgba.length))
    return rgba
  }
  if (image.kind === ImageKind.RGB_24BPP) {
    if (data.length < pixels * 3) throw new Error('Embedded RGB image data is incomplete')
    for (let pixel = 0; pixel < pixels; pixel++) {
      rgba[pixel * 4] = data[pixel * 3]!
      rgba[pixel * 4 + 1] = data[pixel * 3 + 1]!
      rgba[pixel * 4 + 2] = data[pixel * 3 + 2]!
      rgba[pixel * 4 + 3] = 255
    }
    return rgba
  }
  if (image.kind === ImageKind.GRAYSCALE_1BPP) {
    const bytesPerRow = Math.ceil(width / 8)
    if (data.length < bytesPerRow * height) throw new Error('Embedded bitmap data is incomplete')
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const isWhite = (data[y * bytesPerRow + Math.floor(x / 8)]! >> (7 - (x % 8))) & 1
        const value = isWhite ? 255 : 0
        const offset = (y * width + x) * 4
        rgba[offset] = value
        rgba[offset + 1] = value
        rgba[offset + 2] = value
        rgba[offset + 3] = 255
      }
    }
    return rgba
  }
  throw new Error('Embedded image uses an unsupported pixel format')
}

export function encodeGifRgba(rgba: Uint8ClampedArray, width: number, height: number): Uint8Array {
  const palette = quantize(rgba, 256, {
    format: 'rgba4444',
    oneBitAlpha: true,
  })
  const indexed = applyPalette(rgba, palette, 'rgba4444')
  const transparentIndex = palette.findIndex((color) => color[3] === 0)
  const encoder = GIFEncoder()
  encoder.writeFrame(indexed, width, height, {
    palette,
    ...(transparentIndex >= 0 ? { transparent: true, transparentIndex } : {}),
  })
  encoder.finish()
  return encoder.bytes()
}

export function flattenRgbaOnWhite(rgba: Uint8ClampedArray): Uint8ClampedArray {
  const flattened = new Uint8ClampedArray(rgba)
  for (let offset = 0; offset < flattened.length; offset += 4) {
    const alpha = flattened[offset + 3]!
    if (alpha < 255) {
      flattened[offset] = Math.round((flattened[offset]! * alpha + 255 * (255 - alpha)) / 255)
      flattened[offset + 1] = Math.round(
        (flattened[offset + 1]! * alpha + 255 * (255 - alpha)) / 255,
      )
      flattened[offset + 2] = Math.round(
        (flattened[offset + 2]! * alpha + 255 * (255 - alpha)) / 255,
      )
      flattened[offset + 3] = 255
    }
  }
  return flattened
}

function canvasBytes(
  canvas: HTMLCanvasElement,
  mimeType: 'image/png' | 'image/jpeg',
  quality?: number,
): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    canvas.toBlob(
      (blob) => {
        if (!blob) return reject(new Error('Could not encode extracted image'))
        void blob.arrayBuffer().then((buffer) => resolve(new Uint8Array(buffer)), reject)
      },
      mimeType,
      quality,
    )
  })
}

async function imageToRgba(image: PdfJsImageData): Promise<Uint8ClampedArray> {
  if (!image.bitmap) return packedPdfImageToRgba(image)
  if (image.width * image.height > MAX_IMAGE_PIXELS) {
    throw new Error('Embedded image is too large to extract')
  }
  const canvas = document.createElement('canvas')
  canvas.width = image.width
  canvas.height = image.height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas is unavailable')
  context.drawImage(image.bitmap, 0, 0, image.width, image.height)
  return context.getImageData(0, 0, image.width, image.height).data
}

async function encodeImage(
  image: PdfJsImageData,
  format: PdfExtractImageFormat,
): Promise<{ bytes: Uint8Array; rgba: Uint8ClampedArray }> {
  const rgba = await imageToRgba(image)
  if (format === 'gif') {
    return { bytes: encodeGifRgba(rgba, image.width, image.height), rgba }
  }
  const canvas = document.createElement('canvas')
  canvas.width = image.width
  canvas.height = image.height
  const context = canvas.getContext('2d')
  if (!context) throw new Error('Canvas is unavailable')
  const imageData = context.createImageData(image.width, image.height)
  imageData.data.set(format === 'jpg' ? flattenRgbaOnWhite(rgba) : rgba)
  context.putImageData(imageData, 0, 0)
  return {
    bytes: await canvasBytes(canvas, format === 'png' ? 'image/png' : 'image/jpeg', 0.92),
    rgba,
  }
}

function waitForPdfObject(page: PDFPageProxy, objectId: string): Promise<PdfJsImageData> {
  return new Promise((resolve, reject) => {
    try {
      const objects = objectId.startsWith('g_') ? page.commonObjs : page.objs
      objects.get(objectId, (value: PdfJsImageData) => resolve(value))
    } catch (error) {
      reject(error)
    }
  })
}

function imageFingerprint(rgba: Uint8ClampedArray, width: number, height: number): string {
  let first = 0x811c9dc5
  let second = 0x9e3779b9
  for (const value of rgba) {
    first = Math.imul(first ^ value, 0x01000193)
    second = Math.imul(second ^ value, 0x85ebca6b)
  }
  return `${width}x${height}:${first >>> 0}:${second >>> 0}`
}

async function operatorImage(
  page: PDFPageProxy,
  operator: number,
  args: unknown[],
): Promise<PdfJsImageData | null> {
  if (operator === OPS.paintImageXObject || operator === OPS.paintImageXObjectRepeat) {
    return typeof args[0] === 'string' ? waitForPdfObject(page, args[0]) : null
  }
  if (operator === OPS.paintInlineImageXObject || operator === OPS.paintInlineImageXObjectGroup) {
    return (args[0] as PdfJsImageData | undefined) ?? null
  }
  return null
}

export async function extractEmbeddedPdfImages(
  sourceDocument: PDFDocumentProxy,
  format: PdfExtractImageFormat,
): Promise<PdfExtractedImage[]> {
  if (!['png', 'jpg', 'gif'].includes(format)) throw new Error('Unsupported image format')
  const seenFingerprints = new Set<string>()
  const images: PdfExtractedImage[] = []

  for (let pageIndex = 0; pageIndex < sourceDocument.numPages; pageIndex++) {
    const page = await sourceDocument.getPage(pageIndex + 1)
    const operators = await page.getOperatorList()
    let imageNumber = 1
    for (let index = 0; index < operators.fnArray.length; index++) {
      const image = await operatorImage(
        page,
        operators.fnArray[index]!,
        operators.argsArray[index]!,
      )
      if (!image) continue
      const encoded = await encodeImage(image, format)
      const fingerprint = imageFingerprint(encoded.rgba, image.width, image.height)
      if (seenFingerprints.has(fingerprint)) continue
      seenFingerprints.add(fingerprint)
      if (images.length >= MAX_IMAGE_COUNT)
        throw new Error('PDF contains too many images to extract')
      images.push({ pageNumber: pageIndex + 1, imageNumber: imageNumber++, bytes: encoded.bytes })
    }
  }
  return images
}
