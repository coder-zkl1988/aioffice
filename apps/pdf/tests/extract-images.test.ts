import { ImageKind } from 'pdfjs-dist/legacy/build/pdf.mjs'
import { describe, expect, it } from 'vitest'
import {
  encodeGifRgba,
  flattenRgbaOnWhite,
  packedPdfImageToRgba,
} from '../src/renderer/extract-images'

describe('embedded PDF image decoding', () => {
  it('expands RGB pixels to RGBA', () => {
    expect(
      Array.from(
        packedPdfImageToRgba({
          width: 2,
          height: 1,
          kind: ImageKind.RGB_24BPP,
          data: new Uint8Array([255, 0, 0, 0, 128, 255]),
        }),
      ),
    ).toEqual([255, 0, 0, 255, 0, 128, 255, 255])
  })

  it('preserves RGBA pixels and alpha', () => {
    expect(
      Array.from(
        packedPdfImageToRgba({
          width: 1,
          height: 1,
          kind: ImageKind.RGBA_32BPP,
          data: new Uint8Array([10, 20, 30, 40]),
        }),
      ),
    ).toEqual([10, 20, 30, 40])
  })

  it('expands one-bit grayscale rows with PDF.js bit order', () => {
    expect(
      Array.from(
        packedPdfImageToRgba({
          width: 3,
          height: 1,
          kind: ImageKind.GRAYSCALE_1BPP,
          data: new Uint8Array([0b1010_0000]),
        }),
      ),
    ).toEqual([255, 255, 255, 255, 0, 0, 0, 255, 255, 255, 255, 255])
  })

  it('rejects incomplete and unsupported image data', () => {
    expect(() =>
      packedPdfImageToRgba({
        width: 2,
        height: 1,
        kind: ImageKind.RGB_24BPP,
        data: new Uint8Array(3),
      }),
    ).toThrow('Embedded RGB image data is incomplete')
    expect(() =>
      packedPdfImageToRgba({ width: 1, height: 1, kind: 99, data: new Uint8Array(4) }),
    ).toThrow('unsupported pixel format')
  })
})

describe('JPEG alpha flattening', () => {
  it('composites transparent and translucent pixels onto white', () => {
    const source = new Uint8ClampedArray([20, 40, 60, 0, 255, 0, 0, 128])
    expect(Array.from(flattenRgbaOnWhite(source))).toEqual([255, 255, 255, 255, 255, 127, 127, 255])
    expect(Array.from(source)).toEqual([20, 40, 60, 0, 255, 0, 0, 128])
  })
})

describe('GIF extraction encoding', () => {
  it('encodes a valid single-frame GIF', () => {
    const bytes = encodeGifRgba(new Uint8ClampedArray([255, 0, 0, 255]), 1, 1)
    expect(new TextDecoder().decode(bytes.subarray(0, 6))).toBe('GIF89a')
    expect(bytes.at(-1)).toBe(0x3b)
  })
})
