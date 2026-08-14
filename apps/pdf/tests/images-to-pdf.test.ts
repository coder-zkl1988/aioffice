import { describe, expect, it } from 'vitest'
import { PDFDict, PDFDocument, PDFName, PDFRawStream } from 'pdf-lib'
import {
  convertImagePixelsForPdf,
  prepareImagesForPdf,
  sanitizeSvgForPdf,
} from '../src/renderer/images-to-pdf'

describe('convertImagePixelsForPdf', () => {
  it('preserves color pixels and alpha', () => {
    const pixels = new Uint8ClampedArray([10, 20, 30, 77])
    expect(convertImagePixelsForPdf(pixels, 'color')).toBe(pixels)
    expect(Array.from(pixels)).toEqual([10, 20, 30, 77])
  })

  it('converts pixels to luminance grayscale while preserving alpha', () => {
    const pixels = new Uint8ClampedArray([255, 0, 0, 99, 0, 255, 0, 255])
    expect(Array.from(convertImagePixelsForPdf(pixels, 'greyscale'))).toEqual([
      76, 76, 76, 99, 150, 150, 150, 255,
    ])
  })

  it('applies a stable black-and-white threshold', () => {
    const pixels = new Uint8ClampedArray([127, 127, 127, 255, 128, 128, 128, 123])
    expect(Array.from(convertImagePixelsForPdf(pixels, 'blackwhite'))).toEqual([
      0, 0, 0, 255, 255, 255, 255, 123,
    ])
  })
})

describe('SVG preparation for PDF', () => {
  it('converts color SVG files to vector PDF pages', async () => {
    const [prepared] = await prepareImagesForPdf(
      [
        new File(
          [
            '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360"><rect x="20" y="30" width="180" height="90" fill="#1677ff"/><path d="M20 180 L300 210 L500 120" fill="none" stroke="#111" stroke-width="8"/></svg>',
          ],
          'vector.svg',
          { type: 'image/svg+xml' },
        ),
      ],
      'color',
    )

    expect(prepared).toMatchObject({ kind: 'vectorPdf' })
    if (!prepared || !('kind' in prepared)) throw new Error('Expected a vector PDF page')
    const document = await PDFDocument.load(prepared.pdf)
    expect(document.getPage(0).getSize()).toEqual({ width: 640, height: 360 })
    const resources = document.getPage(0).node.lookupMaybe(PDFName.of('Resources'), PDFDict)
    const xObjects = resources?.lookupMaybe(PDFName.of('XObject'), PDFDict)
    const subtypes =
      xObjects?.keys().map((name) => {
        const object = document.context.lookup(xObjects.get(name))
        return object instanceof PDFRawStream
          ? object.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
          : undefined
      }) ?? []
    expect(subtypes).not.toContain('Image')
  })

  it('uses viewBox dimensions and strips executable or external content', () => {
    const result = sanitizeSvgForPdf(`
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 640 360" onclick="alert(1)">
        <script>alert(1)</script>
        <style>.safe { fill: #123456 } .unsafe { fill: url(https://example.com/a.svg) }</style>
        <image href="https://example.com/tracker.png" width="10" height="10" />
        <animate href="#shape" attributeName="href" to="https://example.com/late.svg" />
        <use href="#shape" />
        <rect id="shape" class="safe" width="640" height="360" />
      </svg>
    `)
    expect(result.width).toBe(640)
    expect(result.height).toBe(360)
    expect(result.markup).not.toMatch(/script|onclick|animate|https:\/\/example\.com/iu)
    expect(result.markup).toContain('href="#shape"')
  })

  it('keeps gradient fills and gives unsupported gradient strokes a vector-safe color', () => {
    const result = sanitizeSvgForPdf(`
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 100">
        <defs>
          <linearGradient id="accent"><stop stop-color="#0f766e"/><stop style="stop-color: #2563eb"/></linearGradient>
        </defs>
        <style>.styled { stroke: url(#accent); fill: none }</style>
        <rect width="80" height="80" fill="url(#accent)"/>
        <path d="M0 90 L200 10" stroke="url(#accent)"/>
        <path class="styled" style="stroke: url('#accent')" d="M0 10 L200 90"/>
      </svg>
    `)
    expect(result.markup).toContain('fill="url(#accent)"')
    expect(result.markup).not.toMatch(/stroke\s*[:=]\s*["']?url\(/iu)
    expect(result.markup.match(/#2563eb/giu)?.length).toBeGreaterThanOrEqual(3)
  })

  it('converts physical dimensions to CSS pixels and creates a viewBox', () => {
    const result = sanitizeSvgForPdf(
      '<svg xmlns="http://www.w3.org/2000/svg" width="25.4mm" height="1in"><rect width="100%" height="100%"/></svg>',
    )
    expect(result.width).toBeCloseTo(96)
    expect(result.height).toBeCloseTo(96)
    expect(result.markup).toContain('viewBox="0 0 96 96"')
  })

  it('rejects malformed, active-doctype, and oversized SVG files', () => {
    expect(() => sanitizeSvgForPdf('<svg><path></svg>')).toThrow(/invalid/i)
    expect(() =>
      sanitizeSvgForPdf('<!DOCTYPE svg><svg xmlns="http://www.w3.org/2000/svg"/>'),
    ).toThrow(/invalid/i)
    expect(() =>
      sanitizeSvgForPdf(
        '<svg xmlns="http://www.w3.org/2000/svg" width="20000" height="100"></svg>',
      ),
    ).toThrow(/large/i)
  })
})
