import { describe, expect, it, vi } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { PDF_AUTO_SPLIT_QR_CONTENT, pdfAutoSplitDividerQrModules } from '@genoffice/pdf-tools'
import {
  decodePdfAutoSplitCanvas,
  decodePdfAutoSplitQr,
  detectPdfAutoSplitDividerPages,
  isPdfAutoSplitDividerData,
} from '../src/renderer/auto-split'

function qrPixels(modules: boolean[][], cellSize = 4, quietZone = 4) {
  const width = (modules.length + quietZone * 2) * cellSize
  const data = new Uint8ClampedArray(width * width * 4).fill(255)
  for (let row = 0; row < modules.length; row++) {
    for (let column = 0; column < modules.length; column++) {
      if (!modules[row]![column]) continue
      for (let y = 0; y < cellSize; y++) {
        for (let x = 0; x < cellSize; x++) {
          const pixelX = (column + quietZone) * cellSize + x
          const pixelY = (row + quietZone) * cellSize + y
          const offset = (pixelY * width + pixelX) * 4
          data[offset] = 0
          data[offset + 1] = 0
          data[offset + 2] = 0
          data[offset + 3] = 255
        }
      }
    }
  }
  return { data, width }
}

describe('PDF QR auto split detection', () => {
  it('recognizes all compatible Stirling divider payloads', () => {
    expect(isPdfAutoSplitDividerData('https://github.com/Stirling-Tools/Stirling-PDF')).toBe(true)
    expect(isPdfAutoSplitDividerData('https://github.com/Frooodle/Stirling-PDF')).toBe(true)
    expect(isPdfAutoSplitDividerData('https://stirlingpdf.com')).toBe(true)
    expect(isPdfAutoSplitDividerData('https://example.com')).toBe(false)
  })

  it('decodes the locally generated divider QR matrix', () => {
    const pixels = qrPixels(pdfAutoSplitDividerQrModules())
    expect(decodePdfAutoSplitQr(pixels.data, pixels.width, pixels.width)).toBe(
      PDF_AUTO_SPLIT_QR_CONTENT,
    )
  })

  it('retries QR decoding at multiple scales for high-resolution pages', () => {
    const canvas = document.createElement('canvas')
    canvas.width = 1200
    canvas.height = 1700
    const context = {
      drawImage: vi.fn(),
      getImageData: () => ({ data: new Uint8ClampedArray(canvas.width * canvas.height * 4) }),
    }
    const getContext = vi
      .spyOn(HTMLCanvasElement.prototype, 'getContext')
      .mockReturnValue(context as unknown as CanvasRenderingContext2D)
    const decode = vi
      .fn<(_: Uint8ClampedArray, width: number, height: number) => string | null>()
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(PDF_AUTO_SPLIT_QR_CONTENT)
    expect(decodePdfAutoSplitCanvas(canvas, decode)).toBe(PDF_AUTO_SPLIT_QR_CONTENT)
    expect(decode.mock.calls.map((call) => call.slice(1))).toEqual([
      [1200, 1700],
      [900, 1275],
    ])
    getContext.mockRestore()
  })

  it('scans every page, records compatible dividers, and cleans resources', async () => {
    const cleanup = vi.fn()
    const render = vi.fn(() => ({ promise: Promise.resolve() }))
    const getContext = vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({
      drawImage: vi.fn(),
      getImageData: () => ({ data: new Uint8ClampedArray(100 * 100 * 4) }),
    } as unknown as CanvasRenderingContext2D)
    const document = {
      numPages: 3,
      getPage: vi.fn(async () => ({
        getViewport: ({ scale }: { scale: number }) => ({
          width: 100 * scale,
          height: 100 * scale,
        }),
        render,
        cleanup,
      })),
    } as unknown as PDFDocumentProxy
    const decode = vi
      .fn<(_: Uint8ClampedArray, __: number, ___: number) => string | null>()
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(null)
      .mockReturnValueOnce(PDF_AUTO_SPLIT_QR_CONTENT)
      .mockReturnValueOnce('https://example.com')
    await expect(detectPdfAutoSplitDividerPages(document, 72, decode)).resolves.toEqual([1])
    expect(document.getPage).toHaveBeenCalledTimes(3)
    expect(render).toHaveBeenCalledTimes(3)
    expect(cleanup).toHaveBeenCalledTimes(3)
    getContext.mockRestore()
  })

  it('cleans page resources when rendering fails', async () => {
    const cleanup = vi.fn()
    const getContext = vi
      .spyOn(HTMLCanvasElement.prototype, 'getContext')
      .mockReturnValue({} as CanvasRenderingContext2D)
    const document = {
      numPages: 1,
      getPage: async () => ({
        getViewport: () => ({ width: 100, height: 100 }),
        render: () => ({ promise: Promise.reject(new Error('render failure')) }),
        cleanup,
      }),
    } as unknown as PDFDocumentProxy
    await expect(detectPdfAutoSplitDividerPages(document)).rejects.toThrow('render failure')
    expect(cleanup).toHaveBeenCalledOnce()
    getContext.mockRestore()
  })
})
