// eslint-disable-next-line @typescript-eslint/triple-slash-reference
/// <reference path="./fontkit.d.ts" />

import * as modernFontkit from 'fontkit'
import type { PDFDocument } from 'pdf-lib'

type PdfLibFontkit = Parameters<PDFDocument['registerFontkit']>[0]
type StreamEvent = 'data' | 'end' | 'error'
type StreamHandlers = {
  data?: (bytes: Uint8Array) => void
  end?: () => void
  error?: (error: unknown) => void
}

interface PdfLibSubsetStream {
  on(event: StreamEvent, handler: (value?: unknown) => void): PdfLibSubsetStream
}

interface CompatibleSubset extends modernFontkit.FontSubset {
  encodeStream?: () => PdfLibSubsetStream
}

interface CompatibleFont extends modernFontkit.Font {
  createSubset(): CompatibleSubset
}

function subsetStream(subset: CompatibleSubset): PdfLibSubsetStream {
  const handlers: StreamHandlers = {}
  let scheduled = false
  const stream: PdfLibSubsetStream = {
    on(event, handler) {
      handlers[event] = handler as never
      if (!scheduled) {
        scheduled = true
        queueMicrotask(() => {
          try {
            handlers.data?.(subset.encode())
            handlers.end?.()
          } catch (error) {
            handlers.error?.(error)
          }
        })
      }
      return stream
    },
  }
  return stream
}

export const pdfLibFontkit: PdfLibFontkit = {
  create(data, postscriptName) {
    const font = modernFontkit.create(data, postscriptName) as CompatibleFont
    const createSubset = font.createSubset.bind(font)
    font.createSubset = () => {
      const subset = createSubset()
      // pdf-lib expects the streaming API removed in fontkit 2; encoding is still synchronous.
      subset.encodeStream = () => subsetStream(subset)
      return subset
    }
    return font as never
  },
}
