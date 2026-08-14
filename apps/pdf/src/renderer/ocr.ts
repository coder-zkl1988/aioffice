import type { PDFDocumentProxy } from 'pdfjs-dist'
import { createWorker, OEM } from 'tesseract.js'
import type {
  PdfOcrLanguage,
  PdfOcrMode,
  PdfOcrPageText,
  PdfOcrTextLayer,
} from '@genoffice/pdf-tools'

export interface PdfOcrProgress {
  stage: 'analyzing' | 'loading' | 'recognizing' | 'finishing'
  pageNumber?: number
  pageCount: number
  progress: number
}

export interface PreparePdfOcrOptions {
  mode: PdfOcrMode
  languages: PdfOcrLanguage[]
  renderDpi: number
  clean: boolean
  sidecar: boolean
}

export interface PreparedPdfOcr {
  textLayers: PdfOcrTextLayer[]
  pageTexts: PdfOcrPageText[]
  skippedPageIndexes: number[]
}

function pageText(items: unknown[]): string {
  return items
    .map((item) => {
      if (typeof item !== 'object' || !item || !('str' in item)) return ''
      const lineBreak = 'hasEOL' in item && item.hasEOL ? '\n' : ' '
      return `${String(item.str ?? '')}${lineBreak}`
    })
    .join('')
    .replace(/[\t ]+/gu, ' ')
    .replace(/ *\n */gu, '\n')
    .replace(/\n{3,}/gu, '\n\n')
    .trim()
}

function pageHasText(text: string): boolean {
  return text.replace(/\s+/gu, '').length >= 3
}

export async function pdfOcrPagePlan(
  pdfDocument: PDFDocumentProxy,
  mode: PdfOcrMode,
): Promise<{
  pageIndexes: number[]
  skippedPageIndexes: number[]
  existingPageTexts: PdfOcrPageText[]
}> {
  const pagesWithText: number[] = []
  const existingPageTexts: PdfOcrPageText[] = []
  for (let pageIndex = 0; pageIndex < pdfDocument.numPages; pageIndex++) {
    const page = await pdfDocument.getPage(pageIndex + 1)
    try {
      const content = await page.getTextContent()
      const text = pageText(content.items)
      if (pageHasText(text)) {
        pagesWithText.push(pageIndex)
        existingPageTexts.push({ pageIndex, text, source: 'existing' })
      }
    } finally {
      page.cleanup()
    }
  }
  if (mode === 'strict' && pagesWithText.length > 0) {
    throw new Error(`OCR stopped because page ${pagesWithText[0]! + 1} already contains text`)
  }
  if (mode === 'force') {
    return {
      pageIndexes: Array.from({ length: pdfDocument.numPages }, (_, pageIndex) => pageIndex),
      skippedPageIndexes: [],
      existingPageTexts,
    }
  }
  const textPages = new Set(pagesWithText)
  return {
    pageIndexes: Array.from({ length: pdfDocument.numPages }, (_, pageIndex) => pageIndex).filter(
      (pageIndex) => !textPages.has(pageIndex),
    ),
    skippedPageIndexes: pagesWithText,
    existingPageTexts,
  }
}

export function cleanOcrCanvas(source: HTMLCanvasElement): HTMLCanvasElement {
  const output = document.createElement('canvas')
  output.width = source.width
  output.height = source.height
  const context = output.getContext('2d', { willReadFrequently: true })
  if (!context) throw new Error('Canvas is unavailable for OCR cleanup')
  context.drawImage(source, 0, 0)
  const image = context.getImageData(0, 0, output.width, output.height)
  const histogram = new Uint32Array(256)
  for (let offset = 0; offset < image.data.length; offset += 4) {
    const gray = Math.round(
      image.data[offset]! * 0.299 +
        image.data[offset + 1]! * 0.587 +
        image.data[offset + 2]! * 0.114,
    )
    histogram[gray]++
  }
  const pixelCount = output.width * output.height
  const percentile = (ratio: number): number => {
    const target = pixelCount * ratio
    let count = 0
    for (let value = 0; value < histogram.length; value++) {
      count += histogram[value]!
      if (count >= target) return value
    }
    return 255
  }
  const black = percentile(0.01)
  const white = Math.max(black + 1, percentile(0.99))
  for (let offset = 0; offset < image.data.length; offset += 4) {
    const gray = Math.round(
      image.data[offset]! * 0.299 +
        image.data[offset + 1]! * 0.587 +
        image.data[offset + 2]! * 0.114,
    )
    const normalized = Math.max(0, Math.min(255, ((gray - black) * 255) / (white - black)))
    const enhanced = normalized < 245 ? Math.max(0, normalized * 0.92) : 255
    image.data[offset] = enhanced
    image.data[offset + 1] = enhanced
    image.data[offset + 2] = enhanced
  }
  context.putImageData(image, 0, 0)
  return output
}

export async function preparePdfOcr(
  pdfDocument: PDFDocumentProxy,
  options: PreparePdfOcrOptions,
  onProgress: (progress: PdfOcrProgress) => void = () => {},
): Promise<PreparedPdfOcr> {
  onProgress({ stage: 'analyzing', pageCount: pdfDocument.numPages, progress: 0 })
  const plan = await pdfOcrPagePlan(pdfDocument, options.mode)
  if (plan.pageIndexes.length === 0) throw new Error('All pages already contain searchable text')

  let currentPageNumber = plan.pageIndexes[0]! + 1
  onProgress({
    stage: 'loading',
    pageNumber: currentPageNumber,
    pageCount: pdfDocument.numPages,
    progress: 0,
  })
  const assetBase = new URL('ocr/', window.location.href).href
  const workerWrapperUrl = URL.createObjectURL(
    new Blob(
      [
        `const originalError = console.error.bind(console);\n`,
        `console.error = (...args) => {\n`,
        `  const message = args.map(String).join(' ');\n`,
        `  if (!message.startsWith('Warning: Parameter not found: ')) originalError(...args);\n`,
        `};\n`,
        `importScripts(${JSON.stringify(`${assetBase}worker.min.js`)});\n`,
      ],
      { type: 'text/javascript' },
    ),
  )
  let worker: Awaited<ReturnType<typeof createWorker>> | undefined
  const textLayers: PdfOcrTextLayer[] = []
  const pageTexts: PdfOcrPageText[] =
    options.sidecar && options.mode === 'skipText' ? [...plan.existingPageTexts] : []
  try {
    worker = await createWorker(options.languages.join('+'), OEM.LSTM_ONLY, {
      workerPath: workerWrapperUrl,
      workerBlobURL: false,
      corePath: `${assetBase}core/`,
      langPath: `${assetBase}lang/`,
      logger: (message) => {
        if (message.status !== 'recognizing text') return
        onProgress({
          stage: 'recognizing',
          pageNumber: currentPageNumber,
          pageCount: pdfDocument.numPages,
          progress: message.progress,
        })
      },
    })
    await worker.setParameters({
      user_defined_dpi: String(options.renderDpi),
      preserve_interword_spaces: '1',
    })
    for (const [position, pageIndex] of plan.pageIndexes.entries()) {
      currentPageNumber = pageIndex + 1
      onProgress({
        stage: 'recognizing',
        pageNumber: currentPageNumber,
        pageCount: pdfDocument.numPages,
        progress: 0,
      })
      const page = await pdfDocument.getPage(currentPageNumber)
      const canvas = window.document.createElement('canvas')
      try {
        const viewport = page.getViewport({ scale: options.renderDpi / 72, rotation: 0 })
        canvas.width = Math.max(1, Math.ceil(viewport.width))
        canvas.height = Math.max(1, Math.ceil(viewport.height))
        const context = canvas.getContext('2d', { alpha: false })
        if (!context) throw new Error('Canvas is unavailable for OCR rendering')
        context.fillStyle = '#fff'
        context.fillRect(0, 0, canvas.width, canvas.height)
        await page.render({ canvas, canvasContext: context, viewport }).promise
        const input = options.clean ? cleanOcrCanvas(canvas) : canvas
        const result = await worker.recognize(
          input,
          {
            pdfTitle: `OCR page ${currentPageNumber}`,
            pdfTextOnly: options.mode !== 'force',
          },
          { text: true, pdf: true, blocks: false },
        )
        if (!result.data.pdf) {
          throw new Error(`No text was recognized on page ${currentPageNumber}`)
        }
        const recognizedText = result.data.text.trim()
        if (!recognizedText) continue
        textLayers.push({
          pageIndex,
          bytes: new Uint8Array(result.data.pdf),
          replacePage: options.mode === 'force',
        })
        if (options.sidecar) pageTexts.push({ pageIndex, text: recognizedText, source: 'ocr' })
        if (input !== canvas) {
          input.width = 0
          input.height = 0
        }
      } finally {
        canvas.width = 0
        canvas.height = 0
        page.cleanup()
      }
      onProgress({
        stage: position === plan.pageIndexes.length - 1 ? 'finishing' : 'recognizing',
        pageNumber: currentPageNumber,
        pageCount: pdfDocument.numPages,
        progress: 1,
      })
    }
  } finally {
    try {
      await worker?.terminate()
    } finally {
      URL.revokeObjectURL(workerWrapperUrl)
    }
  }
  if (textLayers.length === 0) throw new Error('No text was recognized in this PDF')
  return { textLayers, pageTexts, skippedPageIndexes: plan.skippedPageIndexes }
}
