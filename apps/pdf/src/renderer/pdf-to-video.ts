import type { PdfToVideoOptions, PdfVideoResolution } from '@genoffice/pdf-tools'
import fixWebmDuration from 'fix-webm-duration'
import { AnnotationMode } from 'pdfjs-dist/legacy/build/pdf.mjs'
import type { PDFDocumentProxy } from 'pdfjs-dist'

export interface PdfVideoProgress {
  pageNumber: number
  pageCount: number
  progress: number
}

const VIDEO_DIMENSIONS: Record<PdfVideoResolution, { width: number; height: number }> = {
  '480p': { width: 854, height: 480 },
  '720p': { width: 1280, height: 720 },
  '1080p': { width: 1920, height: 1080 },
}

export function pdfVideoDimensions(resolution: PdfVideoResolution): {
  width: number
  height: number
} {
  const dimensions = VIDEO_DIMENSIONS[resolution]
  if (!dimensions) throw new Error('PDF video resolution is invalid')
  return dimensions
}

export function validatePdfVideoSettings(options: {
  pageCount: number
  pageIndexes: number[]
  secondsPerPage: number
  transitionSeconds: number
  resolution: PdfVideoResolution
}): void {
  if (
    !Number.isInteger(options.secondsPerPage) ||
    options.secondsPerPage < 1 ||
    options.secondsPerPage > 10
  ) {
    throw new Error('Video page duration must be a whole number from 1 to 10 seconds')
  }
  if (
    !Number.isFinite(options.transitionSeconds) ||
    options.transitionSeconds < 0 ||
    options.transitionSeconds > 1
  ) {
    throw new Error('Video transition must be from 0 to 1 second')
  }
  pdfVideoDimensions(options.resolution)
  if (options.pageIndexes.length === 0) throw new Error('Choose at least one page')
  if (
    options.pageIndexes.some(
      (pageIndex) =>
        !Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= options.pageCount,
    )
  ) {
    throw new Error('Selected page is outside the PDF')
  }
  if (
    options.pageIndexes.length > 100 ||
    options.pageIndexes.length * options.secondsPerPage > 300
  ) {
    throw new Error('PDF video is limited to 100 pages or 5 minutes')
  }
}

function videoMimeType(): string {
  if (typeof MediaRecorder === 'undefined') {
    throw new Error('This browser does not support local video encoding')
  }
  for (const mimeType of ['video/webm;codecs=vp9', 'video/webm;codecs=vp8', 'video/webm']) {
    if (MediaRecorder.isTypeSupported(mimeType)) return mimeType
  }
  throw new Error('This browser cannot encode WebM video')
}

function wait(milliseconds: number): Promise<void> {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds))
}

async function renderVideoPage(
  sourceDocument: PDFDocumentProxy,
  pageIndex: number,
  width: number,
  height: number,
  includeAnnotations: boolean,
): Promise<HTMLCanvasElement> {
  const page = await sourceDocument.getPage(pageIndex + 1)
  try {
    const viewport = page.getViewport({ scale: 1 })
    const scale = Math.min(width / viewport.width, height / viewport.height)
    const renderViewport = page.getViewport({ scale })
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(renderViewport.width))
    canvas.height = Math.max(1, Math.round(renderViewport.height))
    const context = canvas.getContext('2d', { alpha: false })
    if (!context) throw new Error('Canvas is unavailable')
    await page.render({
      canvas,
      viewport: renderViewport,
      intent: 'print',
      annotationMode: includeAnnotations ? AnnotationMode.ENABLE_STORAGE : AnnotationMode.DISABLE,
      ...(includeAnnotations
        ? { printAnnotationStorage: sourceDocument.annotationStorage.print }
        : {}),
      background: '#ffffff',
    }).promise
    return canvas
  } finally {
    page.cleanup()
  }
}

function drawVideoFrame(
  context: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  page: HTMLCanvasElement,
  opacity = 1,
): void {
  context.save()
  context.fillStyle = '#17191d'
  context.fillRect(0, 0, canvas.width, canvas.height)
  context.globalAlpha = opacity
  context.drawImage(page, (canvas.width - page.width) / 2, (canvas.height - page.height) / 2)
  context.restore()
}

function requestCanvasFrame(stream: MediaStream): void {
  const track = stream.getVideoTracks()[0] as CanvasCaptureMediaStreamTrack | undefined
  track?.requestFrame?.()
}

async function holdVideoFrame(
  context: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  stream: MediaStream,
  page: HTMLCanvasElement,
  milliseconds: number,
  onProgress: (progress: number) => void,
): Promise<void> {
  const startedAt = performance.now()
  while (true) {
    const elapsed = performance.now() - startedAt
    const progress = Math.min(1, elapsed / Math.max(1, milliseconds))
    drawVideoFrame(context, canvas, page)
    requestCanvasFrame(stream)
    onProgress(progress)
    if (progress >= 1) return
    await wait(Math.min(100, milliseconds - elapsed))
  }
}

async function fadeVideoFrame(
  context: CanvasRenderingContext2D,
  canvas: HTMLCanvasElement,
  stream: MediaStream,
  previous: HTMLCanvasElement,
  current: HTMLCanvasElement,
  milliseconds: number,
  onProgress: (progress: number) => void,
): Promise<void> {
  if (milliseconds <= 0) {
    drawVideoFrame(context, canvas, current)
    requestCanvasFrame(stream)
    onProgress(1)
    return
  }
  const startedAt = performance.now()
  while (true) {
    const elapsed = performance.now() - startedAt
    const progress = Math.min(1, elapsed / milliseconds)
    drawVideoFrame(context, canvas, previous)
    context.save()
    context.globalAlpha = progress
    context.drawImage(
      current,
      (canvas.width - current.width) / 2,
      (canvas.height - current.height) / 2,
    )
    context.restore()
    requestCanvasFrame(stream)
    onProgress(progress)
    if (progress >= 1) return
    await wait(Math.min(33, milliseconds - elapsed))
  }
}

function recorderEvent(recorder: MediaRecorder, eventName: 'pause' | 'resume'): Promise<void> {
  return new Promise((resolve) =>
    recorder.addEventListener(eventName, () => resolve(), { once: true }),
  )
}

export async function renderPdfToWebm(
  sourceDocument: PDFDocumentProxy,
  options: PdfToVideoOptions,
  onProgress: (progress: PdfVideoProgress) => void = () => undefined,
): Promise<Uint8Array> {
  validatePdfVideoSettings({ ...options, pageCount: sourceDocument.numPages })
  const { width, height } = pdfVideoDimensions(options.resolution)
  const output = document.createElement('canvas')
  output.width = width
  output.height = height
  const context = output.getContext('2d', { alpha: false })
  if (!context || typeof output.captureStream !== 'function') {
    throw new Error('Canvas video capture is unavailable')
  }
  const stream = output.captureStream(30)
  const mimeType = videoMimeType()
  const recorder = new MediaRecorder(stream, {
    mimeType,
    videoBitsPerSecond:
      options.resolution === '1080p'
        ? 8_000_000
        : options.resolution === '720p'
          ? 5_000_000
          : 2_500_000,
  })
  const chunks: Blob[] = []
  recorder.addEventListener('dataavailable', (event) => {
    if (event.data.size > 0) chunks.push(event.data)
  })
  const stopped = new Promise<void>((resolve, reject) => {
    recorder.addEventListener('stop', () => resolve(), { once: true })
    recorder.addEventListener('error', () => reject(new Error('Could not encode PDF video')), {
      once: true,
    })
  })
  let previous = await renderVideoPage(
    sourceDocument,
    options.pageIndexes[0]!,
    width,
    height,
    options.includeAnnotations,
  )
  try {
    drawVideoFrame(context, output, previous)
    recorder.start(1000)
    requestCanvasFrame(stream)
    for (let index = 0; index < options.pageIndexes.length; index++) {
      if (index > 0) {
        const paused = recorderEvent(recorder, 'pause')
        recorder.pause()
        await paused
        const current = await renderVideoPage(
          sourceDocument,
          options.pageIndexes[index]!,
          width,
          height,
          options.includeAnnotations,
        )
        const resumed = recorderEvent(recorder, 'resume')
        recorder.resume()
        await resumed
        await fadeVideoFrame(
          context,
          output,
          stream,
          previous,
          current,
          options.transitionSeconds * 1000,
          (progress) =>
            onProgress({
              pageNumber: index + 1,
              pageCount: options.pageIndexes.length,
              progress: (progress * options.transitionSeconds) / options.secondsPerPage,
            }),
        )
        previous.width = 0
        previous.height = 0
        previous = current
      }
      const holdMilliseconds =
        (options.secondsPerPage - (index === 0 ? 0 : options.transitionSeconds)) * 1000
      await holdVideoFrame(context, output, stream, previous, holdMilliseconds, (progress) =>
        onProgress({
          pageNumber: index + 1,
          pageCount: options.pageIndexes.length,
          progress:
            (index === 0 ? 0 : options.transitionSeconds / options.secondsPerPage) +
            progress * (holdMilliseconds / 1000 / options.secondsPerPage),
        }),
      )
    }
    recorder.stop()
    await stopped
    const blob = new Blob(chunks, { type: mimeType })
    if (blob.size === 0) throw new Error('Local video encoder produced an empty file')
    const fixedBlob = await fixWebmDuration(
      blob,
      options.pageIndexes.length * options.secondsPerPage * 1000,
      { logger: false },
    )
    return new Uint8Array(await fixedBlob.arrayBuffer())
  } finally {
    if (recorder.state !== 'inactive') recorder.stop()
    for (const track of stream.getTracks()) track.stop()
    previous.width = 0
    previous.height = 0
    output.width = 0
    output.height = 0
  }
}
