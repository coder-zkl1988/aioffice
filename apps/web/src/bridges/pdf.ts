import type { ImageEditFailure, PdfApi, TextEditFailure } from '../../../pdf/src/shared/ipc'
import type {
  PdfMobileScannerApi,
  PdfMobileScannerFile,
} from '../../../pdf/src/shared/mobile-scanner'
import { runPdfToolBytes } from '@genoffice/pdf-tools'
import {
  cancelWebAiStream,
  getWebAiSettings,
  onWebAiStream,
  webAiStream,
  webFetchImage,
  webGenerateImage,
  webImageSearch,
} from '../lib/ai'
import {
  consumePendingPath,
  generatedBinaryFileKind,
  getStoredFile,
  pickBrowserFile,
  putStoredFile,
  readLanguage,
  readTheme,
  writeBrowserFile,
  writeBrowserFiles,
  WEB_BINARY_FILE_MIMES,
} from '../lib/files'
import {
  applyWebPdfSave,
  applyWebPdfImageEdits,
  applyWebPdfTextEdits,
  extractWebPdf,
  insertWebPdf,
  insertWebPdfBlankPage,
  listWebPdfEditFonts,
  listWebPdfPageImages,
  renderWebPdfPageImage,
  renderWebPdfPagePreview,
  validateWebPdfTextEdits,
} from '../lib/pdf'

let currentPath: string | null = null

function bytesToArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

async function pdfBytes(path: string): Promise<ArrayBuffer> {
  const file = await getStoredFile(path)
  if (!file || file.kind !== 'pdf' || typeof file.data === 'string') {
    throw new Error('找不到浏览器中的 PDF 文件')
  }
  return file.data
}

async function persist(path: string, data: ArrayBuffer): Promise<void> {
  const previous = await getStoredFile(path)
  await putStoredFile({
    path,
    name: previous?.name || 'Document.pdf',
    kind: 'pdf',
    mime: 'application/pdf',
    updatedAt: Date.now(),
    data,
  })
}

const pdfApi: PdfApi = {
  consumePending: async () => {
    currentPath = consumePendingPath()
    return currentPath
  },
  readFile: pdfBytes,
  save: async (request) => {
    try {
      let source = await pdfBytes(request.path)
      let skippedTextEdits: TextEditFailure[] = []
      let skippedImageEdits: ImageEditFailure[] = []
      if (request.textEdits?.length) {
        const textResult = await applyWebPdfTextEdits(source, request.textEdits)
        source = textResult.data
        skippedTextEdits = textResult.skipped
      }
      if (request.imageEdits?.length) {
        const imageResult = await applyWebPdfImageEdits(source, request.imageEdits)
        source = imageResult.data
        skippedImageEdits = imageResult.skipped
      }
      const result = await applyWebPdfSave(source, { ...request, textEdits: [], imageEdits: [] })
      const data = bytesToArrayBuffer(result.bytes)
      if (request.targetPath) {
        await writeBrowserFile({
          name: request.targetPath.split(/[\\/]/).pop() || 'Document.pdf',
          extension: '.pdf',
          mime: 'application/pdf',
          blob: new Blob([data], { type: 'application/pdf' }),
          forcePicker: true,
        })
      } else {
        await persist(request.path, data)
      }
      return {
        ok: true,
        skippedTextEdits: [...skippedTextEdits, ...result.skippedTextEdits],
        skippedImageEdits: [...skippedImageEdits, ...result.skippedImageEdits],
      }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  validateTextEdits: async ({ path, edits }) =>
    validateWebPdfTextEdits(await pdfBytes(path), edits),
  bulkReplaceText: async (request) => {
    try {
      const result = await applyWebPdfTextEdits(await pdfBytes(request.path), request.edits)
      const appliedCount = request.edits.length - result.skipped.length
      if (appliedCount === 0) return { ok: false, error: '没有可应用的文字替换' }
      const data = result.data
      const saved = await writeBrowserFile({
        name: request.suggestedName,
        extension: '.pdf',
        mime: 'application/pdf',
        blob: new Blob([data], { type: 'application/pdf' }),
        forcePicker: true,
      })
      await putStoredFile({
        path: saved.path,
        name: saved.name,
        kind: 'pdf',
        mime: 'application/pdf',
        updatedAt: Date.now(),
        data,
      })
      return { ok: true, savedPath: saved.path, appliedCount, skipped: result.skipped }
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') return { ok: true, canceled: true }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  listEditFonts: listWebPdfEditFonts,
  listPageImages: async (path) => listWebPdfPageImages(await pdfBytes(path)),
  pageImagePng: async ({ path, pageIndex, rect }) =>
    renderWebPdfPageImage(await pdfBytes(path), pageIndex, rect),
  pagePreviewPng: async ({ path, ...request }) =>
    renderWebPdfPagePreview(await pdfBytes(path), request),
  extractPages: async (request) => {
    try {
      const data = bytesToArrayBuffer(
        await extractWebPdf(await pdfBytes(request.path), request.pages),
      )
      await writeBrowserFile({
        name: request.suggestedName,
        extension: '.pdf',
        mime: 'application/pdf',
        blob: new Blob([data], { type: 'application/pdf' }),
        forcePicker: true,
      })
      return { ok: true, savedPath: request.suggestedName }
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') return { ok: true, canceled: true }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  insertPdf: async (request) => {
    try {
      const picked = await pickBrowserFile('pdf')
      if (!picked) return { ok: true, canceled: true }
      const addition = await picked.file.arrayBuffer()
      const result = await insertWebPdf(
        await pdfBytes(request.path),
        addition,
        request.afterPageIndex,
      )
      await persist(request.path, bytesToArrayBuffer(result.bytes))
      return { ok: true, insertedCount: result.count }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  insertBlankPage: async (request) => {
    try {
      const result = await insertWebPdfBlankPage(
        await pdfBytes(request.path),
        request.afterPageIndex,
        {
          count: request.count,
          pageSize: request.pageSize,
          orientation: request.orientation,
        },
      )
      await persist(request.path, bytesToArrayBuffer(result.bytes))
      return { ok: true, insertedCount: result.count }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  runTool: async (request) => {
    try {
      const outputs = await runPdfToolBytes(await pdfBytes(request.path), request.operation)
      if (outputs.length === 0) return { ok: false, error: 'PDF 工具没有生成结果' }
      const sourceStem = request.baseName.replace(/\.pdf$/i, '') || 'Document'
      const prepared = outputs.map((output) => {
        const data = bytesToArrayBuffer(output.bytes)
        const outputName = output.fileName ?? `${sourceStem}${output.suffix}`
        const extension = output.extension ?? /\.[^.]+$/.exec(outputName)?.[0] ?? '.pdf'
        const mime =
          output.mimeType ?? (extension === '.pdf' ? 'application/pdf' : 'application/octet-stream')
        const storedKind = generatedBinaryFileKind(extension, mime)
        return {
          data,
          storedData: storedKind === 'markdown' ? new TextDecoder().decode(output.bytes) : data,
          storedKind,
          file: {
            name: outputName,
            extension,
            mime,
            blob: new Blob([data], { type: mime }),
          },
        }
      })
      const savedFiles = await writeBrowserFiles(prepared.map(({ file }) => file))
      for (let index = 0; index < savedFiles.length; index++) {
        const saved = savedFiles[index]!
        const storedKind = prepared[index]!.storedKind
        if (!storedKind) continue
        await putStoredFile({
          path: saved.path,
          name: saved.name,
          kind: storedKind,
          mime: WEB_BINARY_FILE_MIMES[storedKind],
          updatedAt: Date.now(),
          data: prepared[index]!.storedData,
        })
      }
      return { ok: true, savedPath: savedFiles[0]!.path, count: outputs.length }
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') return { ok: true, canceled: true }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  fetchWebResource: async (request) => {
    const response = await fetch(new URL('./api/pdf/web-resource', document.baseURI), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(request),
    })
    const body = (await response.json()) as {
      url?: string
      contentType?: string
      base64?: string
      error?: string
    }
    if (!response.ok || !body.url || !body.contentType || !body.base64) {
      throw new Error(body.error || `HTTP ${response.status}`)
    }
    return {
      url: body.url,
      contentType: body.contentType,
      bytes: Uint8Array.from(atob(body.base64), (character) => character.charCodeAt(0)),
    }
  },
  requestTimestampToken: async ({ tsaUrl, request }) => {
    const requestBase64 = btoa(String.fromCharCode(...request))
    const response = await fetch(new URL('./api/pdf/timestamp-token', document.baseURI), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tsaUrl, requestBase64 }),
    })
    const body = (await response.json()) as {
      responseBase64?: string
      error?: string
    }
    if (!response.ok || !body.responseBase64) {
      throw new Error(body.error || `HTTP ${response.status}`)
    }
    return Uint8Array.from(atob(body.responseBase64), (character) => character.charCodeAt(0))
  },
  exportImages: async (request) => {
    try {
      for (let index = 0; index < request.images.length; index++) {
        const binary = atob(request.images[index] || '')
        const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0))
        await writeBrowserFile({
          name: `${request.baseName}-p${request.pageNumbers[index] || index + 1}.png`,
          extension: '.png',
          mime: 'image/png',
          blob: new Blob([bytes], { type: 'image/png' }),
        })
      }
      return { ok: true, savedDir: 'browser-downloads', count: request.images.length }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  imageSearch: webImageSearch,
  fetchImage: webFetchImage,
  generateImage: webGenerateImage,
  setDirty: () => {},
  onCloseSaveRequest: () => () => {},
  sendCloseSaveResult: () => {},
  onSaveAsRequest: () => () => {},
  sendSaveAsResult: () => {},
  onSaveAsFlow: () => () => {},
  getLanguage: async () => readLanguage(),
  onLanguageChanged: () => () => {},
  getTheme: async () => readTheme(),
  onThemeChanged: () => () => {},
  getAiSettings: async () => getWebAiSettings(),
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  onAiStream: onWebAiStream,
}

window.pdfApi = pdfApi

async function mobileScannerJson<T>(path: string, body?: unknown): Promise<T> {
  const response = await fetch(new URL(path, document.baseURI), {
    method: 'POST',
    headers: {
      'X-GenOffice-Client': 'pdf',
      ...(body === undefined ? {} : { 'Content-Type': 'application/json' }),
    },
    body: body === undefined ? undefined : JSON.stringify(body),
  })
  const result = (await response.json()) as T & { error?: string }
  if (!response.ok) throw new Error(result.error || `HTTP ${response.status}`)
  return result
}

const mobileScannerApi: PdfMobileScannerApi = {
  createSession: async () => {
    const result = await mobileScannerJson<{
      sessionId: string
      uploadPath: string
      expiresAt: number
    }>('./api/pdf/mobile-scanner/session')
    return {
      sessionId: result.sessionId,
      uploadUrl: new URL(result.uploadPath, window.location.origin).toString(),
      expiresAt: result.expiresAt,
    }
  },
  pollSession: async (sessionId) => {
    const result = await mobileScannerJson<{
      expiresAt: number
      files: Array<Omit<PdfMobileScannerFile, 'bytes'> & { base64: string }>
    }>('./api/pdf/mobile-scanner/poll', { sessionId })
    return {
      expiresAt: result.expiresAt,
      files: result.files.map(({ base64, ...file }) => ({
        ...file,
        bytes: Uint8Array.from(atob(base64), (character) => character.charCodeAt(0)),
      })),
    }
  },
  closeSession: async (sessionId) => {
    await mobileScannerJson('./api/pdf/mobile-scanner/close', { sessionId })
  },
}

window.pdfMobileScanner = mobileScannerApi
