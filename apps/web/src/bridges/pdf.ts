import type { PdfApi } from '../../../pdf/src/shared/ipc'
import { cancelWebAiStream, getWebAiSettings, onWebAiStream, webAiStream } from '../lib/ai'
import {
  consumePendingPath,
  getStoredFile,
  pickBrowserFile,
  putStoredFile,
  readLanguage,
  readTheme,
  writeBrowserFile,
} from '../lib/files'
import { applyWebPdfSave, extractWebPdf, insertWebPdf } from '../lib/pdf'

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
      const source = await pdfBytes(request.path)
      const result = await applyWebPdfSave(source, request)
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
        skippedTextEdits: result.skippedTextEdits,
        skippedImageEdits: result.skippedImageEdits,
      }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  validateTextEdits: async ({ edits }) =>
    edits.map(() => ({ reason: 'Web 版暂不支持修改 PDF 内容流文本' })),
  listEditFonts: async () => [],
  listPageImages: async () => [],
  pageImagePng: async () => null,
  pagePreviewPng: async () => null,
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
  imageSearch: async () => ({ images: [], method: 'error', error: 'Web 图片搜索尚未配置' }),
  fetchImage: async (url) => {
    try {
      const response = await fetch(url)
      if (!response.ok) return null
      const blob = await response.blob()
      const bytes = new Uint8Array(await blob.arrayBuffer())
      let binary = ''
      for (let offset = 0; offset < bytes.length; offset += 0x8000) {
        binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
      }
      return { base64: btoa(binary), mime: blob.type || 'image/png' }
    } catch {
      return null
    }
  },
  generateImage: async () => ({ error: 'Web 图片生成尚未配置' }),
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
