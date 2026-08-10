import type { MarkdownApi } from '../../../markdown/src/shared/ipc'
import {
  consumePendingPath,
  getStoredFile,
  putStoredFile,
  readLanguage,
  readTheme,
  writeBrowserFile,
} from '../lib/files'
import {
  cancelWebAiStream,
  getWebAiSettings,
  onWebAiStream,
  webAiStream,
  webSearch,
} from '../lib/ai'
let currentPath: string | null = null

function base64ToBlob(base64: string, mime: string): Blob {
  const binary = atob(base64)
  const bytes = new Uint8Array(binary.length)
  for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index)
  return new Blob([bytes], { type: mime })
}

async function persistMarkdown(path: string, name: string, text: string): Promise<void> {
  await putStoredFile({
    path,
    name,
    kind: 'markdown',
    mime: 'text/markdown',
    updatedAt: Date.now(),
    data: text,
  })
}

const markdownApi: MarkdownApi = {
  consumePending: async () => {
    currentPath = consumePendingPath()
    return currentPath
  },
  readFile: async (path) => {
    const file = await getStoredFile(path)
    if (!file || file.kind !== 'markdown' || typeof file.data !== 'string') {
      throw new Error('找不到浏览器中的 Markdown 文件')
    }
    return file.data
  },
  save: async ({ text, mode, suggestedName }) => {
    try {
      const existing = currentPath ? await getStoredFile(currentPath) : null
      const result = await writeBrowserFile({
        path: currentPath,
        name: suggestedName || existing?.name || 'Untitled.md',
        extension: '.md',
        mime: 'text/markdown',
        blob: new Blob([text], { type: 'text/markdown;charset=utf-8' }),
        forcePicker: mode === 'saveAs' || !currentPath,
      })
      currentPath = result.path
      await persistMarkdown(result.path, result.name, text)
      return { ok: true, path: result.path }
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') return { ok: true, canceled: true }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  setDirty: () => {},
  onSaveRequest: () => () => {},
  sendSaveRequestAck: () => {},
  onCloseSaveRequest: () => () => {},
  sendCloseSaveResult: () => {},
  onFileRenamed: () => () => {},
  pickImage: () =>
    new Promise((resolve) => {
      const input = document.createElement('input')
      input.type = 'file'
      input.accept = 'image/png,image/jpeg,image/gif'
      input.onchange = () => {
        const file = input.files?.[0]
        if (!file) return resolve(null)
        const reader = new FileReader()
        reader.onload = () => resolve(String(reader.result))
        reader.onerror = () => resolve(null)
        reader.readAsDataURL(file)
      }
      input.oncancel = () => resolve(null)
      input.click()
    }),
  saveImage: async ({ base64, ext }) => {
    const mime =
      ext === 'jpg' || ext === 'jpeg' ? 'image/jpeg' : ext === 'gif' ? 'image/gif' : 'image/png'
    return `data:${mime};base64,${base64}`
  },
  readImage: async (src) => {
    const match = /^data:(image\/(?:png|jpeg|gif));base64,(.+)$/i.exec(src)
    if (!match) return null
    return {
      mime: match[1].toLowerCase() as 'image/png' | 'image/jpeg' | 'image/gif',
      base64: match[2],
    }
  },
  onExportRequest: () => () => {},
  exportDocx: async ({ base64, suggestedName }) => {
    try {
      const result = await writeBrowserFile({
        name: suggestedName,
        extension: '.docx',
        mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        blob: base64ToBlob(
          base64,
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        ),
        forcePicker: true,
      })
      return { ok: true, path: result.path }
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') return { ok: true, canceled: true }
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  exportPdf: async ({ html }) => {
    const printWindow = window.open('', '_blank', 'noopener,noreferrer')
    if (!printWindow) return { ok: false, error: '浏览器阻止了打印窗口' }
    printWindow.document.open()
    printWindow.document.write(html)
    printWindow.document.close()
    printWindow.addEventListener('load', () => printWindow.print(), { once: true })
    return { ok: true, path: 'browser-print-dialog' }
  },
  getLanguage: async () => readLanguage(),
  onLanguageChanged: () => () => {},
  getTheme: async () => readTheme(),
  onThemeChanged: () => () => {},
  getAiSettings: async () => getWebAiSettings(),
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  onAiStream: onWebAiStream,
  webSearch,
}

window.markdownApi = markdownApi

if (!window.projectApi) {
  Object.defineProperty(window, 'projectApi', { value: undefined, configurable: true })
}
