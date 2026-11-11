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
  onWebAiSettingsChanged,
  onWebAiStream,
  saveWebAiSettings,
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

// The browser port implements the subset of the markdown API it can serve;
// everything else is absent rather than stubbed.
const markdownApi: Partial<MarkdownApi> = {
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
  getDocumentTheme: async () => 'follow' as const,
  onDocumentThemeChanged: () => () => {},
  getMarkdownTheme: async () => 'default' as const,
  setMarkdownTheme: async (theme) => theme,
  getAiSettings: async () => getWebAiSettings(),
  setAiSettings: async (settings) => saveWebAiSettings(settings),
  onAiSettingsChanged: onWebAiSettingsChanged,
  consumeHeadlessExport: async () => null,
  headlessExportDone: () => {},
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  onAiStream: onWebAiStream,
  webSearch,
}

// Same contract as the other web bridges: an unimplemented native member must
// stay callable or one mount-time call unmounts the whole React tree.
const warnedMarkdownMembers = new Set<string>()
const markdownBridge = new Proxy(markdownApi as Record<string, unknown>, {
  get(target, property) {
    if (typeof property !== 'string') return undefined
    const existing = target[property]
    if (existing !== undefined) return existing
    if (!warnedMarkdownMembers.has(property)) {
      warnedMarkdownMembers.add(property)
      console.warn(`[web] markdownApi.${property} is not implemented in the browser bridge`)
    }
    return property.startsWith('on') ? () => () => {} : async () => null
  },
})

window.markdownApi = markdownBridge as unknown as MarkdownApi

if (!window.projectApi) {
  Object.defineProperty(window, 'projectApi', { value: undefined, configurable: true })
}
