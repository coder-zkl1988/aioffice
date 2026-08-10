import type { DesktopApi, OpenFileResult, PickImageResult } from '../../../docs/src/shared/ipc'
import {
  consumePendingPath,
  getFileHandle,
  getStoredFile,
  listStoredFiles,
  pickBrowserFile,
  putStoredFile,
  queuePendingFile,
  readLanguage,
  readTheme,
  writeBrowserFile,
} from '../lib/files'
import {
  cancelWebAiStream,
  getWebAiSettings,
  onWebAiStream,
  saveWebAiSettings,
  webAiChat,
  webAiStream,
} from '../lib/ai'

async function hash(data: ArrayBuffer): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', data)
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('')
}

async function toOpenResult(
  path: string,
  name: string,
  data: ArrayBuffer,
): Promise<OpenFileResult> {
  return { path, name, data, hash: await hash(data) }
}

async function persistDocx(path: string, name: string, data: ArrayBuffer): Promise<void> {
  await putStoredFile({
    path,
    name,
    kind: 'docx',
    mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    updatedAt: Date.now(),
    data,
  })
}

async function openDocx(): Promise<OpenFileResult | null> {
  const picked = await pickBrowserFile('docx')
  if (!picked) return null
  const data = await picked.file.arrayBuffer()
  await persistDocx(picked.path, picked.file.name, data)
  return toOpenResult(picked.path, picked.file.name, data)
}

async function openDocxPath(path: string): Promise<OpenFileResult | null> {
  const stored = await getStoredFile(path)
  if (!stored || stored.kind !== 'docx' || typeof stored.data === 'string') return null
  return toOpenResult(stored.path, stored.name, stored.data)
}

async function saveAs(
  defaultName: string,
  data: ArrayBuffer,
): Promise<{ ok: boolean; path?: string; error?: string }> {
  try {
    const saved = await writeBrowserFile({
      name: defaultName || 'Untitled.docx',
      extension: '.docx',
      mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      blob: new Blob([data]),
      forcePicker: true,
    })
    await persistDocx(saved.path, saved.name, data)
    return { ok: true, path: saved.path }
  } catch (error) {
    if ((error as DOMException).name === 'AbortError') return { ok: false, error: '已取消保存' }
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

function pickImage(): Promise<PickImageResult | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = 'image/png,image/jpeg,image/gif'
    input.onchange = async () => {
      const file = input.files?.[0]
      if (!file) return resolve(null)
      const bytes = new Uint8Array(await file.arrayBuffer())
      let binary = ''
      for (let offset = 0; offset < bytes.length; offset += 0x8000) {
        binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
      }
      resolve({ base64: btoa(binary), mime: file.type as PickImageResult['mime'], name: file.name })
    }
    input.oncancel = () => resolve(null)
    input.click()
  })
}

const desktop: DesktopApi = {
  getLanguage: async () => readLanguage(),
  onLanguageChanged: () => () => {},
  getTheme: async () => readTheme(),
  onThemeChanged: () => () => {},
  openDocx,
  openDocxPath,
  consumePendingOpenDocx: async () => {
    const path = consumePendingPath()
    return path ? openDocxPath(path) : null
  },
  consumeNewBlankDoc: async () => true,
  onOpenDocx: () => () => {},
  onRenamedDocx: () => () => {},
  saveDocx: async (path, data, auto) => {
    try {
      const stored = await getStoredFile(path)
      const name = stored?.name ?? decodeURIComponent(path.split('/').pop() || 'Untitled.docx')
      const handle = getFileHandle(path)
      if (handle) {
        const writable = await handle.createWritable()
        await writable.write(new Blob([data]))
        await writable.close()
      } else if (!auto) {
        await writeBrowserFile({
          path,
          name,
          extension: '.docx',
          mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          blob: new Blob([data]),
        })
      }
      await persistDocx(path, name, data)
      return { ok: true }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  writeRecoveryCopy: async (path, data) => {
    const stored = await getStoredFile(path)
    await persistDocx(path, stored?.name ?? 'Recovered.docx', data)
    return { ok: true }
  },
  onTeardown: () => () => {},
  saveDocxAs: saveAs,
  saveDocxNew: saveAs,
  getRecentFiles: async () => (await listStoredFiles('docx')).map((file) => file.path),
  pickImage,
  getAiSettings: async () => getWebAiSettings(),
  setAiSettings: async (settings) => saveWebAiSettings(settings),
  print: async () => window.print(),
  exportPdf: async () => {
    window.print()
    return { ok: true, path: 'browser-print-dialog' }
  },
  printPdfBuffer: async () => ({ ok: false, error: 'Web 版请使用浏览器打印并选择“另存为 PDF”' }),
  saveMergedPdf: async () => ({ ok: false, error: 'Web 版暂不支持合并打印片段' }),
  aiChat: webAiChat,
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  aiGskStatus: async () => ({ loggedIn: false }),
  aiGskLogin: async () => {
    window.open('https://www.genspark.ai/', '_blank', 'noopener,noreferrer')
  },
  webSearch: async () => ({ results: [], method: 'error', error: 'Web 版搜索服务尚未配置' }),
  imageSearch: async () => ({ images: [], method: 'error', error: 'Web 版图片搜索尚未配置' }),
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
      return { base64: btoa(binary), mime: blob.type }
    } catch {
      return null
    }
  },
  pickAttachments: async () => null,
  addAttachmentPaths: async () => ({ accepted: [], rejected: ['Web 版暂不支持本地路径附件'] }),
  addPastedImage: async () => ({ accepted: [], rejected: ['Web 版暂不支持 AI 图片附件'] }),
  readAttachment: async () => ({ ok: false, error: 'Web 版暂不支持附件读取' }),
  readAttachmentImage: async () => ({ ok: false, error: 'Web 版暂不支持附件读取' }),
  getPathForFile: (file) => file.name,
  openNewTab: async (openPath) => {
    if (openPath) queuePendingFile(openPath)
    window.open('./docs.html', '_blank', 'noopener')
  },
  listDocsTabs: async () => [{ id: 'web-current', title: document.title, focused: true }],
  focusDocsTab: async () => {},
  onAiStream: onWebAiStream,
  onMenuCommand: () => () => {},
  onCloseCheck: () => () => {},
  reportCloseCheck: () => {},
  onCloseSaveRequest: () => () => {},
  reportCloseSaveResult: () => {},
  reportViewMenuState: () => {},
}

window.desktop = desktop

if (!window.projectApi) {
  Object.defineProperty(window, 'projectApi', { value: undefined, configurable: true })
}
