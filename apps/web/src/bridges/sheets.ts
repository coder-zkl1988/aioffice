import type {
  DesktopApi,
  WorkbookFile,
  WorkbookSaveRequest,
  WorkbookSaveResult,
} from '../../../sheets/src/shared/desktop-api'
import {
  cancelWebAiStream,
  getWebAiSettings,
  onWebAiStream,
  saveWebAiSettings,
  webAiChat,
  webAiStream,
} from '../lib/ai'
import {
  consumePendingPath,
  getFileHandle,
  getStoredFile,
  makeWebPath,
  pickBrowserFile,
  putStoredFile,
  readLanguage,
  readTheme,
  rememberFileHandle,
  writeBrowserFile,
} from '../lib/files'

const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
let pendingPath = consumePendingPath()
const sessionFiles = new Map<string, { path: string; name: string }>()

function base64FromBytes(data: ArrayBuffer): string {
  const bytes = new Uint8Array(data)
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function bytesFromBase64(value: string): ArrayBuffer {
  const binary = atob(value)
  return Uint8Array.from(binary, (character) => character.charCodeAt(0)).buffer
}

async function sheetsApi<T>(action: string, body?: unknown): Promise<T> {
  const response = await fetch(new URL(`./api/sheets/${action}`, document.baseURI), {
    method: body === undefined ? 'GET' : 'POST',
    headers: body === undefined ? undefined : { 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body),
  })
  if (!response.ok) {
    const result = (await response.json().catch(() => ({}))) as { error?: string }
    throw new Error(result.error || `Sheets 服务返回 HTTP ${response.status}`)
  }
  return (await response.json()) as T
}

async function persistWorkbook(path: string, name: string, data: ArrayBuffer): Promise<void> {
  await putStoredFile({
    path,
    name,
    kind: 'xlsx',
    mime: XLSX_MIME,
    updatedAt: Date.now(),
    data,
  })
}

async function openStoredWorkbook(
  path: string,
  name: string,
  data: ArrayBuffer,
): Promise<WorkbookFile> {
  const file = await sheetsApi<WorkbookFile>('open', {
    name,
    webPath: path,
    xlsxBase64: base64FromBytes(data),
  })
  sessionFiles.set(file.sessionId, { path, name })
  return { ...file, path, name }
}

async function selectWorkbook(): Promise<WorkbookFile | null> {
  if (pendingPath) {
    const path = pendingPath
    pendingPath = null
    const stored = await getStoredFile(path)
    if (stored?.kind === 'xlsx' && typeof stored.data !== 'string') {
      return openStoredWorkbook(stored.path, stored.name, stored.data)
    }
  }
  const picked = await pickBrowserFile('xlsx')
  if (!picked) return null
  const data = await picked.file.arrayBuffer()
  await persistWorkbook(picked.path, picked.file.name, data)
  return openStoredWorkbook(picked.path, picked.file.name, data)
}

async function chooseSaveTarget(
  request: WorkbookSaveRequest,
  current: { path: string; name: string },
): Promise<{ path: string; name: string } | null> {
  const currentHandle = getFileHandle(current.path)
  if (request.mode !== 'save-as' && currentHandle) return current
  if (!window.showSaveFilePicker) {
    return request.mode === 'save-as'
      ? { path: makeWebPath(current.name), name: current.name }
      : current
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: current.name,
      types: [{ description: 'Excel 工作簿', accept: { [XLSX_MIME]: ['.xlsx'] } }],
    })
    const path = makeWebPath(handle.name)
    rememberFileHandle(path, handle)
    return { path, name: handle.name }
  } catch (error) {
    if ((error as DOMException).name === 'AbortError') return null
    throw error
  }
}

async function saveWorkbookEdits(request: WorkbookSaveRequest): Promise<WorkbookSaveResult> {
  const current = sessionFiles.get(request.sessionId)
  if (!current) throw new Error('工作簿会话不存在或已过期')
  const target = await chooseSaveTarget(request, current)
  if (!target) return { canceled: true }
  const result = await sheetsApi<{
    canceled: false
    file: WorkbookFile
    touchedEntries: string[]
    xlsxBase64: string
  }>('save', { request, name: target.name, webPath: target.path })
  const data = bytesFromBase64(result.xlsxBase64)
  const handle = getFileHandle(target.path)
  if (handle) {
    const writable = await handle.createWritable()
    await writable.write(new Blob([data], { type: XLSX_MIME }))
    await writable.close()
  } else {
    await writeBrowserFile({
      path: target.path,
      name: target.name,
      extension: '.xlsx',
      mime: XLSX_MIME,
      blob: new Blob([data], { type: XLSX_MIME }),
    })
  }
  await persistWorkbook(target.path, target.name, data)
  sessionFiles.delete(request.sessionId)
  sessionFiles.set(result.file.sessionId, target)
  return {
    canceled: false,
    file: { ...result.file, path: target.path, name: target.name },
    touchedEntries: result.touchedEntries,
  }
}

function menuActions(callback: Parameters<DesktopApi['onMenuAction']>[0]): () => void {
  const listener = (event: KeyboardEvent) => {
    if (!(event.metaKey || event.ctrlKey)) return
    const key = event.key.toLowerCase()
    const action =
      key === 'o'
        ? 'open'
        : key === 's' && event.shiftKey
          ? 'save-as'
          : key === 's'
            ? 'save'
            : key === 'z' && event.shiftKey
              ? 'redo'
              : key === 'z'
                ? 'undo'
                : null
    if (!action) return
    event.preventDefault()
    callback(action)
  }
  window.addEventListener('keydown', listener)
  return () => window.removeEventListener('keydown', listener)
}

const desktopApi: DesktopApi = {
  getLanguage: async () => readLanguage(),
  onLanguageChanged: () => () => {},
  getTheme: async () => readTheme(),
  onThemeChanged: () => () => {},
  selectWorkbook,
  readWorkbookRange: (request) => sheetsApi('range', request),
  readWorkbookFormulas: (request) => sheetsApi('formulas', request),
  recalcWorkbook: (request) => sheetsApi('recalc', request),
  readWorkbookMedia: (request) => sheetsApi('media', request),
  readPivotDefinition: (request) => sheetsApi('pivot', request),
  readLocalImage: async () => {
    throw new Error('Web 版不支持通过系统绝对路径读取图片，请使用“插入图片”选择本地文件')
  },
  captureScreenSources: async () => ({ status: 'denied', sources: [] }),
  captureScreenSource: async () => null,
  saveWorkbookEdits,
  writeWorkbookRecovery: (request) => sheetsApi('recovery', request),
  autoRenameWorkbook: async () => ({ renamed: false }),
  exportPdf: async () => {
    window.print()
    return { canceled: false, path: 'browser-print-dialog.pdf' }
  },
  closeWorkbook: async (sessionId) => {
    sessionFiles.delete(sessionId)
    await sheetsApi('close', { sessionId })
  },
  openExternal: async (url) => {
    const parsed = new URL(url)
    if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
      throw new Error('仅支持打开 HTTP(S) 链接')
    }
    window.open(parsed.href, '_blank', 'noopener,noreferrer')
  },
  onMenuAction: menuActions,
  onWorkbookRenamed: () => () => {},
  notifyPendingEdits: () => {},
  onCloseSaveRequest: () => () => {},
  reportCloseSaveResult: () => {},
  consumeNewBlankWorkbook: async () => false,
  hasQueuedWorkbook: async () => pendingPath !== null,
  getAiSettings: async () => getWebAiSettings(),
  setAiSettings: async (settings) => saveWebAiSettings(settings),
  aiChat: webAiChat,
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  aiGskStatus: async () => ({ loggedIn: false }),
  aiGskLogin: async () => {
    window.open('https://www.genspark.ai/', '_blank', 'noopener,noreferrer')
  },
  webSearch: async () => ({ results: [], method: 'error' }),
  onAiStream: onWebAiStream,
  pickAttachments: async () => null,
  addAttachmentPaths: async () => ({ accepted: [], rejected: ['Web 版不支持系统文件路径附件'] }),
  addPastedImage: async () => ({ accepted: [], rejected: ['Web 版暂不支持粘贴图片附件'] }),
  readAttachment: async () => ({ ok: false, error: 'Web 版暂不支持附件读取' }),
  readAttachmentImage: async () => ({ ok: false, error: 'Web 版暂不支持附件读取' }),
  getPathForFile: (file) => file.name,
}

window.desktopApi = desktopApi

window.addEventListener('pagehide', () => {
  for (const sessionId of sessionFiles.keys()) {
    navigator.sendBeacon(
      new URL('./api/sheets/close', document.baseURI),
      JSON.stringify({ sessionId }),
    )
  }
  sessionFiles.clear()
})

if (!window.projectApi) {
  Object.defineProperty(window, 'projectApi', { value: undefined, configurable: true })
}
