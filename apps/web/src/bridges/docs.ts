import html2canvas from 'html2canvas'
import { PDFDocument } from 'pdf-lib'
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
  webFetchImage,
  webImageSearch,
  webSearch,
} from '../lib/ai'
import { createWebAttachments } from '../lib/attachments'

const attachments = createWebAttachments()

function base64FromBytes(data: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < data.length; offset += 0x8000) {
    binary += String.fromCharCode(...data.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function bytesFromBase64(value: string): Uint8Array {
  const binary = atob(value)
  return Uint8Array.from(binary, (character) => character.charCodeAt(0))
}

async function renderVisiblePreviewPages(
  pageWidthTwips: number,
  pageHeightTwips: number,
): Promise<string> {
  const pages = [...document.querySelectorAll<HTMLElement>('.pv-page:not(.pv-print-skip)')]
  if (pages.length === 0) throw new Error('请先打开分页预览后再导出混合纸张 PDF')

  const output = await PDFDocument.create()
  const widthPoints = pageWidthTwips / 20
  const heightPoints = pageHeightTwips / 20
  const scale = Math.min(2, Math.max(1, window.devicePixelRatio || 1))
  for (const element of pages) {
    const canvas = await html2canvas(element, {
      backgroundColor: '#ffffff',
      logging: false,
      scale,
      useCORS: true,
    })
    const image = await output.embedPng(canvas.toDataURL('image/png'))
    const page = output.addPage([widthPoints, heightPoints])
    page.drawImage(image, { x: 0, y: 0, width: widthPoints, height: heightPoints })
  }
  return base64FromBytes(await output.save())
}

async function saveMergedPdf(
  defaultName: string,
  base64Parts: string[],
  outPath?: string,
): Promise<{ ok: boolean; path?: string; error?: string }> {
  try {
    const output = await PDFDocument.create()
    for (const value of base64Parts) {
      const part = await PDFDocument.load(bytesFromBase64(value))
      const pages = await output.copyPages(part, part.getPageIndices())
      for (const page of pages) output.addPage(page)
    }
    const bytes = await output.save()
    const saved = await writeBrowserFile({
      path: outPath,
      name: defaultName.replace(/\.docx$/i, ''),
      extension: '.pdf',
      mime: 'application/pdf',
      blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer], { type: 'application/pdf' }),
      forcePicker: !outPath,
    })
    return { ok: true, path: saved.path }
  } catch (error) {
    if ((error as DOMException).name === 'AbortError') return { ok: false }
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

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
  printPdfBuffer: async (pageWidthTwips, pageHeightTwips) => {
    try {
      return {
        ok: true,
        base64: await renderVisiblePreviewPages(pageWidthTwips, pageHeightTwips),
      }
    } catch (error) {
      return { ok: false, error: error instanceof Error ? error.message : String(error) }
    }
  },
  saveMergedPdf,
  aiChat: webAiChat,
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  aiGskStatus: async () => ({ loggedIn: false }),
  aiGskLogin: async () => {
    window.open('https://www.genspark.ai/', '_blank', 'noopener,noreferrer')
  },
  webSearch,
  imageSearch: webImageSearch,
  fetchImage: webFetchImage,
  pickAttachments: attachments.pickAttachments,
  addAttachmentPaths: attachments.addAttachmentPaths,
  addPastedImage: attachments.addPastedImage,
  readAttachment: attachments.readAttachment,
  readAttachmentImage: attachments.readAttachmentImage,
  getPathForFile: attachments.getPathForFile,
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
