import { PDFDocument } from 'pdf-lib'
import html2canvas from 'html2canvas'
import PptxGenJS from 'pptxgenjs'
import type {
  AudienceNavAction,
  DesktopFilesApi,
  OpenResult,
  ShowInkEvent,
  ShowSyncState,
  SlidesApi,
} from '../../../slides/src/shared/ipc'
import { createWebAttachments } from '../lib/attachments'
import {
  cancelWebAiStream,
  getWebAiSettings,
  hasConfiguredWebAi,
  onWebAiStream,
  saveWebAiSettings,
  webAiChat,
  webAiStream,
  webAnalyzeMedia,
  webFetchImage,
  webGenerateImage,
  webImageSearch,
  webSearch,
} from '../lib/ai'
import {
  buildWebSlideHtmlPrompts,
  sanitizeGeneratedSlideHtml,
  type WebSlideHtmlRequest,
} from '../lib/slide-html'
import {
  consumePendingPath,
  getFileHandle,
  getStoredFile,
  listStoredFiles,
  makeWebPath,
  pickBrowserFile,
  putStoredFile,
  readLanguage,
  readTheme,
  rememberFileHandle,
  writeBrowserFile,
} from '../lib/files'

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
type InsertImageResult = Awaited<ReturnType<SlidesApi['insertImage']>>
type EditImageFillResult = Awaited<ReturnType<SlidesApi['editImageFill']>>
type InsertMediaResult = Awaited<ReturnType<SlidesApi['insertMedia']>>
type InsertModel3dResult = Awaited<ReturnType<SlidesApi['insertModel3d']>>
let pendingPath = consumePendingPath()
let sessionId: string | null = null
let currentFile: { path: string; name: string } | null = null
let exportPdfName = 'presentation.pdf'
const webAttachments = createWebAttachments()
const STYLE_TEMPLATES_KEY = 'genoffice.web.slides.style-templates'
const STYLE_SIDECARS_KEY = 'genoffice.web.slides.style-sidecars'
const PRESENTER_SESSION_KEY = 'genoffice.web.slides.presenter-session'
const PRESENTER_SYNC_KEY = 'genoffice.web.slides.presenter-sync'
const presenterChannel =
  typeof BroadcastChannel === 'undefined' ? null : new BroadcastChannel('genoffice-slides-show')
const audienceNavListeners = new Set<(action: AudienceNavAction) => void>()
const showSyncListeners = new Set<(state: ShowSyncState) => void>()
const showInkListeners = new Set<(event: ShowInkEvent) => void>()
let audienceWindow: Window | null = null
const audienceMode = new URLSearchParams(location.search).get('mode') === 'audience'

if (audienceMode) {
  sessionId = localStorage.getItem(PRESENTER_SESSION_KEY)
}

presenterChannel?.addEventListener(
  'message',
  (
    event: MessageEvent<{
      type: 'nav' | 'sync' | 'ink' | 'end'
      value?: AudienceNavAction | ShowSyncState | ShowInkEvent
    }>,
  ) => {
    if (event.data.type === 'nav') {
      for (const listener of audienceNavListeners) listener(event.data.value as AudienceNavAction)
    } else if (event.data.type === 'sync') {
      for (const listener of showSyncListeners) listener(event.data.value as ShowSyncState)
    } else if (event.data.type === 'ink') {
      for (const listener of showInkListeners) listener(event.data.value as ShowInkEvent)
    } else if (event.data.type === 'end' && audienceMode) {
      window.close()
    }
  },
)

type StyleTemplate = { topic: string; styleSkill: string; createdAt: string }

function readJsonMap<T>(key: string): Record<string, T> {
  try {
    const value = JSON.parse(localStorage.getItem(key) || '{}') as Record<string, T>
    return value && typeof value === 'object' ? value : {}
  } catch {
    return {}
  }
}

function writeJsonMap<T>(key: string, value: Record<string, T>): void {
  localStorage.setItem(key, JSON.stringify(value))
}

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

async function slidesRequest<T>(action: string, body: unknown): Promise<T> {
  const response = await fetch(new URL(`./api/slides/${action}`, document.baseURI), {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  })
  if (!response.ok) {
    const result = (await response.json().catch(() => ({}))) as { error?: string }
    throw new Error(result.error || `Slides 服务返回 HTTP ${response.status}`)
  }
  return (await response.json()) as T
}

async function remoteCall<T>(action: string, args: unknown[]): Promise<T> {
  if (!sessionId) return null as T
  return slidesRequest<T>('call', { sessionId, action, args })
}

async function persistPresentation(path: string, name: string, data: ArrayBuffer): Promise<void> {
  await putStoredFile({
    path,
    name,
    kind: 'pptx',
    mime: PPTX_MIME,
    updatedAt: Date.now(),
    data,
  })
}

async function openPresentation(
  path: string,
  name: string,
  data: ArrayBuffer,
  fitWidthPx: number,
): Promise<OpenResult> {
  const response = await slidesRequest<{ sessionId: string; result: OpenResult }>('open', {
    name,
    webPath: path,
    fitWidthPx,
    pptxBase64: base64FromBytes(data),
  })
  sessionId = response.sessionId
  currentFile = { path, name }
  return response.result
}

async function openPptx(fitWidthPx: number): Promise<OpenResult | null> {
  const picked = await pickBrowserFile('pptx')
  if (!picked) return null
  const data = await picked.file.arrayBuffer()
  await persistPresentation(picked.path, picked.file.name, data)
  return openPresentation(picked.path, picked.file.name, data, fitWidthPx)
}

async function openPptxPath(path: string, fitWidthPx: number): Promise<OpenResult | null> {
  const stored = await getStoredFile(path)
  if (!stored || stored.kind !== 'pptx' || typeof stored.data === 'string') return null
  return openPresentation(stored.path, stored.name, stored.data, fitWidthPx)
}

async function consumePendingOpen(fitWidthPx: number): Promise<OpenResult | null> {
  if (!pendingPath) return null
  const path = pendingPath
  pendingPath = null
  return openPptxPath(path, fitWidthPx)
}

async function newBlank(fitWidthPx: number): Promise<OpenResult> {
  const response = await slidesRequest<{ sessionId: string; result: OpenResult }>('blank', {
    fitWidthPx,
  })
  sessionId = response.sessionId
  currentFile = { path: response.result.path, name: 'Presentation.pptx' }
  return response.result
}

async function chooseSaveTarget(
  defaultName: string,
  forcePicker: boolean,
): Promise<{ path: string; name: string } | null> {
  const current = currentFile ?? { path: makeWebPath(defaultName), name: defaultName }
  if (!forcePicker && getFileHandle(current.path)) return current
  if (!window.showSaveFilePicker) {
    return forcePicker ? { path: makeWebPath(defaultName), name: defaultName } : current
  }
  try {
    const handle = await window.showSaveFilePicker({
      suggestedName: defaultName,
      types: [{ description: 'PowerPoint 演示文稿', accept: { [PPTX_MIME]: ['.pptx'] } }],
    })
    const path = makeWebPath(handle.name)
    rememberFileHandle(path, handle)
    return { path, name: handle.name }
  } catch (error) {
    if ((error as DOMException).name === 'AbortError') return null
    throw error
  }
}

async function savePresentation(
  defaultName: string,
  forcePicker: boolean,
): Promise<{ ok: boolean; path?: string; error?: string; slides?: OpenResult['slides'] }> {
  if (!sessionId) return { ok: false, error: 'no file open' }
  const target = await chooseSaveTarget(defaultName, forcePicker)
  if (!target) return { ok: false }
  try {
    const result = await slidesRequest<{
      ok: true
      path: string
      slides: OpenResult['slides']
      pptxBase64: string
    }>('save', { sessionId, name: target.name, webPath: target.path })
    const data = bytesFromBase64(result.pptxBase64)
    const handle = getFileHandle(target.path)
    if (handle) {
      const writable = await handle.createWritable()
      await writable.write(new Blob([data], { type: PPTX_MIME }))
      await writable.close()
    } else {
      await writeBrowserFile({
        path: target.path,
        name: target.name,
        extension: '.pptx',
        mime: PPTX_MIME,
        blob: new Blob([data], { type: PPTX_MIME }),
      })
    }
    await persistPresentation(target.path, target.name, data)
    currentFile = target
    return { ok: true, path: target.path, slides: result.slides }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

function downloadBase64(name: string, mime: string, value: string): void {
  const link = document.createElement('a')
  link.download = name
  link.href = `data:${mime};base64,${value}`
  link.click()
}

function pickSingleFile(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.oncancel = () => resolve(null)
    input.click()
  })
}

function fileExtension(file: File): string {
  return file.name.split('.').pop()?.toLowerCase() || file.type.split('/').pop() || 'bin'
}

async function imageSize(file: File): Promise<{ width: number; height: number }> {
  try {
    const bitmap = await createImageBitmap(file)
    const size = { width: bitmap.width, height: bitmap.height }
    bitmap.close()
    return size
  } catch {
    return { width: 4, height: 3 }
  }
}

async function insertImage(slideIndex: number, fitWidthPx: number): Promise<InsertImageResult> {
  const file = await pickSingleFile('.png,.jpg,.jpeg,.gif,.bmp,.webp,.tif,.tiff')
  if (!file) return null
  const natural = await imageSize(file)
  const maxWidth = fitWidthPx / 2
  const maxHeight = (fitWidthPx * 9) / 32
  const scale = Math.min(maxWidth / natural.width, maxHeight / natural.height)
  const width = Math.max(1, natural.width * scale)
  const height = Math.max(1, natural.height * scale)
  return remoteCall<InsertImageResult>('addImageBytes', [
    {
      slideIndex,
      base64: base64FromBytes(await file.arrayBuffer()),
      ext: fileExtension(file),
      xPx: (fitWidthPx - width) / 2,
      yPx: ((fitWidthPx * 9) / 16 - height) / 2,
      wPx: width,
      hPx: height,
      fitWidthPx,
      name: file.name,
    },
  ])
}

async function editImageFill(op: {
  slideIndex: number
  sourceId: string
}): Promise<EditImageFillResult> {
  const file = await pickSingleFile('.png,.jpg,.jpeg,.gif,.bmp,.webp,.tif,.tiff')
  if (!file) return null
  return remoteCall<EditImageFillResult>('editImageFillBytes', [
    {
      ...op,
      base64: base64FromBytes(await file.arrayBuffer()),
      ext: fileExtension(file),
    },
  ])
}

async function insertMedia(
  slideIndex: number,
  kind: 'video' | 'audio',
  fitWidthPx: number,
): Promise<InsertMediaResult> {
  const accept =
    kind === 'video' ? '.mp4,.m4v,.mov,.webm,.avi,video/*' : '.mp3,.wav,.m4a,.aac,.ogg,audio/*'
  const file = await pickSingleFile(accept)
  if (!file) return null
  return remoteCall<InsertMediaResult>('addMediaBytes', [
    {
      slideIndex,
      kind,
      base64: base64FromBytes(await file.arrayBuffer()),
      ext: fileExtension(file),
      fitWidthPx,
      name: file.name,
    },
  ])
}

async function insertModel3d(slideIndex: number, fitWidthPx: number): Promise<InsertModel3dResult> {
  const file = await pickSingleFile('.glb,.gltf,model/gltf-binary,model/gltf+json')
  if (!file) return null
  return remoteCall<InsertModel3dResult>('addModel3dBytes', [
    {
      slideIndex,
      base64: base64FromBytes(await file.arrayBuffer()),
      ext: fileExtension(file),
      fitWidthPx,
      name: file.name,
    },
  ])
}

async function fetchRemoteImage(
  url: string,
): Promise<{ base64: string; ext: string; mime: string }> {
  const image = await webFetchImage(url)
  if (!image) throw new Error('无法读取远程图片')
  return image
}

async function insertImageUrl(op: {
  slideIndex: number
  url: string
  xPx: number
  yPx: number
  wPx: number
  hPx: number
  fitWidthPx: number
}): Promise<Awaited<ReturnType<SlidesApi['insertImageUrl']>>> {
  const image = await fetchRemoteImage(op.url)
  return remoteCall('addImageBytes', [{ ...op, base64: image.base64, ext: image.ext }])
}

async function replacePictureUrl(op: {
  slideIndex: number
  sourceId: string
  url: string
  keepSrcRect?: boolean
}): Promise<Awaited<ReturnType<SlidesApi['replacePictureUrl']>>> {
  const image = await fetchRemoteImage(op.url)
  return remoteCall('replacePictureBytes', [
    {
      slideIndex: op.slideIndex,
      sourceId: op.sourceId,
      base64: image.base64,
      ext: image.ext,
      keepSrcRect: op.keepSrcRect,
    },
  ])
}

async function renderHtmlPage(html: string, width: number, height: number): Promise<string> {
  const frame = document.createElement('iframe')
  frame.sandbox.add('allow-same-origin')
  frame.style.cssText = `position:fixed;left:-100000px;top:0;width:${width}px;height:${height}px;border:0`
  frame.width = String(width)
  frame.height = String(height)
  document.body.append(frame)
  try {
    await new Promise<void>((resolve, reject) => {
      frame.onload = () => resolve()
      frame.onerror = () => reject(new Error('HTML 页面加载失败'))
      frame.srcdoc = html
    })
    const frameDocument = frame.contentDocument
    if (!frameDocument) throw new Error('无法读取 HTML 页面')
    await frameDocument.fonts?.ready
    await Promise.all(
      [...frameDocument.images].map((image) =>
        image.complete ? image.decode().catch(() => undefined) : Promise.resolve(),
      ),
    )
    const canvas = await html2canvas(frameDocument.body, {
      width,
      height,
      windowWidth: width,
      windowHeight: height,
      scale: 1,
      useCORS: true,
      backgroundColor: '#FFFFFF',
      logging: false,
    })
    return canvas.toDataURL('image/png')
  } finally {
    frame.remove()
  }
}

async function htmlToPptx(
  pagesHtml: string[],
  fitWidthPx: number,
  mode: 'replace' | 'append' | 'replace_at' | 'insert_at' = 'replace',
  atIndex?: number,
  deckName = 'Generated.pptx',
): Promise<OpenResult | { error: string }> {
  if (pagesHtml.length === 0 || pagesHtml.length > 100) return { error: '生成页数无效' }
  try {
    if (!sessionId) await newBlank(fitWidthPx)
    const width = 1280
    const height = 720
    const images = await Promise.all(pagesHtml.map((html) => renderHtmlPage(html, width, height)))
    const presentation = new PptxGenJS()
    presentation.layout = 'LAYOUT_WIDE'
    presentation.author = 'GenOffice Web'
    presentation.subject = deckName
    for (const data of images) {
      presentation.addSlide().addImage({ data, x: 0, y: 0, w: 13.333333, h: 7.5 })
    }
    const output = await presentation.write({ outputType: 'arraybuffer', compression: true })
    const bytes =
      output instanceof ArrayBuffer
        ? output
        : output instanceof Uint8Array
          ? Uint8Array.from(output).buffer
          : await (output as Blob).arrayBuffer()
    const result = await remoteCall<OpenResult>('importGenerated', [
      {
        pptxBase64: base64FromBytes(bytes),
        fitWidthPx,
        mode,
        atIndex,
        deckName,
      },
    ])
    currentFile = {
      path: result.path,
      name: deckName.toLowerCase().endsWith('.pptx') ? deckName : `${deckName}.pptx`,
    }
    return result
  } catch (error) {
    return { error: error instanceof Error ? error.message : String(error) }
  }
}

async function generateWebSlidePage(
  request: WebSlideHtmlRequest,
): Promise<{ ok: boolean; marker?: string; error?: string }> {
  const settings = getWebAiSettings()
  if (!hasConfiguredWebAi(settings)) {
    return { ok: false, error: '请先在工作台设置中配置自定义 AI 模型' }
  }
  const prompts = buildWebSlideHtmlPrompts(request)
  const result = await webAiChat({ settings, ...prompts })
  if (!result.ok || !result.content) {
    return { ok: false, error: result.error || '模型未返回幻灯片内容' }
  }
  try {
    return {
      ok: true,
      marker: sanitizeGeneratedSlideHtml(result.content, request),
    }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

function menuCommands(callback: Parameters<SlidesApi['onMenuCommand']>[0]): () => void {
  const listener = (event: KeyboardEvent) => {
    if (!(event.metaKey || event.ctrlKey)) return
    const key = event.key.toLowerCase()
    const command =
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
    if (!command) return
    event.preventDefault()
    callback(command)
  }
  window.addEventListener('keydown', listener)
  return () => window.removeEventListener('keydown', listener)
}

async function exportPdf(op: {
  pngsBase64: string[]
  widthPx: number
  heightPx: number
}): Promise<{ ok: boolean; path?: string; error?: string }> {
  try {
    const document = await PDFDocument.create()
    for (const png of op.pngsBase64) {
      const image = await document.embedPng(bytesFromBase64(png))
      const page = document.addPage([op.widthPx, op.heightPx])
      page.drawImage(image, { x: 0, y: 0, width: op.widthPx, height: op.heightPx })
    }
    const bytes = await document.save()
    await writeBrowserFile({
      name: exportPdfName,
      extension: '.pdf',
      mime: 'application/pdf',
      blob: new Blob([new Uint8Array(bytes).buffer as ArrayBuffer], { type: 'application/pdf' }),
      forcePicker: true,
    })
    return { ok: true, path: exportPdfName }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

const localApi: Partial<SlidesApi> = {
  getLanguage: async () => readLanguage(),
  onLanguageChanged: () => () => {},
  getTheme: async () => readTheme(),
  onThemeChanged: () => () => {},
  openPptx,
  openPptxPath,
  consumePendingOpen,
  newBlank,
  save: () => savePresentation(currentFile?.name ?? 'Presentation.pptx', false),
  saveAs: (defaultName) => savePresentation(defaultName, true),
  getRecentFiles: async () => (await listStoredFiles('pptx')).map((file) => file.path),
  onMenuCommand: menuCommands,
  onOpened: () => () => {},
  onRenamed: () => () => {},
  onCloseSaveRequest: () => () => {},
  reportCloseSaveResult: () => {},
  setAutoSavePref: () => {},
  getAiSettings: async () => getWebAiSettings(),
  setAiSettings: async (settings) => saveWebAiSettings(settings),
  aiStream: webAiStream,
  aiStreamCancel: cancelWebAiStream,
  onAiStream: onWebAiStream,
  aiGskStatus: async () => ({ loggedIn: hasConfiguredWebAi() }),
  aiGskLogin: async () => {
    window.open('https://www.genspark.ai/', '_blank', 'noopener,noreferrer')
  },
  gskStatus: async () => ({ available: hasConfiguredWebAi() }),
  webSearch,
  imageSearch: webImageSearch,
  generateImage: webGenerateImage,
  analyzeMedia: webAnalyzeMedia,
  htmlToPptx,
  cloudGenStatus: async () => ({ enabled: hasConfiguredWebAi() }),
  cloudGeneratePage: generateWebSlidePage,
  clipboardExternal: async () => {
    try {
      if (navigator.clipboard.read) {
        for (const item of await navigator.clipboard.read()) {
          const imageType = item.types.find((type) => type.startsWith('image/'))
          if (!imageType) continue
          const blob = await item.getType(imageType)
          return {
            kind: 'image' as const,
            base64: base64FromBytes(await blob.arrayBuffer()),
            ext: imageType.split('/')[1]?.replace('jpeg', 'jpg') || 'png',
          }
        }
      }
      const text = await navigator.clipboard.readText()
      return text ? { kind: 'text' as const, text } : { kind: 'none' as const }
    } catch {
      return { kind: 'none' as const }
    }
  },
  nativeClipboard: async (operation) => {
    document.execCommand(operation)
  },
  insertImage,
  insertImageUrl,
  replacePictureUrl,
  editImageFill,
  insertMedia,
  insertModel3d,
  pickExportDir: async () => 'browser-downloads',
  exportImages: async (op) => {
    const paths = op.pngsBase64.map((png, index) => {
      const name = `${op.baseName}-${String(index + 1).padStart(2, '0')}.png`
      downloadBase64(name, 'image/png', png)
      return name
    })
    return { ok: true, paths }
  },
  pickExportPdfPath: async (defaultName) => {
    exportPdfName = defaultName
    return defaultName
  },
  exportPdf,
  printSlides: async () => {
    window.print()
    return { ok: true }
  },
  saveStyleSidecar: async (data) => {
    const sidecars = readJsonMap<StyleTemplate>(STYLE_SIDECARS_KEY)
    sidecars[currentFile?.path ?? 'unsaved'] = data
    writeJsonMap(STYLE_SIDECARS_KEY, sidecars)
    return { ok: true }
  },
  saveStyleTemplate: async (name, data) => {
    const safeName = name.trim().slice(0, 120)
    if (!safeName) return { ok: false, error: '模板名称不能为空' }
    const templates = readJsonMap<StyleTemplate>(STYLE_TEMPLATES_KEY)
    templates[safeName] = data
    writeJsonMap(STYLE_TEMPLATES_KEY, templates)
    return { ok: true }
  },
  listStyleTemplates: async () => {
    const templates = readJsonMap<StyleTemplate>(STYLE_TEMPLATES_KEY)
    return Object.entries(templates)
      .map(([name, data]) => ({ name, topic: data.topic, createdAt: data.createdAt }))
      .sort((left, right) => right.createdAt.localeCompare(left.createdAt))
  },
  loadStyleTemplate: async (name) => {
    const data = readJsonMap<StyleTemplate>(STYLE_TEMPLATES_KEY)[name]
    return data ? { ok: true, ...data } : { ok: false, error: '模板不存在' }
  },
  presenterStart: async () => {
    if (!sessionId) return { audience: false }
    localStorage.setItem(PRESENTER_SESSION_KEY, sessionId)
    const url = new URL('./slides.html?mode=audience', document.baseURI)
    audienceWindow = window.open(url, 'genoffice-slides-audience', 'popup=yes')
    return { audience: !!audienceWindow }
  },
  presenterSync: (state) => {
    localStorage.setItem(PRESENTER_SYNC_KEY, JSON.stringify(state))
    presenterChannel?.postMessage({ type: 'sync', value: state })
  },
  presenterInk: (event) => presenterChannel?.postMessage({ type: 'ink', value: event }),
  presenterSwap: async () => false,
  presenterEnd: async () => {
    presenterChannel?.postMessage({ type: 'end' })
    audienceWindow?.close()
    audienceWindow = null
    localStorage.removeItem(PRESENTER_SESSION_KEY)
    localStorage.removeItem(PRESENTER_SYNC_KEY)
  },
  audienceReady: async () => {
    try {
      return JSON.parse(localStorage.getItem(PRESENTER_SYNC_KEY) || 'null') as ShowSyncState | null
    } catch {
      return null
    }
  },
  audienceNav: (action) => presenterChannel?.postMessage({ type: 'nav', value: action }),
  onAudienceNav: (listener) => {
    audienceNavListeners.add(listener)
    return () => audienceNavListeners.delete(listener)
  },
  onShowSync: (listener) => {
    showSyncListeners.add(listener)
    return () => showSyncListeners.delete(listener)
  },
  onShowInk: (listener) => {
    showInkListeners.add(listener)
    return () => showInkListeners.delete(listener)
  },
}

window.slidesApi = new Proxy(localApi, {
  get(target, property) {
    const local = target[property as keyof SlidesApi]
    if (local) return local
    if (typeof property !== 'string') return undefined
    return (...args: unknown[]) => remoteCall(property, args)
  },
}) as SlidesApi

const desktopFiles: DesktopFilesApi = {
  pickAttachments: webAttachments.pickAttachments,
  addAttachmentPaths: webAttachments.addAttachmentPaths,
  addPastedImage: webAttachments.addPastedImage,
  readAttachment: webAttachments.readAttachment,
  readAttachmentImage: webAttachments.readAttachmentImage,
  getPathForFile: webAttachments.getPathForFile,
}

window.desktop = desktopFiles as unknown as Window['desktop']

window.addEventListener('pagehide', () => {
  presenterChannel?.close()
  if (!sessionId || audienceMode) return
  navigator.sendBeacon(
    new URL('./api/slides/close', document.baseURI),
    JSON.stringify({ sessionId }),
  )
  sessionId = null
})
