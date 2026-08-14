import { createHash } from 'node:crypto'
import { createReadStream, existsSync, statSync } from 'node:fs'
import { createServer, type IncomingMessage, type ServerResponse } from 'node:http'
import { extname, join, normalize, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
  AiCreditsError,
  AiTimeoutError,
  chatForProvider,
  streamForProvider,
  type AiChatRequest,
  type AiProviderConfig,
  type AiReasoningEffort,
  type AiStreamChunk,
  type AiStreamRequest,
} from '@genoffice/ai-provider'
import { imageSearch, webSearch } from '@genoffice/ai-search'
import { docxToText, pdfToText, pptxToText, xlsxToText } from '@genoffice/file-parse'
import {
  applyImageEdits,
  listPageImages,
  renderImagePng,
  renderPagePreviewPng,
} from '../../pdf/src/main/image-edit'
import { applyTextEdits, listEditFonts, validateTextEdits } from '../../pdf/src/main/text-edit'
import { requestPdfTimestampToken } from '../../pdf/src/main/timestamp'
import { fetchPdfWebResource } from '../../pdf/src/main/web-resource'
import {
  handleMobileScannerPublicRequest,
  MobileScannerError,
  MobileScannerHub,
} from '../../pdf/src/main/mobile-scanner'
import { isPdfMobileScannerSessionId } from '../../pdf/src/shared/mobile-scanner'
import type {
  ImageEditInput,
  PagePreviewRequest,
  PdfWebResourceRequest,
  TextEditInput,
} from '../../pdf/src/shared/ipc'
import { normalizeProviderModels } from './ai-models'
import {
  validateProviderBaseUrl,
  validatePublicResourceUrl,
  webContentSecurityPolicy,
} from './security'
import { SheetsWebService } from './sheets'
import { SlidesWebService } from './slides'

const serverDirectory = fileURLToPath(new URL('.', import.meta.url))
const staticRoot = resolve(serverDirectory, '../dist')
const port = Number(process.env.PORT || 80)
const basePath = normalizeBasePath(process.env.WEB_BASE_PATH || '/')
const maxRequestBytes = Number(process.env.WEB_MAX_REQUEST_BYTES || 128 * 1024 * 1024)
const sheetsService = new SheetsWebService(
  process.env.XLSX_SIDECAR_PATH || join(serverDirectory, 'native', 'xlsx-sidecar'),
  Number(process.env.WEB_MAX_WORKBOOK_BYTES || 64 * 1024 * 1024),
)
const slidesService = new SlidesWebService(
  Number(process.env.WEB_MAX_PRESENTATION_BYTES || 96 * 1024 * 1024),
  Number(process.env.WEB_MAX_SLIDES_SESSIONS || 64),
  Number(process.env.WEB_SLIDES_SESSION_TTL_MS || 30 * 60 * 1000),
)
const maxRemoteImageBytes = Number(process.env.WEB_MAX_REMOTE_IMAGE_BYTES || 20 * 1024 * 1024)
const maxAttachmentBytes = Number(process.env.WEB_MAX_ATTACHMENT_BYTES || 50 * 1024 * 1024)
const attachmentTextCache = new Map<string, string>()
const mobileScannerHub = new MobileScannerHub()
const attachmentTextExtensions = new Set([
  'txt',
  'md',
  'markdown',
  'csv',
  'tsv',
  'json',
  'yaml',
  'yml',
  'xml',
  'html',
  'htm',
  'log',
  'js',
  'ts',
  'tsx',
  'jsx',
  'py',
  'java',
  'c',
  'h',
  'cpp',
  'go',
  'rs',
  'rb',
  'sh',
  'sql',
  'css',
])

const mimeTypes: Record<string, string> = {
  '.css': 'text/css; charset=utf-8',
  '.gif': 'image/gif',
  '.html': 'text/html; charset=utf-8',
  '.jpeg': 'image/jpeg',
  '.jpg': 'image/jpeg',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ttf': 'font/ttf',
  '.wasm': 'application/wasm',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
}

function normalizeBasePath(value: string): string {
  const path = `/${value}`.replace(/\/{2,}/g, '/').replace(/\/$/, '')
  return path || '/'
}

function applySecurityHeaders(response: ServerResponse): void {
  response.setHeader('X-Content-Type-Options', 'nosniff')
  response.setHeader('Referrer-Policy', 'same-origin')
  response.setHeader('Permissions-Policy', 'camera=(self), microphone=(), geolocation=()')
  response.setHeader('Content-Security-Policy', webContentSecurityPolicy)
}

function json(response: ServerResponse, status: number, body: unknown): void {
  response.statusCode = status
  response.setHeader('Content-Type', 'application/json; charset=utf-8')
  response.end(JSON.stringify(body))
}

async function readJson<T>(request: IncomingMessage): Promise<T> {
  const chunks: Buffer[] = []
  let size = 0
  for await (const chunk of request) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    size += buffer.length
    if (size > maxRequestBytes) throw new Error('请求内容过大')
    chunks.push(buffer)
  }
  try {
    return JSON.parse(Buffer.concat(chunks).toString('utf8')) as T
  } catch {
    throw new Error('请求 JSON 无效')
  }
}

async function handleAttachmentText(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<{
    name?: unknown
    base64?: unknown
    offset?: unknown
    maxChars?: unknown
  }>(request)
  if (typeof body.name !== 'string' || body.name.length === 0 || body.name.length > 255) {
    throw new Error('附件名称无效')
  }
  if (
    typeof body.base64 !== 'string' ||
    body.base64.length > Math.ceil((maxAttachmentBytes * 4) / 3) + 8
  ) {
    throw new Error('附件内容无效或过大')
  }
  const bytes = Buffer.from(body.base64, 'base64')
  if (bytes.length === 0 || bytes.length > maxAttachmentBytes) throw new Error('附件内容无效或过大')
  const ext = extname(body.name).slice(1).toLowerCase()
  const supported =
    attachmentTextExtensions.has(ext) || ['docx', 'pdf', 'pptx', 'xlsx'].includes(ext)
  if (!supported) throw new Error(`不支持 .${ext || 'unknown'} 附件`)

  const cacheKey = `${ext}:${createHash('sha256').update(bytes).digest('hex')}`
  let text = attachmentTextCache.get(cacheKey)
  if (text === undefined) {
    if (attachmentTextExtensions.has(ext)) text = bytes.toString('utf8')
    else if (ext === 'docx') text = await docxToText(bytes)
    else if (ext === 'pdf') text = await pdfToText(bytes)
    else if (ext === 'pptx') text = await pptxToText(bytes)
    else text = await xlsxToText(bytes)
    attachmentTextCache.set(cacheKey, text)
    if (attachmentTextCache.size > 8) {
      const oldest = attachmentTextCache.keys().next().value
      if (oldest) attachmentTextCache.delete(oldest)
    }
  }

  const offset = Math.max(0, Math.floor(Number(body.offset) || 0))
  const maxChars = Math.min(100_000, Math.max(1, Math.floor(Number(body.maxChars) || 20_000)))
  json(response, 200, {
    ok: true,
    name: body.name,
    totalChars: text.length,
    text: text.slice(offset, offset + maxChars),
    offset,
  })
}

async function handlePdfTimestampToken(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<{ tsaUrl?: unknown; requestBase64?: unknown }>(request)
  if (typeof body.tsaUrl !== 'string' || typeof body.requestBase64 !== 'string') {
    throw new Error('时间戳请求无效')
  }
  if (body.requestBase64.length === 0 || body.requestBase64.length > 96 * 1024) {
    throw new Error('时间戳请求为空或过大')
  }
  const requestBytes = Uint8Array.from(Buffer.from(body.requestBase64, 'base64'))
  const timestamp = await requestPdfTimestampToken(body.tsaUrl, requestBytes)
  json(response, 200, { ok: true, responseBase64: Buffer.from(timestamp).toString('base64') })
}

async function customConfig(settings: AiStreamRequest['settings']): Promise<AiProviderConfig> {
  if (settings?.provider !== 'custom') throw new Error('Web 版请先配置自定义模型')
  const source = settings.providers?.custom
  if (!source?.apiKey || source.apiKey.length > 8192) throw new Error('请填写 API Key')
  if (!source.model || source.model.length > 256) throw new Error('请填写模型名称')
  const reasoningEffort = source.reasoningEffort || 'default'
  if (!['default', 'low', 'medium', 'high'].includes(reasoningEffort)) {
    throw new Error('思考强度无效')
  }
  return {
    apiKey: source.apiKey,
    model: source.model,
    baseUrl: await validateProviderBaseUrl(source.baseUrl),
    reasoningEffort: reasoningEffort as AiReasoningEffort,
  }
}

function providerEndpoint(baseUrl: string, path: string): string {
  return `${baseUrl.replace(/\/$/, '')}/${path.replace(/^\//, '')}`
}

function aiLogContext(config: AiProviderConfig, requestId: string, startedAt: number) {
  let provider = 'custom'
  try {
    const url = new URL(config.baseUrl || '')
    provider = `${url.host}${url.pathname.replace(/\/$/, '')}`
  } catch {
    // customConfig already validates this URL; retain a non-sensitive fallback for diagnostics.
  }
  return {
    requestId,
    provider,
    model: config.model,
    reasoningEffort: config.reasoningEffort || 'default',
    durationMs: Date.now() - startedAt,
  }
}

async function providerError(response: Response): Promise<string> {
  try {
    const body = (await response.json()) as {
      error?: string | { message?: string }
      message?: string
    }
    if (typeof body.error === 'string') return body.error
    return body.error?.message || body.message || `HTTP ${response.status}`
  } catch {
    return `HTTP ${response.status}`
  }
}

async function handleModels(request: IncomingMessage, response: ServerResponse): Promise<void> {
  const body = await readJson<{ baseUrl?: unknown; apiKey?: unknown }>(request)
  if (typeof body.apiKey !== 'string' || !body.apiKey || body.apiKey.length > 8192) {
    throw new Error('请填写 API Key')
  }
  const baseUrl = await validateProviderBaseUrl(body.baseUrl)
  const result = await fetch(providerEndpoint(baseUrl, 'models'), {
    headers: { Authorization: `Bearer ${body.apiKey}`, Accept: 'application/json' },
    redirect: 'error',
    signal: AbortSignal.timeout(20_000),
  })
  if (!result.ok) return json(response, 502, { error: await providerError(result) })
  const models = normalizeProviderModels(await result.json())
  if (models.length === 0) return json(response, 502, { error: '接口未返回可用模型' })
  json(response, 200, { models })
}

async function handleChat(request: IncomingMessage, response: ServerResponse): Promise<void> {
  const body = await readJson<AiChatRequest>(request)
  const config = await customConfig(body.settings)
  const result = await chatForProvider(
    'custom',
    config,
    String(body.system || ''),
    String(body.user || ''),
  )
  json(response, result.ok ? 200 : 502, result)
}

async function handleStream(request: IncomingMessage, response: ServerResponse): Promise<void> {
  const body = await readJson<AiStreamRequest>(request)
  const config = await customConfig(body.settings)
  const requestId = String(body.requestId || '')
  if (!requestId || requestId.length > 128) throw new Error('requestId 无效')
  const startedAt = Date.now()

  const controller = new AbortController()
  request.once('aborted', () => controller.abort())
  response.once('close', () => {
    if (!response.writableEnded) controller.abort()
  })
  response.statusCode = 200
  response.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8')
  response.setHeader('Cache-Control', 'no-store')
  response.setHeader('X-Accel-Buffering', 'no')
  response.flushHeaders()

  const send = (chunk: AiStreamChunk) => {
    if (!response.destroyed && !response.writableEnded) response.write(`${JSON.stringify(chunk)}\n`)
  }
  let lastPing = 0
  const ping = () => {
    const now = Date.now()
    if (now - lastPing >= 5_000) {
      lastPing = now
      send({ requestId, type: 'ping' })
    }
  }
  ping()
  const keepAliveTimer = setInterval(ping, 5_000)
  keepAliveTimer.unref()

  try {
    let stopReason: string | undefined
    await streamForProvider(
      'custom',
      config,
      String(body.system || ''),
      Array.isArray(body.messages) ? body.messages : [],
      Array.isArray(body.tools) ? body.tools : [],
      Math.min(Math.max(Number(body.maxTokens) || 8192, 1), 32768),
      {
        signal: controller.signal,
        onDelta: (text) => send({ requestId, type: 'delta', text }),
        onToolCall: (toolCall) => send({ requestId, type: 'tool-call', toolCall }),
        onActivity: ping,
        onStopReason: (reason) => {
          stopReason = reason
        },
      },
    )
    send({ requestId, type: 'done', stopReason })
    console.info(
      JSON.stringify({
        event: 'ai_stream_complete',
        ...aiLogContext(config, requestId, startedAt),
      }),
    )
  } catch (error) {
    console.error(
      JSON.stringify({
        event: 'ai_stream_error',
        ...aiLogContext(config, requestId, startedAt),
        error: error instanceof Error ? error.message : String(error),
      }),
    )
    if (controller.signal.aborted) {
      send({ requestId, type: 'done' })
    } else {
      send({
        requestId,
        type: 'error',
        error: error instanceof Error ? error.message : String(error),
        ...(error instanceof AiTimeoutError
          ? { errorCode: 'timeout' as const }
          : error instanceof AiCreditsError
            ? { errorCode: 'credits' as const }
            : {}),
      })
    }
  } finally {
    clearInterval(keepAliveTimer)
    response.end()
  }
}

const remoteImageExtensions: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/gif': 'gif',
  'image/webp': 'webp',
  'image/bmp': 'bmp',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
}

async function downloadRemoteImage(source: unknown): Promise<{
  base64: string
  ext: string
  mime: string
}> {
  let url = await validatePublicResourceUrl(source)
  for (let redirect = 0; redirect <= 3; redirect += 1) {
    const result = await fetch(url, {
      redirect: 'manual',
      signal: AbortSignal.timeout(15_000),
      headers: { Accept: 'image/*' },
    })
    if (result.status >= 300 && result.status < 400) {
      const location = result.headers.get('location')
      if (!location || redirect === 3) throw new Error('远程图片重定向无效')
      url = await validatePublicResourceUrl(new URL(location, url).toString())
      continue
    }
    if (!result.ok || !result.body) throw new Error(`远程图片返回 HTTP ${result.status}`)
    const mime = result.headers.get('content-type')?.split(';')[0]?.trim().toLowerCase() || ''
    const ext = remoteImageExtensions[mime]
    if (!ext) throw new Error('远程资源不是支持的图片格式')
    const declared = Number(result.headers.get('content-length') || 0)
    if (declared > maxRemoteImageBytes) throw new Error('远程图片过大')
    const reader = result.body.getReader()
    const chunks: Uint8Array[] = []
    let total = 0
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      total += value.byteLength
      if (total > maxRemoteImageBytes) {
        await reader.cancel()
        throw new Error('远程图片过大')
      }
      chunks.push(value)
    }
    return { base64: Buffer.concat(chunks).toString('base64'), ext, mime }
  }
  throw new Error('远程图片重定向过多')
}

interface AiImageRequest {
  settings: AiStreamRequest['settings']
  prompt: string
  model?: string
  referenceImageUrls?: string[]
  aspectRatio?: string
  imageSize?: string
}

function requestedImageSize(body: AiImageRequest): string {
  if (body.imageSize && /^\d{3,4}x\d{3,4}$/.test(body.imageSize)) return body.imageSize
  if (body.aspectRatio === '16:9') return '1536x1024'
  if (body.aspectRatio === '9:16') return '1024x1536'
  return '1024x1024'
}

async function handleImageGeneration(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<AiImageRequest>(request)
  const config = await customConfig(body.settings)
  const prompt = String(body.prompt || '').trim()
  if (!prompt || prompt.length > 20_000) throw new Error('图片生成提示词无效')
  const headers = { Authorization: `Bearer ${config.apiKey}` }
  let result: Response
  if (body.referenceImageUrls?.length) {
    const form = new FormData()
    form.set('prompt', prompt)
    form.set('model', body.model || config.model)
    form.set('size', requestedImageSize(body))
    for (const [index, source] of body.referenceImageUrls.slice(0, 4).entries()) {
      const image = await downloadRemoteImage(source)
      form.append(
        'image[]',
        new Blob([Buffer.from(image.base64, 'base64')], { type: image.mime }),
        `reference-${index + 1}.${image.ext}`,
      )
    }
    result = await fetch(providerEndpoint(config.baseUrl || '', 'images/edits'), {
      method: 'POST',
      headers,
      body: form,
      signal: AbortSignal.timeout(120_000),
    })
  } else {
    result = await fetch(providerEndpoint(config.baseUrl || '', 'images/generations'), {
      method: 'POST',
      headers: { ...headers, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        prompt,
        model: body.model || config.model,
        size: requestedImageSize(body),
        n: 1,
      }),
      signal: AbortSignal.timeout(120_000),
    })
  }
  if (!result.ok) return json(response, 502, { error: await providerError(result) })
  const payload = (await result.json()) as {
    data?: Array<{ url?: string; b64_json?: string; revised_prompt?: string }>
  }
  const image = payload.data?.[0]
  if (image?.url) return json(response, 200, { url: image.url })
  if (image?.b64_json)
    return json(response, 200, { url: `data:image/png;base64,${image.b64_json}` })
  json(response, 502, { error: '图片模型未返回图片' })
}

interface AiMediaRequest {
  settings: AiStreamRequest['settings']
  mediaUrls: string[]
  requirements: string
}

async function handleMediaAnalysis(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<AiMediaRequest>(request)
  const config = await customConfig(body.settings)
  const mediaUrls = Array.isArray(body.mediaUrls) ? body.mediaUrls.slice(0, 12) : []
  const requirements = String(body.requirements || '').trim()
  if (!requirements || requirements.length > 20_000 || mediaUrls.length === 0) {
    throw new Error('媒体分析请求无效')
  }
  for (const url of mediaUrls) {
    if (!url.startsWith('data:image/')) await validatePublicResourceUrl(url)
  }
  const result = await fetch(providerEndpoint(config.baseUrl || '', 'chat/completions'), {
    method: 'POST',
    headers: { Authorization: `Bearer ${config.apiKey}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: config.model,
      ...(config.reasoningEffort && config.reasoningEffort !== 'default'
        ? { reasoning_effort: config.reasoningEffort }
        : {}),
      messages: [
        {
          role: 'user',
          content: [
            { type: 'text', text: requirements },
            ...mediaUrls.map((url) => ({ type: 'image_url', image_url: { url } })),
          ],
        },
      ],
    }),
    signal: AbortSignal.timeout(120_000),
  })
  if (!result.ok) return json(response, 502, { error: await providerError(result) })
  const payload = (await result.json()) as {
    choices?: Array<{ message?: { content?: string | Array<{ type?: string; text?: string }> } }>
  }
  const content = payload.choices?.[0]?.message?.content
  const text =
    typeof content === 'string'
      ? content
      : content
          ?.map((item) => item.text || '')
          .filter(Boolean)
          .join('\n')
  json(response, text ? 200 : 502, text ? { text } : { error: '媒体模型未返回分析结果' })
}

interface SearchRequest {
  query?: unknown
  maxResults?: unknown
}

function searchInput(body: SearchRequest, fallback: number): { query: string; maxResults: number } {
  const query = typeof body.query === 'string' ? body.query.trim() : ''
  if (!query || query.length > 500) throw new Error('搜索关键词无效')
  return {
    query,
    maxResults: Math.min(Math.max(Number(body.maxResults) || fallback, 1), 20),
  }
}

function plainSnippet(value: unknown): string {
  return String(value || '')
    .replace(/<[^>]+>/g, '')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&amp;/g, '&')
}

async function wikipediaSearch(query: string, maxResults: number) {
  try {
    const language = /[\u3400-\u9fff]/.test(query) ? 'zh' : 'en'
    const endpoint = new URL(`https://${language}.wikipedia.org/w/api.php`)
    endpoint.search = new URLSearchParams({
      action: 'query',
      list: 'search',
      srsearch: query,
      srlimit: String(maxResults),
      format: 'json',
      origin: '*',
    }).toString()
    const response = await fetch(endpoint, { signal: AbortSignal.timeout(15_000) })
    if (!response.ok) return { results: [], method: 'wikipedia' }
    const payload = (await response.json()) as {
      query?: { search?: Array<{ title?: string; snippet?: string }> }
    }
    return {
      results: (payload.query?.search || []).map((item) => ({
        title: String(item.title || ''),
        url: `https://${language}.wikipedia.org/wiki/${encodeURIComponent(String(item.title || '').replace(/ /g, '_'))}`,
        snippet: plainSnippet(item.snippet),
      })),
      method: 'wikipedia',
    }
  } catch {
    return { results: [], method: 'error' }
  }
}

async function wikimediaImageSearch(query: string, maxResults: number) {
  try {
    const endpoint = new URL('https://commons.wikimedia.org/w/api.php')
    endpoint.search = new URLSearchParams({
      action: 'query',
      generator: 'search',
      gsrsearch: query,
      gsrnamespace: '6',
      gsrlimit: String(maxResults),
      prop: 'imageinfo',
      iiprop: 'url|size',
      iiurlwidth: '1600',
      format: 'json',
      origin: '*',
    }).toString()
    const response = await fetch(endpoint, { signal: AbortSignal.timeout(15_000) })
    if (!response.ok) return { images: [], method: 'wikimedia' }
    const payload = (await response.json()) as {
      query?: {
        pages?: Record<
          string,
          {
            title?: string
            imageinfo?: Array<{
              url?: string
              thumburl?: string
              descriptionurl?: string
              width?: number
              height?: number
            }>
          }
        >
      }
    }
    const images = Object.values(payload.query?.pages || {}).flatMap((page) => {
      const image = page.imageinfo?.[0]
      const imageUrl = image?.thumburl || image?.url
      if (!imageUrl) return []
      return [
        {
          title: String(page.title || '').replace(/^File:/, ''),
          imageUrl,
          sourceUrl: image.descriptionurl || image.url || imageUrl,
          source: 'Wikimedia Commons',
          ...(typeof image.width === 'number' ? { width: image.width } : {}),
          ...(typeof image.height === 'number' ? { height: image.height } : {}),
        },
      ]
    })
    return { images: images.slice(0, maxResults), method: 'wikimedia' }
  } catch {
    return { images: [], method: 'error' }
  }
}

interface PdfRequest {
  pdfBase64: string
}

interface PdfTextRequest extends PdfRequest {
  edits: TextEditInput[]
}

function pdfPayload(body: PdfRequest): Uint8Array {
  if (typeof body.pdfBase64 !== 'string') throw new Error('PDF 请求无效')
  const bytes = Buffer.from(body.pdfBase64, 'base64')
  if (bytes.length < 5 || bytes.subarray(0, 5).toString('ascii') !== '%PDF-') {
    throw new Error('PDF 文件无效')
  }
  return bytes
}

function pdfTextPayload(body: PdfTextRequest): { bytes: Uint8Array; edits: TextEditInput[] } {
  if (!Array.isArray(body.edits)) throw new Error('PDF 文字编辑请求无效')
  return { bytes: pdfPayload(body), edits: body.edits }
}

async function handlePdfTextValidate(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const { bytes, edits } = pdfTextPayload(await readJson<PdfTextRequest>(request))
  json(response, 200, { validations: await validateTextEdits(bytes, edits) })
}

async function handlePdfTextApply(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const { bytes, edits } = pdfTextPayload(await readJson<PdfTextRequest>(request))
  const result = await applyTextEdits(bytes, edits)
  json(response, 200, {
    pdfBase64: Buffer.from(result.bytes).toString('base64'),
    skipped: result.skipped,
  })
}

async function handlePdfImagesList(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const bytes = pdfPayload(await readJson<PdfRequest>(request))
  json(response, 200, { images: await listPageImages(bytes) })
}

interface PdfImageRenderRequest extends PdfRequest {
  pageIndex: number
  rect: [number, number, number, number]
}

async function handlePdfImageRender(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<PdfImageRenderRequest>(request)
  if (!Number.isInteger(body.pageIndex) || !Array.isArray(body.rect)) {
    throw new Error('PDF 图片渲染请求无效')
  }
  json(response, 200, {
    pngBase64: await renderImagePng(pdfPayload(body), body.pageIndex, body.rect),
  })
}

interface PdfPagePreviewRequest extends PdfRequest, Omit<PagePreviewRequest, 'path'> {}

async function handlePdfPagePreview(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<PdfPagePreviewRequest>(request)
  const { pageIndex, excludeRects, clip, pxWidth, rotate } = body
  if (
    !Number.isInteger(pageIndex) ||
    !Array.isArray(excludeRects) ||
    !clip ||
    typeof pxWidth !== 'number' ||
    typeof rotate !== 'number'
  ) {
    throw new Error('PDF 页面预览请求无效')
  }
  json(response, 200, {
    pngBase64: await renderPagePreviewPng(pdfPayload(body), {
      pageIndex,
      excludeRects,
      clip,
      pxWidth,
      rotate,
    }),
  })
}

interface PdfImageApplyRequest extends PdfRequest {
  edits: ImageEditInput[]
}

async function handlePdfImageApply(
  request: IncomingMessage,
  response: ServerResponse,
): Promise<void> {
  const body = await readJson<PdfImageApplyRequest>(request)
  if (!Array.isArray(body.edits)) throw new Error('PDF 图片编辑请求无效')
  const result = await applyImageEdits(pdfPayload(body), body.edits)
  json(response, 200, {
    pdfBase64: Buffer.from(result.bytes).toString('base64'),
    skipped: result.skipped,
  })
}

function requestPath(url: URL): string | null {
  if (basePath === '/') return url.pathname
  if (url.pathname === basePath) return ''
  return url.pathname.startsWith(`${basePath}/`) ? url.pathname.slice(basePath.length) : null
}

function serveStatic(
  pathname: string,
  response: ServerResponse,
  headOnly = false,
  acceptEncoding = '',
): void {
  if (pathname === '/favicon.ico') {
    response.statusCode = 204
    response.end()
    return
  }
  const requested = pathname === '' || pathname === '/' ? '/index.html' : pathname
  let decoded: string
  try {
    decoded = decodeURIComponent(requested)
  } catch {
    return json(response, 400, { error: '路径无效' })
  }
  const safePath = normalize(decoded).replace(/^(\.\.(\/|\\|$))+/, '')
  const filePath = resolve(join(staticRoot, safePath))
  if (
    !filePath.startsWith(`${staticRoot}/`) ||
    !existsSync(filePath) ||
    !statSync(filePath).isFile()
  ) {
    return json(response, 404, { error: 'Not found' })
  }
  response.statusCode = 200
  response.setHeader('Content-Type', mimeTypes[extname(filePath)] || 'application/octet-stream')
  response.setHeader(
    'Cache-Control',
    filePath.endsWith('.html') ? 'no-cache' : 'public, max-age=31536000, immutable',
  )
  const gzipPath = `${filePath}.gz`
  const servedPath =
    /(?:^|,)\s*gzip\s*(?:,|$)/i.test(acceptEncoding) && existsSync(gzipPath) ? gzipPath : filePath
  if (servedPath === gzipPath) {
    response.setHeader('Content-Encoding', 'gzip')
    response.setHeader('Vary', 'Accept-Encoding')
  }
  response.setHeader('Content-Length', statSync(servedPath).size)
  if (headOnly) {
    response.end()
    return
  }
  createReadStream(servedPath).pipe(response)
}

const server = createServer(async (request, response) => {
  applySecurityHeaders(response)
  try {
    const url = new URL(request.url || '/', 'http://localhost')
    if (url.pathname === '/favicon.ico') {
      response.statusCode = 204
      return response.end()
    }
    const pathname = requestPath(url)
    if (pathname === null) return json(response, 404, { error: 'Not found' })
    if (
      await handleMobileScannerPublicRequest({
        request,
        response,
        pathname,
        url,
        hub: mobileScannerHub,
        clientBasePath: basePath,
      })
    ) {
      return
    }
    if (request.method === 'POST' && pathname === '/api/pdf/mobile-scanner/session') {
      if (request.headers['x-genoffice-client'] !== 'pdf') {
        return json(response, 403, { error: 'Forbidden' })
      }
      try {
        const session = mobileScannerHub.createSession()
        const prefix = basePath === '/' ? '' : basePath
        return json(response, 200, {
          ...session,
          uploadPath: `${prefix}/mobile-scanner/${session.sessionId}`,
        })
      } catch (error) {
        return json(response, error instanceof MobileScannerError ? error.status : 400, {
          error: error instanceof Error ? error.message : String(error),
        })
      }
    }
    if (request.method === 'POST' && pathname === '/api/pdf/mobile-scanner/poll') {
      if (request.headers['x-genoffice-client'] !== 'pdf') {
        return json(response, 403, { error: 'Forbidden' })
      }
      const body = await readJson<{ sessionId?: unknown }>(request)
      if (!isPdfMobileScannerSessionId(body.sessionId)) throw new Error('扫码会话无效')
      const result = mobileScannerHub.takeFiles(body.sessionId)
      return json(response, 200, {
        expiresAt: result.expiresAt,
        files: result.files.map(({ bytes, ...file }) => ({
          ...file,
          base64: Buffer.from(bytes).toString('base64'),
        })),
      })
    }
    if (request.method === 'POST' && pathname === '/api/pdf/mobile-scanner/close') {
      if (request.headers['x-genoffice-client'] !== 'pdf') {
        return json(response, 403, { error: 'Forbidden' })
      }
      const body = await readJson<{ sessionId?: unknown }>(request)
      if (isPdfMobileScannerSessionId(body.sessionId)) mobileScannerHub.closeSession(body.sessionId)
      return json(response, 200, { ok: true })
    }
    if (request.method === 'POST' && pathname === '/api/ai/models') {
      return await handleModels(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/ai/chat') {
      return await handleChat(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/ai/stream') {
      return await handleStream(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/ai/image') {
      return await handleImageGeneration(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/ai/analyze-media') {
      return await handleMediaAnalysis(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/ai/web-search') {
      const input = searchInput(await readJson<SearchRequest>(request), 6)
      const result = await webSearch(input.query, input.maxResults)
      return json(
        response,
        200,
        result.results.length ? result : await wikipediaSearch(input.query, input.maxResults),
      )
    }
    if (request.method === 'POST' && pathname === '/api/ai/image-search') {
      const input = searchInput(await readJson<SearchRequest>(request), 8)
      const result = await imageSearch(input.query, input.maxResults)
      return json(
        response,
        200,
        result.images.length ? result : await wikimediaImageSearch(input.query, input.maxResults),
      )
    }
    if (
      request.method === 'POST' &&
      (pathname === '/api/slides/fetch-image' || pathname === '/api/files/fetch-image')
    ) {
      const body = await readJson<{ url?: unknown }>(request)
      return json(response, 200, await downloadRemoteImage(body.url))
    }
    if (request.method === 'POST' && pathname === '/api/attachments/text') {
      return await handleAttachmentText(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/text/validate') {
      return await handlePdfTextValidate(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/text/apply') {
      return await handlePdfTextApply(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/timestamp-token') {
      return await handlePdfTimestampToken(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/web-resource') {
      const resource = await fetchPdfWebResource(await readJson<PdfWebResourceRequest>(request))
      return json(response, 200, {
        url: resource.url,
        contentType: resource.contentType,
        base64: Buffer.from(resource.bytes).toString('base64'),
      })
    }
    if (request.method === 'GET' && pathname === '/api/pdf/fonts') {
      return json(response, 200, { fonts: listEditFonts() })
    }
    if (request.method === 'POST' && pathname === '/api/pdf/images/list') {
      return await handlePdfImagesList(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/images/render') {
      return await handlePdfImageRender(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/images/preview') {
      return await handlePdfPagePreview(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/images/apply') {
      return await handlePdfImageApply(request, response)
    }
    if (request.method === 'GET' && pathname === '/api/sheets/blank') {
      return json(response, 200, await sheetsService.blank())
    }
    if (request.method === 'POST' && pathname === '/api/sheets/open') {
      return json(response, 200, await sheetsService.open(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/range') {
      return json(response, 200, await sheetsService.readRange(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/formulas') {
      return json(response, 200, await sheetsService.readFormulas(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/recalc') {
      return json(response, 200, await sheetsService.recalc(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/media') {
      return json(response, 200, await sheetsService.readMedia(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/pivot') {
      return json(response, 200, await sheetsService.readPivot(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/save') {
      return json(response, 200, await sheetsService.save(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/recovery') {
      return json(response, 200, await sheetsService.writeRecovery(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/sheets/close') {
      const body = await readJson<{ sessionId?: unknown }>(request)
      await sheetsService.close(body.sessionId)
      return json(response, 200, { ok: true })
    }
    if (request.method === 'POST' && pathname === '/api/slides/open') {
      return json(response, 200, await slidesService.open(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/slides/blank') {
      const body = await readJson<{ fitWidthPx?: unknown }>(request)
      return json(response, 200, await slidesService.blank(Number(body.fitWidthPx) || 1280))
    }
    if (request.method === 'POST' && pathname === '/api/slides/call') {
      return json(response, 200, await slidesService.call(await readJson(request)))
    }
    if (request.method === 'POST' && pathname === '/api/slides/save') {
      const body = await readJson<{
        sessionId?: unknown
        name?: unknown
        webPath?: unknown
      }>(request)
      if (typeof body.sessionId !== 'string') throw new Error('演示文稿会话无效')
      return json(
        response,
        200,
        await slidesService.save(
          body.sessionId,
          typeof body.name === 'string' ? body.name : undefined,
          typeof body.webPath === 'string' ? body.webPath : undefined,
        ),
      )
    }
    if (request.method === 'POST' && pathname === '/api/slides/close') {
      const body = await readJson<{ sessionId?: unknown }>(request)
      slidesService.close(body.sessionId)
      return json(response, 200, { ok: true })
    }
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      response.setHeader('Allow', 'GET, HEAD, POST')
      return json(response, 405, { error: 'Method not allowed' })
    }
    serveStatic(
      pathname,
      response,
      request.method === 'HEAD',
      request.headers['accept-encoding'] || '',
    )
  } catch (error) {
    if (!response.headersSent) {
      json(response, 400, { error: error instanceof Error ? error.message : String(error) })
    } else if (!response.writableEnded) {
      response.end()
    }
  }
})

server.listen(port, '0.0.0.0', () => {
  console.log(`GenOffice Web listening on :${port} at ${basePath}`)
})
