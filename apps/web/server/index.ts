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
  type AiStreamChunk,
  type AiStreamRequest,
} from '@genoffice/ai-provider'
import {
  applyImageEdits,
  listPageImages,
  renderImagePng,
  renderPagePreviewPng,
} from '../../pdf/src/main/image-edit'
import { applyTextEdits, listEditFonts, validateTextEdits } from '../../pdf/src/main/text-edit'
import type { ImageEditInput, PagePreviewRequest, TextEditInput } from '../../pdf/src/shared/ipc'
import { validateProviderBaseUrl } from './security'

const serverDirectory = fileURLToPath(new URL('.', import.meta.url))
const staticRoot = resolve(serverDirectory, '../dist')
const port = Number(process.env.PORT || 80)
const basePath = normalizeBasePath(process.env.WEB_BASE_PATH || '/')
const maxRequestBytes = Number(process.env.WEB_MAX_REQUEST_BYTES || 128 * 1024 * 1024)

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
  response.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=()')
  response.setHeader(
    'Content-Security-Policy',
    "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; font-src 'self' data:; connect-src 'self'; worker-src 'self' blob:; frame-src 'self' blob:; object-src 'none'; base-uri 'self'; form-action 'self'",
  )
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

async function customConfig(settings: AiStreamRequest['settings']): Promise<AiProviderConfig> {
  if (settings?.provider !== 'custom') throw new Error('Web 版请先配置自定义模型')
  const source = settings.providers?.custom
  if (!source?.apiKey || source.apiKey.length > 8192) throw new Error('请填写 API Key')
  if (!source.model || source.model.length > 256) throw new Error('请填写模型名称')
  return {
    apiKey: source.apiKey,
    model: source.model,
    baseUrl: await validateProviderBaseUrl(source.baseUrl),
  }
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
  } catch (error) {
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
    response.end()
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

function serveStatic(pathname: string, response: ServerResponse, headOnly = false): void {
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
  if (headOnly) {
    response.end()
    return
  }
  createReadStream(filePath).pipe(response)
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
    if (request.method === 'POST' && pathname === '/api/ai/chat') {
      return await handleChat(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/ai/stream') {
      return await handleStream(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/text/validate') {
      return await handlePdfTextValidate(request, response)
    }
    if (request.method === 'POST' && pathname === '/api/pdf/text/apply') {
      return await handlePdfTextApply(request, response)
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
    if (request.method !== 'GET' && request.method !== 'HEAD') {
      response.setHeader('Allow', 'GET, HEAD, POST')
      return json(response, 405, { error: 'Method not allowed' })
    }
    serveStatic(pathname, response, request.method === 'HEAD')
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
