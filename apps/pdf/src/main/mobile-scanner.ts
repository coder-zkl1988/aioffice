import { randomBytes, randomUUID } from 'node:crypto'
import { networkInterfaces } from 'node:os'
import { createServer, type IncomingMessage, type Server, type ServerResponse } from 'node:http'
import type {
  PdfMobileScannerFile,
  PdfMobileScannerPollResult,
  PdfMobileScannerSession,
} from '../shared/mobile-scanner'
import { isPdfMobileScannerSessionId } from '../shared/mobile-scanner'

const SESSION_TTL_MS = 10 * 60 * 1000
const MAX_FILE_BYTES = 20 * 1024 * 1024
const MAX_SESSION_BYTES = 80 * 1024 * 1024
const MAX_TOTAL_BYTES = 256 * 1024 * 1024
const MAX_SESSION_FILES = 40
const MAX_SESSIONS = 64

interface ScannerSessionData {
  sessionId: string
  createdAt: number
  expiresAt: number
  bytes: number
  files: PdfMobileScannerFile[]
}

export class MobileScannerError extends Error {
  constructor(
    message: string,
    readonly status = 400,
  ) {
    super(message)
  }
}

export class MobileScannerHub {
  private readonly sessions = new Map<string, ScannerSessionData>()
  private totalBytes = 0

  createSession(now = Date.now()): Omit<PdfMobileScannerSession, 'uploadUrl'> {
    this.cleanup(now)
    if (this.sessions.size >= MAX_SESSIONS) {
      throw new MobileScannerError('Too many active scanner sessions', 503)
    }
    const sessionId = randomBytes(24).toString('base64url')
    const session: ScannerSessionData = {
      sessionId,
      createdAt: now,
      expiresAt: now + SESSION_TTL_MS,
      bytes: 0,
      files: [],
    }
    this.sessions.set(sessionId, session)
    return { sessionId, expiresAt: session.expiresAt }
  }

  hasSession(sessionId: string, now = Date.now()): boolean {
    this.cleanup(now)
    return this.sessions.has(sessionId)
  }

  addFile(
    sessionId: string,
    input: { name: string; declaredType?: string; bytes: Uint8Array },
    now = Date.now(),
  ): PdfMobileScannerFile {
    const session = this.requireSession(sessionId, now)
    const size = input.bytes.byteLength
    if (size === 0) throw new MobileScannerError('The uploaded image is empty')
    if (size > MAX_FILE_BYTES) throw new MobileScannerError('Each image must be 20 MB or smaller')
    if (session.files.length >= MAX_SESSION_FILES) {
      throw new MobileScannerError(`A scanner session accepts at most ${MAX_SESSION_FILES} images`)
    }
    if (session.bytes + size > MAX_SESSION_BYTES) {
      throw new MobileScannerError('Scanner session images must total 80 MB or less')
    }
    if (this.totalBytes + size > MAX_TOTAL_BYTES) {
      throw new MobileScannerError('Scanner storage is temporarily full', 503)
    }

    const type = detectImageType(input.bytes)
    if (!type) throw new MobileScannerError('Only valid JPEG, PNG, and WebP images are accepted')
    const declaredType = input.declaredType?.split(';', 1)[0]?.trim().toLowerCase()
    if (declaredType?.startsWith('image/') && declaredType !== type) {
      throw new MobileScannerError('The uploaded image type does not match its content')
    }

    const name = safeImageName(input.name, type)
    const file: PdfMobileScannerFile = {
      id: randomUUID(),
      name: uniqueImageName(session.files, name),
      type,
      size,
      bytes: input.bytes.slice(),
    }
    session.files.push(file)
    session.bytes += size
    this.totalBytes += size
    return file
  }

  takeFiles(sessionId: string, now = Date.now()): PdfMobileScannerPollResult {
    const session = this.requireSession(sessionId, now)
    const files = session.files.splice(0)
    const releasedBytes = files.reduce((total, file) => total + file.size, 0)
    session.bytes -= releasedBytes
    this.totalBytes -= releasedBytes
    return { files, expiresAt: session.expiresAt }
  }

  closeSession(sessionId: string): void {
    const session = this.sessions.get(sessionId)
    if (!session) return
    this.totalBytes -= session.bytes
    this.sessions.delete(sessionId)
  }

  cleanup(now = Date.now()): void {
    for (const session of this.sessions.values()) {
      if (session.expiresAt <= now) this.closeSession(session.sessionId)
    }
  }

  private requireSession(sessionId: string, now: number): ScannerSessionData {
    if (!isPdfMobileScannerSessionId(sessionId)) {
      throw new MobileScannerError('Scanner session is invalid', 404)
    }
    this.cleanup(now)
    const session = this.sessions.get(sessionId)
    if (!session) throw new MobileScannerError('Scanner session was not found or has expired', 404)
    return session
  }
}

function detectImageType(bytes: Uint8Array): string | undefined {
  if (bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) {
    return 'image/jpeg'
  }
  if (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  ) {
    return 'image/png'
  }
  if (
    bytes.length >= 12 &&
    String.fromCharCode(...bytes.subarray(0, 4)) === 'RIFF' &&
    String.fromCharCode(...bytes.subarray(8, 12)) === 'WEBP'
  ) {
    return 'image/webp'
  }
  return undefined
}

function safeImageName(value: string, type: string): string {
  const fallbackExtension = type === 'image/png' ? '.png' : type === 'image/webp' ? '.webp' : '.jpg'
  const normalized = String(value || 'scan')
    .normalize('NFKC')
    .replace(/[\p{Cc}/\\:*?"<>|]/gu, '_')
    .replace(/\s+/gu, ' ')
    .trim()
    .slice(0, 120)
  const withName = normalized && !/^\.+$/u.test(normalized) ? normalized : 'scan'
  return /\.[A-Za-z0-9]{1,8}$/u.test(withName) ? withName : `${withName}${fallbackExtension}`
}

function uniqueImageName(files: PdfMobileScannerFile[], requested: string): string {
  const used = new Set(files.map((file) => file.name.toLowerCase()))
  if (!used.has(requested.toLowerCase())) return requested
  const match = /^(.*?)(\.[^.]+)?$/u.exec(requested)
  const stem = match?.[1] || 'scan'
  const extension = match?.[2] || ''
  for (let index = 2; index <= MAX_SESSION_FILES + 1; index += 1) {
    const candidate = `${stem}-${index}${extension}`
    if (!used.has(candidate.toLowerCase())) return candidate
  }
  return `${stem}-${randomUUID().slice(0, 8)}${extension}`
}

function normalizeClientBasePath(value: string): string {
  const normalized = `/${value}`.replace(/\/{2,}/gu, '/').replace(/\/$/u, '')
  return normalized === '/' ? '' : normalized
}

function htmlAttribute(value: string): string {
  return value
    .replace(/&/gu, '&amp;')
    .replace(/"/gu, '&quot;')
    .replace(/</gu, '&lt;')
    .replace(/>/gu, '&gt;')
}

export function mobileScannerUploaderHtml(sessionId: string, clientBasePath = ''): string {
  if (!isPdfMobileScannerSessionId(sessionId)) throw new MobileScannerError('Invalid session ID')
  const basePath = normalizeClientBasePath(clientBasePath)
  const uploadPath = htmlAttribute(`${basePath}/api/pdf/mobile-scanner/upload/${sessionId}`)
  const scriptPath = htmlAttribute(`${basePath}/mobile-scanner/app.js`)
  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
  <meta name="color-scheme" content="light dark">
  <title>aiOffice 手机扫描</title>
  <style>
    :root{font-family:Inter,"PingFang SC","Microsoft YaHei",sans-serif;color:#202124;background:#f5f6f7;color-scheme:light}
    *{box-sizing:border-box}body{margin:0;min-height:100vh;background:#f5f6f7}main{width:min(100%,560px);margin:0 auto;padding:28px 20px 44px}
    header{padding:0 0 24px;border-bottom:1px solid #dfe1e5}h1{margin:0;font-size:24px;font-weight:650;letter-spacing:0}p{margin:8px 0 0;color:#5f6368;line-height:1.55}
    section{padding:24px 0;border-bottom:1px solid #dfe1e5}.picker{display:flex;min-height:52px;align-items:center;justify-content:center;padding:12px 18px;border:1px solid #1a73e8;border-radius:6px;background:#1a73e8;color:#fff;font-size:16px;font-weight:600;cursor:pointer}
    .picker input{position:absolute;width:1px;height:1px;overflow:hidden;opacity:0}.status{min-height:48px;padding-top:16px;color:#3c4043;line-height:1.5}.status.error{color:#b3261e}.privacy{font-size:13px;color:#6f7378}
    progress{width:100%;height:6px;margin-top:14px;accent-color:#1a73e8}@media(prefers-color-scheme:dark){:root,body{color:#e8eaed;background:#202124}header,section{border-color:#3c4043}p,.privacy{color:#bdc1c6}.status{color:#e8eaed}}
  </style>
</head>
<body data-session-id="${sessionId}" data-upload-path="${uploadPath}">
  <main>
    <header><h1>aiOffice 手机扫描</h1><p>拍摄或选择页面，图片会直接发送到已打开的 PDF 工具。</p></header>
    <section>
      <label class="picker"><input id="files" type="file" accept="image/*" capture="environment" multiple>拍摄或选择图片</label>
      <progress id="progress" max="1" value="0" hidden></progress>
      <div id="status" class="status">等待选择图片</div>
    </section>
    <p class="privacy">图片会在此设备上转为 JPEG 并移除照片元数据；会话 10 分钟后自动失效。</p>
  </main>
  <script src="${scriptPath}" defer></script>
</body>
</html>`
}

export const mobileScannerUploaderScript = String.raw`(() => {
  const body = document.body
  const input = document.getElementById('files')
  const status = document.getElementById('status')
  const progress = document.getElementById('progress')
  const uploadPath = body.dataset.uploadPath
  let sequence = 0

  function message(text, error = false) {
    status.textContent = text
    status.classList.toggle('error', error)
  }

  function loadImage(file) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(file)
      const image = new Image()
      image.onload = () => { URL.revokeObjectURL(url); resolve(image) }
      image.onerror = () => { URL.revokeObjectURL(url); reject(new Error('无法读取这张图片')) }
      image.src = url
    })
  }

  async function normalizedImage(file) {
    const image = await loadImage(file)
    const maxEdge = 2560
    const scale = Math.min(1, maxEdge / Math.max(image.naturalWidth, image.naturalHeight))
    const canvas = document.createElement('canvas')
    canvas.width = Math.max(1, Math.round(image.naturalWidth * scale))
    canvas.height = Math.max(1, Math.round(image.naturalHeight * scale))
    const context = canvas.getContext('2d')
    if (!context) throw new Error('此浏览器无法处理图片')
    context.fillStyle = '#fff'
    context.fillRect(0, 0, canvas.width, canvas.height)
    context.drawImage(image, 0, 0, canvas.width, canvas.height)
    const blob = await new Promise((resolve, reject) => canvas.toBlob(
      value => value ? resolve(value) : reject(new Error('图片转换失败')),
      'image/jpeg',
      0.92,
    ))
    sequence += 1
    const stamp = new Date().toISOString().replace(/[:.]/g, '-').replace('T', '_').replace('Z', '')
    return { blob, name: 'mobile-scan-' + stamp + '-' + String(sequence).padStart(2, '0') + '.jpg' }
  }

  async function upload(file, index, total) {
    message('正在处理第 ' + (index + 1) + ' / ' + total + ' 张…')
    const prepared = await normalizedImage(file)
    const response = await fetch(uploadPath + '?name=' + encodeURIComponent(prepared.name), {
      method: 'POST',
      headers: { 'Content-Type': prepared.blob.type },
      body: prepared.blob,
    })
    let result = {}
    try { result = await response.json() } catch {}
    if (!response.ok) throw new Error(result.error || ('上传失败（HTTP ' + response.status + '）'))
    progress.value = (index + 1) / total
  }

  input.addEventListener('change', async () => {
    const files = Array.from(input.files || [])
    if (!files.length) return
    input.disabled = true
    progress.hidden = false
    progress.value = 0
    try {
      for (let index = 0; index < files.length; index += 1) await upload(files[index], index, files.length)
      message('已发送 ' + files.length + ' 张图片，可以继续拍摄。')
    } catch (error) {
      message(error instanceof Error ? error.message : String(error), true)
    } finally {
      input.value = ''
      input.disabled = false
    }
  })
})()`

function json(response: ServerResponse, status: number, body: unknown): void {
  response.statusCode = status
  response.setHeader('Content-Type', 'application/json; charset=utf-8')
  response.setHeader('Cache-Control', 'no-store')
  response.end(JSON.stringify(body))
}

function scannerHeaders(response: ServerResponse): void {
  response.setHeader('X-Content-Type-Options', 'nosniff')
  response.setHeader('Referrer-Policy', 'no-referrer')
  response.setHeader('Cache-Control', 'no-store')
  response.setHeader('Permissions-Policy', 'camera=(self), microphone=(), geolocation=()')
  response.setHeader(
    'Content-Security-Policy',
    "default-src 'self'; script-src 'self'; style-src 'unsafe-inline'; img-src 'self' blob: data:; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'; form-action 'self'",
  )
}

async function readUploadBytes(request: IncomingMessage): Promise<Uint8Array> {
  const declaredLength = Number(request.headers['content-length'] || 0)
  if (Number.isFinite(declaredLength) && declaredLength > MAX_FILE_BYTES) {
    throw new MobileScannerError('Each image must be 20 MB or smaller', 413)
  }
  const chunks: Buffer[] = []
  let size = 0
  for await (const chunk of request) {
    const bytes = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)
    size += bytes.length
    if (size > MAX_FILE_BYTES)
      throw new MobileScannerError('Each image must be 20 MB or smaller', 413)
    chunks.push(bytes)
  }
  return new Uint8Array(Buffer.concat(chunks))
}

export async function handleMobileScannerPublicRequest(options: {
  request: IncomingMessage
  response: ServerResponse
  pathname: string
  url: URL
  hub: MobileScannerHub
  clientBasePath?: string
}): Promise<boolean> {
  const { request, response, pathname, url, hub, clientBasePath = '' } = options
  if (request.method === 'GET' && pathname === '/mobile-scanner/app.js') {
    scannerHeaders(response)
    response.statusCode = 200
    response.setHeader('Content-Type', 'text/javascript; charset=utf-8')
    response.end(mobileScannerUploaderScript)
    return true
  }

  const pageMatch = /^\/mobile-scanner\/([A-Za-z0-9_-]{24,64})$/u.exec(pathname)
  if (request.method === 'GET' && pageMatch) {
    scannerHeaders(response)
    if (!hub.hasSession(pageMatch[1]!)) {
      json(response, 404, { error: 'Scanner session was not found or has expired' })
      return true
    }
    response.statusCode = 200
    response.setHeader('Content-Type', 'text/html; charset=utf-8')
    response.end(mobileScannerUploaderHtml(pageMatch[1]!, clientBasePath))
    return true
  }

  const uploadMatch = /^\/api\/pdf\/mobile-scanner\/upload\/([A-Za-z0-9_-]{24,64})$/u.exec(pathname)
  if (request.method === 'POST' && uploadMatch) {
    scannerHeaders(response)
    try {
      const file = hub.addFile(uploadMatch[1]!, {
        name: url.searchParams.get('name') || 'mobile-scan.jpg',
        declaredType: request.headers['content-type'],
        bytes: await readUploadBytes(request),
      })
      json(response, 200, { ok: true, file: { id: file.id, name: file.name, size: file.size } })
    } catch (error) {
      const status = error instanceof MobileScannerError ? error.status : 400
      json(response, status, { error: error instanceof Error ? error.message : String(error) })
    }
    return true
  }
  return false
}

function preferredLanAddress(): string | undefined {
  const candidates = Object.entries(networkInterfaces()).flatMap(([name, entries]) =>
    (entries || [])
      .filter((entry) => entry.family === 'IPv4' && !entry.internal)
      .map((entry) => ({ name, address: entry.address })),
  )
  const interfaceRank = (name: string): number =>
    /^(en0|eth0|wlan0|wi-fi)$/iu.test(name) ? 0 : /^(en|eth|wlan)/iu.test(name) ? 1 : 2
  const addressRank = (address: string): number =>
    /^(192\.168\.|10\.|172\.(1[6-9]|2\d|3[01])\.)/u.test(address) ? 0 : 1
  candidates.sort(
    (left, right) =>
      addressRank(left.address) - addressRank(right.address) ||
      interfaceRank(left.name) - interfaceRank(right.name) ||
      left.name.localeCompare(right.name),
  )
  return candidates[0]?.address
}

export class DesktopMobileScannerService {
  private readonly hub = new MobileScannerHub()
  private server?: Server
  private port?: number

  async createSession(): Promise<PdfMobileScannerSession> {
    const host = preferredLanAddress()
    if (!host) throw new MobileScannerError('No LAN address is available for phone scanning')
    await this.ensureServer()
    const session = this.hub.createSession()
    return {
      ...session,
      uploadUrl: `http://${host}:${this.port}/mobile-scanner/${session.sessionId}`,
    }
  }

  pollSession(sessionId: string): PdfMobileScannerPollResult {
    return this.hub.takeFiles(sessionId)
  }

  closeSession(sessionId: string): void {
    this.hub.closeSession(sessionId)
  }

  close(): void {
    this.server?.close()
    this.server = undefined
    this.port = undefined
  }

  private async ensureServer(): Promise<void> {
    if (this.server && this.port) return
    const server = createServer(async (request, response) => {
      try {
        const url = new URL(request.url || '/', 'http://localhost')
        if (
          await handleMobileScannerPublicRequest({
            request,
            response,
            pathname: url.pathname,
            url,
            hub: this.hub,
          })
        ) {
          return
        }
        json(response, 404, { error: 'Not found' })
      } catch (error) {
        if (!response.headersSent) {
          json(response, error instanceof MobileScannerError ? error.status : 400, {
            error: error instanceof Error ? error.message : String(error),
          })
        } else if (!response.writableEnded) {
          response.end()
        }
      }
    })
    await new Promise<void>((resolve, reject) => {
      server.once('error', reject)
      server.listen(0, '0.0.0', () => {
        server.off('error', reject)
        resolve()
      })
    })
    const address = server.address()
    if (!address || typeof address === 'string') {
      server.close()
      throw new MobileScannerError('Could not start the phone scanner server')
    }
    server.unref()
    this.server = server
    this.port = address.port
  }
}
