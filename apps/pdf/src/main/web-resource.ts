import { fetchWithSsrfGuardResult, type FetchWithSsrfGuardOptions } from '@genoffice/electron-utils'
import type { PdfWebResourceKind, PdfWebResourceRequest, PdfWebResourceResult } from '../shared/ipc'

const MAX_RESOURCE_BYTES: Record<PdfWebResourceKind, number> = {
  document: 20 * 1024 * 1024,
  stylesheet: 5 * 1024 * 1024,
  image: 20 * 1024 * 1024,
}

const ACCEPT_HEADERS: Record<PdfWebResourceKind, string> = {
  document: 'text/html,application/xhtml+xml;q=0.9',
  stylesheet: 'text/css,text/plain;q=0.5',
  image: 'image/png,image/jpeg,image/gif,image/webp,image/bmp,image/svg+xml;q=0.9',
}

const ALLOWED_CONTENT_TYPES: Record<PdfWebResourceKind, ReadonlySet<string>> = {
  document: new Set(['text/html', 'application/xhtml+xml']),
  stylesheet: new Set(['text/css', 'text/plain']),
  image: new Set([
    'image/png',
    'image/jpeg',
    'image/gif',
    'image/webp',
    'image/bmp',
    'image/svg+xml',
  ]),
}

export interface FetchPdfWebResourceOptions {
  fetchImpl?: typeof fetch
  signal?: AbortSignal
}

async function limitedResponseBytes(response: Response, limit: number): Promise<Uint8Array> {
  const declared = Number(response.headers.get('content-length') || 0)
  if (Number.isFinite(declared) && declared > limit)
    throw new Error('Remote web resource is too large')
  if (!response.body) throw new Error('Remote web resource has no content')
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    total += value.byteLength
    if (total > limit) {
      await reader.cancel()
      throw new Error('Remote web resource is too large')
    }
    chunks.push(value)
  }
  if (total === 0) throw new Error('Remote web resource is empty')
  const output = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    output.set(chunk, offset)
    offset += chunk.byteLength
  }
  return output
}

export async function fetchPdfWebResource(
  request: PdfWebResourceRequest,
  options: FetchPdfWebResourceOptions = {},
): Promise<PdfWebResourceResult> {
  if (
    !request ||
    typeof request.url !== 'string' ||
    request.url.length > 4096 ||
    !['document', 'stylesheet', 'image'].includes(request.kind)
  ) {
    throw new Error('Remote web resource request is invalid')
  }
  let parsed: URL
  try {
    parsed = new URL(request.url.trim())
  } catch {
    throw new Error('Website URL is invalid')
  }
  if (parsed.username || parsed.password) throw new Error('Website URL cannot contain credentials')
  parsed.hash = ''

  const guardOptions: FetchWithSsrfGuardOptions = {
    headers: {
      Accept: ACCEPT_HEADERS[request.kind],
      'User-Agent': 'Mozilla/5.0 (compatible; GenOffice-PDF/1.0)',
    },
    signal: options.signal ?? AbortSignal.timeout(20_000),
    ...(options.fetchImpl ? { fetchImpl: options.fetchImpl } : {}),
  }
  const fetched = await fetchWithSsrfGuardResult(parsed.toString(), guardOptions)
  if (!fetched) throw new Error('Website URL must point to a public HTTP or HTTPS resource')
  if (!fetched.response.ok)
    throw new Error(`Remote web resource returned HTTP ${fetched.response.status}`)
  const contentTypeHeader = fetched.response.headers.get('content-type')?.trim().toLowerCase() || ''
  const mediaType = contentTypeHeader.split(';')[0]?.trim() || ''
  if (!ALLOWED_CONTENT_TYPES[request.kind].has(mediaType)) {
    throw new Error(`Remote resource is not a supported ${request.kind}`)
  }
  return {
    url: fetched.url,
    contentType: contentTypeHeader,
    bytes: await limitedResponseBytes(fetched.response, MAX_RESOURCE_BYTES[request.kind]),
  }
}
