import { isAllowedPdfTimestampTsaUrl } from '@genoffice/pdf-tools'

const MAX_TIMESTAMP_REQUEST_BYTES = 64 * 1024
const MAX_TIMESTAMP_RESPONSE_BYTES = 1024 * 1024

async function limitedResponseBytes(response: Response): Promise<Uint8Array> {
  const declaredLength = Number(response.headers.get('content-length'))
  if (Number.isFinite(declaredLength) && declaredLength > MAX_TIMESTAMP_RESPONSE_BYTES) {
    throw new Error('The TSA response is too large')
  }
  if (!response.body) throw new Error('The TSA returned an empty response')
  const reader = response.body.getReader()
  const chunks: Uint8Array[] = []
  let total = 0
  while (true) {
    const { value, done } = await reader.read()
    if (done) break
    total += value.length
    if (total > MAX_TIMESTAMP_RESPONSE_BYTES) {
      await reader.cancel()
      throw new Error('The TSA response is too large')
    }
    chunks.push(value)
  }
  if (total === 0) throw new Error('The TSA returned an empty response')
  const output = new Uint8Array(total)
  let offset = 0
  for (const chunk of chunks) {
    output.set(chunk, offset)
    offset += chunk.length
  }
  return output
}

export async function requestPdfTimestampToken(
  tsaUrl: string,
  requestBytes: Uint8Array,
): Promise<Uint8Array> {
  if (!isAllowedPdfTimestampTsaUrl(tsaUrl)) throw new Error('The TSA URL is not allowed')
  if (requestBytes.length === 0 || requestBytes.length > MAX_TIMESTAMP_REQUEST_BYTES) {
    throw new Error('The timestamp request is empty or too large')
  }
  const response = await fetch(tsaUrl, {
    method: 'POST',
    headers: {
      Accept: 'application/timestamp-reply',
      'Content-Type': 'application/timestamp-query',
    },
    body: requestBytes.buffer.slice(
      requestBytes.byteOffset,
      requestBytes.byteOffset + requestBytes.byteLength,
    ) as ArrayBuffer,
    redirect: 'error',
    signal: AbortSignal.timeout(30_000),
  })
  if (!response.ok) throw new Error(`The TSA returned HTTP ${response.status}`)
  return limitedResponseBytes(response)
}
