import {
  defaultAiSettings,
  resolveAiSettings,
  type AiChatRequest,
  type AiChatResponse,
  type AiSettings,
  type AiStreamChunk,
  type AiStreamRequest,
} from '@genoffice/ai-provider'

export const AI_SETTINGS_STORAGE_KEY = 'genoffice.web.ai-settings'

const listeners = new Set<(chunk: AiStreamChunk) => void>()
const streams = new Map<string, AbortController>()

function apiUrl(action: 'chat' | 'stream'): URL {
  return new URL(`./api/ai/${action}`, document.baseURI)
}

async function responseError(response: Response): Promise<string> {
  try {
    const body = (await response.json()) as { error?: string }
    return body.error || `HTTP ${response.status}`
  } catch {
    return `HTTP ${response.status}`
  }
}

function emit(chunk: AiStreamChunk): void {
  for (const listener of listeners) listener(chunk)
}

export function getWebAiSettings(): AiSettings {
  const defaults = defaultAiSettings()
  const stored = localStorage.getItem(AI_SETTINGS_STORAGE_KEY)
  if (!stored) return defaults
  try {
    return resolveAiSettings(JSON.parse(stored) as Partial<AiSettings>, defaults)
  } catch {
    return defaults
  }
}

export function saveWebAiSettings(settings: AiSettings): void {
  localStorage.setItem(AI_SETTINGS_STORAGE_KEY, JSON.stringify(settings))
}

export async function webAiChat(request: AiChatRequest): Promise<AiChatResponse> {
  try {
    const response = await fetch(apiUrl('chat'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(request),
    })
    if (!response.ok) return { ok: false, error: await responseError(response) }
    return (await response.json()) as AiChatResponse
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

export async function webAiStream(request: AiStreamRequest): Promise<void> {
  const controller = new AbortController()
  streams.get(request.requestId)?.abort()
  streams.set(request.requestId, controller)
  try {
    const response = await fetch(apiUrl('stream'), {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(request),
      signal: controller.signal,
    })
    if (!response.ok || !response.body) throw new Error(await responseError(response))

    const reader = response.body.getReader()
    const decoder = new TextDecoder()
    let buffer = ''
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      buffer += decoder.decode(value, { stream: true })
      const lines = buffer.split('\n')
      buffer = lines.pop() || ''
      for (const line of lines) {
        if (line.trim()) emit(JSON.parse(line) as AiStreamChunk)
      }
    }
    if (buffer.trim()) emit(JSON.parse(buffer) as AiStreamChunk)
  } catch (error) {
    emit(
      controller.signal.aborted
        ? { requestId: request.requestId, type: 'done' }
        : {
            requestId: request.requestId,
            type: 'error',
            error: error instanceof Error ? error.message : String(error),
          },
    )
  } finally {
    if (streams.get(request.requestId) === controller) streams.delete(request.requestId)
  }
}

export async function cancelWebAiStream(requestId: string): Promise<void> {
  streams.get(requestId)?.abort()
}

export function onWebAiStream(handler: (chunk: AiStreamChunk) => void): () => void {
  listeners.add(handler)
  return () => listeners.delete(handler)
}

export async function testWebAiConnection(settings: AiSettings): Promise<AiChatResponse> {
  return webAiChat({
    settings,
    system: 'Reply with exactly OK.',
    user: 'Connection test',
  })
}
