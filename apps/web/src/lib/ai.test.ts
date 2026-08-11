// @vitest-environment jsdom

import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { hasConfiguredWebAi, webAiChat } from './ai'

function configuredSettings(): AiSettings {
  const settings = defaultAiSettings()
  settings.provider = 'custom'
  settings.providers.custom = {
    apiKey: 'test-key',
    model: 'test-model',
    baseUrl: 'https://ai.example.com/v1',
  }
  return settings
}

afterEach(() => vi.unstubAllGlobals())

describe('hasConfiguredWebAi', () => {
  it('enables Web AI capabilities only for a complete custom provider configuration', () => {
    const settings = defaultAiSettings()
    expect(hasConfiguredWebAi(settings)).toBe(false)

    settings.provider = 'custom'
    settings.providers.custom = {
      apiKey: 'test-key',
      model: 'test-model',
      baseUrl: 'https://ai.example.com/v1',
    }
    expect(hasConfiguredWebAi(settings)).toBe(true)

    settings.providers.custom.model = ''
    expect(hasConfiguredWebAi(settings)).toBe(false)
  })

  it('uses the keepalive stream endpoint for one-shot chat requests', async () => {
    const fetchMock = vi
      .fn()
      .mockResolvedValue(
        new Response(
          [
            JSON.stringify({ requestId: 'ignored', type: 'delta', text: 'wrong' }),
            JSON.stringify({ requestId: 'request-1', type: 'ping' }),
            JSON.stringify({ requestId: 'request-1', type: 'delta', text: 'slide ' }),
            JSON.stringify({ requestId: 'request-1', type: 'delta', text: 'content' }),
            JSON.stringify({ requestId: 'request-1', type: 'done' }),
          ].join('\n'),
          { status: 200 },
        ),
      )
    vi.stubGlobal('crypto', { randomUUID: () => 'request-1' })
    vi.stubGlobal('fetch', fetchMock)

    const result = await webAiChat({
      settings: configuredSettings(),
      system: 'system',
      user: 'user',
    })

    expect(result).toEqual({ ok: true, content: 'slide content' })
    const [url, init] = fetchMock.mock.calls[0]!
    expect(String(url)).toContain('/api/ai/stream')
    expect(JSON.parse(String(init.body))).toMatchObject({
      requestId: 'request-1',
      messages: [{ role: 'user', text: 'user' }],
      tools: [],
    })
  })

  it('surfaces the iSpace proxy error message instead of a bare HTTP status', async () => {
    vi.stubGlobal('crypto', { randomUUID: () => 'request-2' })
    vi.stubGlobal(
      'fetch',
      vi.fn().mockResolvedValue(
        new Response(JSON.stringify({ code: 'UPSTREAM_ERROR', message: '后端服务无响应。' }), {
          status: 502,
          headers: { 'Content-Type': 'application/json' },
        }),
      ),
    )

    await expect(
      webAiChat({ settings: configuredSettings(), system: 'system', user: 'user' }),
    ).resolves.toEqual({ ok: false, error: '后端服务无响应。' })
  })
})
