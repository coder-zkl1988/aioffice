import { afterEach, describe, expect, it, vi } from 'vitest'
import { requestPdfTimestampToken } from '../src/main/timestamp'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('PDF timestamp proxy', () => {
  it('posts timestamp requests with the RFC 3161 media types', async () => {
    const fetchMock = vi.fn(async (_url: string, init?: RequestInit) => {
      expect(init).toMatchObject({
        method: 'POST',
        redirect: 'error',
        headers: {
          Accept: 'application/timestamp-reply',
          'Content-Type': 'application/timestamp-query',
        },
      })
      expect(new Uint8Array(init?.body as ArrayBuffer)).toEqual(new Uint8Array([1, 2, 3]))
      return new Response(new Uint8Array([4, 5, 6]))
    })
    vi.stubGlobal('fetch', fetchMock)

    await expect(
      requestPdfTimestampToken('http://timestamp.digicert.com', new Uint8Array([1, 2, 3])),
    ).resolves.toEqual(new Uint8Array([4, 5, 6]))
    expect(fetchMock).toHaveBeenCalledOnce()
  })

  it('rejects unknown authorities before making a network request', async () => {
    const fetchMock = vi.fn()
    vi.stubGlobal('fetch', fetchMock)

    await expect(
      requestPdfTimestampToken('https://example.com/tsa', new Uint8Array([1])),
    ).rejects.toThrow('not allowed')
    expect(fetchMock).not.toHaveBeenCalled()
  })

  it('rejects empty, oversized, and oversized-response requests', async () => {
    const fetchMock = vi.fn(
      async () =>
        new Response(new Uint8Array([1]), {
          headers: { 'Content-Length': String(1024 * 1024 + 1) },
        }),
    )
    vi.stubGlobal('fetch', fetchMock)

    await expect(
      requestPdfTimestampToken('http://timestamp.digicert.com', new Uint8Array()),
    ).rejects.toThrow('empty or too large')
    await expect(
      requestPdfTimestampToken('http://timestamp.digicert.com', new Uint8Array(64 * 1024 + 1)),
    ).rejects.toThrow('empty or too large')
    await expect(
      requestPdfTimestampToken('http://timestamp.digicert.com', new Uint8Array([1])),
    ).rejects.toThrow('response is too large')
  })
})
