import { describe, expect, it, vi } from 'vitest'
import { fetchPdfWebResource } from '../src/main/web-resource'

describe('PDF remote web resource fetcher', () => {
  it('returns validated HTML bytes and the final redirect URL', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(null, {
          status: 302,
          headers: { location: 'https://1.1.1.1/article/index.html' },
        }),
      )
      .mockResolvedValueOnce(
        new Response('<h1>Article</h1>', {
          status: 200,
          headers: { 'content-type': 'text/html; charset=utf-8' },
        }),
      )

    const result = await fetchPdfWebResource(
      { url: 'https://8.8.8.8/start', kind: 'document' },
      { fetchImpl },
    )

    expect(result.url).toBe('https://1.1.1.1/article/index.html')
    expect(result.contentType).toBe('text/html; charset=utf-8')
    expect(new TextDecoder().decode(result.bytes)).toContain('Article')
  })

  it('rejects private targets, credentials, mismatched types, and oversized resources', async () => {
    await expect(
      fetchPdfWebResource(
        { url: 'http://127.0.0.1/secret', kind: 'document' },
        { fetchImpl: vi.fn() },
      ),
    ).rejects.toThrow('public')
    await expect(
      fetchPdfWebResource({ url: 'https://user:pass@8.8.8.8/', kind: 'document' }),
    ).rejects.toThrow('credentials')
    await expect(
      fetchPdfWebResource(
        { url: 'https://8.8.8.8/file.pdf', kind: 'document' },
        {
          fetchImpl: vi.fn().mockResolvedValue(
            new Response('pdf', {
              status: 200,
              headers: { 'content-type': 'application/pdf' },
            }),
          ),
        },
      ),
    ).rejects.toThrow('supported document')
    await expect(
      fetchPdfWebResource(
        { url: 'https://8.8.8.8/huge.css', kind: 'stylesheet' },
        {
          fetchImpl: vi.fn().mockResolvedValue(
            new Response('body{}', {
              status: 200,
              headers: {
                'content-type': 'text/css',
                'content-length': String(6 * 1024 * 1024),
              },
            }),
          ),
        },
      ),
    ).rejects.toThrow('too large')
  })
})
