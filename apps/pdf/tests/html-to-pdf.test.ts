import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import type { PdfWebResourceRequest, PdfWebResourceResult } from '../src/shared/ipc'
import { parseLocalWebDocument, parseRemoteWebDocument } from '../src/renderer/html-to-pdf'

const options = {
  includePageNumbers: true,
  zoom: 1,
  untitledLabel: 'Untitled web document',
}

const onePixelPng = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z2S8AAAAASUVORK5CYII=',
    'base64',
  ),
)

async function webpageArchive(): Promise<Uint8Array> {
  const archive = new JSZip()
  archive.file('docs/other.html', '<h1>Other page</h1>')
  archive.file(
    'site/index.html',
    `<!doctype html><html lang="zh-CN"><head><title>Local website</title><link rel="stylesheet" href="assets/site.css"><script>fetch('https://tracker.invalid/script')</script></head><body><main class="report"><h1>Local website</h1><p id="summary">Printable text</p><img src="assets/cover.png" alt="Cover"><img src="https://tracker.invalid/pixel"><iframe src="https://tracker.invalid/frame"></iframe></main></body></html>`,
  )
  archive.file(
    'site/assets/site.css',
    `.report { color: rgb(18, 52, 86); display: grid; gap: 12px; background-image: url(https://tracker.invalid/bg); position: fixed; } #summary { font-weight: 700; } @import url(https://tracker.invalid/import.css);`,
  )
  archive.file('site/assets/cover.png', onePixelPng)
  return archive.generateAsync({ type: 'uint8array' })
}

describe('HTML and ZIP document parsing', () => {
  it('selects the shallowest index page and embeds local resources', async () => {
    const parsed = await parseLocalWebDocument(await webpageArchive(), 'website.zip', options)

    expect(parsed).toMatchObject({
      title: 'Local website',
      entryPath: 'site/index.html',
      resourceCount: 1,
    })
    expect(parsed.html).toContain('Printable text')
    expect(parsed.html).toContain('data:image/png;base64,')
    expect(parsed.html).toContain('.report{')
    expect(parsed.html).toContain('font-weight:700')
  })

  it('removes remote, active and dangerous style content', async () => {
    const parsed = await parseLocalWebDocument(await webpageArchive(), 'website.zip', options)

    expect(parsed.html).not.toMatch(
      /script|iframe|tracker\.invalid|background-image|position:fixed/iu,
    )
    expect(parsed.html).toContain("default-src 'none'")
  })

  it('supports standalone HTML and safe data images', async () => {
    const html = `<!doctype html><html><head><title>Standalone</title></head><body><h1>Standalone</h1><img src="data:image/png;base64,${Buffer.from(onePixelPng).toString('base64')}"><form><input value="secret"></form></body></html>`
    const parsed = await parseLocalWebDocument(new TextEncoder().encode(html), 'page.htm', {
      ...options,
      zoom: 1.25,
    })

    expect(parsed.title).toBe('Standalone')
    expect(parsed.entryPath).toBeUndefined()
    expect(parsed.html).toContain('zoom:1.25')
    expect(parsed.html).toContain('data:image/png;base64,')
    expect(parsed.html).not.toMatch(/<form|<input|secret/iu)
  })

  it('rejects invalid inputs and archives without HTML', async () => {
    await expect(parseLocalWebDocument(new Uint8Array(), 'empty.html', options)).rejects.toThrow(
      'empty',
    )
    await expect(
      parseLocalWebDocument(new Uint8Array([1, 2, 3]), 'broken.zip', options),
    ).rejects.toThrow('invalid')
    const archive = new JSZip()
    archive.file('readme.txt', 'nothing to print')
    await expect(
      parseLocalWebDocument(
        await archive.generateAsync({ type: 'uint8array' }),
        'no-html.zip',
        options,
      ),
    ).rejects.toThrow('no HTML')
    await expect(
      parseLocalWebDocument(new TextEncoder().encode('<h1>Test</h1>'), 'page.html', {
        ...options,
        zoom: 3,
      }),
    ).rejects.toThrow('zoom')
  })
})

describe('remote website parsing', () => {
  it('resolves remote styles and images from the final document URL', async () => {
    const requests: PdfWebResourceRequest[] = []
    const resources = new Map<string, { contentType: string; body: string | Uint8Array }>([
      [
        'https://example.com/articles/start',
        {
          contentType: 'text/html; charset=utf-8',
          body: `<!doctype html><html lang="zh-CN"><head><title>Remote article</title><base href="/assets/"><link rel="stylesheet" href="print.css"><script>alert(1)</script></head><body><main class="article"><h1>Remote article</h1><p>Printable text</p><img src="cover.png"><iframe src="https://tracker.invalid"></iframe></main></body></html>`,
        },
      ],
      [
        'https://example.com/assets/print.css',
        {
          contentType: 'text/css',
          body: '.article { color: rgb(10, 20, 30); background-image: url(secret.png); }',
        },
      ],
      ['https://example.com/assets/cover.png', { contentType: 'image/png', body: onePixelPng }],
    ])
    const loadResource = async (request: PdfWebResourceRequest): Promise<PdfWebResourceResult> => {
      requests.push(request)
      const resource = resources.get(request.url)
      if (!resource) throw new Error('missing fixture')
      return {
        url: request.url,
        contentType: resource.contentType,
        bytes:
          typeof resource.body === 'string'
            ? new TextEncoder().encode(resource.body)
            : resource.body,
      }
    }

    const parsed = await parseRemoteWebDocument('example.com/articles/start', loadResource, options)

    expect(parsed).toMatchObject({
      title: 'Remote article',
      url: 'https://example.com/articles/start',
      resourceCount: 2,
    })
    expect(requests).toContainEqual({
      url: 'https://example.com/assets/print.css',
      kind: 'stylesheet',
    })
    expect(requests).toContainEqual({
      url: 'https://example.com/assets/cover.png',
      kind: 'image',
    })
    expect(parsed.html).toContain('Printable text')
    expect(parsed.html).toContain('data:image/png;base64,')
    expect(parsed.html).toContain('.article{color:rgb(10, 20, 30)}')
    expect(parsed.html).not.toMatch(/script|iframe|tracker\.invalid|background-image/iu)
  })

  it('keeps converting when an optional remote resource fails', async () => {
    const parsed = await parseRemoteWebDocument(
      'https://example.com/',
      async ({ url, kind }) => {
        if (kind !== 'document') throw new Error('offline')
        return {
          url,
          contentType: 'text/html',
          bytes: new TextEncoder().encode(
            '<title>Fallback</title><p>Text stays</p><img src="missing.png">',
          ),
        }
      },
      options,
    )

    expect(parsed.title).toBe('Fallback')
    expect(parsed.html).toContain('Text stays')
    expect(parsed.html).not.toContain('<img')
  })
})
