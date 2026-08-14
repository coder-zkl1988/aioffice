import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseEpubDocument } from '../src/renderer/epub-to-pdf'

const options = {
  embedAllFonts: false,
  includeTableOfContents: true,
  includePageNumbers: true,
  optimizeForEbook: false,
  tableOfContentsLabel: 'Contents',
  untitledLabel: 'Untitled chapter',
}

async function epubBytes(): Promise<Uint8Array> {
  const archive = new JSZip()
  archive.file('mimetype', 'application/epub+zip', { compression: 'STORE' })
  archive.file(
    'META-INF/container.xml',
    '<?xml version="1.0"?><container xmlns="urn:oasis:names:tc:opendocument:xmlns:container" version="1.0"><rootfiles><rootfile full-path="OPS/package.opf" media-type="application/oebps-package+xml"/></rootfiles></container>',
  )
  archive.file(
    'OPS/package.opf',
    `<?xml version="1.0"?><package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>Local EPUB</dc:title><dc:creator>Alice Author</dc:creator><dc:language>zh-CN</dc:language></metadata><manifest><item id="chapter-two" href="text/chapter-2.xhtml" media-type="application/xhtml+xml"/><item id="chapter-one" href="text/chapter-1.xhtml" media-type="application/xhtml+xml"/><item id="cover" href="images/cover.png" media-type="image/png"/></manifest><spine><itemref idref="chapter-one"/><itemref idref="chapter-two"/></spine></package>`,
  )
  archive.file(
    'OPS/text/chapter-1.xhtml',
    `<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>First</title><script>fetch('https://tracker.invalid/script')</script></head><body><h1>First chapter</h1><p style="color:red;position:fixed;background:url(https://tracker.invalid/bg)">Safe text</p><img src="../images/cover.png" alt="Cover"/><img src="https://tracker.invalid/pixel"/><iframe src="https://tracker.invalid/frame"></iframe></body></html>`,
  )
  archive.file(
    'OPS/text/chapter-2.xhtml',
    '<?xml version="1.0"?><html xmlns="http://www.w3.org/1999/xhtml"><head><title>Second</title></head><body><h1>Second chapter</h1><p>Ending.</p></body></html>',
  )
  archive.file(
    'OPS/images/cover.png',
    Uint8Array.from(
      Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z2S8AAAAASUVORK5CYII=',
        'base64',
      ),
    ),
  )
  return archive.generateAsync({ type: 'uint8array' })
}

describe('EPUB document parsing', () => {
  it('uses package metadata and spine reading order', async () => {
    const parsed = await parseEpubDocument(await epubBytes(), options)

    expect(parsed).toMatchObject({
      title: 'Local EPUB',
      author: 'Alice Author',
      chapterCount: 2,
    })
    expect(parsed.html.indexOf('First chapter')).toBeLessThan(parsed.html.indexOf('Second chapter'))
    expect(parsed.html).toContain('<h1>Contents</h1>')
    expect(parsed.html).toContain('<li>First</li>')
    expect(parsed.html).toContain('data:image/png;base64,')
  })

  it('removes active and remote content before rendering', async () => {
    const parsed = await parseEpubDocument(await epubBytes(), options)

    expect(parsed.html).toContain('color:red')
    expect(parsed.html).not.toMatch(/script|iframe|position:fixed|background:url/iu)
    expect(parsed.html).not.toContain('tracker.invalid')
    expect(parsed.html).toContain("default-src 'none'")
  })

  it('omits table of contents when disabled', async () => {
    const parsed = await parseEpubDocument(await epubBytes(), {
      ...options,
      includeTableOfContents: false,
    })

    expect(parsed.html).not.toContain('<section class="ebook-toc">')
    expect(parsed.html).toContain('First chapter')
  })

  it('rejects empty, invalid and incomplete EPUB archives', async () => {
    await expect(parseEpubDocument(new Uint8Array(), options)).rejects.toThrow('empty')
    await expect(parseEpubDocument(new Uint8Array([1, 2, 3]), options)).rejects.toThrow('invalid')

    const archive = new JSZip()
    archive.file('mimetype', 'application/epub+zip')
    await expect(
      parseEpubDocument(await archive.generateAsync({ type: 'uint8array' }), options),
    ).rejects.toThrow('container')
  })
})
