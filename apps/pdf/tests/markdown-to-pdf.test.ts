import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseMarkdownDocument } from '../src/renderer/markdown-to-pdf'

const options = {
  includePageNumbers: true,
  zoom: 1,
  untitledLabel: 'Untitled Markdown document',
}

const onePixelPng = Uint8Array.from(
  Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Z2S8AAAAASUVORK5CYII=',
    'base64',
  ),
)

async function markdownArchive(): Promise<Uint8Array> {
  const archive = new JSZip()
  archive.file('notes/other.md', '# Other page')
  archive.file(
    'guide/index.md',
    `# Local Markdown

This is **bold** and [linked](https://example.com).

| Feature | Status |
| --- | --- |
| GFM table | Ready |

![Local cover](images/cover.png)

\`\`\`ts
const answer = 42
\`\`\`

<script>fetch('https://tracker.invalid/script')</script>
<img src="https://tracker.invalid/pixel">
`,
  )
  archive.file('guide/images/cover.png', onePixelPng)
  return archive.generateAsync({ type: 'uint8array' })
}

describe('Markdown document parsing', () => {
  it('renders GFM structure and embeds ZIP images', async () => {
    const parsed = await parseMarkdownDocument(await markdownArchive(), 'guide.zip', options)

    expect(parsed).toMatchObject({
      title: 'Local Markdown',
      entryPath: 'guide/index.md',
      resourceCount: 1,
    })
    expect(parsed.html).toContain('<table>')
    expect(parsed.html).toContain('<strong>bold</strong>')
    expect(parsed.html).toContain('const answer = 42')
    expect(parsed.html).toContain('data:image/png;base64,')
  })

  it('removes active and remote raw HTML', async () => {
    const parsed = await parseMarkdownDocument(await markdownArchive(), 'guide.zip', options)

    expect(parsed.html).not.toMatch(/script|tracker\.invalid/iu)
    expect(parsed.html).toContain("default-src 'none'")
  })

  it('supports standalone markdown and derives the title', async () => {
    const parsed = await parseMarkdownDocument(
      new TextEncoder().encode('# Standalone\n\n- first\n- second'),
      'notes.markdown',
      { ...options, zoom: 1.2 },
    )

    expect(parsed.title).toBe('Standalone')
    expect(parsed.html).toContain('<li>first</li>')
    expect(parsed.html).toContain('zoom:1.2')
  })

  it('rejects invalid inputs and archives without markdown', async () => {
    await expect(parseMarkdownDocument(new Uint8Array(), 'empty.md', options)).rejects.toThrow(
      'empty',
    )
    await expect(
      parseMarkdownDocument(new Uint8Array([1, 2, 3]), 'broken.zip', options),
    ).rejects.toThrow('invalid')
    const archive = new JSZip()
    archive.file('readme.txt', 'nothing to render')
    await expect(
      parseMarkdownDocument(
        await archive.generateAsync({ type: 'uint8array' }),
        'no-markdown.zip',
        options,
      ),
    ).rejects.toThrow('no Markdown')
  })
})
