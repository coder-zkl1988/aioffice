import JSZip from 'jszip'
import { marked } from 'marked'
import type { PdfRasterPage } from '@genoffice/pdf-tools'
import { parseLocalWebDocument, type PrepareHtmlToPdfOptions } from './html-to-pdf'
import { renderLocalHtmlPages } from './local-html-pages'
import { localResourceDirectory, normalizeLocalResourcePath } from './local-web-content'

const MAX_INPUT_BYTES = 100 * 1024 * 1024
const MAX_MARKDOWN_BYTES = 20 * 1024 * 1024
const MAX_ARCHIVE_ENTRIES = 5_000
const MAX_ARCHIVE_ENTRY_BYTES = 50 * 1024 * 1024
const MAX_ARCHIVE_TOTAL_BYTES = 250 * 1024 * 1024
const MAX_OUTPUT_PAGES = 100

export interface PreparedMarkdownDocument {
  html: string
  pages: PdfRasterPage[]
  title: string
  entryPath?: string
  resourceCount: number
}

function markdownHtml(markdown: string): string {
  const body = marked.parse(markdown.replace(/\r\n?/g, '\n'), {
    async: false,
    breaks: false,
    gfm: true,
  })
  return `<!doctype html><html><head><meta charset="utf-8"><style>h1{font-size:30px;line-height:1.25;margin:0 0 24px}h2{font-size:23px;line-height:1.3;margin:1.5em 0 .65em}h3{font-size:19px;line-height:1.35;margin:1.35em 0 .55em}p{margin:.75em 0}blockquote{margin:16px 0;padding:2px 0 2px 16px;border-left:4px solid #c8d0d8;color:#56616c}code{padding:2px 4px;border-radius:3px;background-color:#f0f2f4;font-family:"SFMono-Regular",Consolas,"Liberation Mono",monospace}pre{padding:14px 16px;border:1px solid #dde2e7;border-radius:5px;background-color:#f6f8fa}pre code{padding:0;background-color:transparent}table{width:100%;margin:18px 0;border-collapse:collapse}th{background-color:#f2f5f7}th,td{padding:8px 10px;border:1px solid #d7dde2;text-align:left}hr{margin:26px 0;border:0;border-top:1px solid #d7dde2}ul,ol{padding-left:1.6em}li{margin:.25em 0}</style></head><body>${body}</body></html>`
}

function archiveEntries(zip: JSZip): Map<string, JSZip.JSZipObject> {
  const values = Object.values(zip.files)
  if (values.length > MAX_ARCHIVE_ENTRIES) throw new Error('ZIP contains too many files')
  const output = new Map<string, JSZip.JSZipObject>()
  let totalBytes = 0
  for (const entry of values) {
    if (entry.dir) continue
    const path = normalizeLocalResourcePath(entry.name, '', 'ZIP')
    if (output.has(path)) throw new Error('ZIP contains duplicate file paths')
    const declaredSize = (entry as JSZip.JSZipObject & { _data?: { uncompressedSize?: number } })
      ._data?.uncompressedSize
    if (typeof declaredSize === 'number') {
      if (declaredSize > MAX_ARCHIVE_ENTRY_BYTES) throw new Error('ZIP contains an oversized file')
      totalBytes += declaredSize
    }
    output.set(path, entry)
  }
  if (totalBytes > MAX_ARCHIVE_TOTAL_BYTES) throw new Error('ZIP expanded content is too large')
  return output
}

function mainMarkdownPath(entries: Map<string, JSZip.JSZipObject>): string {
  const paths = [...entries.keys()].filter((path) => /\.(?:md|markdown)$/i.test(path))
  if (paths.length === 0) throw new Error('ZIP contains no Markdown document')
  return paths.sort((left, right) => {
    const leftName = left.split('/').at(-1)!.toLowerCase()
    const rightName = right.split('/').at(-1)!.toLowerCase()
    const priority = (name: string) =>
      name === 'index.md' || name === 'index.markdown'
        ? 0
        : name === 'readme.md' || name === 'readme.markdown'
          ? 1
          : 2
    return (
      priority(leftName) - priority(rightName) ||
      left.split('/').length - right.split('/').length ||
      left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' })
    )
  })[0]!
}

async function checkedMarkdown(entry: JSZip.JSZipObject, label: string): Promise<string> {
  const bytes = await entry.async('uint8array')
  if (bytes.length === 0) throw new Error(`${label} is empty`)
  if (bytes.length > MAX_MARKDOWN_BYTES) throw new Error(`${label} must be 20 MB or smaller`)
  return new TextDecoder().decode(bytes)
}

export async function parseMarkdownDocument(
  input: Uint8Array | ArrayBuffer,
  fileName: string,
  options: PrepareHtmlToPdfOptions,
): Promise<Omit<PreparedMarkdownDocument, 'pages'>> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  if (bytes.length === 0) throw new Error('Markdown input is empty')
  if (bytes.length > MAX_INPUT_BYTES) {
    throw new Error('Markdown or ZIP input must be 100 MB or smaller')
  }
  if (/\.zip$/i.test(fileName)) {
    let archive: JSZip
    try {
      archive = await JSZip.loadAsync(bytes)
    } catch {
      throw new Error('ZIP file is invalid or damaged')
    }
    const entries = archiveEntries(archive)
    const entryPath = mainMarkdownPath(entries)
    const markdown = await checkedMarkdown(entries.get(entryPath)!, entryPath)
    for (const entry of Object.values(archive.files)) {
      if (!entry.dir && /\.(?:html?|md|markdown)$/i.test(entry.name)) archive.remove(entry.name)
    }
    const htmlPath = `${localResourceDirectory(entryPath)}index.html`
    archive.file(htmlPath, markdownHtml(markdown), { createFolders: true })
    const parsed = await parseLocalWebDocument(
      await archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
      fileName,
      options,
    )
    return { ...parsed, entryPath }
  }
  if (!/\.(?:md|markdown)$/i.test(fileName)) {
    throw new Error('Choose a Markdown or ZIP file')
  }
  if (bytes.length > MAX_MARKDOWN_BYTES) throw new Error('Markdown file must be 20 MB or smaller')
  return parseLocalWebDocument(
    new TextEncoder().encode(markdownHtml(new TextDecoder().decode(bytes))),
    fileName.replace(/\.(?:md|markdown)$/i, '.html'),
    options,
  )
}

export async function prepareMarkdownDocumentForPdf(
  file: File,
  options: PrepareHtmlToPdfOptions,
): Promise<PreparedMarkdownDocument> {
  const parsed = await parseMarkdownDocument(
    new Uint8Array(await file.arrayBuffer()),
    file.name,
    options,
  )
  return {
    ...parsed,
    pages: await renderLocalHtmlPages(parsed.html, {
      maxPages: MAX_OUTPUT_PAGES,
      includePageNumbers: options.includePageNumbers,
    }),
  }
}
