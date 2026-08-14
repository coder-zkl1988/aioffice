import JSZip from 'jszip'
import type { PdfRasterPage } from '@genoffice/pdf-tools'
import { sanitizeSvgForPdf } from './images-to-pdf'
import { renderLocalHtmlPages } from './local-html-pages'
import {
  localBytesToBase64,
  localResourceDirectory,
  normalizeLocalResourcePath,
  sanitizeLocalHtmlFragment,
} from './local-web-content'

const MAX_EPUB_BYTES = 100 * 1024 * 1024
const MAX_EPUB_ENTRIES = 5_000
const MAX_EPUB_ENTRY_BYTES = 50 * 1024 * 1024
const MAX_EPUB_TOTAL_BYTES = 250 * 1024 * 1024
const MAX_EPUB_CHAPTERS = 200
const MAX_EPUB_PAGES = 200
const MAX_EMBEDDED_IMAGE_BYTES = 20 * 1024 * 1024
const MAX_EMBEDDED_FONT_BYTES = 10 * 1024 * 1024
const MAX_TOTAL_RESOURCE_BYTES = 100 * 1024 * 1024

const IMAGE_MEDIA_TYPES = new Set([
  'image/bmp',
  'image/gif',
  'image/jpeg',
  'image/png',
  'image/svg+xml',
  'image/webp',
])
const FONT_MEDIA_TYPES = new Set([
  'application/font-sfnt',
  'application/vnd.ms-opentype',
  'font/otf',
  'font/ttf',
  'font/woff',
  'font/woff2',
])

export interface PrepareEpubOptions {
  embedAllFonts: boolean
  includeTableOfContents: boolean
  includePageNumbers: boolean
  optimizeForEbook: boolean
  tableOfContentsLabel: string
  untitledLabel: string
}

export interface PreparedEpubDocument {
  html: string
  pages: PdfRasterPage[]
  title: string
  author?: string
  chapterCount: number
}

interface EpubManifestItem {
  id: string
  path: string
  mediaType: string
}

interface EpubChapter {
  path: string
  title: string
  html: string
}

interface EpubModel {
  title: string
  author?: string
  language?: string
  chapters: EpubChapter[]
  fontCss: string
  fontFamilies: string[]
}

interface EpubResourceState {
  dataUrls: Map<string, string | undefined>
  totalBytes: number
}

function normalizedArchivePath(value: string, basePath = ''): string {
  return normalizeLocalResourcePath(value, basePath, 'EPUB')
}

function directoryPath(path: string): string {
  return localResourceDirectory(path)
}

function xmlDocument(markup: string, label: string): Document {
  const document = new DOMParser().parseFromString(markup, 'application/xml')
  if (document.querySelector('parsererror')) throw new Error(`${label} is invalid`)
  return document
}

function localNameElements(root: ParentNode, name: string): Element[] {
  return [...root.querySelectorAll('*')].filter(
    (element) => element.localName.toLowerCase() === name.toLowerCase(),
  )
}

function firstLocalName(root: ParentNode, name: string): Element | undefined {
  return localNameElements(root, name)[0]
}

function archiveEntries(zip: JSZip): Map<string, JSZip.JSZipObject> {
  const entries = Object.values(zip.files)
  if (entries.length > MAX_EPUB_ENTRIES) throw new Error('EPUB contains too many files')
  const output = new Map<string, JSZip.JSZipObject>()
  for (const entry of entries) {
    if (entry.dir) continue
    const path = normalizedArchivePath(entry.name)
    output.set(path, entry)
  }
  return output
}

async function checkedEntryBytes(
  entry: JSZip.JSZipObject,
  label: string,
  maxBytes = MAX_EPUB_ENTRY_BYTES,
): Promise<Uint8Array> {
  const bytes = await entry.async('uint8array')
  if (bytes.length === 0) throw new Error(`${label} is empty`)
  if (bytes.length > maxBytes) throw new Error(`${label} is too large`)
  return bytes
}

async function checkedEntryText(entry: JSZip.JSZipObject, label: string): Promise<string> {
  return new TextDecoder().decode(await checkedEntryBytes(entry, label))
}

async function resourceDataUrl(
  pathValue: string,
  chapterPath: string,
  manifestByPath: Map<string, EpubManifestItem>,
  entries: Map<string, JSZip.JSZipObject>,
  resources: EpubResourceState,
): Promise<string | undefined> {
  if (!pathValue || /^(?:data:|https?:|file:|javascript:)/i.test(pathValue.trim())) return undefined
  let path: string
  try {
    path = normalizedArchivePath(pathValue, directoryPath(chapterPath))
  } catch {
    return undefined
  }
  const item = manifestByPath.get(path)
  const entry = entries.get(path)
  const mediaType = item?.mediaType.toLowerCase().split(';')[0]!.trim()
  if (!entry || !mediaType || !IMAGE_MEDIA_TYPES.has(mediaType)) return undefined
  if (resources.dataUrls.has(path)) return resources.dataUrls.get(path)
  const bytes = await checkedEntryBytes(entry, path, MAX_EMBEDDED_IMAGE_BYTES)
  resources.totalBytes += bytes.length
  if (resources.totalBytes > MAX_TOTAL_RESOURCE_BYTES) {
    throw new Error('EPUB embedded resources are too large')
  }
  let dataUrl: string
  if (mediaType === 'image/svg+xml') {
    const sanitized = sanitizeSvgForPdf(new TextDecoder().decode(bytes))
    dataUrl = `data:image/svg+xml;base64,${localBytesToBase64(new TextEncoder().encode(sanitized.markup))}`
  } else {
    dataUrl = `data:${mediaType};base64,${localBytesToBase64(bytes)}`
  }
  resources.dataUrls.set(path, dataUrl)
  return dataUrl
}

async function chapterMarkup(
  markup: string,
  chapterPath: string,
  manifestByPath: Map<string, EpubManifestItem>,
  entries: Map<string, JSZip.JSZipObject>,
  resources: EpubResourceState,
): Promise<{ title: string; html: string }> {
  return sanitizeLocalHtmlFragment(markup, {
    basePath: directoryPath(chapterPath),
    resolveImage: (source) =>
      resourceDataUrl(source, chapterPath, manifestByPath, entries, resources),
  })
}

async function embeddedFontCss(
  manifest: EpubManifestItem[],
  entries: Map<string, JSZip.JSZipObject>,
  enabled: boolean,
  resources: EpubResourceState,
): Promise<{ css: string; families: string[] }> {
  if (!enabled) return { css: '', families: [] }
  const rules: string[] = []
  const families: string[] = []
  let index = 0
  for (const item of manifest) {
    const mediaType = item.mediaType.toLowerCase().split(';')[0]!.trim()
    if (!FONT_MEDIA_TYPES.has(mediaType)) continue
    const entry = entries.get(item.path)
    if (!entry) continue
    const bytes = await checkedEntryBytes(entry, item.path, MAX_EMBEDDED_FONT_BYTES)
    resources.totalBytes += bytes.length
    if (resources.totalBytes > MAX_TOTAL_RESOURCE_BYTES) {
      throw new Error('EPUB embedded resources are too large')
    }
    const family = `GenOfficeEpubFont${++index}`
    families.push(family)
    rules.push(
      `@font-face{font-family:"${family}";src:url(data:${mediaType};base64,${localBytesToBase64(bytes)}) format("${mediaType.includes('woff2') ? 'woff2' : mediaType.includes('woff') ? 'woff' : mediaType.includes('opentype') || mediaType.includes('otf') ? 'opentype' : 'truetype'}");font-style:normal;font-weight:100 900;font-display:block}`,
    )
  }
  return { css: rules.join(''), families }
}

function epubDocument(model: EpubModel, options: PrepareEpubOptions): string {
  const escape = (value: string) =>
    value.replace(/[&<>"']/g, (character) =>
      character === '&'
        ? '&amp;'
        : character === '<'
          ? '&lt;'
          : character === '>'
            ? '&gt;'
            : character === '"'
              ? '&quot;'
              : '&#39;',
    )
  const toc = options.includeTableOfContents
    ? `<section class="ebook-toc"><h1>${escape(options.tableOfContentsLabel)}</h1><ol>${model.chapters.map((chapter) => `<li>${escape(chapter.title)}</li>`).join('')}</ol></section>`
    : ''
  const chapters = model.chapters
    .map(
      (chapter, index) =>
        `<article class="ebook-chapter"${index > 0 || toc ? ' data-pdf-page-break-before' : ''}>${chapter.html}</article>`,
    )
    .join('')
  const bookFonts = model.fontFamilies.map((family) => `"${family}"`).join(',')
  const bodyFonts = `${bookFonts ? `${bookFonts},` : ''}Georgia,"Noto Serif SC","Songti SC",serif`
  return `<!doctype html><html lang="${escape(model.language || 'en')}"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; font-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><meta name="viewport" content="width=794"><title>${escape(model.title)}</title><style>${model.fontCss}*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;color:#222}body{width:794px;font:15px/1.72 ${bodyFonts};overflow-wrap:anywhere}.ebook-toc,.ebook-chapter{padding:48px 58px 58px}.ebook-toc h1{margin:0 0 26px;font:700 27px/1.25 -apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans SC",sans-serif}.ebook-toc ol{margin:0;padding-left:1.6em}.ebook-toc li{padding:5px 0}.ebook-chapter h1,.ebook-chapter h2,.ebook-chapter h3{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans SC",sans-serif;line-height:1.3}.ebook-chapter h1{font-size:28px;margin:0 0 24px}.ebook-chapter h2{font-size:22px;margin:1.5em 0 .65em}.ebook-chapter h3{font-size:18px;margin:1.4em 0 .55em}.ebook-chapter p{margin:.8em 0;text-align:justify}.ebook-chapter img{display:block;max-width:100%;height:auto;margin:18px auto}.ebook-chapter table{max-width:100%;border-collapse:collapse}.ebook-chapter td,.ebook-chapter th{padding:6px 8px;border:1px solid #d4d4d4;vertical-align:top}.ebook-chapter pre{max-width:100%;padding:12px;background:#f5f5f5;white-space:pre-wrap}.ebook-chapter blockquote{margin:14px 0;padding-left:16px;border-left:3px solid #c9c9c9;color:#555}.ebook-chapter a{color:#185d94;text-decoration:underline}</style></head><body>${toc}${chapters}</body></html>`
}

export async function parseEpubDocument(
  input: Uint8Array | ArrayBuffer,
  options: PrepareEpubOptions,
): Promise<{ html: string; title: string; author?: string; chapterCount: number }> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  if (bytes.length === 0) throw new Error('EPUB file is empty')
  if (bytes.length > MAX_EPUB_BYTES) throw new Error('EPUB file must be 100 MB or smaller')

  let zip: JSZip
  try {
    zip = await JSZip.loadAsync(bytes)
  } catch {
    throw new Error('EPUB file is invalid or damaged')
  }
  const entries = archiveEntries(zip)
  let totalBytes = 0
  for (const entry of entries.values()) {
    const declaredSize = (entry as JSZip.JSZipObject & { _data?: { uncompressedSize?: number } })
      ._data?.uncompressedSize
    if (typeof declaredSize === 'number') {
      if (declaredSize > MAX_EPUB_ENTRY_BYTES) throw new Error('EPUB contains an oversized file')
      totalBytes += declaredSize
    }
  }
  if (totalBytes > MAX_EPUB_TOTAL_BYTES) throw new Error('EPUB expanded content is too large')

  const mimetype = entries.get('mimetype')
  if (
    !mimetype ||
    (await checkedEntryText(mimetype, 'EPUB mimetype')).trim() !== 'application/epub+zip'
  ) {
    throw new Error('EPUB mimetype is invalid')
  }
  const containerEntry = entries.get('META-INF/container.xml')
  if (!containerEntry) throw new Error('EPUB container is missing')
  const container = xmlDocument(
    await checkedEntryText(containerEntry, 'EPUB container'),
    'EPUB container',
  )
  const rootFile = firstLocalName(container, 'rootfile')?.getAttribute('full-path')
  if (!rootFile) throw new Error('EPUB package path is missing')
  const packagePath = normalizedArchivePath(rootFile)
  const packageEntry = entries.get(packagePath)
  if (!packageEntry) throw new Error('EPUB package is missing')
  const packageDocument = xmlDocument(
    await checkedEntryText(packageEntry, 'EPUB package'),
    'EPUB package',
  )
  const packageBase = directoryPath(packagePath)
  const metadata = firstLocalName(packageDocument, 'metadata')
  const title =
    metadata && firstLocalName(metadata, 'title')?.textContent?.trim()
      ? firstLocalName(metadata, 'title')!.textContent!.trim().slice(0, 500)
      : options.untitledLabel
  const author = metadata && firstLocalName(metadata, 'creator')?.textContent?.trim()
  const language = metadata && firstLocalName(metadata, 'language')?.textContent?.trim()
  const manifest = localNameElements(packageDocument, 'item').flatMap((element) => {
    const id = element.getAttribute('id')?.trim()
    const href = element.getAttribute('href')?.trim()
    const mediaType = element.getAttribute('media-type')?.trim()
    if (!id || !href || !mediaType) return []
    return [{ id, path: normalizedArchivePath(href, packageBase), mediaType }]
  })
  const manifestById = new Map(manifest.map((item) => [item.id, item]))
  const manifestByPath = new Map(manifest.map((item) => [item.path, item]))
  const spineItems = localNameElements(packageDocument, 'itemref')
    .map((element) => element.getAttribute('idref')?.trim())
    .filter((value): value is string => Boolean(value))
    .map((id) => manifestById.get(id))
    .filter((item): item is EpubManifestItem => Boolean(item))
    .filter((item) => /^(?:application\/xhtml\+xml|text\/html)$/i.test(item.mediaType))
  if (spineItems.length === 0) throw new Error('EPUB reading order is empty')
  if (spineItems.length > MAX_EPUB_CHAPTERS) {
    throw new Error(`EPUB may contain no more than ${MAX_EPUB_CHAPTERS} chapters`)
  }

  const chapters: EpubChapter[] = []
  const resources: EpubResourceState = { dataUrls: new Map(), totalBytes: 0 }
  let resourceBytes = 0
  for (const [index, item] of spineItems.entries()) {
    const entry = entries.get(item.path)
    if (!entry) throw new Error(`EPUB chapter is missing: ${item.path}`)
    const chapter = await chapterMarkup(
      await checkedEntryText(entry, item.path),
      item.path,
      manifestByPath,
      entries,
      resources,
    )
    if (!chapter.html.trim()) continue
    chapters.push({
      path: item.path,
      title: chapter.title || `${options.untitledLabel} ${index + 1}`,
      html: chapter.html,
    })
    resourceBytes += chapter.html.length
    if (resourceBytes > MAX_TOTAL_RESOURCE_BYTES)
      throw new Error('EPUB embedded resources are too large')
  }
  if (chapters.length === 0) throw new Error('EPUB contains no readable chapters')
  const fonts = await embeddedFontCss(manifest, entries, options.embedAllFonts, resources)
  const model: EpubModel = {
    title,
    author: author?.slice(0, 500),
    language: language?.slice(0, 50),
    chapters,
    fontCss: fonts.css,
    fontFamilies: fonts.families,
  }
  return {
    html: epubDocument(model, options),
    title: model.title,
    author: model.author,
    chapterCount: chapters.length,
  }
}

export async function prepareEpubDocumentForPdf(
  file: File,
  options: PrepareEpubOptions,
): Promise<PreparedEpubDocument> {
  const parsed = await parseEpubDocument(new Uint8Array(await file.arrayBuffer()), options)
  return {
    ...parsed,
    pages: await renderLocalHtmlPages(parsed.html, {
      maxPages: MAX_EPUB_PAGES,
      includePageNumbers: options.includePageNumbers,
      optimizeForEbook: options.optimizeForEbook,
      rasterScale: options.optimizeForEbook ? 1.25 : 1.5,
    }),
  }
}
