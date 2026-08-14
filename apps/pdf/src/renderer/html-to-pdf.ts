import JSZip from 'jszip'
import type { PdfRasterPage } from '@genoffice/pdf-tools'
import type { PdfWebResourceRequest, PdfWebResourceResult } from '../shared/ipc'
import { sanitizeSvgForPdf } from './images-to-pdf'
import { renderLocalHtmlPages } from './local-html-pages'
import {
  localBytesToBase64,
  localResourceDirectory,
  normalizeLocalResourcePath,
  sanitizeLocalHtmlFragment,
  sanitizeLocalStylesheet,
} from './local-web-content'

const MAX_INPUT_BYTES = 100 * 1024 * 1024
const MAX_HTML_BYTES = 20 * 1024 * 1024
const MAX_ARCHIVE_ENTRIES = 5_000
const MAX_ARCHIVE_ENTRY_BYTES = 50 * 1024 * 1024
const MAX_ARCHIVE_TOTAL_BYTES = 250 * 1024 * 1024
const MAX_STYLESHEET_BYTES = 5 * 1024 * 1024
const MAX_IMAGE_BYTES = 20 * 1024 * 1024
const MAX_TOTAL_RESOURCE_BYTES = 100 * 1024 * 1024
const MAX_OUTPUT_PAGES = 100
const MAX_REMOTE_RESOURCES = 80
const MAX_REMOTE_STYLESHEETS = 20

const IMAGE_MEDIA_TYPES: Record<string, string> = {
  '.bmp': 'image/bmp',
  '.gif': 'image/gif',
  '.jpeg': 'image/jpeg',
  '.jpg': 'image/jpeg',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.webp': 'image/webp',
}

export interface PrepareHtmlToPdfOptions {
  includePageNumbers: boolean
  zoom: number
  untitledLabel: string
}

export interface ParsedLocalWebDocument {
  html: string
  title: string
  entryPath?: string
  resourceCount: number
}

export interface PreparedLocalWebDocument extends ParsedLocalWebDocument {
  pages: PdfRasterPage[]
}

export interface ParsedRemoteWebDocument extends ParsedLocalWebDocument {
  url: string
}

export interface PreparedRemoteWebDocument extends ParsedRemoteWebDocument {
  pages: PdfRasterPage[]
}

export type PdfWebResourceLoader = (request: PdfWebResourceRequest) => Promise<PdfWebResourceResult>

interface ResourceState {
  dataUrls: Map<string, string | undefined>
  totalBytes: number
}

interface RemoteResourceState {
  dataUrls: Map<string, string | undefined>
  stylesheets: Map<string, string | undefined>
  totalBytes: number
  count: number
}

function validateHtmlOptions(options: PrepareHtmlToPdfOptions): void {
  if (!Number.isFinite(options.zoom) || options.zoom < 0.5 || options.zoom > 2) {
    throw new Error('HTML zoom must be between 50% and 200%')
  }
}

function fileExtension(path: string): string {
  const match = path.toLowerCase().match(/\.[a-z0-9]+$/)
  return match?.[0] ?? ''
}

function archiveEntries(zip: JSZip): Map<string, JSZip.JSZipObject> {
  const values = Object.values(zip.files)
  if (values.length > MAX_ARCHIVE_ENTRIES) throw new Error('ZIP contains too many files')
  const output = new Map<string, JSZip.JSZipObject>()
  for (const entry of values) {
    if (entry.dir) continue
    const path = normalizeLocalResourcePath(entry.name, '', 'ZIP')
    if (output.has(path)) throw new Error('ZIP contains duplicate file paths')
    output.set(path, entry)
  }
  return output
}

async function checkedEntryBytes(
  entry: JSZip.JSZipObject,
  label: string,
  maxBytes = MAX_ARCHIVE_ENTRY_BYTES,
): Promise<Uint8Array> {
  const bytes = await entry.async('uint8array')
  if (bytes.length === 0) throw new Error(`${label} is empty`)
  if (bytes.length > maxBytes) throw new Error(`${label} is too large`)
  return bytes
}

async function checkedEntryText(
  entry: JSZip.JSZipObject,
  label: string,
  maxBytes = MAX_HTML_BYTES,
): Promise<string> {
  return new TextDecoder().decode(await checkedEntryBytes(entry, label, maxBytes))
}

function mainHtmlPath(entries: Map<string, JSZip.JSZipObject>): string {
  const htmlPaths = [...entries.keys()].filter((path) => /\.html?$/i.test(path))
  if (htmlPaths.length === 0) throw new Error('ZIP contains no HTML document')
  return htmlPaths.sort((left, right) => {
    const leftName = left.split('/').at(-1)!.toLowerCase()
    const rightName = right.split('/').at(-1)!.toLowerCase()
    const leftIndex = leftName === 'index.html' ? 0 : leftName === 'index.htm' ? 1 : 2
    const rightIndex = rightName === 'index.html' ? 0 : rightName === 'index.htm' ? 1 : 2
    return (
      leftIndex - rightIndex ||
      left.split('/').length - right.split('/').length ||
      left.localeCompare(right, undefined, { numeric: true, sensitivity: 'base' })
    )
  })[0]!
}

function checkedArchiveSize(entries: Map<string, JSZip.JSZipObject>): void {
  let totalBytes = 0
  for (const entry of entries.values()) {
    const declaredSize = (entry as JSZip.JSZipObject & { _data?: { uncompressedSize?: number } })
      ._data?.uncompressedSize
    if (typeof declaredSize !== 'number') continue
    if (declaredSize > MAX_ARCHIVE_ENTRY_BYTES) throw new Error('ZIP contains an oversized file')
    totalBytes += declaredSize
  }
  if (totalBytes > MAX_ARCHIVE_TOTAL_BYTES) throw new Error('ZIP expanded content is too large')
}

function safeDataImage(source: string): string | undefined {
  const value = source.trim()
  if (!/^data:image\/(?:bmp|gif|jpeg|png|webp);base64,[a-z0-9+/=\s]+$/i.test(value)) {
    return undefined
  }
  if (value.length > Math.ceil((MAX_IMAGE_BYTES * 4) / 3) + 200) {
    throw new Error('HTML contains an oversized image')
  }
  return value.replace(/\s+/g, '')
}

function remoteMediaType(contentType: string): string {
  return contentType.split(';')[0]?.trim().toLowerCase() || ''
}

function decodeRemoteText(resource: PdfWebResourceResult): string {
  const charset = /(?:^|;)\s*charset\s*=\s*["']?([^;"'\s]+)/i.exec(resource.contentType)?.[1]
  if (charset) {
    try {
      return new TextDecoder(charset).decode(resource.bytes)
    } catch {
      // Fall back to UTF-8 for invalid or unsupported labels.
    }
  }
  return new TextDecoder().decode(resource.bytes)
}

function absoluteRemoteUrl(source: string, baseUrl: string): string | undefined {
  const value = source.trim()
  if (!value || /^(?:data:|blob:|file:|javascript:|ftp:)/i.test(value)) return undefined
  try {
    const url = new URL(value, baseUrl)
    if (url.protocol !== 'http:' && url.protocol !== 'https:') return undefined
    url.hash = ''
    return url.toString()
  } catch {
    return undefined
  }
}

function addRemoteResource(state: RemoteResourceState, bytes: Uint8Array): void {
  state.count += 1
  state.totalBytes += bytes.length
  if (state.count > MAX_REMOTE_RESOURCES) throw new Error('Website contains too many resources')
  if (state.totalBytes > MAX_TOTAL_RESOURCE_BYTES) {
    throw new Error('Website embedded resources are too large')
  }
}

async function remoteImageDataUrl(
  source: string,
  baseUrl: string,
  loadResource: PdfWebResourceLoader,
  state: RemoteResourceState,
): Promise<string | undefined> {
  if (/^data:/i.test(source.trim())) return safeDataImage(source)
  const url = absoluteRemoteUrl(source, baseUrl)
  if (!url) return undefined
  if (state.dataUrls.has(url)) return state.dataUrls.get(url)
  state.dataUrls.set(url, undefined)
  let resource: PdfWebResourceResult
  try {
    resource = await loadResource({ url, kind: 'image' })
  } catch {
    return undefined
  }
  addRemoteResource(state, resource.bytes)
  const mediaType = remoteMediaType(resource.contentType)
  if (!Object.values(IMAGE_MEDIA_TYPES).includes(mediaType)) return undefined
  const dataUrl =
    mediaType === 'image/svg+xml'
      ? `data:image/svg+xml;base64,${localBytesToBase64(new TextEncoder().encode(sanitizeSvgForPdf(decodeRemoteText(resource)).markup))}`
      : `data:${mediaType};base64,${localBytesToBase64(resource.bytes)}`
  state.dataUrls.set(url, dataUrl)
  return dataUrl
}

async function remoteDocumentStyles(
  markup: string,
  baseUrl: string,
  loadResource: PdfWebResourceLoader,
  state: RemoteResourceState,
): Promise<string> {
  const parsed = new DOMParser().parseFromString(markup, 'text/html')
  const styles = [...parsed.querySelectorAll('style')].map((element) => element.textContent ?? '')
  const links = [...parsed.querySelectorAll('link[rel~="stylesheet" i][href]')].slice(
    0,
    MAX_REMOTE_STYLESHEETS,
  )
  for (const link of links) {
    const url = absoluteRemoteUrl(link.getAttribute('href') ?? '', baseUrl)
    if (!url) continue
    let stylesheet = state.stylesheets.get(url)
    if (!state.stylesheets.has(url)) {
      state.stylesheets.set(url, undefined)
      let resource: PdfWebResourceResult
      try {
        resource = await loadResource({ url, kind: 'stylesheet' })
      } catch {
        continue
      }
      addRemoteResource(state, resource.bytes)
      stylesheet = decodeRemoteText(resource)
      state.stylesheets.set(url, stylesheet)
    }
    if (stylesheet) styles.push(stylesheet)
  }
  return sanitizeLocalStylesheet(styles.join('\n'))
}

function normalizedWebsiteUrl(value: string): string {
  const trimmed = value.trim()
  if (!trimmed) throw new Error('Enter a website URL')
  const candidate = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`
  let url: URL
  try {
    url = new URL(candidate)
  } catch {
    throw new Error('Website URL is invalid')
  }
  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new Error('Website URL must use HTTP or HTTPS')
  }
  return url.toString()
}

async function resourceDataUrl(
  source: string,
  basePath: string,
  entries: Map<string, JSZip.JSZipObject> | undefined,
  resources: ResourceState,
): Promise<string | undefined> {
  if (!source) return undefined
  if (/^data:/i.test(source)) return safeDataImage(source)
  if (!entries || /^(?:https?:|file:|blob:|javascript:|ftp:|\/\/)/i.test(source.trim())) {
    return undefined
  }
  let path: string
  try {
    path = normalizeLocalResourcePath(source, basePath, 'HTML')
  } catch {
    return undefined
  }
  if (resources.dataUrls.has(path)) return resources.dataUrls.get(path)
  const mediaType = IMAGE_MEDIA_TYPES[fileExtension(path)]
  const entry = entries.get(path)
  if (!entry || !mediaType) return undefined
  const bytes = await checkedEntryBytes(entry, path, MAX_IMAGE_BYTES)
  resources.totalBytes += bytes.length
  if (resources.totalBytes > MAX_TOTAL_RESOURCE_BYTES) {
    throw new Error('HTML embedded resources are too large')
  }
  const dataUrl =
    mediaType === 'image/svg+xml'
      ? `data:image/svg+xml;base64,${localBytesToBase64(new TextEncoder().encode(sanitizeSvgForPdf(new TextDecoder().decode(bytes)).markup))}`
      : `data:${mediaType};base64,${localBytesToBase64(bytes)}`
  resources.dataUrls.set(path, dataUrl)
  return dataUrl
}

async function documentStyles(
  markup: string,
  htmlPath: string,
  entries: Map<string, JSZip.JSZipObject> | undefined,
): Promise<string> {
  const parsed = new DOMParser().parseFromString(markup, 'text/html')
  const styles = [...parsed.querySelectorAll('style')].map((element) => element.textContent ?? '')
  if (entries) {
    for (const link of parsed.querySelectorAll('link[rel~="stylesheet" i][href]')) {
      const href = link.getAttribute('href')?.trim()
      if (!href || /^(?:https?:|file:|data:|javascript:|\/\/)/i.test(href)) continue
      let path: string
      try {
        path = normalizeLocalResourcePath(href, localResourceDirectory(htmlPath), 'HTML')
      } catch {
        continue
      }
      const entry = entries.get(path)
      if (!entry || fileExtension(path) !== '.css') continue
      styles.push(await checkedEntryText(entry, path, MAX_STYLESHEET_BYTES))
    }
  }
  return sanitizeLocalStylesheet(styles.join('\n'))
}

function localWebHtml(
  body: string,
  title: string,
  stylesheet: string,
  language: string,
  zoom: number,
): string {
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
  return `<!doctype html><html lang="${escape(language)}"><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><meta name="viewport" content="width=794"><title>${escape(title)}</title><style>*{box-sizing:border-box}html,body{margin:0;padding:0;background:#fff;color:#222}body{width:794px;padding:48px 58px 58px;font:15px/1.55 -apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans SC",Arial,sans-serif;overflow-wrap:anywhere}.local-web-document{zoom:${zoom}}img{max-width:100%;height:auto}table{max-width:100%;border-collapse:collapse}td,th{vertical-align:top}pre{max-width:100%;white-space:pre-wrap}a{color:#185d94;text-decoration:underline}${stylesheet}</style></head><body><main class="local-web-document">${body}</main></body></html>`
}

export async function parseLocalWebDocument(
  input: Uint8Array | ArrayBuffer,
  fileName: string,
  options: PrepareHtmlToPdfOptions,
): Promise<ParsedLocalWebDocument> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  if (bytes.length === 0) throw new Error('HTML input is empty')
  if (bytes.length > MAX_INPUT_BYTES) throw new Error('HTML or ZIP input must be 100 MB or smaller')
  validateHtmlOptions(options)

  const isArchive = /\.zip$/i.test(fileName)
  const isHtml = /\.html?$/i.test(fileName)
  if (!isArchive && !isHtml) throw new Error('Choose an HTML, HTM, or ZIP file')

  let entries: Map<string, JSZip.JSZipObject> | undefined
  let entryPath = ''
  let markup: string
  if (isArchive) {
    let archive: JSZip
    try {
      archive = await JSZip.loadAsync(bytes)
    } catch {
      throw new Error('ZIP file is invalid or damaged')
    }
    entries = archiveEntries(archive)
    checkedArchiveSize(entries)
    entryPath = mainHtmlPath(entries)
    markup = await checkedEntryText(entries.get(entryPath)!, entryPath)
  } else {
    if (bytes.length > MAX_HTML_BYTES) throw new Error('HTML file must be 20 MB or smaller')
    markup = new TextDecoder().decode(bytes)
  }

  const parsedSource = new DOMParser().parseFromString(markup, 'text/html')
  const language = parsedSource.documentElement.lang.trim().slice(0, 35) || 'en'
  const resources: ResourceState = { dataUrls: new Map(), totalBytes: 0 }
  const basePath = entryPath ? localResourceDirectory(entryPath) : ''
  const sanitized = await sanitizeLocalHtmlFragment(markup, {
    basePath,
    preserveSelectors: true,
    resolveImage: (source, resourceBase) =>
      resourceDataUrl(source, resourceBase, entries, resources),
  })
  if (!sanitized.html.trim()) throw new Error('HTML contains no printable content')
  const title =
    sanitized.title || fileName.replace(/\.(?:html?|zip)$/i, '') || options.untitledLabel
  const stylesheet = await documentStyles(markup, entryPath, entries)
  return {
    html: localWebHtml(sanitized.html, title, stylesheet, language, options.zoom),
    title,
    entryPath: entryPath || undefined,
    resourceCount: resources.dataUrls.size,
  }
}

export async function parseRemoteWebDocument(
  websiteUrl: string,
  loadResource: PdfWebResourceLoader,
  options: PrepareHtmlToPdfOptions,
): Promise<ParsedRemoteWebDocument> {
  validateHtmlOptions(options)
  const documentResource = await loadResource({
    url: normalizedWebsiteUrl(websiteUrl),
    kind: 'document',
  })
  const markup = decodeRemoteText(documentResource)
  if (!markup.trim()) throw new Error('Website returned no printable content')
  const parsedSource = new DOMParser().parseFromString(markup, 'text/html')
  const declaredBase = parsedSource.querySelector('base[href]')?.getAttribute('href') ?? ''
  const baseUrl = absoluteRemoteUrl(declaredBase, documentResource.url) ?? documentResource.url
  const language = parsedSource.documentElement.lang.trim().slice(0, 35) || 'en'
  const resources: RemoteResourceState = {
    dataUrls: new Map(),
    stylesheets: new Map(),
    totalBytes: 0,
    count: 0,
  }
  const sanitized = await sanitizeLocalHtmlFragment(markup, {
    basePath: baseUrl,
    preserveSelectors: true,
    resolveImage: (source, resourceBase) =>
      remoteImageDataUrl(source, resourceBase, loadResource, resources),
  })
  if (!sanitized.html.trim()) throw new Error('Website contains no printable content')
  const title = sanitized.title || new URL(documentResource.url).hostname || options.untitledLabel
  const stylesheet = await remoteDocumentStyles(markup, baseUrl, loadResource, resources)
  return {
    html: localWebHtml(sanitized.html, title, stylesheet, language, options.zoom),
    title,
    url: documentResource.url,
    resourceCount: resources.count,
  }
}

export async function prepareLocalWebDocumentForPdf(
  file: File,
  options: PrepareHtmlToPdfOptions,
): Promise<PreparedLocalWebDocument> {
  const parsed = await parseLocalWebDocument(
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

export async function prepareRemoteWebDocumentForPdf(
  websiteUrl: string,
  loadResource: PdfWebResourceLoader,
  options: PrepareHtmlToPdfOptions,
): Promise<PreparedRemoteWebDocument> {
  const parsed = await parseRemoteWebDocument(websiteUrl, loadResource, options)
  return {
    ...parsed,
    pages: await renderLocalHtmlPages(parsed.html, {
      maxPages: MAX_OUTPUT_PAGES,
      includePageNumbers: options.includePageNumbers,
    }),
  }
}
