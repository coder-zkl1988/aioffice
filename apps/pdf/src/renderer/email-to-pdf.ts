import type { Address, Attachment, Email } from 'postal-mime'
import type { PdfAttachmentInput, PdfRasterPage } from '@genoffice/pdf-tools'
import { renderLocalHtmlPages } from './local-html-pages'

const MAX_EMAIL_BYTES = 50 * 1024 * 1024
const MAX_INLINE_IMAGE_BYTES = 10 * 1024 * 1024
const MAX_INLINE_IMAGE_TOTAL_BYTES = 30 * 1024 * 1024
const EMAIL_PAGE_WIDTH_CSS = 794

const SAFE_HTML_TAGS = new Set([
  'a',
  'b',
  'blockquote',
  'br',
  'code',
  'del',
  'div',
  'em',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'hr',
  'i',
  'img',
  'li',
  'ol',
  'p',
  'pre',
  's',
  'span',
  'strong',
  'sub',
  'sup',
  'table',
  'tbody',
  'td',
  'tfoot',
  'th',
  'thead',
  'tr',
  'u',
  'ul',
])
const DROP_HTML_TAGS = new Set([
  'applet',
  'audio',
  'base',
  'embed',
  'frame',
  'frameset',
  'head',
  'iframe',
  'link',
  'meta',
  'noscript',
  'object',
  'script',
  'style',
  'svg',
  'template',
  'video',
])
const SAFE_STYLE_PROPERTIES = new Set([
  'background-color',
  'border',
  'border-bottom',
  'border-color',
  'border-left',
  'border-right',
  'border-style',
  'border-top',
  'border-width',
  'color',
  'font-size',
  'font-style',
  'font-weight',
  'line-height',
  'list-style-type',
  'margin',
  'margin-bottom',
  'margin-left',
  'margin-right',
  'margin-top',
  'max-width',
  'padding',
  'padding-bottom',
  'padding-left',
  'padding-right',
  'padding-top',
  'text-align',
  'text-decoration',
  'vertical-align',
  'white-space',
  'width',
  'word-break',
])
const SAFE_INLINE_IMAGE_TYPES = new Set([
  'image/bmp',
  'image/gif',
  'image/jpeg',
  'image/png',
  'image/webp',
])

export interface EmailDocumentLabels {
  from: string
  to: string
  cc: string
  bcc: string
  date: string
  attachments: string
  emptyBody: string
  remoteImageRemoved: string
  untitled: string
}

export interface PrepareEmailDocumentOptions {
  includeAttachments: boolean
  maxAttachmentSizeMb: number
  includeAllRecipients: boolean
  labels: EmailDocumentLabels
  locale: string
}

export interface PreparedEmailDocument {
  html: string
  pages: PdfRasterPage[]
  attachments: PdfAttachmentInput[]
}

interface PreparedEmailMarkup {
  html: string
  attachments: PdfAttachmentInput[]
}

function attachmentBytes(attachment: Attachment): Uint8Array {
  if (attachment.content instanceof Uint8Array) return attachment.content
  if (attachment.content instanceof ArrayBuffer) return new Uint8Array(attachment.content)
  if (attachment.encoding === 'base64') {
    const binary = atob(attachment.content.replace(/\s+/g, ''))
    return Uint8Array.from(binary, (character) => character.charCodeAt(0))
  }
  return new TextEncoder().encode(attachment.content)
}

function normalizedContentId(value: string | undefined): string {
  return (value ?? '').trim().replace(/^<|>$/g, '').toLowerCase()
}

function bytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  const chunkSize = 0x8000
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize))
  }
  return btoa(binary)
}

function safeInlineImages(attachments: Attachment[]): Map<string, string> {
  const images = new Map<string, string>()
  let totalBytes = 0
  for (const attachment of attachments) {
    const contentId = normalizedContentId(attachment.contentId)
    const mimeType = attachment.mimeType.toLowerCase().split(';')[0]!.trim()
    if (!contentId || !SAFE_INLINE_IMAGE_TYPES.has(mimeType)) continue
    const bytes = attachmentBytes(attachment)
    if (bytes.length === 0 || bytes.length > MAX_INLINE_IMAGE_BYTES) continue
    totalBytes += bytes.length
    if (totalBytes > MAX_INLINE_IMAGE_TOTAL_BYTES) break
    images.set(contentId, `data:${mimeType};base64,${bytesToBase64(bytes)}`)
  }
  return images
}

function safeLink(value: string): string | undefined {
  const trimmed = value.trim()
  if (/^(?:https?:|mailto:)/i.test(trimmed)) return trimmed
  return undefined
}

function sanitizedStyle(source: Element): string {
  const probe = document.createElement('span')
  probe.setAttribute('style', source.getAttribute('style') ?? '')
  const output: string[] = []
  for (const property of Array.from(probe.style)) {
    if (!SAFE_STYLE_PROPERTIES.has(property)) continue
    const value = probe.style.getPropertyValue(property).trim()
    if (!value || /(?:expression|url|javascript|@import|var)\s*\(/i.test(value)) continue
    if (/^-/.test(value) && /^(?:margin|max-width|padding|width)/.test(property)) continue
    output.push(`${property}:${value}`)
  }
  return output.join(';')
}

function appendSanitizedNode(
  source: Node,
  parent: Node,
  inlineImages: Map<string, string>,
  labels: EmailDocumentLabels,
): void {
  if (source.nodeType === Node.TEXT_NODE) {
    parent.appendChild(document.createTextNode(source.textContent ?? ''))
    return
  }
  if (!(source instanceof Element)) return
  const tag = source.tagName.toLowerCase()
  if (DROP_HTML_TAGS.has(tag)) return
  if (!SAFE_HTML_TAGS.has(tag)) {
    for (const child of source.childNodes) appendSanitizedNode(child, parent, inlineImages, labels)
    return
  }

  if (tag === 'img') {
    const rawSource = source.getAttribute('src')?.trim() ?? ''
    const cid = rawSource.toLowerCase().startsWith('cid:')
      ? normalizedContentId(rawSource.slice(4))
      : ''
    const imageSource = cid ? inlineImages.get(cid) : undefined
    if (!imageSource) {
      const replacement = document.createElement('span')
      replacement.className = 'email-remote-image'
      const alt = source.getAttribute('alt')?.trim()
      replacement.textContent = alt
        ? `${alt} (${labels.remoteImageRemoved})`
        : labels.remoteImageRemoved
      parent.appendChild(replacement)
      return
    }
    const image = document.createElement('img')
    image.src = imageSource
    image.alt = source.getAttribute('alt')?.slice(0, 300) ?? ''
    image.loading = 'eager'
    image.decoding = 'sync'
    parent.appendChild(image)
    return
  }

  const output = document.createElement(tag)
  const style = sanitizedStyle(source)
  if (style) output.setAttribute('style', style)
  const direction = source.getAttribute('dir')?.toLowerCase()
  if (direction === 'ltr' || direction === 'rtl' || direction === 'auto') {
    output.setAttribute('dir', direction)
  }
  if (tag === 'a') {
    const href = safeLink(source.getAttribute('href') ?? '')
    if (href) {
      output.setAttribute('href', href)
      output.setAttribute('rel', 'noreferrer noopener')
    }
  }
  if (tag === 'td' || tag === 'th') {
    for (const attribute of ['colspan', 'rowspan'] as const) {
      const value = Number(source.getAttribute(attribute))
      if (Number.isInteger(value) && value >= 1 && value <= 100) {
        output.setAttribute(attribute, String(value))
      }
    }
  }
  for (const child of source.childNodes) appendSanitizedNode(child, output, inlineImages, labels)
  parent.appendChild(output)
}

export function sanitizeEmailHtml(
  markup: string,
  attachments: Attachment[],
  labels: EmailDocumentLabels,
): string {
  const parsed = new DOMParser().parseFromString(markup, 'text/html')
  const output = document.createElement('div')
  const inlineImages = safeInlineImages(attachments)
  for (const child of parsed.body.childNodes) {
    appendSanitizedNode(child, output, inlineImages, labels)
  }
  return output.innerHTML
}

function escapeHtml(value: string): string {
  return value.replace(/[&<>"']/g, (character) => {
    if (character === '&') return '&amp;'
    if (character === '<') return '&lt;'
    if (character === '>') return '&gt;'
    if (character === '"') return '&quot;'
    return '&#39;'
  })
}

function mailboxLabel(address: Address): string {
  if (Array.isArray(address.group)) {
    const group = address.group.map(mailboxLabel).filter(Boolean).join(', ')
    return address.name ? `${address.name}: ${group}` : group
  }
  if (address.name && address.address) return `${address.name} <${address.address}>`
  return address.name || address.address || ''
}

function addressList(addresses: Address[] | undefined): string {
  return (addresses ?? []).map(mailboxLabel).filter(Boolean).join(', ')
}

function formattedDate(value: string | undefined, locale: string): string {
  if (!value) return ''
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return value
  return new Intl.DateTimeFormat(locale, {
    dateStyle: 'medium',
    timeStyle: 'medium',
  }).format(date)
}

function safeAttachmentName(value: string | null, index: number): string {
  const simpleName = (value ?? '').replace(/\\/g, '/').split('/').pop()?.trim() ?? ''
  const sanitized = simpleName.replace(/[\p{Cc}:*?"<>|]/gu, '_')
  return sanitized && sanitized !== '.' && sanitized !== '..'
    ? sanitized.slice(0, 180)
    : `attachment_${index + 1}.bin`
}

function uniqueAttachmentInputs(
  attachments: Attachment[],
  maxAttachmentSizeMb: number,
): PdfAttachmentInput[] {
  const maxBytes = Math.round(maxAttachmentSizeMb * 1024 * 1024)
  const names = new Set<string>()
  const inputs: PdfAttachmentInput[] = []
  for (const [index, attachment] of attachments.entries()) {
    if (attachment.disposition === 'inline' || attachment.related || attachment.contentId) continue
    const bytes = attachmentBytes(attachment)
    if (bytes.length === 0 || bytes.length > maxBytes) continue
    const initialName = safeAttachmentName(attachment.filename, index)
    const extensionIndex = initialName.lastIndexOf('.')
    const stem = extensionIndex > 0 ? initialName.slice(0, extensionIndex) : initialName
    const extension = extensionIndex > 0 ? initialName.slice(extensionIndex) : ''
    let name = initialName
    let suffix = 1
    while (names.has(name.toLowerCase())) name = `${stem}_${suffix++}${extension}`
    names.add(name.toLowerCase())
    inputs.push({ name, bytes, mimeType: attachment.mimeType })
  }
  return inputs
}

function headerRow(label: string, value: string): string {
  if (!value) return ''
  return `<dt>${escapeHtml(label)}</dt><dd>${escapeHtml(value)}</dd>`
}

function attachmentRows(attachments: Attachment[]): string {
  return attachments
    .filter(
      (attachment) =>
        attachment.disposition !== 'inline' && !attachment.related && !attachment.contentId,
    )
    .map((attachment, index) => {
      const name = safeAttachmentName(attachment.filename, index)
      const size = attachmentBytes(attachment).length
      const sizeText =
        size < 1024 * 1024
          ? `${Math.ceil(size / 1024)} KB`
          : `${(size / 1024 / 1024).toFixed(1)} MB`
      return `<li><span>${escapeHtml(name)}</span><small>${escapeHtml(sizeText)}</small></li>`
    })
    .join('')
}

function emailStyles(): string {
  return `
    *{box-sizing:border-box}
    html,body{margin:0;padding:0;background:#fff;color:#242424}
    body{width:${EMAIL_PAGE_WIDTH_CSS}px;font:14px/1.65 -apple-system,BlinkMacSystemFont,"Segoe UI","Noto Sans SC","Microsoft YaHei",Arial,sans-serif;overflow-wrap:anywhere}
    .email-document{padding:42px 52px 54px}
    .email-header{padding-bottom:22px;border-bottom:1px solid #d9d9d9}
    .email-subject{margin:0 0 18px;font-size:25px;line-height:1.3;font-weight:650;color:#1d1d1f}
    .email-meta{display:grid;grid-template-columns:72px minmax(0,1fr);gap:6px 12px;margin:0}
    .email-meta dt{margin:0;color:#737373;font-weight:600}
    .email-meta dd{margin:0;color:#2f2f2f}
    .email-body{padding:28px 0 12px;min-height:120px}
    .email-body img{display:block;max-width:100%;height:auto;margin:14px 0}
    .email-body table{max-width:100%;border-collapse:collapse}
    .email-body td,.email-body th{padding:6px 8px;border:1px solid #d7d7d7;vertical-align:top}
    .email-body pre{max-width:100%;padding:12px;background:#f5f5f5;white-space:pre-wrap}
    .email-body blockquote{margin:14px 0;padding-left:16px;border-left:3px solid #c9c9c9;color:#5f5f5f}
    .email-body a{color:#1264a3;text-decoration:underline}
    .email-remote-image{display:inline-block;padding:4px 8px;border:1px solid #dedede;background:#f7f7f7;color:#787878;font-size:12px}
    .email-attachments{margin-top:24px;padding-top:20px;border-top:1px solid #d9d9d9}
    .email-attachments h2{margin:0 0 10px;font-size:14px;color:#606060}
    .email-attachments ul{list-style:none;margin:0;padding:0}
    .email-attachments li{display:flex;justify-content:space-between;gap:20px;padding:7px 0;border-bottom:1px solid #eeeeee}
    .email-attachments small{flex:none;color:#777}
  `
    .replace(/\s+/g, ' ')
    .trim()
}

export async function prepareEmailMarkup(
  input: Uint8Array | ArrayBuffer,
  options: PrepareEmailDocumentOptions,
): Promise<PreparedEmailMarkup> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input)
  if (bytes.length === 0) throw new Error('Email file is empty')
  if (bytes.length > MAX_EMAIL_BYTES) throw new Error('Email file must be 50 MB or smaller')
  if (
    !Number.isFinite(options.maxAttachmentSizeMb) ||
    options.maxAttachmentSizeMb < 1 ||
    options.maxAttachmentSizeMb > 50
  ) {
    throw new Error('Attachment size limit must be from 1 to 50 MB')
  }

  const { default: PostalMime } = await import('postal-mime')
  let email: Email
  try {
    email = await PostalMime.parse(bytes, {
      attachmentEncoding: 'arraybuffer',
      maxHeadersSize: 2 * 1024 * 1024,
      maxNestingDepth: 20,
      maxRfc822NestingDepth: 5,
    })
  } catch {
    throw new Error('Email file is invalid or damaged')
  }

  const subject = email.subject?.trim() || options.labels.untitled
  const body = email.html
    ? sanitizeEmailHtml(email.html, email.attachments, options.labels)
    : `<div style="white-space:pre-wrap">${escapeHtml(email.text?.trim() || options.labels.emptyBody)}</div>`
  const rows = [
    headerRow(options.labels.from, email.from ? mailboxLabel(email.from) : ''),
    headerRow(options.labels.to, addressList(email.to)),
    options.includeAllRecipients ? headerRow(options.labels.cc, addressList(email.cc)) : '',
    options.includeAllRecipients ? headerRow(options.labels.bcc, addressList(email.bcc)) : '',
    headerRow(options.labels.date, formattedDate(email.date, options.locale)),
  ].join('')
  const listedAttachments = attachmentRows(email.attachments)
  const attachmentSection = listedAttachments
    ? `<section class="email-attachments"><h2>${escapeHtml(options.labels.attachments)}</h2><ul>${listedAttachments}</ul></section>`
    : ''
  const html = `<!doctype html><html><head><meta charset="utf-8"><meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src data:; style-src 'unsafe-inline'; base-uri 'none'; form-action 'none'"><meta name="viewport" content="width=${EMAIL_PAGE_WIDTH_CSS}"><title>${escapeHtml(subject)}</title><style>${emailStyles()}</style></head><body><main class="email-document"><header class="email-header"><h1 class="email-subject">${escapeHtml(subject)}</h1><dl class="email-meta">${rows}</dl></header><article class="email-body">${body}</article>${attachmentSection}</main></body></html>`

  return {
    html,
    attachments: options.includeAttachments
      ? uniqueAttachmentInputs(email.attachments, options.maxAttachmentSizeMb)
      : [],
  }
}

export async function prepareEmailDocumentForPdf(
  file: File,
  options: PrepareEmailDocumentOptions,
): Promise<PreparedEmailDocument> {
  const prepared = await prepareEmailMarkup(new Uint8Array(await file.arrayBuffer()), options)
  return {
    html: prepared.html,
    pages: await renderLocalHtmlPages(prepared.html, { maxPages: 40 }),
    attachments: prepared.attachments,
  }
}
