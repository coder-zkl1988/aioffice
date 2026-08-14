const SAFE_LOCAL_TAGS = new Set([
  'a',
  'abbr',
  'address',
  'article',
  'b',
  'blockquote',
  'br',
  'caption',
  'cite',
  'code',
  'dd',
  'del',
  'details',
  'dfn',
  'div',
  'dl',
  'dt',
  'em',
  'figcaption',
  'figure',
  'footer',
  'h1',
  'h2',
  'h3',
  'h4',
  'h5',
  'h6',
  'header',
  'hr',
  'i',
  'img',
  'kbd',
  'li',
  'main',
  'mark',
  'nav',
  'ol',
  'p',
  'pre',
  'q',
  's',
  'samp',
  'section',
  'small',
  'span',
  'strong',
  'sub',
  'summary',
  'sup',
  'table',
  'tbody',
  'td',
  'tfoot',
  'th',
  'thead',
  'time',
  'tr',
  'u',
  'ul',
  'var',
])

const DROP_LOCAL_TAGS = new Set([
  'applet',
  'audio',
  'base',
  'button',
  'canvas',
  'embed',
  'form',
  'frame',
  'frameset',
  'head',
  'iframe',
  'input',
  'link',
  'meta',
  'noscript',
  'object',
  'option',
  'script',
  'select',
  'source',
  'style',
  'svg',
  'template',
  'textarea',
  'track',
  'video',
])

const SAFE_STYLE_PROPERTIES = new Set([
  'align-content',
  'align-items',
  'align-self',
  'background-color',
  'border',
  'border-bottom',
  'border-collapse',
  'border-color',
  'border-left',
  'border-radius',
  'border-right',
  'border-spacing',
  'border-style',
  'border-top',
  'border-width',
  'box-shadow',
  'box-sizing',
  'break-after',
  'break-before',
  'break-inside',
  'clear',
  'color',
  'column-count',
  'column-gap',
  'display',
  'flex',
  'flex-basis',
  'flex-direction',
  'flex-flow',
  'flex-grow',
  'flex-shrink',
  'flex-wrap',
  'float',
  'font-family',
  'font-size',
  'font-style',
  'font-variant',
  'font-weight',
  'gap',
  'grid-auto-columns',
  'grid-auto-flow',
  'grid-auto-rows',
  'grid-column',
  'grid-row',
  'grid-template-columns',
  'grid-template-rows',
  'height',
  'justify-content',
  'justify-items',
  'justify-self',
  'letter-spacing',
  'line-height',
  'list-style-position',
  'list-style-type',
  'margin',
  'margin-bottom',
  'margin-left',
  'margin-right',
  'margin-top',
  'max-height',
  'max-width',
  'min-height',
  'min-width',
  'object-fit',
  'object-position',
  'opacity',
  'overflow',
  'overflow-wrap',
  'overflow-x',
  'overflow-y',
  'padding',
  'padding-bottom',
  'padding-left',
  'padding-right',
  'padding-top',
  'page-break-after',
  'page-break-before',
  'page-break-inside',
  'position',
  'table-layout',
  'text-align',
  'text-decoration',
  'text-indent',
  'text-overflow',
  'text-transform',
  'transform',
  'transform-origin',
  'vertical-align',
  'visibility',
  'white-space',
  'width',
  'word-break',
  'word-spacing',
])

const UNSAFE_STYLE_VALUE = /(?:expression|url|javascript|@import|var)\s*\(/i
const SAFE_SELECTOR = /^[\w\s#.:,[\]="'~+>^$*|()-]+$/u

export interface LocalHtmlSanitizeOptions {
  basePath: string
  preserveSelectors?: boolean
  resolveImage: (source: string, basePath: string) => Promise<string | undefined>
}

export interface SanitizedLocalHtml {
  html: string
  title: string
}

export function normalizeLocalResourcePath(
  value: string,
  basePath = '',
  label = 'Document',
): string {
  let decoded: string
  try {
    decoded = decodeURIComponent(value.split('#')[0]!.replace(/\\/g, '/'))
  } catch {
    throw new Error(`${label} contains an unsafe resource path`)
  }
  const raw = decoded.startsWith('/') ? decoded.slice(1) : `${basePath}${decoded}`
  const segments: string[] = []
  for (const segment of raw.split('/')) {
    if (!segment || segment === '.') continue
    if (segment === '..') {
      if (segments.length === 0) throw new Error(`${label} contains an unsafe resource path`)
      segments.pop()
      continue
    }
    if (segment.includes('\0')) throw new Error(`${label} contains an unsafe resource path`)
    segments.push(segment)
  }
  const output = segments.join('/')
  if (!output || /^[a-z]:\//i.test(output)) {
    throw new Error(`${label} contains an unsafe resource path`)
  }
  return output
}

export function localResourceDirectory(path: string): string {
  const slash = path.lastIndexOf('/')
  return slash < 0 ? '' : path.slice(0, slash + 1)
}

export function localBytesToBase64(bytes: Uint8Array): string {
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function sanitizedStyleDeclaration(style: CSSStyleDeclaration): string {
  const output: string[] = []
  for (const property of Array.from(style)) {
    if (!SAFE_STYLE_PROPERTIES.has(property)) continue
    const value = style.getPropertyValue(property).trim()
    if (!value || UNSAFE_STYLE_VALUE.test(value)) continue
    if (/^-/.test(value) && /^(?:height|margin|max-|min-|padding|width)/.test(property)) continue
    if (property === 'position' && /^(?:fixed|sticky)$/i.test(value)) continue
    output.push(`${property}:${value}`)
  }
  return output.join(';')
}

export function sanitizeLocalStyle(value: string): string {
  const probe = document.createElement('span')
  probe.setAttribute('style', value)
  return sanitizedStyleDeclaration(probe.style)
}

function safeSelector(selector: string): boolean {
  return (
    selector.length <= 500 &&
    SAFE_SELECTOR.test(selector) &&
    !/(?:@|:has\(|:visited|:active|:focus|:target)/i.test(selector)
  )
}

export function sanitizeLocalStylesheet(css: string): string {
  if (!css.trim()) return ''
  const style = document.createElement('style')
  style.textContent = css
    .slice(0, 5 * 1024 * 1024)
    .replace(/@import\s+(?:url\([^)]*\)|["'][^"']*["'])[^;]*;?/giu, '')
    .replace(/url\([^)]*\)/giu, 'none')
  document.head.appendChild(style)
  try {
    const rules = Array.from(style.sheet?.cssRules ?? [])
    return rules
      .flatMap((rule) => {
        const candidate = rule as CSSRule & {
          selectorText?: string
          style?: CSSStyleDeclaration
        }
        if (!candidate.selectorText || !candidate.style || !safeSelector(candidate.selectorText)) {
          return []
        }
        const declaration = sanitizedStyleDeclaration(candidate.style)
        return declaration ? [`${candidate.selectorText}{${declaration}}`] : []
      })
      .join('')
  } catch {
    return ''
  } finally {
    style.remove()
  }
}

function safeClassName(value: string): string {
  return value
    .split(/\s+/)
    .filter((token) => /^[a-z0-9_-]{1,80}$/i.test(token))
    .slice(0, 20)
    .join(' ')
}

async function appendSanitizedNode(
  source: Node,
  parent: Node,
  options: LocalHtmlSanitizeOptions,
): Promise<void> {
  if (source.nodeType === Node.TEXT_NODE) {
    parent.appendChild(document.createTextNode(source.textContent ?? ''))
    return
  }
  if (!(source instanceof Element)) return
  const tag = source.localName.toLowerCase()
  if (DROP_LOCAL_TAGS.has(tag)) return
  if (!SAFE_LOCAL_TAGS.has(tag)) {
    for (const child of source.childNodes) await appendSanitizedNode(child, parent, options)
    return
  }
  if (tag === 'img') {
    const url = await options.resolveImage(
      source.getAttribute('src') ?? source.getAttribute('href') ?? '',
      options.basePath,
    )
    if (!url) return
    const image = document.createElement('img')
    image.src = url
    image.alt = source.getAttribute('alt')?.slice(0, 300) ?? ''
    image.loading = 'eager'
    image.decoding = 'sync'
    const style = sanitizeLocalStyle(source.getAttribute('style') ?? '')
    if (style) image.setAttribute('style', style)
    if (options.preserveSelectors) {
      const className = safeClassName(source.getAttribute('class') ?? '')
      const id = source.getAttribute('id') ?? ''
      if (className) image.className = className
      if (/^[a-z][a-z0-9_-]{0,79}$/i.test(id)) image.id = id
    }
    parent.appendChild(image)
    return
  }
  const output = document.createElement(tag)
  const style = sanitizeLocalStyle(source.getAttribute('style') ?? '')
  if (style) output.setAttribute('style', style)
  if (options.preserveSelectors) {
    const className = safeClassName(source.getAttribute('class') ?? '')
    const id = source.getAttribute('id') ?? ''
    if (className) output.className = className
    if (/^[a-z][a-z0-9_-]{0,79}$/i.test(id)) output.id = id
  }
  const direction = source.getAttribute('dir')?.toLowerCase()
  if (direction === 'ltr' || direction === 'rtl' || direction === 'auto') {
    output.setAttribute('dir', direction)
  }
  const language = source.getAttribute('lang')?.trim()
  if (language && /^[a-z0-9-]{1,35}$/i.test(language)) output.setAttribute('lang', language)
  if (tag === 'a') {
    const href = source.getAttribute('href')?.trim() ?? ''
    if (/^(?:https?:|mailto:)/i.test(href)) output.setAttribute('href', href)
  }
  if (tag === 'td' || tag === 'th') {
    for (const attribute of ['colspan', 'rowspan'] as const) {
      const value = Number(source.getAttribute(attribute))
      if (Number.isInteger(value) && value >= 1 && value <= 100) {
        output.setAttribute(attribute, String(value))
      }
    }
  }
  const pageBreak = `${source.getAttribute('style') ?? ''}`.match(
    /(?:break-before|page-break-before)\s*:\s*(?:page|always)/i,
  )
  if (pageBreak) output.setAttribute('data-pdf-page-break-before', '')
  for (const child of source.childNodes) await appendSanitizedNode(child, output, options)
  parent.appendChild(output)
}

export async function sanitizeLocalHtmlFragment(
  markup: string,
  options: LocalHtmlSanitizeOptions,
): Promise<SanitizedLocalHtml> {
  const parsed = new DOMParser().parseFromString(markup, 'text/html')
  const output = document.createElement('div')
  for (const child of parsed.body.childNodes) {
    await appendSanitizedNode(child, output, options)
  }
  const title =
    parsed.querySelector('title')?.textContent?.trim() ||
    output.querySelector('h1,h2,h3')?.textContent?.trim() ||
    ''
  return { title: title.slice(0, 500), html: output.innerHTML }
}
