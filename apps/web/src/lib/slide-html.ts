export interface WebSlideHtmlRequest {
  brief: string
  title?: string
  styleSkill?: string
  deckContext?: Record<string, unknown>
  images?: { url: string; caption?: string }[]
  width?: number
  height?: number
}

export interface WebSlideHtmlPrompts {
  system: string
  user: string
}

function boundedDimension(value: number | undefined, fallback: number): number {
  if (!Number.isFinite(value)) return fallback
  return Math.min(3840, Math.max(320, Math.round(value!)))
}

export function webSlideDimensions(request: WebSlideHtmlRequest): {
  width: number
  height: number
} {
  return {
    width: boundedDimension(request.width, 1280),
    height: boundedDimension(request.height, 720),
  }
}

export function buildWebSlideHtmlPrompts(request: WebSlideHtmlRequest): WebSlideHtmlPrompts {
  const { width, height } = webSlideDimensions(request)
  const images = (request.images ?? [])
    .filter((image) => /^https?:\/\//i.test(image.url))
    .slice(0, 8)
  const approvedImages = images.length
    ? images
        .map(
          (image, index) =>
            `${index + 1}. ${image.url}${image.caption ? ` - ${image.caption}` : ''}`,
        )
        .join('\n')
    : '(none)'

  return {
    system: [
      'You are a senior presentation designer and frontend engineer.',
      `Create exactly one polished presentation slide as a standalone HTML document sized ${width}x${height}px.`,
      'Return only the complete HTML document. Do not use Markdown fences or add explanations.',
      'Use embedded CSS only. Do not use JavaScript, forms, iframes, external stylesheets, external fonts, SVG scripts, or animation.',
      'Keep every visible element inside the fixed canvas. The body must not scroll or overflow.',
      'Use clear information hierarchy, concise copy, strong alignment, sufficient contrast, and presentation-scale typography.',
      'Preserve the language of the supplied content. Do not invent facts, figures, names, sources, or claims.',
      'Use only the approved image URLs supplied by the user. Put images in img elements with object-fit; do not invent asset URLs.',
      'When no image URL is supplied, create the composition with typography, solid colors, borders, and simple CSS shapes.',
      'The final body must contain meaningful visible content and must be ready for a browser screenshot without interaction.',
    ].join('\n'),
    user: [
      `Canvas: ${width}x${height}px`,
      `Slide title: ${request.title?.trim() || '(derive a concise title from the brief)'}`,
      `Slide brief:\n${request.brief.trim()}`,
      `Deck design system:\n${request.styleSkill?.trim() || '(use a clean, modern professional system)'}`,
      `Deck context:\n${JSON.stringify(request.deckContext ?? {}, null, 2)}`,
      `Approved image URLs:\n${approvedImages}`,
    ].join('\n\n'),
  }
}

function extractHtmlSource(value: string): string {
  const trimmed = value.trim()
  if (!trimmed) throw new Error('模型未返回幻灯片 HTML')
  if (trimmed.length > 750_000) throw new Error('模型返回的幻灯片 HTML 过大')

  const fenced = /```(?:html)?\s*([\s\S]*?)```/i.exec(trimmed)
  const candidate = (fenced?.[1] ?? trimmed).trim()
  const starts = [
    candidate.search(/<!doctype\s+html/i),
    candidate.search(/<html\b/i),
    candidate.search(/<body\b/i),
  ].filter((index) => index >= 0)
  if (starts.length === 0) return candidate
  const start = Math.min(...starts)
  const htmlEnd = candidate.toLowerCase().lastIndexOf('</html>')
  return candidate.slice(start, htmlEnd >= start ? htmlEnd + 7 : undefined)
}

function unquoteCssUrl(value: string): string {
  const trimmed = value.trim()
  if (
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
  ) {
    return trimmed.slice(1, -1).trim()
  }
  return trimmed
}

function isAllowedAssetUrl(value: string, allowedImages: Set<string>): boolean {
  return /^data:image\/(?:png|jpeg|jpg|gif|webp);base64,/i.test(value) || allowedImages.has(value)
}

function isAllowedCssUrl(value: string, allowedImages: Set<string>): boolean {
  return /^#[a-z][\w:.-]*$/i.test(value) || isAllowedAssetUrl(value, allowedImages)
}

function sanitizeCss(css: string, allowedImages: Set<string>): string {
  return css
    .replace(/@import\s+(?:url\([^)]*\)|["'][^"']*["'])\s*;?/gi, '')
    .replace(/url\(([^)]*)\)/gi, (match, rawUrl: string) =>
      isAllowedCssUrl(unquoteCssUrl(rawUrl), allowedImages) ? match : 'none',
    )
}

export function sanitizeGeneratedSlideHtml(
  value: string,
  request: Pick<WebSlideHtmlRequest, 'width' | 'height' | 'images'> = {},
): string {
  if (typeof DOMParser === 'undefined') throw new Error('当前浏览器不支持解析幻灯片 HTML')
  const { width, height } = webSlideDimensions({ brief: '', ...request })
  const allowedImages = new Set(
    (request.images ?? [])
      .map((image) => image.url.trim())
      .filter((url) => /^https?:\/\//i.test(url)),
  )
  const documentNode = new DOMParser().parseFromString(extractHtmlSource(value), 'text/html')

  documentNode
    .querySelectorAll(
      'script, iframe, frame, object, embed, base, link, form, input, button, textarea, select, video, audio, source, track, foreignObject',
    )
    .forEach((element) => element.remove())
  documentNode.querySelectorAll('meta[http-equiv]').forEach((element) => element.remove())

  for (const element of documentNode.querySelectorAll<HTMLElement>('*')) {
    for (const attribute of [...element.attributes]) {
      const name = attribute.name.toLowerCase()
      const attributeValue = attribute.value.trim()
      if (name.startsWith('on') || name === 'srcdoc' || name === 'srcset') {
        element.removeAttribute(attribute.name)
      } else if (name === 'style') {
        element.setAttribute('style', sanitizeCss(attribute.value, allowedImages))
      } else if (name === 'href' || name.endsWith(':href') || name === 'poster') {
        element.removeAttribute(attribute.name)
      } else if (name === 'src' && !isAllowedAssetUrl(attributeValue, allowedImages)) {
        element.removeAttribute(attribute.name)
      } else if (attributeValue.includes('url(')) {
        element.setAttribute(attribute.name, sanitizeCss(attribute.value, allowedImages))
      }
    }
  }

  for (const style of documentNode.querySelectorAll('style')) {
    style.textContent = sanitizeCss(style.textContent || '', allowedImages)
  }

  if (documentNode.body.childElementCount === 0 || !documentNode.body.textContent?.trim()) {
    throw new Error('模型返回的幻灯片没有有效内容')
  }

  const reset = documentNode.createElement('style')
  reset.setAttribute('data-genoffice-slide-reset', '')
  reset.textContent = `html,body{width:${width}px!important;height:${height}px!important;margin:0!important;padding:0!important;overflow:hidden!important}body{position:relative!important;box-sizing:border-box!important}*,*::before,*::after{box-sizing:border-box}`
  documentNode.head.prepend(reset)
  return `<!doctype html>\n${documentNode.documentElement.outerHTML}`
}
