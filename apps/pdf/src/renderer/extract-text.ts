import type { PDFDocumentProxy } from 'pdfjs-dist'
import type {
  PdfExtractedTextBlock,
  PdfExtractedTextLink,
  PdfExtractedTextPage,
  PdfJsonPage,
  PdfJsonTextRun,
} from '@genoffice/pdf-tools'

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
  fontName?: string
}

interface RawTextStyle {
  fontFamily?: string
}

interface RawLinkAnnotation {
  subtype?: string
  rect?: number[]
  url?: string
}

interface TextRun {
  text: string
  x: number
  y: number
  width: number
  height: number
  bold: boolean
}

interface TextLine {
  text: string
  x: number
  y: number
  width: number
  height: number
  bold: boolean
}

const BULLET = /^[•▪◦‣⁃*-]\s*/u
const SENTENCE_END = /[.!?。！？：:；;]$/u
const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u
const CLOSING_PUNCTUATION = /^[,.;:!?，。！？；：、)\]}>》」』】]/u
const OPENING_PUNCTUATION = /[([{<《「『【]$/u

function median(values: number[]): number {
  if (values.length === 0) return 0
  const sorted = [...values].sort((left, right) => left - right)
  const middle = Math.floor(sorted.length / 2)
  return sorted.length % 2 === 0 ? (sorted[middle - 1]! + sorted[middle]!) / 2 : sorted[middle]!
}

function shouldInsertSpace(left: string, right: string): boolean {
  if (!left || !right || /\s$/u.test(left) || /^\s/u.test(right)) return false
  const leftCharacter = left.at(-1) ?? ''
  const rightCharacter = right.at(0) ?? ''
  if (CLOSING_PUNCTUATION.test(rightCharacter) || OPENING_PUNCTUATION.test(leftCharacter)) {
    return false
  }
  return !(CJK.test(leftCharacter) && CJK.test(rightCharacter))
}

function joinText(left: string, right: string, dehyphenate = false): string {
  if (!left) return right
  if (!right) return left
  if (dehyphenate && /[\p{L}\p{N}]-$/u.test(left) && /^\p{Ll}/u.test(right)) {
    return `${left.slice(0, -1)}${right}`
  }
  return `${left}${shouldInsertSpace(left, right) ? ' ' : ''}${right}`
}

function lineFromRuns(runs: TextRun[]): TextLine {
  const ordered = [...runs].sort((left, right) => left.x - right.x)
  let text = ''
  let previous: TextRun | undefined
  let boldWidth = 0
  for (const run of ordered) {
    if (previous && run.x - (previous.x + previous.width) > 0.5) {
      text = joinText(text, run.text)
    } else {
      text += run.text
    }
    if (run.bold) boldWidth += Math.max(run.width, run.text.length)
    previous = run
  }
  const x = Math.min(...ordered.map((run) => run.x))
  const right = Math.max(...ordered.map((run) => run.x + run.width))
  const height = Math.max(...ordered.map((run) => run.height))
  return {
    text: text.replace(/\s+/gu, ' ').trim(),
    x,
    y: ordered.reduce((sum, run) => sum + run.y, 0) / ordered.length,
    width: Math.max(0, right - x),
    height,
    bold: boldWidth >= Math.max(right - x, text.length) * 0.6,
  }
}

export function groupPdfTextRuns(runs: TextRun[]): TextLine[] {
  const horizontal = runs.filter((run) => run.text.trim() && run.height > 0)
  const groups: TextRun[][] = []
  for (const run of [...horizontal].sort((left, right) => right.y - left.y || left.x - right.x)) {
    const group = groups.find((candidate) => {
      const anchor = candidate[0]!
      return Math.abs(anchor.y - run.y) <= Math.max(2, Math.min(anchor.height, run.height) * 0.35)
    })
    if (group) group.push(run)
    else groups.push([run])
  }
  return groups.map(lineFromRuns).filter((line) => line.text.length > 0)
}

function columnOrderedLines(lines: TextLine[]): TextLine[] {
  if (lines.length < 8) return [...lines].sort((a, b) => b.y - a.y || a.x - b.x)
  const minX = Math.min(...lines.map((line) => line.x))
  const maxX = Math.max(...lines.map((line) => line.x + line.width))
  if (maxX - minX < 200) return [...lines].sort((a, b) => b.y - a.y || a.x - b.x)

  let best: { gutter: number; crossing: number; left: number; right: number } | undefined
  for (
    let gutter = minX + (maxX - minX) * 0.35;
    gutter <= minX + (maxX - minX) * 0.65;
    gutter += 4
  ) {
    let crossing = 0
    let left = 0
    let right = 0
    for (const line of lines) {
      const lineRight = line.x + line.width
      if (line.x < gutter - 5 && lineRight > gutter + 5) crossing++
      else if (lineRight <= gutter) left++
      else right++
    }
    if (!best || crossing < best.crossing) best = { gutter, crossing, left, right }
  }
  if (!best || best.left < 4 || best.right < 4 || best.crossing > Math.floor(lines.length * 0.25)) {
    return [...lines].sort((a, b) => b.y - a.y || a.x - b.x)
  }
  const spanning = lines.filter(
    (line) => line.x < best.gutter - 5 && line.x + line.width > best.gutter + 5,
  )
  const left = lines.filter((line) => line.x + line.width <= best.gutter)
  const right = lines.filter((line) => line.x >= best.gutter)
  const sort = (items: TextLine[]) => items.sort((a, b) => b.y - a.y || a.x - b.x)
  return [...sort(spanning), ...sort(left), ...sort(right)]
}

export function pdfTextLinesToBlocks(lines: TextLine[]): PdfExtractedTextBlock[] {
  if (lines.length === 0) return []
  const ordered = columnOrderedLines(lines)
  const bodyHeight = median(ordered.map((line) => line.height).filter((height) => height > 0)) || 12
  const blocks: PdfExtractedTextBlock[] = []
  let paragraph = ''
  let previous: TextLine | undefined
  const flush = () => {
    if (!paragraph) return
    blocks.push({ kind: 'paragraph', text: paragraph })
    paragraph = ''
  }

  for (const line of ordered) {
    const ratio = line.height / bodyHeight
    const headingLevel = ratio >= 1.75 ? 1 : ratio >= 1.45 ? 2 : ratio >= 1.25 ? 3 : undefined
    if (headingLevel) {
      flush()
      blocks.push({ kind: 'heading', text: line.text, level: headingLevel })
      previous = line
      continue
    }
    if (BULLET.test(line.text)) {
      flush()
      blocks.push({ kind: 'listItem', text: line.text })
      previous = line
      continue
    }
    const gap = previous ? previous.y - line.y - Math.max(previous.height, line.height) : 0
    const changedColumn = previous
      ? Math.abs(line.x - previous.x) > Math.max(48, bodyHeight * 4)
      : false
    const paragraphBreak =
      paragraph.length > 0 &&
      (gap > Math.max(previous?.height ?? bodyHeight, line.height) * 0.8 ||
        (changedColumn && SENTENCE_END.test(paragraph)))
    if (paragraphBreak) flush()
    paragraph = joinText(paragraph, line.text, true)
    previous = line
  }
  flush()
  return blocks
}

function intersects(run: TextRun, rect: number[]): boolean {
  const left = Math.min(rect[0] ?? 0, rect[2] ?? 0)
  const right = Math.max(rect[0] ?? 0, rect[2] ?? 0)
  const bottom = Math.min(rect[1] ?? 0, rect[3] ?? 0)
  const top = Math.max(rect[1] ?? 0, rect[3] ?? 0)
  return run.x < right && run.x + run.width > left && run.y < top && run.y + run.height > bottom
}

function externalLinks(runs: TextRun[], annotations: RawLinkAnnotation[]): PdfExtractedTextLink[] {
  const links = annotations
    .filter(
      (annotation) =>
        annotation.subtype === 'Link' &&
        typeof annotation.url === 'string' &&
        annotation.url.trim().length > 0,
    )
    .map((annotation) => {
      const linkedRuns = annotation.rect
        ? runs.filter((run) => intersects(run, annotation.rect!))
        : []
      const label = linkedRuns.length > 0 ? lineFromRuns(linkedRuns).text : ''
      return { url: annotation.url!.trim(), ...(label ? { label } : {}) }
    })
  return [...new Map(links.map((link) => [link.url, link])).values()]
}

export function extractPdfTextPageData(
  pageNumber: number,
  items: RawTextItem[],
  styles: Record<string, RawTextStyle>,
  annotations: RawLinkAnnotation[],
): PdfExtractedTextPage {
  const runs: TextRun[] = items.flatMap((item) => {
    if (!item.str?.trim() || !item.transform || item.transform.length < 6) return []
    const height = item.height || Math.hypot(item.transform[2] ?? 0, item.transform[3] ?? 0) || 1
    const font = `${item.fontName ?? ''} ${styles[item.fontName ?? '']?.fontFamily ?? ''}`
    return [
      {
        text: item.str,
        x: item.transform[4] ?? 0,
        y: item.transform[5] ?? 0,
        width: Math.max(0, item.width ?? 0),
        height,
        bold: /bold|black|heavy|semibold|demi/i.test(font),
      },
    ]
  })
  const blocks = pdfTextLinesToBlocks(groupPdfTextRuns(runs))
  return {
    pageNumber,
    text: blocks.map((block) => block.text).join('\n\n'),
    blocks,
    links: externalLinks(runs, annotations),
  }
}

export async function extractPdfTextPages(
  document: PDFDocumentProxy,
  pageIndexes: number[],
): Promise<PdfExtractedTextPage[]> {
  const pages: PdfExtractedTextPage[] = []
  for (const pageIndex of pageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= document.numPages) {
      throw new Error('Text export page indexes contain an invalid page')
    }
    const page = await document.getPage(pageIndex + 1)
    try {
      const [content, annotations] = await Promise.all([
        page.getTextContent(),
        page.getAnnotations({ intent: 'display' }),
      ])
      pages.push(
        extractPdfTextPageData(
          pageIndex + 1,
          content.items as RawTextItem[],
          content.styles as Record<string, RawTextStyle>,
          annotations as RawLinkAnnotation[],
        ),
      )
    } finally {
      page.cleanup()
    }
  }
  return pages
}

export function pdfJsonTextRuns(
  items: RawTextItem[],
  styles: Record<string, RawTextStyle>,
): PdfJsonTextRun[] {
  return items.flatMap((item) => {
    if (!item.str || !item.transform || item.transform.length < 6) return []
    const height = item.height || Math.hypot(item.transform[2] ?? 0, item.transform[3] ?? 0) || 1
    const fontFamily = styles[item.fontName ?? '']?.fontFamily
    const font = `${item.fontName ?? ''} ${fontFamily ?? ''}`
    return [
      {
        text: item.str,
        x: item.transform[4] ?? 0,
        y: item.transform[5] ?? 0,
        width: Math.max(0, item.width ?? 0),
        height,
        fontSize: height,
        ...(fontFamily ? { fontFamily } : {}),
        bold: /bold|black|heavy|semibold|demi/i.test(font),
        italic: /italic|oblique/i.test(font),
      },
    ]
  })
}

export async function extractPdfJsonPages(
  document: PDFDocumentProxy,
  pageIndexes: number[],
  lightweight: boolean,
): Promise<PdfJsonPage[]> {
  const pages: PdfJsonPage[] = []
  for (const pageIndex of pageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= document.numPages) {
      throw new Error('Structured export page indexes contain an invalid page')
    }
    const page = await document.getPage(pageIndex + 1)
    try {
      const [content, annotations] = await Promise.all([
        page.getTextContent(),
        page.getAnnotations({ intent: 'display' }),
      ])
      const items = content.items as RawTextItem[]
      const styles = content.styles as Record<string, RawTextStyle>
      const semantic = extractPdfTextPageData(
        pageIndex + 1,
        items,
        styles,
        annotations as RawLinkAnnotation[],
      )
      const viewport = page.getViewport({ scale: 1 })
      pages.push({
        ...semantic,
        width: viewport.width,
        height: viewport.height,
        rotation: viewport.rotation,
        ...(!lightweight ? { textRuns: pdfJsonTextRuns(items, styles) } : {}),
      })
    } finally {
      page.cleanup()
    }
  }
  return pages
}
