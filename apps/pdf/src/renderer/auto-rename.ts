import type { PdfAutoRenameStrategy } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'

interface RawTextItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
}

export interface PdfAutoRenameLine {
  text: string
  fontSize: number
  pageNumber: number
}

interface TextRun extends PdfAutoRenameLine {
  x: number
  y: number
  width: number
}

const CJK = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}]/u

function needsSpace(left: string, right: string): boolean {
  if (!left || !right || /\s$/u.test(left) || /^\s/u.test(right)) return false
  const leftCharacter = left.at(-1) ?? ''
  const rightCharacter = right.at(0) ?? ''
  if (/^[,.;:!?，。！？；：、)\]}>》」』】]/u.test(rightCharacter)) return false
  return !(CJK.test(leftCharacter) && CJK.test(rightCharacter))
}

function pageLines(items: RawTextItem[], pageNumber: number): PdfAutoRenameLine[] {
  const runs: TextRun[] = items.flatMap((item) => {
    if (!item.str?.trim() || !item.transform || item.transform.length < 6) return []
    const fontSize = item.height || Math.hypot(item.transform[2] ?? 0, item.transform[3] ?? 0) || 1
    return [
      {
        text: item.str,
        fontSize,
        pageNumber,
        x: item.transform[4] ?? 0,
        y: item.transform[5] ?? 0,
        width: Math.max(0, item.width ?? 0),
      },
    ]
  })
  const groups: TextRun[][] = []
  for (const run of [...runs].sort((left, right) => right.y - left.y || left.x - right.x)) {
    const group = groups.find((candidate) => {
      const anchor = candidate[0]!
      return (
        Math.abs(anchor.y - run.y) <= Math.max(1.5, Math.min(anchor.fontSize, run.fontSize) * 0.25)
      )
    })
    if (group) group.push(run)
    else groups.push([run])
  }
  return groups.flatMap((group) => {
    const ordered = [...group].sort((left, right) => left.x - right.x)
    let text = ''
    let previous: TextRun | undefined
    for (const run of ordered) {
      if (previous && run.x - (previous.x + previous.width) > 0.5 && needsSpace(text, run.text)) {
        text += ' '
      }
      text += run.text
      previous = run
    }
    text = text.replace(/\s+/gu, ' ').trim()
    return text
      ? [{ text, fontSize: Math.max(...ordered.map((run) => run.fontSize)), pageNumber }]
      : []
  })
}

export function inferPdfAutoRenameTitleFromLines(
  lines: PdfAutoRenameLine[],
  strategy: PdfAutoRenameStrategy,
): string | undefined {
  const candidates = lines.slice(0, 200).filter((line) => line.text.trim())
  if (candidates.length === 0) return undefined
  if (strategy === 'firstText') return candidates[0]!.text.trim()

  let bestIndex = 0
  for (let index = 1; index < candidates.length; index++) {
    if (candidates[index]!.fontSize > candidates[bestIndex]!.fontSize) bestIndex = index
  }
  const best = candidates[bestIndex]!
  const titleLines = [best.text]
  for (let index = bestIndex + 1; index < candidates.length && titleLines.length < 4; index++) {
    const line = candidates[index]!
    if (line.pageNumber !== best.pageNumber || Math.abs(line.fontSize - best.fontSize) > 0.1) break
    titleLines.push(line.text)
    if (titleLines.join(' ').length >= 240) break
  }
  return titleLines.join(' ').replace(/\s+/gu, ' ').trim()
}

export async function inferPdfAutoRenameTitle(
  document: PDFDocumentProxy,
  strategy: PdfAutoRenameStrategy,
): Promise<string | undefined> {
  const lines: PdfAutoRenameLine[] = []
  for (let pageIndex = 0; pageIndex < document.numPages && lines.length < 200; pageIndex++) {
    const page = await document.getPage(pageIndex + 1)
    try {
      const content = await page.getTextContent()
      lines.push(...pageLines(content.items as RawTextItem[], pageIndex + 1))
    } finally {
      page.cleanup()
    }
  }
  return inferPdfAutoRenameTitleFromLines(lines, strategy)
}
