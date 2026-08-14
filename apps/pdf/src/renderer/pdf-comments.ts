import type { PdfCommentInput } from '@genoffice/pdf-tools'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { buildSearchIndex, searchInIndex } from './search'

export async function resolvePdfCommentAnchors(
  document: PDFDocumentProxy,
  comments: PdfCommentInput[],
): Promise<PdfCommentInput[]> {
  if (!comments.some((comment) => comment.anchorText?.trim())) return comments
  const index = await buildSearchIndex(document)
  return comments.map((comment) => {
    const anchorText = comment.anchorText?.trim()
    const page = index[comment.pageIndex]
    if (!anchorText || !page) return comment
    const match = searchInIndex([page], anchorText)[0]
    const rect = match?.rects[0]
    if (!rect) return comment
    return {
      ...comment,
      x: rect[0],
      y: rect[3] - comment.height,
    }
  })
}
