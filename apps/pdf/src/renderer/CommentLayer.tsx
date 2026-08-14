import { useEffect, useMemo, useState } from 'react'
import type { CSSProperties, ReactElement } from 'react'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { geomDispSize, pdfRectToCss } from './annotations'
import type { PageGeom } from './annotations'

interface RawTextAnnotation {
  id?: string
  subtype?: string
  rect?: number[]
  color?: ArrayLike<number>
  contentsObj?: { str?: string }
  titleObj?: { str?: string }
  subjectObj?: { str?: string }
}

export interface PdfTextComment {
  id: string
  rect: [number, number, number, number]
  color: string
  text: string
  author: string
  subject?: string
}

function annotationColor(color: ArrayLike<number> | undefined): string {
  if (!color || color.length < 3) return 'rgb(255 242 102)'
  return `rgb(${color[0] ?? 255} ${color[1] ?? 242} ${color[2] ?? 102})`
}

export function textCommentsFromAnnotations(annotations: RawTextAnnotation[]): PdfTextComment[] {
  return annotations.flatMap((annotation, index) => {
    if (annotation.subtype !== 'Text' || !annotation.rect || annotation.rect.length < 4) return []
    const rect = annotation.rect.slice(0, 4)
    if (rect.some((value) => !Number.isFinite(value))) return []
    return [
      {
        id: annotation.id ?? `text-comment-${index}`,
        rect: rect as PdfTextComment['rect'],
        color: annotationColor(annotation.color),
        text: annotation.contentsObj?.str ?? '',
        author: annotation.titleObj?.str ?? '',
        ...(annotation.subjectObj?.str ? { subject: annotation.subjectObj.str } : {}),
      },
    ]
  })
}

export function CommentLayer({
  doc,
  pageNo,
  geom,
  scale,
}: {
  doc: PDFDocumentProxy
  pageNo: number
  geom: PageGeom
  scale: number
}): ReactElement | null {
  const [comments, setComments] = useState<PdfTextComment[] | null>(null)
  const [openId, setOpenId] = useState<string | null>(null)

  useEffect(() => {
    let cancelled = false
    void (async () => {
      const page = await doc.getPage(pageNo)
      const annotations = (await page.getAnnotations()) as RawTextAnnotation[]
      if (!cancelled) setComments(textCommentsFromAnnotations(annotations))
    })()
    return () => {
      cancelled = true
    }
  }, [doc, pageNo])

  useEffect(() => {
    if (!openId) return
    const close = (event: PointerEvent) => {
      if (!(event.target as Element | null)?.closest?.('.pdf-existing-comment')) setOpenId(null)
    }
    window.addEventListener('pointerdown', close)
    return () => window.removeEventListener('pointerdown', close)
  }, [openId])

  const displaySize = useMemo(() => geomDispSize(geom), [geom])
  if (!comments || comments.length === 0) return null

  return (
    <div className="pdf-existing-comment-layer">
      {comments.map((comment) => {
        const box = pdfRectToCss(geom, comment.rect, scale)
        const popupWidth = Math.min(260, displaySize.width * scale - 16)
        const popupStyle: CSSProperties = {
          left: Math.max(
            8,
            Math.min(
              box.left + Math.max(box.width, 20) + 8,
              displaySize.width * scale - popupWidth - 8,
            ),
          ),
          top: Math.max(8, box.top),
          width: popupWidth,
        }
        return (
          <div className="pdf-existing-comment" key={comment.id}>
            <button
              className="pdf-existing-comment-pin"
              type="button"
              style={{ ...box, background: comment.color }}
              aria-label={comment.text || 'PDF comment'}
              aria-expanded={openId === comment.id}
              onClick={() => setOpenId((current) => (current === comment.id ? null : comment.id))}
            >
              <svg
                width="11"
                height="11"
                viewBox="0 0 16 16"
                fill="none"
                stroke="#4f4700"
                strokeWidth="1.6"
                aria-hidden
              >
                <path d="M2.5 3.5h11v8h-6l-3 2.5V11.5h-2z" strokeLinejoin="round" />
              </svg>
            </button>
            {openId === comment.id && (
              <div className="pdf-existing-comment-popup" role="note" style={popupStyle}>
                {comment.subject && <strong>{comment.subject}</strong>}
                {comment.author && <span>{comment.author}</span>}
                <p>{comment.text}</p>
              </div>
            )}
          </div>
        )
      })}
    </div>
  )
}
