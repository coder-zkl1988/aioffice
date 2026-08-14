import type { DrawingInput } from '../shared/ipc'
import { pdfRectToCss, viewToPdf, type PageGeom } from './annotations'
import type { SignatureData } from './SignatureDialog'

export interface SignatureFieldTarget {
  pageIndex: number
  fieldName: string
  rect: [number, number, number, number]
}

export function signatureDrawingForField(
  signature: SignatureData,
  target: SignatureFieldTarget,
  geom: PageGeom,
  color: [number, number, number],
): DrawingInput {
  const box = pdfRectToCss(geom, target.rect, 1)
  const padding = Math.min(4, box.width * 0.08, box.height * 0.08)
  const availableWidth = Math.max(1, box.width - padding * 2)
  const availableHeight = Math.max(1, box.height - padding * 2)
  const factor = Math.min(availableWidth / signature.width, availableHeight / signature.height)
  const width = signature.width * factor
  const height = signature.height * factor
  const left = box.left + (box.width - width) / 2
  const top = box.top + (box.height - height) / 2

  if (signature.kind === 'image') {
    const [ax, ay] = viewToPdf(geom, left, top)
    const [bx, by] = viewToPdf(geom, left + width, top + height)
    return {
      kind: 'image',
      pageIndex: target.pageIndex,
      image: signature.image,
      rect: [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)],
    }
  }

  return {
    kind: 'ink',
    pageIndex: target.pageIndex,
    color,
    width: 1.6,
    paths: signature.paths.map((path) => {
      const output: number[] = []
      for (let index = 0; index < path.length; index += 2) {
        output.push(
          ...viewToPdf(geom, left + path[index]! * factor, top + path[index + 1]! * factor),
        )
      }
      return output
    }),
  }
}
