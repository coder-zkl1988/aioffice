import { describe, expect, it } from 'vitest'
import type { SignatureData } from '../src/renderer/SignatureDialog'
import { signatureDrawingForField } from '../src/renderer/signature-placement'

describe('signature field placement', () => {
  it('fits image signatures inside the field while preserving aspect ratio', () => {
    const signature: SignatureData = { kind: 'image', image: 'png', width: 400, height: 100 }
    const drawing = signatureDrawingForField(
      signature,
      { pageIndex: 2, fieldName: 'approver', rect: [100, 200, 300, 260] },
      { pw: 612, ph: 792, rot: 0 },
      [0, 0, 0],
    )

    expect(drawing).toMatchObject({ kind: 'image', pageIndex: 2, image: 'png' })
    if (drawing.kind !== 'image') throw new Error('Expected image signature')
    const width = drawing.rect[2] - drawing.rect[0]
    const height = drawing.rect[3] - drawing.rect[1]
    expect(width / height).toBeCloseTo(4)
    expect(drawing.rect[0]).toBeGreaterThanOrEqual(100)
    expect(drawing.rect[2]).toBeLessThanOrEqual(300)
    expect(drawing.rect[1]).toBeGreaterThanOrEqual(200)
    expect(drawing.rect[3]).toBeLessThanOrEqual(260)
  })

  it('maps stroke signatures into rotated signature fields', () => {
    const signature: SignatureData = {
      kind: 'strokes',
      paths: [[0, 0, 420, 150]],
      width: 420,
      height: 150,
    }
    const drawing = signatureDrawingForField(
      signature,
      { pageIndex: 0, fieldName: 'signer', rect: [40, 50, 200, 110] },
      { pw: 612, ph: 792, rot: 90 },
      [0.2, 0.3, 0.4],
    )

    expect(drawing).toMatchObject({ kind: 'ink', pageIndex: 0, color: [0.2, 0.3, 0.4] })
    if (drawing.kind !== 'ink') throw new Error('Expected ink signature')
    expect(drawing.paths[0]).toHaveLength(4)
    expect(drawing.paths[0]!.every(Number.isFinite)).toBe(true)
  })
})
