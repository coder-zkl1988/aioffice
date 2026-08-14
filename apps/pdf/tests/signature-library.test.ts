import { describe, expect, it } from 'vitest'
import type { SignatureData } from '../src/renderer/SignatureDialog'
import {
  MAX_SAVED_SIGNATURES,
  MAX_SIGNATURE_BYTES,
  MAX_SIGNATURE_LIBRARY_BYTES,
  signatureDataByteSize,
  validateSignatureLibraryAddition,
} from '../src/renderer/signature-library'

const strokeSignature: SignatureData = {
  kind: 'strokes',
  paths: [[0, 0, 10, 10]],
  width: 420,
  height: 150,
}

describe('signature library limits', () => {
  it('measures serialized signature data', () => {
    expect(signatureDataByteSize(strokeSignature)).toBeGreaterThan(20)
  })

  it('accepts a small signature in an empty library', () => {
    expect(validateSignatureLibraryAddition([], strokeSignature)).toBeNull()
  })

  it('limits the number of saved signatures', () => {
    const saved = Array.from({ length: MAX_SAVED_SIGNATURES }, () => ({ byteSize: 1 }))
    expect(validateSignatureLibraryAddition(saved, strokeSignature)).toBe('item-limit')
  })

  it('rejects an oversized signature', () => {
    const signature: SignatureData = {
      kind: 'image',
      image: 'a'.repeat(MAX_SIGNATURE_BYTES),
      width: 100,
      height: 50,
    }
    expect(validateSignatureLibraryAddition([], signature)).toBe('item-too-large')
  })

  it('limits the total library size', () => {
    const saved = [{ byteSize: MAX_SIGNATURE_LIBRARY_BYTES }]
    expect(validateSignatureLibraryAddition(saved, strokeSignature)).toBe('total-limit')
  })
})
