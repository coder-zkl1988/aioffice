import { describe, expect, it } from 'vitest'

import { isProbablyDamagedPdf } from '../src/renderer/damaged-pdf'

const HEADER = '%PDF-1.7\n'
const EOF = '\nstartxref\n0\n%%EOF\n'

/** Assemble bytes from ASCII parts plus a filler run of `size` zero bytes */
function assemble(parts: string[], filler = 0): Uint8Array {
  const bytes = new Uint8Array(parts.join('').length + filler)
  let offset = 0
  for (const part of parts) {
    for (let i = 0; i < part.length; i++) bytes[offset++] = part.charCodeAt(i)
  }
  return bytes
}

function wellFormed(size = 0): Uint8Array {
  return assemble([HEADER, '1 0 obj\n<<>>\nendobj'], size).map((byte, index, all) =>
    index >= all.length - EOF.length ? EOF.charCodeAt(index - (all.length - EOF.length)) : byte,
  )
}

describe('isProbablyDamagedPdf', () => {
  it('accepts a well-formed PDF', () => {
    expect(isProbablyDamagedPdf(wellFormed())).toBe(false)
  })

  it('accepts an ArrayBuffer as well as a typed array', () => {
    const bytes = wellFormed()
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength)
    expect(isProbablyDamagedPdf(buffer as ArrayBuffer)).toBe(false)
  })

  it('flags bytes with no %PDF header', () => {
    expect(isProbablyDamagedPdf(assemble(['not a pdf at all\n', '%%EOF']))).toBe(true)
  })

  it('flags a PDF whose header sits beyond the first 1024 bytes', () => {
    expect(isProbablyDamagedPdf(assemble(['x'.repeat(1025), HEADER, EOF]))).toBe(true)
  })

  it('accepts a header that is not the very first byte but is inside the window', () => {
    expect(isProbablyDamagedPdf(assemble(['\u0000\u0000\u0000', HEADER, EOF]))).toBe(false)
  })

  it('flags a truncated download with no EOF marker', () => {
    expect(isProbablyDamagedPdf(assemble([HEADER, '1 0 obj\n<<>>\nendobj']))).toBe(true)
  })

  it('accepts an EOF marker with a little trailing data', () => {
    expect(
      isProbablyDamagedPdf(assemble([HEADER, '1 0 obj\n<<>>\nendobj', EOF, 'x'.repeat(1024)])),
    ).toBe(false)
  })

  it('flags a file whose EOF marker is buried under more than 2 KB of trailing data', () => {
    expect(isProbablyDamagedPdf(assemble([HEADER, EOF, 'x'.repeat(4096)]))).toBe(true)
  })

  it('treats a well-formed encrypted file as undamaged (wrong password, not corruption)', () => {
    // The /Encrypt dictionary and binary stream bytes carry no structural meaning here;
    // both anchors are intact, so a failed open is a password question, not damage.
    expect(isProbablyDamagedPdf(assemble([HEADER, 'trailer\n<< /Encrypt 4 0 R >>\n', EOF]))).toBe(
      false,
    )
  })

  it('only inspects the head and tail of a multi-megabyte file', () => {
    expect(isProbablyDamagedPdf(wellFormed(5 * 1024 * 1024))).toBe(false)
    expect(isProbablyDamagedPdf(assemble(['x'.repeat(5 * 1024 * 1024)]))).toBe(true)
  })

  it('flags an empty file', () => {
    expect(isProbablyDamagedPdf(new Uint8Array(0))).toBe(true)
  })
})
