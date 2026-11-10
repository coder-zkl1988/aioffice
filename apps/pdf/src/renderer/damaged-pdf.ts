/** Structural damage sniff used when a PDF fails to open.
 *
 * Wrong-password and encryption failures leave a well-formed file behind, so the
 * bytes alone can tell "damaged structure" apart from "needs a password": the
 * opening `%PDF-` header and the closing `%%EOF` marker are the two cheap checks
 * that mirror the preflight findings (`missingPdfHeader` / `missingEofMarker`).
 */

/** Header bytes a reader tolerates before `%PDF-` starts (same window as preflight) */
const HEADER_WINDOW = 1024
/** `%%EOF` may trail a few bytes; anything beyond this means the marker is missing */
const EOF_WINDOW = 2048

/** Case-sensitive ASCII search restricted to [from, to) without decoding the file */
function asciiIndexOf(bytes: Uint8Array, from: number, to: number, needle: string): number {
  const last = Math.min(to, bytes.length) - needle.length
  scan: for (let i = Math.max(0, from); i <= last; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (bytes[i + j] !== needle.charCodeAt(j)) continue scan
    }
    return i
  }
  return -1
}

/**
 * True when the bytes cannot be a well-formed PDF: no `%PDF-` header in the first
 * 1024 bytes, or no `%%EOF` marker in the last 2 KB. Only the head and tail windows
 * are scanned, so multi-megabyte files cost nothing extra.
 */
export function isProbablyDamagedPdf(bytes: Uint8Array | ArrayBuffer): boolean {
  const data = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  if (data.length === 0) return true
  if (asciiIndexOf(data, 0, HEADER_WINDOW, '%PDF-') < 0) return true
  return asciiIndexOf(data, data.length - EOF_WINDOW, data.length, '%%EOF') < 0
}
