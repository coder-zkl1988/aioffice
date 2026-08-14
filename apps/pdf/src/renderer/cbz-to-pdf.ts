import { extractCbzImageEntries, type PdfRasterPage } from '@genoffice/pdf-tools'
import type { ArcFile, ArcHeader, Extractor, FileHeader } from 'node-unrar-js/esm/index.esm.js'
import { prepareImagesForPdf, type PdfImageColorMode } from './images-to-pdf'

/*
 * UnRAR source code may be used in any software to handle RAR archives without
 * limitations free of charge, but cannot be used to develop RAR (WinRAR)
 * compatible archiver and to re-create RAR compression algorithm, which is
 * proprietary. Distribution of modified UnRAR source code in separate form or
 * as a part of other software is permitted, provided that full text of this
 * paragraph, starting from "UnRAR source code" words, is included in license,
 * or in documentation if license is not available, and in source code comments
 * of resulting package.
 */
const MAX_COMIC_ARCHIVE_BYTES = 200 * 1024 * 1024
const MAX_COMIC_ENTRY_BYTES = 50 * 1024 * 1024
const MAX_COMIC_TOTAL_IMAGE_BYTES = 200 * 1024 * 1024
const MAX_COMIC_IMAGES = 200
const MAX_COMIC_ENTRIES = 2_000
const COMIC_IMAGE_EXTENSION = /\.(?:bmp|gif|jpe?g|png|webp)$/i
const COMIC_NATURAL_ORDER = new Intl.Collator('en', { numeric: true, sensitivity: 'base' })

type CbrExtractor = Pick<Extractor<Uint8Array>, 'getFileList' | 'extract'>

let unrarWasmPromise: Promise<ArrayBuffer> | undefined
let cbrExtractionQueue = Promise.resolve()

function imageMimeType(name: string): string {
  const extension = name.split('.').pop()?.toLowerCase()
  if (extension === 'jpg' || extension === 'jpeg') return 'image/jpeg'
  if (extension === 'webp') return 'image/webp'
  if (extension === 'gif') return 'image/gif'
  if (extension === 'bmp') return 'image/bmp'
  return 'image/png'
}

function checkedCbrEntryName(name: string): string {
  const normalized = name.replace(/\\/g, '/')
  if (
    normalized.includes('\0') ||
    normalized.startsWith('/') ||
    /^[a-z]:\//i.test(normalized) ||
    normalized.split('/').includes('..')
  ) {
    throw new Error('CBR archive contains an unsafe file path')
  }
  return normalized
}

function isVisibleComicImage(name: string): boolean {
  if (!COMIC_IMAGE_EXTENSION.test(name)) return false
  const segments = name.split('/')
  return !segments.some((segment) => segment.startsWith('.') || segment === '__MACOSX')
}

function checkedCbrHeaders(arcHeader: ArcHeader, fileHeaders: FileHeader[]): FileHeader[] {
  if (arcHeader.flags.volume) throw new Error('Multi-volume CBR archives are not supported')
  if (fileHeaders.length > MAX_COMIC_ENTRIES) throw new Error('CBR archive contains too many files')

  const imageHeaders = fileHeaders
    .filter((header) => !header.flags.directory)
    .map((header) => ({ ...header, name: checkedCbrEntryName(header.name) }))
    .filter((header) => isVisibleComicImage(header.name))
    .sort((left, right) => COMIC_NATURAL_ORDER.compare(left.name, right.name))

  if (imageHeaders.length === 0) throw new Error('CBR archive contains no supported images')
  if (imageHeaders.length > MAX_COMIC_IMAGES) {
    throw new Error(`CBR archive may contain no more than ${MAX_COMIC_IMAGES} images`)
  }

  let totalBytes = 0
  for (const header of imageHeaders) {
    if (header.flags.encrypted) throw new Error('Password-protected CBR archives are not supported')
    if (header.unpSize > MAX_COMIC_ENTRY_BYTES) {
      throw new Error('Each CBR image must be 50 MB or smaller')
    }
    totalBytes += header.unpSize
  }
  if (totalBytes > MAX_COMIC_TOTAL_IMAGE_BYTES) {
    throw new Error('CBR images must total 200 MB or less')
  }
  return imageHeaders
}

function unrarError(error: unknown): Error {
  const reason = typeof error === 'object' && error && 'reason' in error ? String(error.reason) : ''
  if (reason === 'ERAR_MISSING_PASSWORD' || reason === 'ERAR_BAD_PASSWORD') {
    return new Error('Password-protected CBR archives are not supported')
  }
  if (reason === 'ERAR_EREFERENCE') {
    return new Error('CBR archive uses an unsupported reference record')
  }
  if (reason.startsWith('ERAR_')) return new Error('CBR archive is invalid or damaged')
  return error instanceof Error ? error : new Error(String(error))
}

function loadArrayBuffer(url: string): Promise<ArrayBuffer> {
  return new Promise((resolve, reject) => {
    const request = new XMLHttpRequest()
    request.open('GET', url, true)
    request.responseType = 'arraybuffer'
    request.onload = () => {
      if (
        (request.status >= 200 && request.status < 300) ||
        (request.status === 0 && request.response)
      ) {
        resolve(request.response as ArrayBuffer)
      } else {
        reject(new Error('Unable to load the local CBR decoder'))
      }
    }
    request.onerror = () => reject(new Error('Unable to load the local CBR decoder'))
    request.send()
  })
}

async function loadUnrarWasm(): Promise<ArrayBuffer> {
  unrarWasmPromise ??= import('node-unrar-js/esm/js/unrar.wasm?url').then(({ default: url }) =>
    loadArrayBuffer(url),
  )
  return unrarWasmPromise
}

export function isRarArchive(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 7 &&
    bytes[0] === 0x52 &&
    bytes[1] === 0x61 &&
    bytes[2] === 0x72 &&
    bytes[3] === 0x21 &&
    bytes[4] === 0x1a &&
    bytes[5] === 0x07 &&
    (bytes[6] === 0x00 || bytes[6] === 0x01)
  )
}

export function extractCbrImageEntriesFromExtractor(
  extractor: CbrExtractor,
): { name: string; bytes: Uint8Array }[] {
  const list = extractor.getFileList()
  const imageHeaders = checkedCbrHeaders(list.arcHeader, [...list.fileHeaders])
  const selectedNames = new Set(imageHeaders.map((header) => header.name))
  const extracted = extractor.extract({
    files: (header) => selectedNames.has(checkedCbrEntryName(header.name)),
  })
  const files = [...extracted.files]
  const images: { name: string; bytes: Uint8Array }[] = []
  let totalBytes = 0

  for (const file of files as ArcFile<Uint8Array>[]) {
    const name = checkedCbrEntryName(file.fileHeader.name)
    if (!selectedNames.has(name)) continue
    if (!file.extraction) throw new Error(`${name}: image could not be extracted`)
    if (file.extraction.length > MAX_COMIC_ENTRY_BYTES) {
      throw new Error(`${name}: image must be 50 MB or smaller`)
    }
    totalBytes += file.extraction.length
    if (totalBytes > MAX_COMIC_TOTAL_IMAGE_BYTES) {
      throw new Error('CBR images must total 200 MB or less')
    }
    images.push({ name, bytes: file.extraction })
  }

  if (images.length !== imageHeaders.length) throw new Error('CBR archive is invalid or damaged')
  return images.sort((left, right) => COMIC_NATURAL_ORDER.compare(left.name, right.name))
}

async function extractCbrImageEntriesUnlocked(
  bytes: Uint8Array,
): Promise<{ name: string; bytes: Uint8Array }[]> {
  try {
    const [{ createExtractorFromData }, wasmBinary] = await Promise.all([
      import('node-unrar-js/esm/index.esm.js'),
      loadUnrarWasm(),
    ])
    const data = new Uint8Array(bytes).buffer
    const extractor = await createExtractorFromData({ data, wasmBinary })
    return extractCbrImageEntriesFromExtractor(extractor)
  } catch (error) {
    throw unrarError(error)
  }
}

export function extractCbrImageEntries(
  archiveBytes: Uint8Array | ArrayBuffer,
): Promise<{ name: string; bytes: Uint8Array }[]> {
  const bytes = new Uint8Array(archiveBytes)
  if (bytes.length === 0) throw new Error('CBR archive is empty')
  if (bytes.length > MAX_COMIC_ARCHIVE_BYTES)
    throw new Error('CBR archive must be 200 MB or smaller')
  if (!isRarArchive(bytes)) throw new Error('CBR archive is invalid or damaged')

  const extraction = cbrExtractionQueue.then(() => extractCbrImageEntriesUnlocked(bytes))
  cbrExtractionQueue = extraction.then(
    () => undefined,
    () => undefined,
  )
  return extraction
}

export async function prepareComicArchiveForPdf(
  archiveFile: File,
  colorMode: PdfImageColorMode,
): Promise<PdfRasterPage[]> {
  const archiveBytes = new Uint8Array(await archiveFile.arrayBuffer())
  const entries = isRarArchive(archiveBytes)
    ? await extractCbrImageEntries(archiveBytes)
    : await extractCbzImageEntries(archiveBytes)
  const imageFiles = entries.map(({ name, bytes }) => {
    const buffer = bytes.buffer.slice(
      bytes.byteOffset,
      bytes.byteOffset + bytes.byteLength,
    ) as ArrayBuffer
    return new File([buffer], name, { type: imageMimeType(name) })
  })
  return prepareImagesForPdf(imageFiles, colorMode, false)
}
