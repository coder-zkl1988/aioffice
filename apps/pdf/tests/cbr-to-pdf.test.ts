import { describe, expect, it } from 'vitest'
import type { ArcHeader, FileHeader } from 'node-unrar-js/esm/index.esm.js'
import { extractCbrImageEntriesFromExtractor, isRarArchive } from '../src/renderer/cbz-to-pdf'

const arcHeader: ArcHeader = {
  comment: '',
  flags: {
    volume: false,
    lock: false,
    solid: false,
    authInfo: false,
    recoveryRecord: false,
    headerEncrypted: false,
  },
}

function header(name: string, overrides: Partial<FileHeader> = {}): FileHeader {
  return {
    name,
    flags: { encrypted: false, solid: false, directory: false },
    packSize: 3,
    unpSize: 3,
    crc: 0,
    time: '',
    unpVer: '5.0',
    method: 'Normal',
    comment: '',
    ...overrides,
  }
}

function extractor(headers: FileHeader[]) {
  return {
    getFileList: () => ({ arcHeader, fileHeaders: headers.values() }),
    extract: () => ({
      arcHeader,
      files: headers
        .filter((item) => !item.flags.directory)
        .map((item) => ({ fileHeader: item, extraction: new Uint8Array([1, 2, 3]) }))
        .values(),
    }),
  }
}

describe('CBR comic extraction', () => {
  it('detects RAR 4 and RAR 5 signatures', () => {
    expect(isRarArchive(new Uint8Array([0x52, 0x61, 0x72, 0x21, 0x1a, 0x07, 0x00]))).toBe(true)
    expect(isRarArchive(new Uint8Array([0x52, 0x61, 0x72, 0x21, 0x1a, 0x07, 0x01, 0x00]))).toBe(
      true,
    )
    expect(isRarArchive(new Uint8Array([0x50, 0x4b, 0x03, 0x04]))).toBe(false)
  })

  it('ignores hidden and non-image files and naturally sorts pages', () => {
    const entries = extractCbrImageEntriesFromExtractor(
      extractor([
        header('pages/10.png'),
        header('notes.txt'),
        header('__MACOSX/2.png'),
        header('pages/2.jpg'),
        header('pages/.thumb.png'),
        header('pages/1.webp'),
      ]),
    )
    expect(entries.map((entry) => entry.name)).toEqual([
      'pages/1.webp',
      'pages/2.jpg',
      'pages/10.png',
    ])
  })

  it('rejects unsafe paths and encrypted images', () => {
    expect(() => extractCbrImageEntriesFromExtractor(extractor([header('../page.png')]))).toThrow(
      'unsafe file path',
    )
    expect(() =>
      extractCbrImageEntriesFromExtractor(
        extractor([
          header('page.png', {
            flags: { encrypted: true, solid: false, directory: false },
          }),
        ]),
      ),
    ).toThrow('Password-protected')
  })
})
