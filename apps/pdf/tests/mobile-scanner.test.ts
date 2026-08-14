import { describe, expect, it } from 'vitest'
import { JSDOM } from 'jsdom'
import {
  MobileScannerHub,
  mobileScannerUploaderHtml,
  mobileScannerUploaderScript,
} from '../src/main/mobile-scanner'
import { isPdfMobileScannerSessionId } from '../src/shared/mobile-scanner'

const JPEG = new Uint8Array([0xff, 0xd8, 0xff, 0xdb, 0x00, 0x01])
const PNG = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00])

describe('MobileScannerHub', () => {
  it('creates high-entropy expiring sessions and consumes uploaded images', () => {
    const hub = new MobileScannerHub()
    const session = hub.createSession(1_000)
    expect(isPdfMobileScannerSessionId(session.sessionId)).toBe(true)
    expect(session.expiresAt).toBe(601_000)

    const first = hub.addFile(
      session.sessionId,
      { name: 'scan.jpg', declaredType: 'image/jpeg', bytes: JPEG },
      2_000,
    )
    const second = hub.addFile(
      session.sessionId,
      { name: 'scan.jpg', declaredType: 'image/jpeg', bytes: JPEG },
      2_000,
    )
    expect(first.name).toBe('scan.jpg')
    expect(second.name).toBe('scan-2.jpg')

    const result = hub.takeFiles(session.sessionId, 3_000)
    expect(result.files.map((file) => file.name)).toEqual(['scan.jpg', 'scan-2.jpg'])
    expect(hub.takeFiles(session.sessionId, 4_000).files).toEqual([])
  })

  it('validates image content instead of trusting names or MIME headers', () => {
    const hub = new MobileScannerHub()
    const session = hub.createSession()
    expect(() =>
      hub.addFile(session.sessionId, {
        name: '../../photo.png',
        declaredType: 'image/png',
        bytes: JPEG,
      }),
    ).toThrow(/does not match/iu)
    expect(() =>
      hub.addFile(session.sessionId, {
        name: 'payload.jpg',
        declaredType: 'image/jpeg',
        bytes: new TextEncoder().encode('<script>alert(1)</script>'),
      }),
    ).toThrow(/valid JPEG, PNG, and WebP/iu)

    const file = hub.addFile(session.sessionId, {
      name: '../合同扫描',
      declaredType: 'image/png',
      bytes: PNG,
    })
    expect(file.name).toBe('.._合同扫描.png')
  })

  it('expires sessions and enforces file and session count limits', () => {
    const hub = new MobileScannerHub()
    const session = hub.createSession(10)
    expect(hub.hasSession(session.sessionId, session.expiresAt)).toBe(false)
    expect(() => hub.takeFiles(session.sessionId, session.expiresAt)).toThrow(/expired/iu)

    const active = hub.createSession()
    for (let index = 0; index < 40; index += 1) {
      hub.addFile(active.sessionId, {
        name: `scan-${index}.jpg`,
        declaredType: 'image/jpeg',
        bytes: JPEG,
      })
    }
    expect(() =>
      hub.addFile(active.sessionId, {
        name: 'scan-41.jpg',
        declaredType: 'image/jpeg',
        bytes: JPEG,
      }),
    ).toThrow(/at most 40/iu)

    const sessions = new MobileScannerHub()
    for (let index = 0; index < 64; index += 1) sessions.createSession()
    expect(() => sessions.createSession()).toThrow(/too many/iu)
  })
})

describe('mobile scanner upload page', () => {
  it('uses an external script and escapes environment-derived paths', () => {
    const session = new MobileScannerHub().createSession()
    const html = mobileScannerUploaderHtml(session.sessionId, '/safe" onload="alert(1)')
    expect(html).toContain('data-upload-path="/safe&quot; onload=&quot;alert(1)/api/pdf/')
    expect(html).toContain('script src="/safe&quot; onload=&quot;alert(1)/mobile-scanner/app.js"')
    expect(html).not.toMatch(/<script(?!\s+src=)/iu)
    const document = new JSDOM(html).window.document
    expect(document.body.hasAttribute('onload')).toBe(false)
    expect(document.scripts).toHaveLength(1)
    expect(document.scripts[0]?.getAttribute('src')).toBe(
      '/safe" onload="alert(1)/mobile-scanner/app.js',
    )
  })

  it('normalizes images through canvas before uploading', () => {
    expect(mobileScannerUploaderScript).toContain('canvas.toBlob')
    expect(mobileScannerUploaderScript).toContain("'image/jpeg'")
    expect(mobileScannerUploaderScript).not.toMatch(/eval\s*\(|innerHTML/iu)
  })
})
