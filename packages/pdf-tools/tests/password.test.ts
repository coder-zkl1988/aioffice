import { decryptPDF, isEncrypted } from '@pdfsmaller/pdf-decrypt'
import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { analyzePdfSecurityBytes, runPdfToolBytes, type PdfPasswordPermissions } from '../src/index'

const ALL_PERMISSIONS: PdfPasswordPermissions = {
  allowPrinting: true,
  allowModifying: true,
  allowCopying: true,
  allowAnnotating: true,
  allowFillingForms: true,
  allowExtraction: true,
  allowAssembly: true,
  allowHighQualityPrint: true,
}

async function sourcePdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  document.addPage([240, 320])
  document.setTitle('Password roundtrip')
  return document.save({ useObjectStreams: false })
}

describe('PDF password protection', () => {
  it('reports unrestricted permissions for an unencrypted PDF', async () => {
    await expect(analyzePdfSecurityBytes(await sourcePdf())).resolves.toEqual({
      encrypted: false,
      permissions: ALL_PERMISSIONS,
      restrictedCount: 0,
    })
  })

  it('protects with AES-256 and unlocks with either password', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'password',
      action: 'protect',
      userPassword: '用户密码',
      ownerPassword: '所有者密码',
      algorithm: 'AES-256',
      permissions: { ...ALL_PERMISSIONS, allowCopying: false, allowModifying: false },
    })
    expect(output.suffix).toBe('_protected.pdf')
    await expect(isEncrypted(output.bytes)).resolves.toMatchObject({
      encrypted: true,
      algorithm: 'AES-256',
      keyLength: 256,
    })
    await expect(analyzePdfSecurityBytes(output.bytes)).resolves.toMatchObject({
      encrypted: true,
      algorithm: 'AES-256',
      version: 5,
      revision: 6,
      keyLength: 256,
      permissions: {
        ...ALL_PERMISSIONS,
        allowCopying: false,
        allowModifying: false,
      },
      restrictedCount: 2,
    })
    await expect(decryptPDF(output.bytes, 'wrong')).rejects.toThrow(/incorrect password/i)

    const [unlocked] = await runPdfToolBytes(output.bytes, {
      kind: 'password',
      action: 'unlock',
      password: '所有者密码',
    })
    expect(unlocked.suffix).toBe('_unlocked.pdf')
    await expect(isEncrypted(unlocked.bytes)).resolves.toMatchObject({ encrypted: false })
    const document = await PDFDocument.load(unlocked.bytes)
    expect(document.getPageCount()).toBe(1)
    expect(document.getTitle()).toBe('Password roundtrip')
  })

  it('supports RC4-128 for legacy readers', async () => {
    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'password',
      action: 'protect',
      userPassword: 'legacy-user',
      ownerPassword: 'legacy-owner',
      algorithm: 'RC4',
      permissions: ALL_PERMISSIONS,
    })
    await expect(isEncrypted(output.bytes)).resolves.toMatchObject({
      encrypted: true,
      algorithm: 'RC4',
      keyLength: 128,
    })
    const decrypted = await decryptPDF(output.bytes, 'legacy-user')
    await expect(PDFDocument.load(decrypted)).resolves.toBeInstanceOf(PDFDocument)
  })

  it('requires a separate owner password for restricted permissions', async () => {
    await expect(
      runPdfToolBytes(await sourcePdf(), {
        kind: 'password',
        action: 'protect',
        userPassword: 'same',
        ownerPassword: 'same',
        algorithm: 'AES-256',
        permissions: { ...ALL_PERMISSIONS, allowPrinting: false },
      }),
    ).rejects.toThrow(/different owner password/i)
  })

  it('rejects protection of an already encrypted PDF', async () => {
    const [encrypted] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'password',
      action: 'protect',
      userPassword: 'open',
      ownerPassword: 'owner',
      algorithm: 'AES-256',
      permissions: ALL_PERMISSIONS,
    })
    await expect(
      runPdfToolBytes(encrypted.bytes, {
        kind: 'password',
        action: 'protect',
        userPassword: 'new-open',
        ownerPassword: 'new-owner',
        algorithm: 'AES-256',
        permissions: ALL_PERMISSIONS,
      }),
    ).rejects.toThrow(/unlock.*before/i)
  })
})
