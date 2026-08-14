import forge from 'node-forge'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { auditPdfSignaturesBytes, runPdfToolBytes, signPdfWithP12Bytes } from '../src/index'

const PASSWORD = 'local-signing-password'

async function testPdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  const page = document.addPage([400, 240])
  const font = await document.embedFont(StandardFonts.Helvetica)
  page.drawText('GenOffice certificate signing test', { x: 36, y: 180, font, size: 20 })
  return document.save({ useObjectStreams: false })
}

function testP12(password = PASSWORD): Uint8Array {
  const keys = forge.pki.rsa.generateKeyPair(2048)
  const certificate = forge.pki.createCertificate()
  certificate.publicKey = keys.publicKey
  certificate.serialNumber = '01'
  certificate.validity.notBefore = new Date(Date.now() - 60_000)
  certificate.validity.notAfter = new Date(Date.now() + 86_400_000)
  const attributes = [
    { name: 'commonName', value: 'GenOffice Local Signer' },
    { name: 'organizationName', value: 'GenOffice' },
  ]
  certificate.setSubject(attributes)
  certificate.setIssuer(attributes)
  certificate.sign(keys.privateKey, forge.md.sha256.create())
  const asn1 = forge.pkcs12.toPkcs12Asn1(keys.privateKey, [certificate], password, {
    algorithm: '3des',
  })
  return Uint8Array.from(forge.util.binary.raw.decode(forge.asn1.toDer(asn1).getBytes()))
}

const options = () => ({
  certificate: testP12(),
  password: PASSWORD,
  signerName: 'Yang Ke',
  reason: 'Document approval',
  location: 'Shandong',
  contactInfo: 'yang@example.com',
})

describe('P12/PFX PDF signing', () => {
  it('creates a detached signature that the local validator accepts', async () => {
    const result = await signPdfWithP12Bytes(await testPdf(), options())
    const audit = await auditPdfSignaturesBytes(result.bytes)
    expect(audit.validCount).toBe(1)
    expect(audit.invalidCount).toBe(0)
    expect(audit.signatures[0]).toMatchObject({
      signerName: 'GenOffice Local Signer',
      reason: 'Document approval',
      location: 'Shandong',
      coversEntireDocument: true,
      cryptographicStatus: 'valid',
    })
    expect(result.certificate.subject).toContain('CN=GenOffice Local Signer')
  })

  it('rejects an incorrect certificate password', async () => {
    await expect(
      signPdfWithP12Bytes(await testPdf(), { ...options(), password: 'incorrect' }),
    ).rejects.toThrow(/password is incorrect|file is invalid/i)
  })

  it('detects changes after signing', async () => {
    const result = await signPdfWithP12Bytes(await testPdf(), options())
    const changed = Uint8Array.from(result.bytes)
    expect(new TextDecoder().decode(changed.subarray(0, 8))).toMatch(/^%PDF-1\./)
    changed[7] = changed[7] === 55 ? 54 : 55
    const audit = await auditPdfSignaturesBytes(changed)
    expect(audit.validCount).toBe(0)
    expect(audit.invalidCount).toBe(1)
  })

  it('does not rewrite a PDF that already has a signature', async () => {
    const result = await signPdfWithP12Bytes(await testPdf(), options())
    await expect(signPdfWithP12Bytes(result.bytes, options())).rejects.toThrow(
      /already contains a digital signature/i,
    )
  })

  it('runs through the shared tool operation', async () => {
    const [output] = await runPdfToolBytes(await testPdf(), {
      kind: 'certificateSign',
      ...options(),
    })
    expect(output.suffix).toBe('_signed.pdf')
    expect((await auditPdfSignaturesBytes(output.bytes)).validCount).toBe(1)
  })
})
