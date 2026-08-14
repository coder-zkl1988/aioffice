import * as asn1js from 'asn1js'
import {
  AlgorithmIdentifier,
  AttributeTypeAndValue,
  Certificate,
  ContentInfo,
  EncapsulatedContentInfo,
  IssuerAndSerialNumber,
  SignedData,
  SignerInfo,
} from 'pkijs'
import { PDFDocument } from 'pdf-lib'
import { describe, expect, it } from 'vitest'
import { auditPdfSignaturesBytes } from '../src/index'

const encoder = new TextEncoder()
const decoder = new TextDecoder('latin1')
const CONTENTS_BYTES = 8192
const BYTE_RANGE_PLACEHOLDER = '0 0000000000 0000000000 0000000000'

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer
}

function buildPdf(objects: string[]): Uint8Array {
  let output = '%PDF-1.4\n%\x80\x80\x80\x80\n'
  const offsets = [0]
  objects.forEach((object, index) => {
    offsets[index + 1] = encoder.encode(output).length
    output += `${index + 1} 0 obj\n${object}\nendobj\n`
  })
  const xrefOffset = encoder.encode(output).length
  output += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`
  output += offsets
    .slice(1)
    .map((offset) => `${String(offset).padStart(10, '0')} 00000 n \n`)
    .join('')
  output += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`
  return encoder.encode(output)
}

function signaturePdfTemplate(): Uint8Array {
  return buildPdf([
    '<< /Type /Catalog /Pages 2 0 R /AcroForm 4 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>',
    '<< /Fields [5 0 R] /SigFlags 3 >>',
    '<< /FT /Sig /T (Approval) /V 6 0 R >>',
    `<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [${BYTE_RANGE_PLACEHOLDER}] /Contents <${'0'.repeat(CONTENTS_BYTES * 2)}> /Name (Fallback signer) /Reason (Approval test) /Location (Local) /M (D:20260101000000Z) >>`,
  ])
}

function patchAscii(input: Uint8Array, search: string, replacement: string): Uint8Array {
  expect(replacement).toHaveLength(search.length)
  const text = decoder.decode(input)
  const offset = text.indexOf(search)
  expect(offset).toBeGreaterThanOrEqual(0)
  const output = Uint8Array.from(input)
  output.set(encoder.encode(replacement), offset)
  return output
}

function signatureRanges(input: Uint8Array): { pdf: Uint8Array; signed: Uint8Array } {
  const text = decoder.decode(input)
  const contentsMarker = '/Contents <'
  const contentsStart = text.indexOf(contentsMarker) + contentsMarker.length - 1
  const contentsEnd = contentsStart + CONTENTS_BYTES * 2 + 2
  const secondLength = input.length - contentsEnd
  const byteRange = [0, contentsStart, contentsEnd, secondLength]
  const replacement = byteRange
    .map((value, index) => (index === 0 ? String(value) : String(value).padStart(10, '0')))
    .join(' ')
  const pdf = patchAscii(input, BYTE_RANGE_PLACEHOLDER, replacement)
  const signed = new Uint8Array(contentsStart + secondLength)
  signed.set(pdf.subarray(0, contentsStart), 0)
  signed.set(pdf.subarray(contentsEnd), contentsStart)
  return { pdf, signed }
}

async function createCertificate(): Promise<{
  certificate: Certificate
  privateKey: CryptoKey
}> {
  const keys = await crypto.subtle.generateKey(
    {
      name: 'RSASSA-PKCS1-v1_5',
      modulusLength: 2048,
      publicExponent: new Uint8Array([1, 0, 1]),
      hash: 'SHA-256',
    },
    true,
    ['sign', 'verify'],
  )
  const certificate = new Certificate()
  certificate.version = 2
  certificate.serialNumber = new asn1js.Integer({ value: 1 })
  const commonName = new AttributeTypeAndValue({
    type: '2.5.4.3',
    value: new asn1js.Utf8String({ value: 'GenOffice Test Signer' }),
  })
  certificate.issuer.typesAndValues.push(commonName)
  certificate.subject.typesAndValues.push(commonName)
  certificate.notBefore.value = new Date('2025-01-01T00:00:00.000Z')
  certificate.notAfter.value = new Date('2030-01-01T00:00:00.000Z')
  await certificate.subjectPublicKeyInfo.importKey(keys.publicKey)
  await certificate.sign(keys.privateKey, 'SHA-256')
  return { certificate, privateKey: keys.privateKey }
}

async function createDetachedCms(
  data: Uint8Array,
  certificate: Certificate,
  privateKey: CryptoKey,
): Promise<Uint8Array> {
  const signedData = new SignedData({
    version: 1,
    digestAlgorithms: [
      new AlgorithmIdentifier({
        algorithmId: '2.16.840.1.101.3.4.2.1',
        algorithmParams: new asn1js.Null(),
      }),
    ],
    encapContentInfo: new EncapsulatedContentInfo({
      eContentType: '1.2.840.113549.1.7.1',
    }),
    certificates: [certificate],
    signerInfos: [
      new SignerInfo({
        version: 1,
        sid: new IssuerAndSerialNumber({
          issuer: certificate.issuer,
          serialNumber: certificate.serialNumber,
        }),
      }),
    ],
  })
  await signedData.sign(privateKey, 0, 'SHA-256', exactArrayBuffer(data))
  const contentInfo = new ContentInfo({
    contentType: '1.2.840.113549.1.7.2',
    content: signedData.toSchema(true),
  })
  return new Uint8Array(contentInfo.toSchema().toBER(false))
}

async function signedPdf(): Promise<Uint8Array> {
  const { pdf, signed } = signatureRanges(signaturePdfTemplate())
  const { certificate, privateKey } = await createCertificate()
  const cms = await createDetachedCms(signed, certificate, privateKey)
  expect(cms.length).toBeLessThan(CONTENTS_BYTES)
  return patchAscii(
    pdf,
    '0'.repeat(CONTENTS_BYTES * 2),
    `${Array.from(cms, (value) => value.toString(16).padStart(2, '0')).join('')}${'0'.repeat(
      (CONTENTS_BYTES - cms.length) * 2,
    )}`,
  )
}

describe('auditPdfSignaturesBytes', () => {
  it('reports an unsigned document without findings', async () => {
    const document = await PDFDocument.create()
    document.addPage([200, 200])
    await expect(auditPdfSignaturesBytes(await document.save())).resolves.toEqual({
      signatures: [],
      validCount: 0,
      invalidCount: 0,
      unsupportedCount: 0,
    })
  })

  it('validates a detached CMS signature without implying certificate trust', async () => {
    const audit = await auditPdfSignaturesBytes(await signedPdf())
    expect(audit.validCount).toBe(1)
    expect(audit.invalidCount).toBe(0)
    expect(audit.signatures[0]).toMatchObject({
      fieldName: 'Approval',
      type: 'approval',
      signerName: 'GenOffice Test Signer',
      reason: 'Approval test',
      location: 'Local',
      byteRangeValid: true,
      coversEntireDocument: true,
      unsignedTailBytes: 0,
      cryptographicStatus: 'valid',
      digestAlgorithm: 'SHA-256',
      trustStatus: 'notEvaluated',
      certificate: {
        selfSigned: true,
        validAtSigningTime: true,
      },
    })
  })

  it('detects document changes made after signing', async () => {
    const tampered = patchAscii(await signedPdf(), 'Approval test', 'Tampered test')
    const audit = await auditPdfSignaturesBytes(tampered)
    expect(audit.invalidCount).toBe(1)
    expect(audit.signatures[0].cryptographicStatus).toBe('invalid')
  })

  it('rejects a malformed ByteRange before cryptographic verification', async () => {
    const malformed = patchAscii(
      signaturePdfTemplate(),
      BYTE_RANGE_PLACEHOLDER,
      '0 9999999999 0000000001 0000000001',
    )
    const audit = await auditPdfSignaturesBytes(malformed)
    expect(audit.invalidCount).toBe(1)
    expect(audit.signatures[0]).toMatchObject({
      byteRangeValid: false,
      cryptographicStatus: 'notChecked',
    })
  })
})
