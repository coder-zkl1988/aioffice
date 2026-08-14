import * as asn1js from 'asn1js'
import {
  AlgorithmIdentifier,
  AttributeTypeAndValue,
  Certificate,
  ContentInfo,
  EncapsulatedContentInfo,
  IssuerAndSerialNumber,
  MessageImprint,
  PKIStatus,
  PKIStatusInfo,
  SignedData,
  SignerInfo,
  TimeStampReq,
  TimeStampResp,
  TSTInfo,
} from 'pkijs'
import { PDFDocument } from 'pdf-lib'
import { describe, expect, it, vi } from 'vitest'
import {
  auditPdfSignaturesBytes,
  createPdfTimestampRequestBytes,
  isAllowedPdfTimestampTsaUrl,
  runPdfToolBytes,
  timestampPdfBytes,
  timestampTokenFromResponseBytes,
} from '../src/index'

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

async function sourcePdf(): Promise<Uint8Array> {
  const document = await PDFDocument.create()
  const page = document.addPage([400, 300])
  page.drawText('GenOffice timestamp test', { x: 40, y: 240 })
  return document.save()
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
  certificate.serialNumber = new asn1js.Integer({ value: 7 })
  const commonName = new AttributeTypeAndValue({
    type: '2.5.4.3',
    value: new asn1js.Utf8String({ value: 'GenOffice Test TSA' }),
  })
  certificate.issuer.typesAndValues.push(commonName)
  certificate.subject.typesAndValues.push(commonName)
  certificate.notBefore.value = new Date('2025-01-01T00:00:00.000Z')
  certificate.notAfter.value = new Date('2030-01-01T00:00:00.000Z')
  await certificate.subjectPublicKeyInfo.importKey(keys.publicKey)
  await certificate.sign(keys.privateKey, 'SHA-256')
  return { certificate, privateKey: keys.privateKey }
}

async function timestampResponse(
  requestBytes: Uint8Array,
  options: { mismatchNonce?: boolean; mismatchDigest?: boolean } = {},
): Promise<Uint8Array> {
  const parsed = asn1js.fromBER(exactArrayBuffer(requestBytes))
  const request = new TimeStampReq({ schema: parsed.result })
  const { certificate, privateKey } = await createCertificate()
  const messageImprint = options.mismatchDigest
    ? await MessageImprint.create('SHA-256', new Uint8Array([9, 8, 7]))
    : request.messageImprint
  const timestamp = new TSTInfo({
    version: 1,
    policy: '1.3.6.1.4.1.57264.1',
    messageImprint,
    serialNumber: new asn1js.Integer({ value: 42 }),
    genTime: new Date('2026-08-14T01:02:03.000Z'),
    ordering: true,
    nonce: options.mismatchNonce ? new asn1js.Integer({ value: 999 }) : request.nonce,
  })
  const encapContentInfo = new EncapsulatedContentInfo({
    eContentType: '1.2.840.113549.1.9.16.1.4',
  })
  encapContentInfo.eContent = new asn1js.OctetString({
    valueHex: timestamp.toSchema().toBER(false),
  })
  const signedData = new SignedData({
    version: 3,
    digestAlgorithms: [
      new AlgorithmIdentifier({
        algorithmId: '2.16.840.1.101.3.4.2.1',
        algorithmParams: new asn1js.Null(),
      }),
    ],
    encapContentInfo,
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
  await signedData.sign(privateKey, 0, 'SHA-256')
  const token = new ContentInfo({
    contentType: ContentInfo.SIGNED_DATA,
    content: signedData.toSchema(true),
  })
  const response = new TimeStampResp({
    status: new PKIStatusInfo({ status: PKIStatus.granted }),
    timeStampToken: new ContentInfo({ schema: token.toSchema() }),
  })
  return new Uint8Array(response.toSchema().toBER(false))
}

describe('RFC 3161 PDF timestamps', () => {
  it('allows only built-in TSA endpoints', () => {
    expect(isAllowedPdfTimestampTsaUrl('http://timestamp.digicert.com')).toBe(true)
    expect(isAllowedPdfTimestampTsaUrl('https://freetsa.org/tsr/')).toBe(true)
    expect(isAllowedPdfTimestampTsaUrl('https://example.com/tsa')).toBe(false)
    expect(isAllowedPdfTimestampTsaUrl('http://user:pass@timestamp.digicert.com')).toBe(false)
  })

  it('creates a SHA-256 request with a nonce and certificate request flag', async () => {
    const input = new Uint8Array([1, 2, 3, 4])
    const { bytes, request } = await createPdfTimestampRequestBytes(input)
    const parsed = asn1js.fromBER(exactArrayBuffer(bytes))
    const reparsed = new TimeStampReq({ schema: parsed.result })

    expect(request.messageImprint.hashAlgorithm.algorithmId).toBe('2.16.840.1.101.3.4.2.1')
    expect(reparsed.certReq).toBe(true)
    expect(reparsed.nonce).toBeDefined()
    expect(new Uint8Array(reparsed.messageImprint.hashedMessage.valueBlock.valueHex)).toHaveLength(
      32,
    )
  })

  it('timestamps a PDF and verifies the resulting signature locally', async () => {
    const requestToken = vi.fn(timestampResponse)
    const result = await timestampPdfBytes(await sourcePdf(), requestToken)
    const audit = await auditPdfSignaturesBytes(result.bytes)

    expect(requestToken).toHaveBeenCalledOnce()
    expect(result.timestampTime).toBe('2026-08-14T01:02:03.000Z')
    expect(audit.validCount).toBe(1)
    expect(audit.signatures[0]).toMatchObject({
      type: 'documentTimestamp',
      byteRangeValid: true,
      coversEntireDocument: true,
      cryptographicStatus: 'valid',
      timestampImprintValid: true,
      timestampTime: '2026-08-14T01:02:03.000Z',
    })

    const [output] = await runPdfToolBytes(await sourcePdf(), {
      kind: 'timestamp',
      tsaUrl: 'http://timestamp.digicert.com',
      timestampedBytes: result.bytes,
    })
    expect(output?.suffix).toBe('_timestamped.pdf')
    expect(output?.bytes).toEqual(result.bytes)
  })

  it('rejects responses with a mismatched digest or nonce', async () => {
    const { bytes, request } = await createPdfTimestampRequestBytes(new Uint8Array([1, 2, 3]))
    await expect(
      timestampResponse(bytes, { mismatchDigest: true }).then((response) =>
        timestampTokenFromResponseBytes(response, request),
      ),
    ).rejects.toThrow('digest')
    await expect(
      timestampResponse(bytes, { mismatchNonce: true }).then((response) =>
        timestampTokenFromResponseBytes(response, request),
      ),
    ).rejects.toThrow('nonce')
  })

  it('refuses to rewrite a PDF that already has a digital signature', async () => {
    const first = await timestampPdfBytes(await sourcePdf(), timestampResponse)
    const requestToken = vi.fn(timestampResponse)
    await expect(timestampPdfBytes(first.bytes, requestToken)).rejects.toThrow(
      'already contains a digital signature',
    )
    expect(requestToken).not.toHaveBeenCalled()
  })
})
