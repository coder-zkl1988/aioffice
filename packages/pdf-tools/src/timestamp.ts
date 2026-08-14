import { Buffer } from 'buffer'
import * as asn1js from 'asn1js'
import {
  ContentInfo,
  MessageImprint,
  SignedData,
  TimeStampReq,
  TimeStampResp,
  TSTInfo,
} from 'pkijs'
import { PDFDocument, PDFSignature } from 'pdf-lib'
import { auditPdfSignaturesBytes, type PdfSignatureCertificate } from './signature-audit'

export const PDF_TIMESTAMP_TSA_PRESETS = [
  { id: 'digicert', label: 'DigiCert', url: 'http://timestamp.digicert.com' },
  { id: 'sectigo', label: 'Sectigo', url: 'http://timestamp.sectigo.com' },
  { id: 'sslcom', label: 'SSL.com', url: 'http://ts.ssl.com' },
  { id: 'freetsa', label: 'FreeTSA', url: 'https://freetsa.org/tsr' },
  { id: 'mesign', label: 'MeSign', url: 'http://tsa.mesign.com' },
] as const

export interface PdfTimestampResult {
  bytes: Uint8Array
  timestampTime: string
  certificate?: PdfSignatureCertificate
}

export type PdfTimestampTokenRequester = (request: Uint8Array) => Promise<Uint8Array>

type GlobalWithBuffer = typeof globalThis & { Buffer?: typeof Buffer }

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer
}

function installBufferCompatibility(): void {
  const target = globalThis as GlobalWithBuffer
  if (!target.Buffer) target.Buffer = Buffer
}

function normalizedTsaUrl(value: string): string | null {
  try {
    const url = new URL(value)
    if (url.username || url.password || !['http:', 'https:'].includes(url.protocol)) return null
    url.hash = ''
    url.search = ''
    url.pathname = url.pathname.replace(/\/+$/, '') || '/'
    return url.toString().replace(/\/$/, '').toLowerCase()
  } catch {
    return null
  }
}

export function isAllowedPdfTimestampTsaUrl(value: string): boolean {
  const normalized = normalizedTsaUrl(value)
  return (
    normalized !== null &&
    PDF_TIMESTAMP_TSA_PRESETS.some((preset) => normalizedTsaUrl(preset.url) === normalized)
  )
}

export async function createPdfTimestampRequestBytes(data: Uint8Array): Promise<{
  bytes: Uint8Array
  request: TimeStampReq
}> {
  if (data.length === 0) throw new Error('The PDF timestamp input is empty')
  const nonceBytes = crypto.getRandomValues(new Uint8Array(16))
  nonceBytes[0] &= 0x7f
  const request = new TimeStampReq({
    version: 1,
    messageImprint: await MessageImprint.create('SHA-256', exactArrayBuffer(data)),
    nonce: new asn1js.Integer({ valueHex: exactArrayBuffer(nonceBytes) }),
    certReq: true,
  })
  return { bytes: new Uint8Array(request.toSchema().toBER(false)), request }
}

function bytesEqual(left: ArrayBuffer, right: ArrayBuffer): boolean {
  const a = new Uint8Array(left)
  const b = new Uint8Array(right)
  return a.length === b.length && a.every((value, index) => value === b[index])
}

export function timestampTokenFromResponseBytes(
  responseBytes: Uint8Array,
  request: TimeStampReq,
): Uint8Array {
  if (responseBytes.length === 0 || responseBytes.length > 1024 * 1024) {
    throw new Error('The TSA response is empty or too large')
  }
  const parsed = asn1js.fromBER(exactArrayBuffer(responseBytes))
  if (parsed.offset === -1) throw new Error('The TSA response is malformed')
  const response = new TimeStampResp({ schema: parsed.result })
  if (![0, 1].includes(response.status.status)) {
    throw new Error(`The TSA rejected the request with status ${response.status.status}`)
  }
  const token = response.timeStampToken
  if (!token) throw new Error('The TSA response does not contain a timestamp token')
  if (token.contentType !== ContentInfo.SIGNED_DATA) {
    throw new Error('The TSA timestamp token is not CMS SignedData')
  }
  const signedData = new SignedData({ schema: token.content })
  const payload = signedData.encapContentInfo.eContent?.getValue()
  if (!payload) throw new Error('The TSA timestamp token has no TSTInfo payload')
  const payloadSchema = asn1js.fromBER(payload)
  if (payloadSchema.offset === -1) throw new Error('The TSA timestamp payload is malformed')
  const timestamp = new TSTInfo({ schema: payloadSchema.result })
  if (
    timestamp.messageImprint.hashAlgorithm.algorithmId !==
      request.messageImprint.hashAlgorithm.algorithmId ||
    !bytesEqual(
      timestamp.messageImprint.hashedMessage.valueBlock.valueHex,
      request.messageImprint.hashedMessage.valueBlock.valueHex,
    )
  ) {
    throw new Error('The TSA response does not match the requested document digest')
  }
  if (
    request.nonce &&
    (!timestamp.nonce ||
      !bytesEqual(timestamp.nonce.valueBlock.valueHex, request.nonce.valueBlock.valueHex))
  ) {
    throw new Error('The TSA response nonce does not match the request')
  }
  return new Uint8Array(token.toSchema().toBER(false))
}

export async function timestampPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  requestToken: PdfTimestampTokenRequester,
): Promise<PdfTimestampResult> {
  const input =
    bytes instanceof Uint8Array ? Uint8Array.from(bytes) : new Uint8Array(bytes.slice(0))
  if (input.length === 0) throw new Error('The PDF is empty')
  const existingAudit = await auditPdfSignaturesBytes(input)
  if (existingAudit.signatures.length > 0) {
    throw new Error('This PDF already contains a digital signature and cannot be rewritten safely')
  }
  const document = await PDFDocument.load(input, { updateMetadata: false })
  if (
    document
      .getForm()
      .getFields()
      .some((field) => field instanceof PDFSignature)
  ) {
    throw new Error(
      'This PDF already contains a signature field that cannot be timestamped locally',
    )
  }

  installBufferCompatibility()
  const [{ pdflibAddPlaceholder }, { Signer }, { SignPdf }] = await Promise.all([
    import('@signpdf/placeholder-pdf-lib'),
    import('@signpdf/utils'),
    import('@signpdf/signpdf'),
  ])
  class TimestampSigner extends Signer {
    override async sign(pdfBuffer: Buffer): Promise<Buffer> {
      const { bytes: requestBytes, request } = await createPdfTimestampRequestBytes(
        new Uint8Array(pdfBuffer),
      )
      const response = await requestToken(requestBytes)
      return Buffer.from(timestampTokenFromResponseBytes(response, request))
    }
  }

  pdflibAddPlaceholder({
    pdfDoc: document,
    reason: 'RFC 3161 document timestamp',
    contactInfo: '',
    name: 'GenOffice Timestamp',
    location: '',
    signingTime: new Date(),
    signatureLength: 32_768,
    subFilter: 'ETSI.RFC3161',
    appName: 'GenOffice',
  })
  const placeholder = await document.save({
    useObjectStreams: false,
    updateFieldAppearances: false,
  })
  const timestamped = new Uint8Array(
    await new SignPdf().sign(Buffer.from(placeholder), new TimestampSigner()),
  )
  const audit = await auditPdfSignaturesBytes(timestamped)
  const signature = audit.signatures.at(-1)
  if (
    !signature ||
    signature.type !== 'documentTimestamp' ||
    signature.cryptographicStatus !== 'valid' ||
    signature.timestampImprintValid !== true ||
    !signature.timestampTime
  ) {
    throw new Error(signature?.error || 'The generated document timestamp could not be verified')
  }
  return {
    bytes: timestamped,
    timestampTime: signature.timestampTime,
    ...(signature.certificate ? { certificate: signature.certificate } : {}),
  }
}
