import * as asn1js from 'asn1js'
import {
  Certificate,
  ContentInfo,
  SignedData,
  SignedDataVerifyError,
  TSTInfo,
  type RelativeDistinguishedNames,
} from 'pkijs'
import {
  PDFArray,
  PDFDict,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFObject,
  PDFRef,
  PDFString,
} from 'pdf-lib'

export type PdfSignatureType = 'approval' | 'certification' | 'documentTimestamp' | 'unknown'
export type PdfSignatureCryptoStatus = 'valid' | 'invalid' | 'unsupported' | 'error' | 'notChecked'

export interface PdfSignatureCertificate {
  subject: string
  issuer: string
  serialNumber: string
  validFrom: string
  validUntil: string
  selfSigned: boolean
  validAtSigningTime?: boolean
}

export interface PdfSignatureAuditItem {
  fieldName: string
  type: PdfSignatureType
  subFilter?: string
  signerName?: string
  reason?: string
  location?: string
  claimedSigningTime?: string
  timestampTime?: string
  byteRange: number[]
  byteRangeValid: boolean
  signedRevisionLength?: number
  coversEntireDocument: boolean
  unsignedTailBytes: number
  cryptographicStatus: PdfSignatureCryptoStatus
  digestAlgorithm?: string
  signatureAlgorithm?: string
  timestampImprintValid?: boolean
  certificate?: PdfSignatureCertificate
  trustStatus: 'notEvaluated'
  warnings: string[]
  error?: string
}

export interface PdfSignatureAudit {
  signatures: PdfSignatureAuditItem[]
  validCount: number
  invalidCount: number
  unsupportedCount: number
}

interface SignatureDictionaryRecord {
  fieldName: string
  dictionary: PDFDict
}

interface ByteRangeResult {
  values: number[]
  valid: boolean
  signedBytes?: Uint8Array
  signedRevisionLength?: number
  coversEntireDocument: boolean
  unsignedTailBytes: number
  error?: string
}

interface CmsVerificationResult {
  status: PdfSignatureCryptoStatus
  digestAlgorithm?: string
  signatureAlgorithm?: string
  timestampTime?: string
  timestampImprintValid?: boolean
  certificate?: PdfSignatureCertificate
  error?: string
}

const SIGNED_DATA_OID = '1.2.840.113549.1.7.2'
const OID_NAMES: Record<string, string> = {
  '1.2.840.113549.1.1.1': 'RSA',
  '1.2.840.113549.1.1.5': 'SHA-1 with RSA',
  '1.2.840.113549.1.1.10': 'RSA-PSS',
  '1.2.840.113549.1.1.11': 'SHA-256 with RSA',
  '1.2.840.113549.1.1.12': 'SHA-384 with RSA',
  '1.2.840.113549.1.1.13': 'SHA-512 with RSA',
  '1.2.840.10045.4.3.2': 'SHA-256 with ECDSA',
  '1.2.840.10045.4.3.3': 'SHA-384 with ECDSA',
  '1.2.840.10045.4.3.4': 'SHA-512 with ECDSA',
  '1.3.14.3.2.26': 'SHA-1',
  '2.16.840.1.101.3.4.2.1': 'SHA-256',
  '2.16.840.1.101.3.4.2.2': 'SHA-384',
  '2.16.840.1.101.3.4.2.3': 'SHA-512',
}

const RDN_NAMES: Record<string, string> = {
  '2.5.4.3': 'CN',
  '2.5.4.6': 'C',
  '2.5.4.7': 'L',
  '2.5.4.8': 'ST',
  '2.5.4.10': 'O',
  '2.5.4.11': 'OU',
  '1.2.840.113549.1.9.1': 'E',
}

function exactArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  return Uint8Array.from(bytes).buffer
}

function resolvedPdfObject(
  document: PDFDocument,
  value: PDFObject | undefined,
): PDFObject | undefined {
  if (!value) return undefined
  return value instanceof PDFRef ? document.context.lookup(value) : value
}

function decodedPdfText(document: PDFDocument, value: PDFObject | undefined): string | undefined {
  const resolved = resolvedPdfObject(document, value)
  if (
    resolved instanceof PDFString ||
    resolved instanceof PDFHexString ||
    resolved instanceof PDFName
  ) {
    return resolved.decodeText()
  }
  return undefined
}

function collectSignatureDictionaries(document: PDFDocument): SignatureDictionaryRecord[] {
  const output: SignatureDictionaryRecord[] = []
  const visitedFields = new Set<PDFDict>()
  const visitedSignatures = new Set<PDFDict>()

  const addSignature = (fieldName: string, dictionary: PDFDict | undefined) => {
    if (!dictionary || visitedSignatures.has(dictionary)) return
    visitedSignatures.add(dictionary)
    output.push({ fieldName: fieldName || `Signature ${output.length + 1}`, dictionary })
  }

  const visitField = (
    entry: PDFObject | undefined,
    parentName: string,
    inheritedType?: string,
    depth = 0,
  ) => {
    if (depth > 64) return
    const field = resolvedPdfObject(document, entry)
    if (!(field instanceof PDFDict) || visitedFields.has(field)) return
    visitedFields.add(field)
    const partialName = decodedPdfText(document, field.get(PDFName.of('T')))
    const fieldName = partialName
      ? parentName
        ? `${parentName}.${partialName}`
        : partialName
      : parentName
    const fieldType = field.lookupMaybe(PDFName.of('FT'), PDFName)?.decodeText() ?? inheritedType
    if (fieldType === 'Sig') {
      addSignature(
        fieldName,
        resolvedPdfObject(document, field.get(PDFName.of('V'))) as PDFDict | undefined,
      )
    }
    const kids = field.lookupMaybe(PDFName.of('Kids'), PDFArray)
    if (!kids) return
    for (let index = 0; index < kids.size(); index++) {
      visitField(kids.get(index), fieldName, fieldType, depth + 1)
    }
  }

  const acroForm = document.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const fields = acroForm?.lookupMaybe(PDFName.of('Fields'), PDFArray)
  if (fields) {
    for (let index = 0; index < fields.size(); index++) visitField(fields.get(index), '')
  }

  const permissions = document.catalog.lookupMaybe(PDFName.of('Perms'), PDFDict)
  for (const key of ['DocMDP', 'UR', 'UR3']) {
    const value = resolvedPdfObject(document, permissions?.get(PDFName.of(key)))
    addSignature(key, value instanceof PDFDict ? value : undefined)
  }
  return output
}

function signatureType(dictionary: PDFDict, subFilter: string | undefined): PdfSignatureType {
  if (subFilter?.toLowerCase() === 'etsi.rfc3161') return 'documentTimestamp'
  const references = dictionary.lookupMaybe(PDFName.of('Reference'), PDFArray)
  if (references) {
    for (let index = 0; index < references.size(); index++) {
      const reference = references.lookupMaybe(index, PDFDict)
      if (
        reference?.lookupMaybe(PDFName.of('TransformMethod'), PDFName)?.decodeText() === 'DocMDP'
      ) {
        return 'certification'
      }
    }
  }
  return dictionary.has(PDFName.of('ByteRange')) ? 'approval' : 'unknown'
}

function readByteRange(dictionary: PDFDict, input: Uint8Array): ByteRangeResult {
  const array = dictionary.lookupMaybe(PDFName.of('ByteRange'), PDFArray)
  const values: number[] = []
  if (array) {
    for (let index = 0; index < array.size(); index++) {
      const number = array.lookupMaybe(index, PDFNumber)?.asNumber()
      if (number === undefined) {
        return {
          values,
          valid: false,
          coversEntireDocument: false,
          unsignedTailBytes: input.length,
          error: 'ByteRange contains a non-numeric value.',
        }
      }
      values.push(number)
    }
  }
  if (values.length < 4 || values.length % 2 !== 0) {
    return {
      values,
      valid: false,
      coversEntireDocument: false,
      unsignedTailBytes: input.length,
      error: 'ByteRange must contain at least two offset-length pairs.',
    }
  }
  let previousEnd = 0
  let totalLength = 0
  for (let index = 0; index < values.length; index += 2) {
    const offset = values[index]
    const length = values[index + 1]
    const end = offset + length
    if (
      !Number.isSafeInteger(offset) ||
      !Number.isSafeInteger(length) ||
      offset < 0 ||
      length < 0 ||
      offset < previousEnd ||
      end > input.length ||
      (index === 0 && offset !== 0)
    ) {
      return {
        values,
        valid: false,
        coversEntireDocument: false,
        unsignedTailBytes: input.length,
        error: 'ByteRange offsets are invalid or outside the file.',
      }
    }
    previousEnd = end
    totalLength += length
  }
  const signedBytes = new Uint8Array(totalLength)
  let writeOffset = 0
  for (let index = 0; index < values.length; index += 2) {
    const offset = values[index]
    const length = values[index + 1]
    signedBytes.set(input.subarray(offset, offset + length), writeOffset)
    writeOffset += length
  }
  const signedRevisionLength = values.at(-2)! + values.at(-1)!
  return {
    values,
    valid: true,
    signedBytes,
    signedRevisionLength,
    coversEntireDocument: signedRevisionLength === input.length,
    unsignedTailBytes: Math.max(0, input.length - signedRevisionLength),
  }
}

function asn1ValueText(value: asn1js.BaseBlock): string {
  const valueBlock = value.valueBlock as unknown as { value?: unknown }
  if (typeof valueBlock.value === 'string') return valueBlock.value
  const json = value.toJSON() as { valueBlock?: { value?: unknown } }
  if (typeof json.valueBlock?.value === 'string') return json.valueBlock.value
  return value.toString()
}

function rdnText(rdn: RelativeDistinguishedNames): string {
  return rdn.typesAndValues
    .map(({ type, value }) => `${RDN_NAMES[type] ?? type}=${asn1ValueText(value)}`)
    .join(', ')
}

function bytesEqual(left: Uint8Array, right: Uint8Array): boolean {
  return left.length === right.length && left.every((value, index) => value === right[index])
}

async function certificateDetails(
  certificate: Certificate,
  validationTime: Date | undefined,
): Promise<PdfSignatureCertificate> {
  const validFrom = certificate.notBefore.value
  const validUntil = certificate.notAfter.value
  const issuerBytes = new Uint8Array(certificate.issuer.valueBeforeDecode)
  const subjectBytes = new Uint8Array(certificate.subject.valueBeforeDecode)
  const selfIssued = bytesEqual(issuerBytes, subjectBytes)
  const selfSigned = selfIssued ? await certificate.verify(certificate).catch(() => false) : false
  return {
    subject: rdnText(certificate.subject),
    issuer: rdnText(certificate.issuer),
    serialNumber: Array.from(certificate.serialNumber.valueBlock.valueHexView, (value) =>
      value.toString(16).padStart(2, '0'),
    ).join(''),
    validFrom: validFrom.toISOString(),
    validUntil: validUntil.toISOString(),
    selfSigned,
    ...(validationTime
      ? { validAtSigningTime: validationTime >= validFrom && validationTime <= validUntil }
      : {}),
  }
}

function oidName(oid: string | undefined): string | undefined {
  return oid ? (OID_NAMES[oid] ?? oid) : undefined
}

function certificateFromError(error: unknown): Certificate | undefined {
  return error instanceof SignedDataVerifyError && error.signerCertificate
    ? error.signerCertificate
    : undefined
}

async function verifyCmsSignature(
  contents: Uint8Array,
  signedBytes: Uint8Array,
  subFilter: string | undefined,
  claimedSigningTime: Date | undefined,
): Promise<CmsVerificationResult> {
  const normalizedSubFilter = subFilter?.toLowerCase()
  if (normalizedSubFilter === 'adbe.x509.rsa_sha1') {
    return {
      status: 'unsupported',
      error: 'The legacy adbe.x509.rsa_sha1 signature format is not supported.',
    }
  }
  const parsed = asn1js.fromBER(exactArrayBuffer(contents))
  if (parsed.offset === -1) {
    return { status: 'error', error: 'The CMS signature container is malformed.' }
  }
  const contentInfo = new ContentInfo({ schema: parsed.result })
  if (contentInfo.contentType !== SIGNED_DATA_OID) {
    return { status: 'unsupported', error: 'The signature container is not CMS SignedData.' }
  }
  const signedData = new SignedData({ schema: contentInfo.content })
  if (signedData.signerInfos.length === 0) {
    return { status: 'error', error: 'The CMS signature contains no signer.' }
  }
  const digestAlgorithm = oidName(signedData.digestAlgorithms[0]?.algorithmId)
  const signatureAlgorithm = oidName(signedData.signerInfos[0]?.signatureAlgorithm.algorithmId)
  const isTimestamp = normalizedSubFilter === 'etsi.rfc3161'
  const isSha1Container = normalizedSubFilter === 'adbe.pkcs7.sha1'
  let timestampTime: string | undefined
  let timestampImprintValid: boolean | undefined
  let verificationTime = claimedSigningTime

  if (isTimestamp) {
    const timestampContent = signedData.encapContentInfo.eContent?.getValue()
    if (!timestampContent) {
      return {
        status: 'error',
        digestAlgorithm,
        signatureAlgorithm,
        error: 'The RFC 3161 token has no timestamp payload.',
      }
    }
    const timestampSchema = asn1js.fromBER(timestampContent)
    if (timestampSchema.offset === -1) {
      return {
        status: 'error',
        digestAlgorithm,
        signatureAlgorithm,
        error: 'The RFC 3161 timestamp payload is malformed.',
      }
    }
    const timestamp = new TSTInfo({ schema: timestampSchema.result })
    timestampTime = timestamp.genTime.toISOString()
    verificationTime = timestamp.genTime
    timestampImprintValid = await timestamp.verify({ data: exactArrayBuffer(signedBytes) })
  } else if (isSha1Container) {
    const encapsulatedDigest = signedData.encapContentInfo.eContent?.getValue()
    const computedDigest = new Uint8Array(
      await globalThis.crypto.subtle.digest('SHA-1', exactArrayBuffer(signedBytes)),
    )
    timestampImprintValid = encapsulatedDigest
      ? bytesEqual(new Uint8Array(encapsulatedDigest), computedDigest)
      : false
  }

  let signerCertificate: Certificate | undefined
  try {
    for (let signer = 0; signer < signedData.signerInfos.length; signer++) {
      const result = await signedData.verify({
        signer,
        checkChain: false,
        extendedMode: true,
        ...(isSha1Container ? {} : { data: exactArrayBuffer(signedBytes) }),
      })
      signerCertificate ??= result.signerCertificate ?? undefined
      if (result.signatureVerified !== true) {
        return {
          status: 'invalid',
          digestAlgorithm,
          signatureAlgorithm,
          timestampTime,
          timestampImprintValid,
          ...(signerCertificate
            ? { certificate: await certificateDetails(signerCertificate, verificationTime) }
            : {}),
          error: result.message || 'The cryptographic signature is invalid.',
        }
      }
    }
  } catch (error) {
    signerCertificate ??= certificateFromError(error)
    const invalid = error instanceof SignedDataVerifyError && error.signatureVerified === false
    return {
      status: invalid ? 'invalid' : 'error',
      digestAlgorithm,
      signatureAlgorithm,
      timestampTime,
      timestampImprintValid,
      ...(signerCertificate
        ? { certificate: await certificateDetails(signerCertificate, verificationTime) }
        : {}),
      error: error instanceof Error ? error.message : 'CMS verification failed.',
    }
  }

  const status = timestampImprintValid === false ? 'invalid' : 'valid'
  return {
    status,
    digestAlgorithm,
    signatureAlgorithm,
    timestampTime,
    timestampImprintValid,
    ...(signerCertificate
      ? { certificate: await certificateDetails(signerCertificate, verificationTime) }
      : {}),
    ...(status === 'invalid' ? { error: 'The signed document digest does not match.' } : {}),
  }
}

function parsePdfDate(value: string | undefined): { text?: string; date?: Date } {
  if (!value) return {}
  const match =
    /^(?:D:)?(\d{4})(\d{2})?(\d{2})?(\d{2})?(\d{2})?(\d{2})?(?:([Zz])|([+-])(\d{2})'?(\d{2})?'?)?$/.exec(
      value.trim(),
    )
  if (!match) return { text: value }
  const year = Number(match[1])
  const month = Number(match[2] ?? 1) - 1
  const day = Number(match[3] ?? 1)
  const hour = Number(match[4] ?? 0)
  const minute = Number(match[5] ?? 0)
  const second = Number(match[6] ?? 0)
  let milliseconds = Date.UTC(year, month, day, hour, minute, second)
  if (match[8]) {
    const offset = (Number(match[9] ?? 0) * 60 + Number(match[10] ?? 0)) * 60_000
    milliseconds += match[8] === '+' ? -offset : offset
  }
  const date = new Date(milliseconds)
  return Number.isNaN(date.getTime()) ? { text: value } : { text: date.toISOString(), date }
}

export async function auditPdfSignaturesBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfSignatureAudit> {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const document = await PDFDocument.load(input, { updateMetadata: false })
  const records = collectSignatureDictionaries(document)
  const signatures = await Promise.all(
    records.map(async ({ fieldName, dictionary }): Promise<PdfSignatureAuditItem> => {
      const subFilter = dictionary.lookupMaybe(PDFName.of('SubFilter'), PDFName)?.decodeText()
      const claimedDate = parsePdfDate(decodedPdfText(document, dictionary.get(PDFName.of('M'))))
      const byteRange = readByteRange(dictionary, input)
      const contents = resolvedPdfObject(document, dictionary.get(PDFName.of('Contents')))
      const warnings: string[] = []
      if (byteRange.valid && !byteRange.coversEntireDocument) {
        warnings.push('The signature covers an earlier revision; later bytes are not signed.')
      }
      let verification: CmsVerificationResult = {
        status: 'notChecked',
        error: byteRange.error,
      }
      if (byteRange.valid && byteRange.signedBytes) {
        if (!(contents instanceof PDFString || contents instanceof PDFHexString)) {
          verification = { status: 'error', error: 'The signature has no Contents value.' }
        } else {
          verification = await verifyCmsSignature(
            contents.asBytes(),
            byteRange.signedBytes,
            subFilter,
            claimedDate.date,
          )
        }
      }
      const signerName =
        verification.certificate && rdnCommonNameFromText(verification.certificate.subject)
          ? rdnCommonNameFromText(verification.certificate.subject)
          : decodedPdfText(document, dictionary.get(PDFName.of('Name')))
      return {
        fieldName,
        type: signatureType(dictionary, subFilter),
        ...(subFilter ? { subFilter } : {}),
        ...(signerName ? { signerName } : {}),
        ...(decodedPdfText(document, dictionary.get(PDFName.of('Reason')))
          ? { reason: decodedPdfText(document, dictionary.get(PDFName.of('Reason'))) }
          : {}),
        ...(decodedPdfText(document, dictionary.get(PDFName.of('Location')))
          ? { location: decodedPdfText(document, dictionary.get(PDFName.of('Location'))) }
          : {}),
        ...(claimedDate.text ? { claimedSigningTime: claimedDate.text } : {}),
        byteRange: byteRange.values,
        byteRangeValid: byteRange.valid,
        ...(byteRange.signedRevisionLength !== undefined
          ? { signedRevisionLength: byteRange.signedRevisionLength }
          : {}),
        coversEntireDocument: byteRange.coversEntireDocument,
        unsignedTailBytes: byteRange.unsignedTailBytes,
        cryptographicStatus: verification.status,
        trustStatus: 'notEvaluated',
        warnings,
        ...(verification.digestAlgorithm ? { digestAlgorithm: verification.digestAlgorithm } : {}),
        ...(verification.signatureAlgorithm
          ? { signatureAlgorithm: verification.signatureAlgorithm }
          : {}),
        ...(verification.timestampTime ? { timestampTime: verification.timestampTime } : {}),
        ...(verification.timestampImprintValid !== undefined
          ? { timestampImprintValid: verification.timestampImprintValid }
          : {}),
        ...(verification.certificate ? { certificate: verification.certificate } : {}),
        ...(verification.error ? { error: verification.error } : {}),
      }
    }),
  )
  return {
    signatures,
    validCount: signatures.filter(({ cryptographicStatus }) => cryptographicStatus === 'valid')
      .length,
    invalidCount: signatures.filter(
      ({ cryptographicStatus, byteRangeValid }) =>
        !byteRangeValid || cryptographicStatus === 'invalid' || cryptographicStatus === 'error',
    ).length,
    unsupportedCount: signatures.filter(
      ({ cryptographicStatus }) => cryptographicStatus === 'unsupported',
    ).length,
  }
}

function rdnCommonNameFromText(subject: string): string | undefined {
  return subject
    .split(', ')
    .find((entry) => entry.startsWith('CN='))
    ?.slice(3)
}
