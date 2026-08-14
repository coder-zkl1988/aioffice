import { Buffer } from 'buffer'
import { PDFDocument, PDFSignature } from 'pdf-lib'
import { auditPdfSignaturesBytes, type PdfSignatureCertificate } from './signature-audit'

export interface PdfCertificateSignOptions {
  certificate: Uint8Array
  password: string
  signerName: string
  reason: string
  location: string
  contactInfo: string
}

export interface PdfCertificateSignResult {
  bytes: Uint8Array
  certificate: PdfSignatureCertificate
}

type GlobalWithBuffer = typeof globalThis & { Buffer?: typeof Buffer }

function installBufferCompatibility(): void {
  const target = globalThis as GlobalWithBuffer
  if (!target.Buffer) target.Buffer = Buffer
}

function certificateSignError(error: unknown): Error {
  const message = error instanceof Error ? error.message : String(error)
  if (
    /Invalid password|PKCS#12 MAC could not be verified|Unable to decrypt|Cannot read/i.test(
      message,
    )
  ) {
    return new Error('The certificate password is incorrect or the P12/PFX file is invalid')
  }
  if (/certificate that matches the private key/i.test(message)) {
    return new Error('The P12/PFX file does not contain a certificate matching its private key')
  }
  if (/Signature exceeds placeholder length/i.test(message)) {
    return new Error('The certificate chain is too large to embed in this PDF')
  }
  return error instanceof Error ? error : new Error(message)
}

export async function signPdfWithP12Bytes(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfCertificateSignOptions,
): Promise<PdfCertificateSignResult> {
  const input =
    bytes instanceof Uint8Array ? Uint8Array.from(bytes) : new Uint8Array(bytes.slice(0))
  if (input.length === 0) throw new Error('The PDF is empty')
  if (options.certificate.length === 0) throw new Error('Choose a P12 or PFX certificate file')

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
    throw new Error('This PDF already contains a signature field that cannot be filled locally')
  }

  installBufferCompatibility()
  try {
    const [{ pdflibAddPlaceholder }, { P12Signer }, { SignPdf }] = await Promise.all([
      import('@signpdf/placeholder-pdf-lib'),
      import('@signpdf/signer-p12'),
      import('@signpdf/signpdf'),
    ])
    const signingTime = new Date()
    pdflibAddPlaceholder({
      pdfDoc: document,
      reason: options.reason.trim(),
      contactInfo: options.contactInfo.trim(),
      name: options.signerName.trim(),
      location: options.location.trim(),
      signingTime,
      signatureLength: 32_768,
      appName: 'GenOffice',
    })
    const placeholder = await document.save({
      useObjectStreams: false,
      updateFieldAppearances: false,
    })
    const signer = new P12Signer(Buffer.from(options.certificate), {
      passphrase: options.password,
      asn1StrictParsing: false,
    })
    const signed = new Uint8Array(
      await new SignPdf().sign(Buffer.from(placeholder), signer, signingTime),
    )
    const audit = await auditPdfSignaturesBytes(signed)
    const signature = audit.signatures.at(-1)
    if (!signature || signature.cryptographicStatus !== 'valid' || !signature.certificate) {
      throw new Error(signature?.error || 'The generated digital signature could not be verified')
    }
    if (signature.certificate.validAtSigningTime === false) {
      throw new Error('The signing certificate is expired or not valid yet')
    }
    return { bytes: signed, certificate: signature.certificate }
  } catch (error) {
    throw certificateSignError(error)
  }
}
