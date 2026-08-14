import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import qrcode from 'qrcode-generator'
import { Buffer } from 'buffer'
import { XMLParser, XMLValidator } from 'fast-xml-parser'
import {
  buildBlankDocx,
  parseDocx,
  saveDocx,
  type Run,
  type SaveBlock,
  type SectionSettings,
} from '@genoffice/docx-engine'
import { decryptPDF, isEncrypted } from '@pdfsmaller/pdf-decrypt'
import { encryptPDF } from '@pdfsmaller/pdf-encrypt'
import { signPdfWithP12Bytes } from './certificate-sign'
import { pdfLibFontkit } from './pdf-lib-fontkit'
import { SRGB_2014_ICC_BYTES } from './srgb2014'
export * from './signature-audit'
export * from './certificate-sign'
export * from './timestamp'
import {
  PDFArray,
  PDFBool,
  PDFButton,
  PDFCheckBox,
  PDFDict,
  PDFDocument,
  PDFDropdown,
  PDFField,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFObject,
  PDFObjectCopier,
  PDFOptionList,
  PDFPage,
  PDFFont,
  PDFRawStream,
  PDFRadioGroup,
  PDFRef,
  PDFSignature,
  PDFString,
  PDFTextField,
  PDFWidgetAnnotation,
  ParseSpeeds,
  StandardFonts,
  decodePDFRawStream,
  degrees,
  drawObject,
  popGraphicsState,
  pushGraphicsState,
  rgb,
  rotateInPlace,
  translate,
} from 'pdf-lib'

export interface InsertPdfResult {
  merged: Uint8Array
  count: number
}

export interface InsertBlankPageResult {
  merged: Uint8Array
  count: number
}

export interface InsertBlankPageOptions {
  count?: number
  pageSize?: PdfPageSize
  orientation?: PdfOrientation
}

export type PdfPageSize = 'KEEP' | 'A3' | 'A4' | 'A5' | 'LETTER' | 'LEGAL' | 'TABLOID'
export type PdfOrientation = 'portrait' | 'landscape'
export type PdfBookletSpine = 'left' | 'right'
export type PdfDuplexPass = 'both' | 'front' | 'back'
export type PdfSinglePageDirection = 'vertical' | 'horizontal'
export type PdfPageRotation = 90 | 180 | 270
export type PdfFormAction =
  'export' | 'fill' | 'create' | 'modify' | 'flatten' | 'unlock' | 'delete'
export type PdfTextExportFormat = 'txt' | 'markdown' | 'both'
export type PdfTableExportFormat = 'csv' | 'xlsx' | 'both'
export type PdfFlattenMode = 'pages' | 'forms'
export type PdfOverlayMode = 'sequential' | 'interleaved' | 'fixedRepeat'
export type PdfOverlayPosition = 'foreground' | 'background'
export type PdfImageOverlayPosition =
  | 'topLeft'
  | 'topCenter'
  | 'topRight'
  | 'middleLeft'
  | 'center'
  | 'middleRight'
  | 'bottomLeft'
  | 'bottomCenter'
  | 'bottomRight'
  | 'custom'
export type PdfContentFilterCriterion = 'text' | 'image' | 'pageSize' | 'orientation' | 'rotation'
export type PdfContentFilterAction = 'keep' | 'remove'
export type PdfPageFilterComparator = 'less' | 'equal' | 'greater'
export type PdfPageFilterRotation = 0 | 90 | 180 | 270
export type PdfDocumentFilterCriterion =
  'text' | 'image' | 'pageCount' | 'fileSize' | 'pageSize' | 'rotation'
export interface PdfDocumentFilterInput {
  fileName: string
  bytes: Uint8Array
  contentMatched?: boolean
}
export type PdfImagesToPdfFit = 'fillPage' | 'fitDocumentToImage' | 'maintainAspectRatio'
export type PdfMetadataTrapped = '' | 'True' | 'False' | 'Unknown'
export type PdfAutoRenameStrategy = 'largestHeading' | 'firstText'
export type PdfPageNumberPosition =
  | 'topLeft'
  | 'topCenter'
  | 'topRight'
  | 'middleLeft'
  | 'center'
  | 'middleRight'
  | 'bottomLeft'
  | 'bottomCenter'
  | 'bottomRight'
export type PdfPageNumberMargin = 'small' | 'medium' | 'large' | 'xLarge'
export type PdfPageNumberFont = 'helvetica' | 'times' | 'courier'
export type PdfAttachmentAction = 'add' | 'extract' | 'rename' | 'delete'
export type PdfPasswordAlgorithm = 'AES-256' | 'RC4'
export type PdfOcrMode = 'skipText' | 'force' | 'strict'
export type PdfOcrLanguage = 'eng' | 'chi_sim'
export type PdfDocxMode = 'editableText' | 'fidelity'

export interface PdfEncryptionInfo {
  encrypted: boolean
  algorithm?: PdfPasswordAlgorithm
  version?: number
  revision?: number
  keyLength?: number
}

export interface PdfSecurityReport extends PdfEncryptionInfo {
  rawPermissions?: number
  permissions: PdfPasswordPermissions | null
  restrictedCount: number
}
export type PdfExtractImageFormat = 'png' | 'jpg' | 'gif'
export type PdfPageImageFormat = 'png' | 'jpg' | 'gif' | 'webp'
export type PdfPageImageColorMode = 'color' | 'greyscale' | 'blackwhite'
export type PdfPageImageOutputMode = 'single' | 'multiple'
export type PdfVideoResolution = '480p' | '720p' | '1080p'
export type PdfSplitMode =
  'afterPages' | 'fileSize' | 'pagesPerDocument' | 'documentCount' | 'chapters'
export type PdfArchiveFormat = 'PDF/A-2b'
export type PdfArchiveMode = 'auto' | 'raster'
export type PdfSectionArrangement = 'rows' | 'columns'
export type PdfRearrangeMode =
  | 'custom'
  | 'reverse'
  | 'oddEven'
  | 'duplex'
  | 'removeFirst'
  | 'removeLast'
  | 'removeFirstAndLast'
  | 'duplicate'

export interface CropMargins {
  top: number
  right: number
  bottom: number
  left: number
}

export interface PdfPageCropBox {
  x: number
  y: number
  width: number
  height: number
}

export interface NUpOptions {
  rows: number
  columns: number
  orientation: PdfOrientation
  arrangement?: 'rows' | 'columns'
  readingDirection?: 'ltr' | 'rtl'
  innerMargin?: number
  topMargin?: number
  rightMargin?: number
  bottomMargin?: number
  leftMargin?: number
  borderWidth?: number
}

export interface BookletOptions {
  spine: PdfBookletSpine
  gutter: number
  border: boolean
  duplexPass: PdfDuplexPass
  flipOnShortEdge: boolean
}

export interface PosterOptions {
  pageSize: Exclude<PdfPageSize, 'KEEP'>
  rows: number
  columns: number
  readingDirection: 'ltr' | 'rtl'
}

export interface SinglePageOptions {
  direction: PdfSinglePageDirection
}

export interface RearrangeOptions {
  mode: PdfRearrangeMode
  customOrder?: number[]
  duplicateCount?: number
}

export interface SanitizeOptions {
  removeJavaScript: boolean
  removeEmbeddedFiles: boolean
  removeXmpMetadata: boolean
  removeMetadata: boolean
  removeLinks: boolean
}

export type PdfPipelineStep =
  | ({ kind: 'sanitize' } & SanitizeOptions)
  | { kind: 'removeAnnotations' }
  | { kind: 'removeSignatures' }
  | { kind: 'flattenForms' }
  | { kind: 'repair' }
  | { kind: 'decompress' }

export interface OverlayOptions {
  overlayDocuments: Uint8Array[]
  mode: PdfOverlayMode
  position: PdfOverlayPosition
  opacity: number
  repeatCounts?: number[]
}

export interface ImageOverlayOptions {
  image: Uint8Array
  pageIndexes: number[]
  position: PdfImageOverlayPosition
  widthPercent: number
  margin: number
  opacity: number
  layer: PdfOverlayPosition
  x?: number
  y?: number
}

export interface PdfImageOverlayPlacement {
  x: number
  y: number
  width: number
  height: number
  rotation: 0 | 90 | 180 | 270
}

export interface PdfPageNumberOptions {
  pageIndexes: number[]
  position: PdfPageNumberPosition
  margin: PdfPageNumberMargin
  fontSize: number
  font: PdfPageNumberFont
  fontColor: string
  startingNumber: number
  zeroPad: number
  textPattern: string
  baseName?: string
}

export interface PdfPageNumberLabel {
  pageIndex: number
  text: string
}

export interface PdfAttachmentInput {
  name: string
  bytes: Uint8Array
  mimeType?: string
}

export interface PdfAttachmentInfo {
  name: string
  size: number
  mimeType?: string
  description?: string
}

export interface PdfBookmark {
  title: string
  pageNumber: number
  children: PdfBookmark[]
}

export interface PdfChapterSplitOutput {
  title: string
  bytes: Uint8Array
}

export interface PdfSectionSplitOutput {
  sourcePageNumber: number
  sectionNumber: number
  bytes: Uint8Array
}

export interface PdfPageAnalysis {
  pageNumber: number
  width: number
  height: number
  rotation: number
}

export interface PdfFontInfo {
  name: string
  subtype: string
  embedded: boolean
  subset: boolean
  encoding?: string
  hasToUnicode: boolean
  pages: number[]
}

export interface PdfFontReport {
  fontCount: number
  embeddedCount: number
  subsetCount: number
  fonts: PdfFontInfo[]
}

export interface PdfaPreservationReport {
  eligible: boolean
  fontCount: number
  embeddedFontCount: number
  unembeddedFonts: string[]
}

export interface PdfToPdfaOptions {
  format: PdfArchiveFormat
  archiveMode: PdfArchiveMode
  renderDpi: number
  imageQuality: number
  pageImages?: Uint8Array[]
}

export interface PdfAnnotationInfo {
  pageNumber: number
  annotationNumber: number
  subtype: string
  author?: string
  subject?: string
  contents?: string
  modifiedAt?: string
  name?: string
  flags?: number
  rectangle?: {
    x: number
    y: number
    width: number
    height: number
  }
}

export interface PdfAnnotationReport {
  totalCount: number
  typeBreakdown: Record<string, number>
  annotations: PdfAnnotationInfo[]
}

export interface PdfAnalysis {
  pageCount: number
  pdfVersion?: string
  fileSize: number
  isEncrypted: boolean
  properties: {
    title?: string
    author?: string
    subject?: string
    keywords?: string
    creator?: string
    producer?: string
    creationDate?: string
    modificationDate?: string
    trapped?: Exclude<PdfMetadataTrapped, ''>
    custom: PdfMetadataCustomField[]
  }
  pages: PdfPageAnalysis[]
  form: {
    fieldCount: number
    hasXfa: boolean
    signatureCount: number
  }
  annotations: {
    totalCount: number
    typeBreakdown: Record<string, number>
  }
  fonts: string[]
  imageCount: number
  attachmentCount: number
  bookmarkCount: number
}

export type PdfPreflightSeverity = 'error' | 'warning' | 'info'

export type PdfPreflightFindingCode =
  | 'missingPdfHeader'
  | 'headerNotAtStart'
  | 'missingEofMarker'
  | 'largeTrailingData'
  | 'missingStartXref'
  | 'startXrefOutOfRange'
  | 'startXrefTargetInvalid'
  | 'strictParseFailed'
  | 'emptyPageTree'
  | 'invalidPageBox'
  | 'invalidXmpMetadata'
  | 'standardDeclaredOnly'
  | 'pdfaMissingOutputIntent'
  | 'pdfaAttachmentsRisk'
  | 'pdfaJavaScriptRisk'
  | 'pdfuaNotTagged'
  | 'pdfuaMissingLanguage'
  | 'javascriptPresent'
  | 'attachmentsPresent'
  | 'xfaPresent'
  | 'signaturesPresent'

export interface PdfPreflightFinding {
  code: PdfPreflightFindingCode
  severity: PdfPreflightSeverity
  detail?: string
  pageNumber?: number
}

export interface PdfStandardDeclaration {
  family: 'PDF/A' | 'PDF/UA'
  part: string
  conformance?: string
  revision?: string
  label: string
}

export interface PdfPreflightReport {
  schema: 'genoffice.pdf.preflight'
  version: 1
  status: 'pass' | 'warning' | 'error'
  fileSize: number
  pdfVersion?: string
  parseable: boolean
  strictParsing: boolean
  pageCount: number
  structure: {
    headerOffset: number
    eofMarkerCount: number
    trailingBytes: number
    startXrefOffset?: number
    startXrefInRange: boolean
    startXrefTargetValid: boolean
    incrementalUpdates: number
  }
  standards: PdfStandardDeclaration[]
  features: {
    hasXmpMetadata: boolean
    xmpValid: boolean
    tagged: boolean
    marked: boolean
    language?: string
    outputIntentCount: number
    javaScriptActionCount: number
    attachmentCount: number
    formFieldCount: number
    hasXfa: boolean
    signatureCount: number
    encrypted: boolean
  }
  findings: PdfPreflightFinding[]
  disclaimer: 'local-structural-preflight'
}

export interface PdfMetadataCustomField {
  key: string
  value: string
}

export interface PdfMetadataValues {
  title: string
  author: string
  subject: string
  keywords: string
  creator: string
  producer: string
  creationDate: string
  modificationDate: string
  trapped: PdfMetadataTrapped
  custom: PdfMetadataCustomField[]
}

export interface PdfMetadataOptions {
  deleteAll: boolean
  metadata: PdfMetadataValues
}

export const PDF_CLASSIFICATION_METADATA_KEY = 'GenOfficeClassification'

export type PdfClassificationSensitivity = 'standard' | 'internal' | 'confidential' | 'restricted'

export interface PdfClassificationLabel {
  id: string
  name: string
}

export interface PdfClassificationMetadata {
  labels: PdfClassificationLabel[]
  sensitivity: PdfClassificationSensitivity
}

export type PdfJavaScriptSource = 'named' | 'document' | 'page' | 'annotation' | 'form'

export interface PdfJavaScriptAction {
  source: PdfJavaScriptSource
  trigger: string
  code: string
  name?: string
  pageNumber?: number
  annotationNumber?: number
  fieldName?: string
}

export interface PdfJavaScriptAudit {
  actions: PdfJavaScriptAction[]
  uniqueScriptCount: number
  totalCodeBytes: number
}

export type PdfFormFieldType =
  'text' | 'checkbox' | 'radio' | 'dropdown' | 'optionList' | 'button' | 'signature' | 'unknown'

export interface PdfFormFieldInfo {
  name: string
  label?: string
  type: PdfFormFieldType
  readOnly: boolean
  required: boolean
  multiline?: boolean
  multiselect?: boolean
  editable?: boolean
  value?: string | boolean | string[]
  options?: string[]
}

export interface PdfFormFieldValue {
  name: string
  value: string | boolean | string[]
}

export interface PdfFormFieldModification {
  name: string
  newName?: string
  label?: string
  readOnly?: boolean
  required?: boolean
  options?: string[]
  multiselect?: boolean
}

export type PdfCreatableFormFieldType = 'text' | 'checkbox' | 'radio' | 'dropdown' | 'optionList'

export interface PdfFormFieldCreation {
  name: string
  label?: string
  type: PdfCreatableFormFieldType
  pageIndex: number
  x: number
  y: number
  width: number
  height: number
  required?: boolean
  readOnly?: boolean
  multiline?: boolean
  multiselect?: boolean
  options?: string[]
  defaultValue?: string | boolean | string[]
  optionSpacing?: number
}

export type PdfTextBlockKind = 'heading' | 'paragraph' | 'listItem'

export interface PdfExtractedTextBlock {
  kind: PdfTextBlockKind
  text: string
  level?: 1 | 2 | 3
}

export interface PdfExtractedTextLink {
  url: string
  label?: string
}

export interface PdfExtractedTextPage {
  pageNumber: number
  text: string
  blocks: PdfExtractedTextBlock[]
  links: PdfExtractedTextLink[]
}

export interface PdfToMarkdownOptions {
  pageIndexes: number[]
  includePageBreaks: boolean
  baseName?: string
  pages?: PdfExtractedTextPage[]
}

export interface PdfJsonTextRun {
  text: string
  x: number
  y: number
  width: number
  height: number
  fontSize: number
  fontFamily?: string
  bold: boolean
  italic: boolean
}

export interface PdfJsonPage extends PdfExtractedTextPage {
  width: number
  height: number
  rotation: number
  textRuns?: PdfJsonTextRun[]
}

export interface PdfToJsonOptions {
  pageIndexes: number[]
  lightweight: boolean
  baseName?: string
  pages?: PdfJsonPage[]
}

export interface PdfJsonImportFonts {
  regular: Uint8Array
  bold: Uint8Array
  italic: Uint8Array
  boldItalic: Uint8Array
  unicode: Uint8Array
}

export interface JsonToPdfOptions {
  jsonBytes: Uint8Array
  fonts?: PdfJsonImportFonts
  baseName?: string
}

export type PdfToXmlOptions = PdfToJsonOptions

export interface PdfToVideoOptions {
  pageIndexes: number[]
  secondsPerPage: number
  resolution: PdfVideoResolution
  transitionSeconds: number
  includeAnnotations: boolean
  baseName?: string
  videoBytes?: Uint8Array
}

export type PdfEpubMode = 'reflowable' | 'fixed'

export interface PdfEpubPage extends PdfExtractedTextPage {
  width: number
  height: number
  imageBytes?: Uint8Array
}

export interface PdfToEpubOptions {
  pageCount: number
  pageIndexes: number[]
  mode: PdfEpubMode
  renderDpi: number
  includeAnnotations: boolean
  baseName?: string
  pages?: PdfEpubPage[]
}

export type PdfPptxMode = 'editableText' | 'fidelity'

export interface PdfPptxTextRun {
  text: string
  x: number
  y: number
  width: number
  height: number
  fontSize: number
  angle: number
  fontFamily?: string
  bold: boolean
  italic: boolean
}

export interface PdfPptxPage {
  pageNumber: number
  width: number
  height: number
  textRuns: PdfPptxTextRun[]
  imageBytes?: Uint8Array
}

export interface PdfToPptxOptions {
  pageCount: number
  pageIndexes: number[]
  mode: PdfPptxMode
  renderDpi: number
  includeAnnotations: boolean
  baseName?: string
  pages?: PdfPptxPage[]
}

export interface PdfDocxTextRun extends PdfPptxTextRun {
  hasEol?: boolean
}

export interface PdfDocxPage {
  pageNumber: number
  width: number
  height: number
  textRuns: PdfDocxTextRun[]
  imageBytes?: Uint8Array
  imageWidth?: number
  imageHeight?: number
}

export interface PdfToDocxOptions {
  pageCount: number
  pageIndexes: number[]
  mode: PdfDocxMode
  renderDpi: number
  includeAnnotations: boolean
  baseName?: string
  pages?: PdfDocxPage[]
}

export type PdfToOdtOptions = PdfToDocxOptions

export interface PdfToRtfOptions {
  pageCount: number
  pageIndexes: number[]
  baseName?: string
  pages?: PdfDocxPage[]
}

export interface PdfExtractedTable {
  pageNumber: number
  tableNumber: number
  rows: string[][]
}

export interface PdfPageRange {
  firstPage: number
  lastPage: number
}

export const PDF_AUTO_SPLIT_QR_CONTENTS = [
  'https://github.com/Stirling-Tools/Stirling-PDF',
  'https://github.com/Frooodle/Stirling-PDF',
  'https://stirlingpdf.com',
] as const

export const PDF_AUTO_SPLIT_QR_CONTENT = PDF_AUTO_SPLIT_QR_CONTENTS[2]

export interface SplitSectionsOptions {
  pageIndexes: number[]
  rows: number
  columns: number
  merge: boolean
  arrangement?: PdfSectionArrangement
}

export interface PdfColorAdjustments {
  contrast: number
  brightness: number
  saturation: number
  red: number
  green: number
  blue: number
}

export interface PdfLineArtOptions {
  threshold: number
  edgeLevel: 1 | 2 | 3
}

export interface PdfBlankPageDetectionOptions {
  threshold: number
  whitePercent: number
  includeBlankPages: boolean
}

export interface PdfRedactionOptions {
  mode?: 'text' | 'areas'
  patterns: string[]
  useRegex: boolean
  wholeWord: boolean
  color: string
  padding: number
  renderDpi: number
  areas?: PdfRedactionArea[]
  pages?: PdfRedactedPage[]
  pageImages?: Uint8Array[]
}

export interface PdfRedactionArea {
  pageIndex: number
  x: number
  y: number
  width: number
  height: number
}

export interface PdfRedactedPage {
  pageIndex: number
  image: Uint8Array
}

export interface PdfCommentInput {
  pageIndex: number
  x: number
  y: number
  width: number
  height: number
  text: string
  author?: string
  subject?: string
  anchorText?: string
}

export interface PdfPageRotationCorrection {
  pageIndex: number
  angle: PdfPageRotation
}

export interface PdfDeskewPage {
  pageIndex: number
  angle: number
  image: Uint8Array
}

export interface PdfDeskewOptions {
  pageIndexes: number[]
  maxAngle: number
  renderDpi: number
  includeAnnotations: boolean
  pages?: PdfDeskewPage[]
}

export type PdfScannerColorspace = 'grayscale' | 'color'
export type PdfScannerQuality = 'high' | 'medium' | 'low' | 'custom'
export type PdfScannerRotation = 'none' | 'slight' | 'moderate' | 'severe'

export interface PdfScannerEffectOptions {
  quality: PdfScannerQuality
  rotation: PdfScannerRotation
  colorspace: PdfScannerColorspace
  border: number
  rotate: number
  rotateVariance: number
  brightness: number
  contrast: number
  blur: number
  noise: number
  yellowish: boolean
  renderDpi: number
  seed: number
  pageImages?: Uint8Array[]
}

export interface PdfRasterPage {
  image: Uint8Array
  width: number
  height: number
}

export interface PdfVectorPage {
  kind: 'vectorPdf'
  pdf: Uint8Array
}

export type PdfImagesToPdfPage = PdfRasterPage | PdfVectorPage

export type PdfComparisonPage = PdfRasterPage
export type PdfScannedImagePage = PdfRasterPage

export interface PdfImagesToPdfOptions {
  images: PdfImagesToPdfPage[]
  fitOption: PdfImagesToPdfFit
  autoRotate: boolean
  appendToCurrent: boolean
}

export interface CbzToPdfOptions {
  images: PdfRasterPage[]
  fitOption: PdfImagesToPdfFit
  autoRotate: boolean
  baseName?: string
}

export type EmailDocumentOutputFormat = 'pdf' | 'html'

export interface EmailToPdfOptions {
  outputFormat: EmailDocumentOutputFormat
  pages?: PdfRasterPage[]
  htmlBytes?: Uint8Array
  attachments?: PdfAttachmentInput[]
  baseName?: string
}

export interface EpubToPdfOptions {
  pages: PdfRasterPage[]
  baseName?: string
  title?: string
  author?: string
}

export interface HtmlToPdfOptions {
  pages: PdfRasterPage[]
  baseName?: string
  title?: string
}

export interface CreatePdfOptions {
  pages: PdfRasterPage[]
  baseName?: string
  title?: string
}

export interface MarkdownToPdfOptions {
  pages: PdfRasterPage[]
  baseName?: string
  title?: string
}

export interface CbzImageEntry {
  name: string
  bytes: Uint8Array
}

export interface PdfComparisonOptions {
  comparisonDocument: Uint8Array
  renderDpi: number
  threshold: number
  pages?: PdfComparisonPage[]
}

export interface PdfScannerImageSplitOptions {
  angleThreshold: number
  tolerance: number
  minArea: number
  minContourArea: number
  borderSize: number
  renderDpi: number
  pages?: PdfScannedImagePage[]
}

export interface PdfPasswordPermissions {
  allowPrinting: boolean
  allowModifying: boolean
  allowCopying: boolean
  allowAnnotating: boolean
  allowFillingForms: boolean
  allowExtraction: boolean
  allowAssembly: boolean
  allowHighQualityPrint: boolean
}

export interface PdfOcrTextLayer {
  pageIndex: number
  bytes: Uint8Array
  replacePage?: boolean
}

export interface PdfOcrPageText {
  pageIndex: number
  text: string
  source: 'existing' | 'ocr'
}

export interface PdfOcrOptions {
  mode: PdfOcrMode
  languages: PdfOcrLanguage[]
  renderDpi: number
  clean: boolean
  sidecar: boolean
  baseName?: string
  textLayers?: PdfOcrTextLayer[]
  pageTexts?: PdfOcrPageText[]
  skippedPageIndexes?: number[]
}

export interface PdfCertificateSignOperation {
  certificate: Uint8Array
  password: string
  signerName: string
  reason: string
  location: string
  contactInfo: string
}

export interface PdfTimestampOperation {
  tsaUrl: string
  timestampedBytes?: Uint8Array
}

export type PdfPasswordOptions =
  | {
      action: 'protect'
      userPassword: string
      ownerPassword: string
      algorithm: PdfPasswordAlgorithm
      permissions: PdfPasswordPermissions
    }
  | { action: 'unlock'; password: string }

export async function pdfEncryptionInfoBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfEncryptionInfo> {
  return isEncrypted(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes))
}

const ALL_PDF_PERMISSIONS: PdfPasswordPermissions = {
  allowPrinting: true,
  allowHighQualityPrint: true,
  allowModifying: true,
  allowCopying: true,
  allowAnnotating: true,
  allowFillingForms: true,
  allowExtraction: true,
  allowAssembly: true,
}

function pdfPermissionsFromFlags(flags: number): PdfPasswordPermissions {
  return {
    allowPrinting: (flags & 0x0000_0004) !== 0,
    allowModifying: (flags & 0x0000_0008) !== 0,
    allowCopying: (flags & 0x0000_0010) !== 0,
    allowAnnotating: (flags & 0x0000_0020) !== 0,
    allowFillingForms: (flags & 0x0000_0100) !== 0,
    allowExtraction: (flags & 0x0000_0200) !== 0,
    allowAssembly: (flags & 0x0000_0400) !== 0,
    allowHighQualityPrint: (flags & 0x0000_0800) !== 0,
  }
}

export async function analyzePdfSecurityBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfSecurityReport> {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const encryption = await pdfEncryptionInfoBytes(input)
  if (!encryption.encrypted) {
    return { ...encryption, permissions: { ...ALL_PDF_PERMISSIONS }, restrictedCount: 0 }
  }
  const document = await PDFDocument.load(input, {
    ignoreEncryption: true,
    updateMetadata: false,
  })
  const encryptObject = document.context.trailerInfo.Encrypt
  const encryptDictionary = encryptObject
    ? document.context.lookup(encryptObject, PDFDict)
    : undefined
  const rawPermissions = encryptDictionary?.lookupMaybe(PDFName.of('P'), PDFNumber)?.asNumber()
  const permissions = rawPermissions === undefined ? null : pdfPermissionsFromFlags(rawPermissions)
  return {
    ...encryption,
    ...(rawPermissions === undefined ? {} : { rawPermissions }),
    permissions,
    restrictedCount: permissions
      ? Object.values(permissions).filter((allowed) => !allowed).length
      : 0,
  }
}

export interface PdfExtractedImage {
  pageNumber: number
  imageNumber: number
  bytes: Uint8Array
}

export interface PdfRenderedPageImage {
  pageNumber: number
  bytes: Uint8Array
}

export interface PdfToImagesOptions {
  pageCount: number
  pageIndexes: number[]
  format: PdfPageImageFormat
  outputMode: PdfPageImageOutputMode
  renderDpi: number
  imageQuality: number
  colorMode: PdfPageImageColorMode
  includeAnnotations: boolean
  baseName?: string
  images?: PdfRenderedPageImage[]
}

export type PdfComicImageFormat = Extract<PdfPageImageFormat, 'png' | 'jpg' | 'webp'>

export interface PdfToCbzOptions {
  pageCount: number
  pageIndexes: number[]
  format: PdfComicImageFormat
  renderDpi: number
  imageQuality: number
  colorMode: PdfPageImageColorMode
  includeAnnotations: boolean
  baseName?: string
  images?: PdfRenderedPageImage[]
}

export interface PdfHtmlTextRun {
  text: string
  x: number
  y: number
  fontSize: number
  angle: number
  fontFamily?: string
  bold?: boolean
  italic?: boolean
}

export interface PdfHtmlPage {
  pageNumber: number
  width: number
  height: number
  imageBytes: Uint8Array
  text: string
  textRuns: PdfHtmlTextRun[]
}

export interface PdfToHtmlOptions {
  pageCount: number
  pageIndexes: number[]
  renderDpi: number
  includeAnnotations: boolean
  baseName?: string
  pages?: PdfHtmlPage[]
}

export type AttachmentOptions =
  | { action: 'add'; attachments: PdfAttachmentInput[] }
  | { action: 'extract' }
  | { action: 'rename'; attachmentName: string; newName: string }
  | { action: 'delete'; attachmentName: string }

export type SplitOptions =
  | { mode: 'afterPages'; splitAfterPages: number[] }
  | { mode: 'fileSize'; maxBytes: number }
  | { mode: 'pagesPerDocument'; value: number }
  | { mode: 'documentCount'; value: number }
  | { mode: 'chapters'; bookmarkLevel: number; allowDuplicates: boolean }

export type PdfToolOperation =
  | ({ kind: 'split' } & SplitOptions)
  | { kind: 'merge'; documents: Uint8Array[]; currentDocumentIndex: number }
  | ({ kind: 'compare' } & PdfComparisonOptions)
  | { kind: 'extractPages'; pageIndexes: number[] }
  | ({ kind: 'splitSections' } & SplitSectionsOptions)
  | { kind: 'crop'; mode: 'manual'; margins: CropMargins }
  | {
      kind: 'crop'
      mode: 'auto'
      whiteThreshold: number
      padding: number
      pageBoxes?: PdfPageCropBox[]
    }
  | {
      kind: 'scale'
      pageSize: PdfPageSize
      orientation: PdfOrientation
      scaleFactor: number
    }
  | ({ kind: 'nup' } & NUpOptions)
  | ({ kind: 'booklet' } & BookletOptions)
  | ({ kind: 'poster' } & PosterOptions)
  | ({ kind: 'singlePage' } & SinglePageOptions)
  | { kind: 'rotatePages'; pageIndexes: number[]; angle: PdfPageRotation }
  | {
      kind: 'autoRotate'
      inferUndetected: boolean
      pageRotations?: PdfPageRotationCorrection[]
    }
  | ({ kind: 'deskew' } & PdfDeskewOptions)
  | ({ kind: 'scannerEffect' } & PdfScannerEffectOptions)
  | ({ kind: 'scannerImageSplit' } & PdfScannerImageSplitOptions)
  | { kind: 'autoSplit'; action: 'divider' }
  | {
      kind: 'autoSplit'
      action: 'split'
      duplexMode: boolean
      baseName?: string
      dividerPageIndexes?: number[]
    }
  | { kind: 'removePages'; pageIndexes: number[] }
  | { kind: 'removeImages'; pageIndexes: number[] }
  | {
      kind: 'extractText'
      format: PdfTextExportFormat
      pageIndexes: number[]
      pages?: PdfExtractedTextPage[]
    }
  | ({ kind: 'pdfToMarkdown' } & PdfToMarkdownOptions)
  | {
      kind: 'extractTables'
      format: PdfTableExportFormat
      pageIndexes: number[]
      includeTwoColumnTextTables: boolean
      baseName?: string
      tables?: PdfExtractedTable[]
    }
  | {
      kind: 'pdfToXlsx'
      pageIndexes: number[]
      includeTwoColumnTextTables: boolean
      baseName?: string
      tables?: PdfExtractedTable[]
    }
  | {
      kind: 'extractImages'
      format: PdfExtractImageFormat
      baseName?: string
      images?: PdfExtractedImage[]
    }
  | { kind: 'removeAnnotations' }
  | ({ kind: 'removeBlanks'; blankPageIndexes?: number[] } & PdfBlankPageDetectionOptions)
  | { kind: 'invertColors'; pageIndexes: number[] }
  | {
      kind: 'replaceColors'
      pageIndexes: number[]
      textColor: string
      backgroundColor: string
    }
  | ({
      kind: 'adjustColors'
      pageIndexes: number[]
      pageImages?: Uint8Array[]
    } & PdfColorAdjustments)
  | ({ kind: 'rearrange' } & RearrangeOptions)
  | ({ kind: 'redact' } & PdfRedactionOptions)
  | { kind: 'comments'; comments: PdfCommentInput[] }
  | {
      kind: 'compress'
      renderDpi: number
      imageQuality: number
      lineArt?: boolean
      lineArtThreshold?: number
      lineArtEdgeLevel?: 1 | 2 | 3
      pageImages?: Uint8Array[]
    }
  | { kind: 'flatten'; mode: 'forms' }
  | { kind: 'flatten'; mode: 'pages'; renderDpi: number; pageImages?: Uint8Array[] }
  | {
      kind: 'forms'
      action: PdfFormAction
      fields?: PdfFormFieldValue[]
      fieldNames?: string[]
      modifications?: PdfFormFieldModification[]
      creations?: PdfFormFieldCreation[]
    }
  | { kind: 'repair' }
  | { kind: 'decompress' }
  | { kind: 'removeSignatures' }
  | ({ kind: 'certificateSign' } & PdfCertificateSignOperation)
  | ({ kind: 'timestamp' } & PdfTimestampOperation)
  | ({ kind: 'password' } & PdfPasswordOptions)
  | ({ kind: 'ocr' } & PdfOcrOptions)
  | ({ kind: 'sanitize' } & SanitizeOptions)
  | { kind: 'pipeline'; steps: PdfPipelineStep[] }
  | ({ kind: 'overlay' } & OverlayOptions)
  | ({ kind: 'overlayImage' } & ImageOverlayOptions)
  | ({ kind: 'imagesToPdf' } & PdfImagesToPdfOptions)
  | ({ kind: 'cbzToPdf' } & CbzToPdfOptions)
  | ({ kind: 'emailToPdf' } & EmailToPdfOptions)
  | ({ kind: 'epubToPdf' } & EpubToPdfOptions)
  | ({ kind: 'createPdf' } & CreatePdfOptions)
  | ({ kind: 'htmlToPdf' } & HtmlToPdfOptions)
  | ({ kind: 'markdownToPdf' } & MarkdownToPdfOptions)
  | ({ kind: 'pdfToImages' } & PdfToImagesOptions)
  | ({ kind: 'pdfToVideo' } & PdfToVideoOptions)
  | ({ kind: 'pdfToCbz' } & PdfToCbzOptions)
  | ({ kind: 'pdfToHtml' } & PdfToHtmlOptions)
  | ({ kind: 'pdfToEpub' } & PdfToEpubOptions)
  | ({ kind: 'pdfToPptx' } & PdfToPptxOptions)
  | ({ kind: 'pdfToDocx' } & PdfToDocxOptions)
  | ({ kind: 'pdfToOdt' } & PdfToOdtOptions)
  | ({ kind: 'pdfToRtf' } & PdfToRtfOptions)
  | {
      kind: 'pdfToPdfa'
      format: PdfArchiveFormat
      archiveMode: PdfArchiveMode
      renderDpi: number
      imageQuality: number
      pageImages?: Uint8Array[]
    }
  | ({ kind: 'pdfToJson' } & PdfToJsonOptions)
  | ({ kind: 'jsonToPdf' } & JsonToPdfOptions)
  | ({ kind: 'pdfToXml' } & PdfToXmlOptions)
  | ({ kind: 'metadata' } & PdfMetadataOptions)
  | {
      kind: 'autoRename'
      strategy: PdfAutoRenameStrategy
      suggestedName?: string
    }
  | ({ kind: 'pageNumbers' } & PdfPageNumberOptions)
  | {
      kind: 'filterPages'
      criterion: PdfContentFilterCriterion
      action: PdfContentFilterAction
      pageIndexes: number[]
      text?: string
      caseSensitive: boolean
      wholeWord: boolean
      pageSize?: Exclude<PdfPageSize, 'KEEP'>
      orientation?: PdfOrientation
      rotation?: PdfPageFilterRotation
      comparator?: PdfPageFilterComparator
      matchedPageIndexes?: number[]
    }
  | {
      kind: 'filterDocuments'
      currentFileName: string
      documents: PdfDocumentFilterInput[]
      criterion: PdfDocumentFilterCriterion
      comparator: PdfPageFilterComparator
      text?: string
      caseSensitive: boolean
      wholeWord: boolean
      currentContentMatched?: boolean
      pageCount?: number
      fileSizeBytes?: number
      pageSize?: Exclude<PdfPageSize, 'KEEP'>
      rotation?: PdfPageFilterRotation
    }
  | ({ kind: 'attachments' } & AttachmentOptions)
  | { kind: 'bookmarks'; bookmarks: PdfBookmark[] }

export interface PdfToolOutput {
  suffix: string
  bytes: Uint8Array
  fileName?: string
  mimeType?: string
  extension?: string
}

export function pdfToVideoOutput(options: PdfToVideoOptions): PdfToolOutput {
  if (!options.videoBytes || options.videoBytes.length < 16) {
    throw new Error('Encoded PDF video is required')
  }
  if (
    options.videoBytes[0] !== 0x1a ||
    options.videoBytes[1] !== 0x45 ||
    options.videoBytes[2] !== 0xdf ||
    options.videoBytes[3] !== 0xa3
  ) {
    throw new Error('Encoded PDF video is not a valid WebM file')
  }
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  return {
    suffix: '_slideshow.webm',
    fileName: `${baseName}_slideshow.webm`,
    bytes: options.videoBytes,
    mimeType: 'video/webm',
    extension: '.webm',
  }
}

const RESERVED_FILE_STEMS = /^(?:con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i

export function pdfAutoRenameFileName(value: string): string {
  let stem = value
    .normalize('NFKC')
    .replace(/\.pdf$/i, '')
    .replace(/[\p{Cc}/\\?%*:|"<>]/gu, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/[. ]+$/g, '')
  if (!stem) throw new Error('No usable PDF title was found')
  if (RESERVED_FILE_STEMS.test(stem)) stem = `_${stem}`

  const encoder = new TextEncoder()
  while (stem && encoder.encode(`${stem}.pdf`).length > 240) stem = stem.slice(0, -1)
  stem = stem.replace(/[. ]+$/g, '')
  if (!stem) throw new Error('No usable PDF title was found')
  return `${stem}.pdf`
}

const PDF_PAGE_NUMBER_POSITIONS = new Set<PdfPageNumberPosition>([
  'topLeft',
  'topCenter',
  'topRight',
  'middleLeft',
  'center',
  'middleRight',
  'bottomLeft',
  'bottomCenter',
  'bottomRight',
])

const PDF_PAGE_NUMBER_MARGINS: Record<PdfPageNumberMargin, number> = {
  small: 0.02,
  medium: 0.035,
  large: 0.05,
  xLarge: 0.075,
}

const PDF_PAGE_NUMBER_FONTS: Record<PdfPageNumberFont, StandardFonts> = {
  helvetica: StandardFonts.Helvetica,
  times: StandardFonts.TimesRoman,
  courier: StandardFonts.Courier,
}

function checkedPdfPageNumberOptions(
  pageCount: number,
  options: PdfPageNumberOptions,
): PdfPageNumberOptions & { pageIndexes: number[] } {
  const pageIndexes = checkedPageIndexes(pageCount, options.pageIndexes).sort(
    (left, right) => left - right,
  )
  if (!PDF_PAGE_NUMBER_POSITIONS.has(options.position)) {
    throw new Error('Page number position is invalid')
  }
  if (!(options.margin in PDF_PAGE_NUMBER_MARGINS)) {
    throw new Error('Page number margin is invalid')
  }
  if (!(options.font in PDF_PAGE_NUMBER_FONTS)) throw new Error('Page number font is invalid')
  if (!Number.isFinite(options.fontSize) || options.fontSize < 1 || options.fontSize > 200) {
    throw new Error('Page number font size must be between 1 and 200')
  }
  if (!Number.isSafeInteger(options.startingNumber) || options.startingNumber < 1) {
    throw new Error('Page number starting value must be a positive integer')
  }
  if (!Number.isInteger(options.zeroPad) || options.zeroPad < 0 || options.zeroPad > 12) {
    throw new Error('Page number padding must be between 0 and 12')
  }
  pdfRgbColor(options.fontColor, 'Page number color')
  return { ...options, pageIndexes }
}

export function pdfPageNumberLabels(
  pageCount: number,
  options: PdfPageNumberOptions,
): PdfPageNumberLabel[] {
  const checked = checkedPdfPageNumberOptions(pageCount, options)
  const pattern = checked.textPattern || '{n}'
  const baseName = (checked.baseName || 'document').replace(/\.pdf$/i, '')
  return checked.pageIndexes.map((pageIndex, sequence) => {
    const number = String(checked.startingNumber + sequence).padStart(checked.zeroPad, '0')
    return {
      pageIndex,
      text: pattern
        .replaceAll('{n}', number)
        .replaceAll('{total}', String(pageCount))
        .replaceAll('{filename}', baseName),
    }
  })
}

export async function addPdfPageNumbersBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfPageNumberOptions,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const checked = checkedPdfPageNumberOptions(document.getPageCount(), options)
  const labels = pdfPageNumberLabels(document.getPageCount(), checked)
  const font = await document.embedFont(PDF_PAGE_NUMBER_FONTS[checked.font])
  const [red, green, blue] = pdfRgbColor(checked.fontColor, 'Page number color')
  const color = rgb(red / 255, green / 255, blue / 255)
  const marginFactor = PDF_PAGE_NUMBER_MARGINS[checked.margin]
  for (const label of labels) {
    const page = document.getPage(label.pageIndex)
    const box = page.getMediaBox()
    const textWidth = font.widthOfTextAtSize(label.text, checked.fontSize)
    const textHeight = font.heightAtSize(checked.fontSize, { descender: true })
    const column = checked.position.endsWith('Left')
      ? 'left'
      : checked.position.endsWith('Right')
        ? 'right'
        : 'center'
    const row = checked.position.startsWith('top')
      ? 'top'
      : checked.position.startsWith('bottom')
        ? 'bottom'
        : 'middle'
    const left = box.x + box.width * marginFactor
    const centerX = box.x + box.width / 2
    const right = box.x + box.width * (1 - marginFactor)
    const bottom = box.y + box.height * marginFactor
    const middleY = box.y + box.height / 2
    const top = box.y + box.height * (1 - marginFactor)
    const x =
      column === 'left' ? left : column === 'right' ? right - textWidth : centerX - textWidth / 2
    const y =
      row === 'top' ? top - textHeight : row === 'bottom' ? bottom : middleY - textHeight / 2
    page.drawText(label.text, {
      x,
      y,
      size: checked.fontSize,
      font,
      color,
    })
  }
  return document.save({ useObjectStreams: false, updateFieldAppearances: false })
}

const PAGE_SIZES: Record<Exclude<PdfPageSize, 'KEEP'>, readonly [number, number]> = {
  A3: [841.89, 1190.55],
  A4: [595.28, 841.89],
  A5: [419.53, 595.28],
  LETTER: [612, 792],
  LEGAL: [612, 1008],
  TABLOID: [792, 1224],
}

function finiteNonNegative(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0) throw new Error(`${name} must be non-negative`)
  return value
}

function finiteNumber(value: number, name: string): number {
  if (!Number.isFinite(value)) throw new Error(`${name} must be finite`)
  return value
}

function finitePositive(value: number, name: string): number {
  if (!Number.isFinite(value) || value <= 0) throw new Error(`${name} must be positive`)
  return value
}

function positiveInteger(value: number, name: string): number {
  if (!Number.isInteger(value) || value <= 0) throw new Error(`${name} must be a positive integer`)
  return value
}

function unitInterval(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0 || value > 1) {
    throw new Error(`${name} must be between 0 and 1`)
  }
  return value
}

function orientedSize(
  size: readonly [number, number],
  orientation: PdfOrientation,
): readonly [number, number] {
  const [width, height] = size
  return orientation === 'landscape'
    ? [Math.max(width, height), Math.min(width, height)]
    : [Math.min(width, height), Math.max(width, height)]
}

function copyMetadata(source: PDFDocument, output: PDFDocument): void {
  const title = source.getTitle()
  const author = source.getAuthor()
  const subject = source.getSubject()
  const keywords = source.getKeywords()
  const creator = source.getCreator()
  const producer = source.getProducer()
  const created = source.getCreationDate()
  if (title) output.setTitle(title)
  if (author) output.setAuthor(author)
  if (subject) output.setSubject(subject)
  if (keywords) output.setKeywords(keywords.split(/[,;]/).map((value) => value.trim()))
  if (creator) output.setCreator(creator)
  if (producer) output.setProducer(producer)
  if (created) output.setCreationDate(created)
  output.setModificationDate(new Date())
}

function ensureEmbeddablePages(document: PDFDocument): void {
  for (const page of document.getPages()) {
    page.drawRectangle({ x: 0, y: 0, width: 0, height: 0, opacity: 0 })
  }
}

async function embedCroppedPages(source: PDFDocument, output: PDFDocument) {
  const pages = source.getPages()
  const boundingBoxes = pages.map((page) => {
    const box = page.getCropBox()
    return {
      left: box.x,
      bottom: box.y,
      right: box.x + box.width,
      top: box.y + box.height,
    }
  })
  ensureEmbeddablePages(source)
  return output.embedPages(pages, boundingBoxes)
}

function drawFittedPage(
  page: ReturnType<PDFDocument['addPage']>,
  embeddedPage: Awaited<ReturnType<PDFDocument['embedPage']>>,
  box: { x: number; y: number; width: number; height: number },
  borderWidth = 0,
): void {
  const scale = Math.min(box.width / embeddedPage.width, box.height / embeddedPage.height)
  const width = embeddedPage.width * scale
  const height = embeddedPage.height * scale
  const x = box.x + (box.width - width) / 2
  const y = box.y + (box.height - height) / 2
  page.drawPage(embeddedPage, { x, y, width, height })
  if (borderWidth > 0) {
    page.drawRectangle({
      x,
      y,
      width,
      height,
      borderColor: rgb(0, 0, 0),
      borderWidth,
    })
  }
}

function validPageIndexes(document: PDFDocument, pageIndexes: number[]): number[] {
  const pageCount = document.getPageCount()
  return pageIndexes.filter(
    (pageIndex) => Number.isInteger(pageIndex) && pageIndex >= 0 && pageIndex < pageCount,
  )
}

function checkedPageIndexes(
  pageCount: number,
  pageIndexes: number[],
  options: { allowDuplicates?: boolean; allowEmpty?: boolean } = {},
): number[] {
  if (!Array.isArray(pageIndexes)) throw new Error('pageIndexes must be an array')
  if (
    pageIndexes.some(
      (pageIndex) => !Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pageCount,
    )
  ) {
    throw new Error('pageIndexes contain an invalid page')
  }
  const result = options.allowDuplicates ? [...pageIndexes] : [...new Set(pageIndexes)]
  if (!options.allowEmpty && result.length === 0) throw new Error('At least one page is required')
  return result
}

export async function extractPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const pages = await output.copyPages(source, validPageIndexes(source, pageIndexes))
  for (const page of pages) output.addPage(page)
  return output.save({ useObjectStreams: false })
}

export async function insertPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  otherBytes: Uint8Array | ArrayBuffer,
  afterPageIndex: number,
): Promise<InsertPdfResult> {
  const destination = await PDFDocument.load(bytes, { updateMetadata: false })
  const source = await PDFDocument.load(otherBytes, { updateMetadata: false })
  const pages = await destination.copyPages(source, source.getPageIndices())
  let insertionIndex = Math.min(
    Math.max(Math.trunc(afterPageIndex) + 1, 0),
    destination.getPageCount(),
  )
  for (const page of pages) destination.insertPage(insertionIndex++, page)
  return {
    merged: await destination.save({ useObjectStreams: false }),
    count: pages.length,
  }
}

export async function insertBlankPageBytes(
  bytes: Uint8Array | ArrayBuffer,
  afterPageIndex: number,
  options: InsertBlankPageOptions = {},
): Promise<InsertBlankPageResult> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = document.getPageCount()
  if (pageCount === 0) throw new Error('The PDF must contain at least one page')
  const count = options.count ?? 1
  if (!Number.isInteger(count) || count < 1 || count > 100) {
    throw new Error('Blank page count must be a whole number from 1 to 100')
  }
  const pageSize = options.pageSize ?? 'KEEP'
  if (pageSize !== 'KEEP' && !(pageSize in PAGE_SIZES)) {
    throw new Error('Blank page size is invalid')
  }
  const orientation = options.orientation ?? 'portrait'
  if (orientation !== 'portrait' && orientation !== 'landscape') {
    throw new Error('Blank page orientation is invalid')
  }
  const normalizedIndex = Number.isFinite(afterPageIndex) ? Math.trunc(afterPageIndex) : -1
  const insertionIndex = Math.min(Math.max(normalizedIndex + 1, 0), pageCount)
  const referenceIndex = Math.min(Math.max(normalizedIndex, 0), pageCount - 1)
  const reference = document.getPage(referenceIndex)
  const size =
    pageSize === 'KEEP'
      ? ([reference.getWidth(), reference.getHeight()] as const)
      : orientation === 'portrait'
        ? PAGE_SIZES[pageSize]
        : ([PAGE_SIZES[pageSize][1], PAGE_SIZES[pageSize][0]] as const)
  for (let offset = 0; offset < count; offset++) {
    const blank = document.insertPage(insertionIndex + offset, [size[0], size[1]])
    if (pageSize === 'KEEP') blank.setRotation(degrees(reference.getRotation().angle))
  }
  return {
    merged: await document.save({ useObjectStreams: false }),
    count,
  }
}

export async function mergePdfBytes(
  documents: ReadonlyArray<Uint8Array | ArrayBuffer>,
): Promise<Uint8Array> {
  if (documents.length === 0) throw new Error('At least one PDF is required')
  let merged: Uint8Array = new Uint8Array(documents[0]!)
  for (const document of documents.slice(1)) {
    merged = (await insertPdfBytes(merged, document, Number.MAX_SAFE_INTEGER)).merged
  }
  return merged
}

export async function splitPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  splitAfterPages: number[],
): Promise<Uint8Array[]> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = source.getPageCount()
  if (pageCount === 0) return []
  const splitPoints = [
    ...new Set(
      splitAfterPages.filter(
        (pageIndex) => Number.isInteger(pageIndex) && pageIndex >= 0 && pageIndex < pageCount,
      ),
    ),
  ].sort((left, right) => left - right)
  return splitPdfAtBoundariesBytes(bytes, splitPoints)
}

export function pdfAutoSplitDividerQrModules(): boolean[][] {
  const code = qrcode(0, 'H')
  code.addData(PDF_AUTO_SPLIT_QR_CONTENT)
  code.make()
  return Array.from({ length: code.getModuleCount() }, (_, row) =>
    Array.from({ length: code.getModuleCount() }, (_, column) => code.isDark(row, column)),
  )
}

export async function createPdfAutoSplitDividerBytes(): Promise<Uint8Array> {
  const document = await PDFDocument.create({ updateMetadata: false })
  const page = document.addPage([595.28, 841.89])
  const regular = await document.embedFont(StandardFonts.Helvetica)
  const bold = await document.embedFont(StandardFonts.HelveticaBold)
  const modules = pdfAutoSplitDividerQrModules()
  const quietZone = 4
  const qrSize = 300
  const cellSize = qrSize / (modules.length + quietZone * 2)
  const x = (page.getWidth() - qrSize) / 2
  const y = (page.getHeight() - qrSize) / 2 + 20
  const titleWidth = bold.widthOfTextAtSize('AUTO SPLIT DIVIDER', 28)
  page.drawText('AUTO SPLIT DIVIDER', {
    x: (page.getWidth() - titleWidth) / 2,
    y: y + qrSize + 78,
    size: 28,
    font: bold,
    color: rgb(0.08, 0.08, 0.08),
  })
  page.drawRectangle({ x, y, width: qrSize, height: qrSize, color: rgb(1, 1, 1) })
  for (let row = 0; row < modules.length; row++) {
    for (let column = 0; column < modules.length; column++) {
      if (!modules[row]![column]) continue
      page.drawRectangle({
        x: x + (column + quietZone) * cellSize,
        y: y + (modules.length - row - 1 + quietZone) * cellSize,
        width: cellSize + 0.02,
        height: cellSize + 0.02,
        color: rgb(0, 0, 0),
      })
    }
  }
  const instruction = 'Place this page between documents before scanning.'
  const detail = 'The divider page is removed automatically during local PDF splitting.'
  page.drawText(instruction, {
    x: (page.getWidth() - bold.widthOfTextAtSize(instruction, 15)) / 2,
    y: y - 52,
    size: 15,
    font: bold,
    color: rgb(0.12, 0.12, 0.12),
  })
  page.drawText(detail, {
    x: (page.getWidth() - regular.widthOfTextAtSize(detail, 11)) / 2,
    y: y - 78,
    size: 11,
    font: regular,
    color: rgb(0.35, 0.35, 0.35),
  })
  return document.save({ useObjectStreams: false })
}

export function autoSplitPdfPageRanges(
  pageCount: number,
  dividerPageIndexes: number[],
  duplexMode: boolean,
): PdfPageRange[] {
  if (!Number.isInteger(pageCount) || pageCount < 0) {
    throw new Error('pageCount must be a non-negative integer')
  }
  if (!Array.isArray(dividerPageIndexes)) {
    throw new Error('dividerPageIndexes must be an array')
  }
  const dividers = new Set<number>()
  for (const pageIndex of dividerPageIndexes) {
    if (!Number.isInteger(pageIndex) || pageIndex < 0 || pageIndex >= pageCount) {
      throw new Error('dividerPageIndexes contain an invalid page')
    }
    dividers.add(pageIndex)
  }
  const excluded = new Set(dividers)
  if (duplexMode) {
    for (const pageIndex of dividers) {
      if (pageIndex + 1 < pageCount) excluded.add(pageIndex + 1)
    }
  }
  const ranges: PdfPageRange[] = []
  let firstPage: number | null = null
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex++) {
    if (excluded.has(pageIndex)) {
      if (firstPage !== null) {
        ranges.push({ firstPage, lastPage: pageIndex - 1 })
        firstPage = null
      }
      continue
    }
    firstPage ??= pageIndex
  }
  if (firstPage !== null) ranges.push({ firstPage, lastPage: pageCount - 1 })
  return ranges
}

export async function autoSplitPdfZipBytes(
  bytes: Uint8Array | ArrayBuffer,
  dividerPageIndexes: number[],
  duplexMode: boolean,
  baseName: string,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const ranges = autoSplitPdfPageRanges(document.getPageCount(), dividerPageIndexes, duplexMode)
  if (dividerPageIndexes.length === 0) {
    throw new Error('No QR divider pages were detected')
  }
  if (ranges.length === 0) throw new Error('No document pages remain after removing dividers')
  const archive = new JSZip()
  const stem = safeArchiveStem(baseName)
  for (let index = 0; index < ranges.length; index++) {
    const range = ranges[index]!
    archive.file(
      `${stem}_${index + 1}.pdf`,
      await extractPdfPageRangeBytes(bytes, range.firstPage, range.lastPage),
    )
  }
  return archive.generateAsync({ type: 'uint8array', compression: 'STORE' })
}

function splitRangeBoundariesByPageCount(pageCount: number, pagesPerDocument: number): number[] {
  if (!Number.isInteger(pagesPerDocument) || pagesPerDocument <= 0) {
    throw new Error('pagesPerDocument must be a positive integer')
  }
  const boundaries: number[] = []
  for (let lastPage = pagesPerDocument - 1; lastPage < pageCount; lastPage += pagesPerDocument) {
    boundaries.push(lastPage)
  }
  return boundaries
}

function splitRangeBoundariesByDocumentCount(pageCount: number, documentCount: number): number[] {
  if (!Number.isInteger(documentCount) || documentCount <= 0) {
    throw new Error('documentCount must be a positive integer')
  }
  const outputCount = Math.min(pageCount, documentCount)
  const pagesPerDocument = Math.floor(pageCount / outputCount)
  const extraPages = pageCount % outputCount
  const boundaries: number[] = []
  let lastPage = -1
  for (let index = 0; index < outputCount; index++) {
    lastPage += pagesPerDocument + (index < extraPages ? 1 : 0)
    boundaries.push(lastPage)
  }
  return boundaries
}

function collectLiveWidgetObjects(document: PDFDocument): Set<PDFObject> {
  const liveWidgets = new Set<PDFObject>()
  for (const page of document.getPages()) {
    const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annotations) continue
    for (let index = 0; index < annotations.size(); index++) {
      const entry = annotations.get(index)
      liveWidgets.add(entry)
      const resolved = document.context.lookup(entry)
      if (resolved) liveWidgets.add(resolved)
    }
  }
  return liveWidgets
}

function isWidgetDictionary(dictionary: PDFDict): boolean {
  return dictionary.lookupMaybe(PDFName.of('Subtype'), PDFName)?.toString() === '/Widget'
}

function isLiveWidgetEntry(
  document: PDFDocument,
  entry: PDFObject,
  dictionary: PDFDict,
  liveWidgets: Set<PDFObject>,
): boolean {
  const resolved = entry instanceof PDFRef ? document.context.lookup(entry) : undefined
  return (
    liveWidgets.has(entry) ||
    liveWidgets.has(dictionary) ||
    (resolved !== undefined && liveWidgets.has(resolved))
  )
}

function prunePdfFieldEntry(
  document: PDFDocument,
  entry: PDFObject,
  liveWidgets: Set<PDFObject>,
): boolean {
  const dictionary = document.context.lookupMaybe(entry, PDFDict)
  if (!dictionary) return true
  const kids = dictionary.lookupMaybe(PDFName.of('Kids'), PDFArray)
  if (kids) {
    for (let index = kids.size() - 1; index >= 0; index--) {
      const childEntry = kids.get(index)
      const childDictionary = document.context.lookupMaybe(childEntry, PDFDict)
      const keepChild = childDictionary
        ? isWidgetDictionary(childDictionary)
          ? isLiveWidgetEntry(document, childEntry, childDictionary, liveWidgets)
          : prunePdfFieldEntry(document, childEntry, liveWidgets)
        : true
      if (!keepChild) kids.remove(index)
    }
    if (kids.size() > 0) return true
  }
  return isWidgetDictionary(dictionary)
    ? isLiveWidgetEntry(document, entry, dictionary, liveWidgets)
    : false
}

function pruneOrphanedPdfFields(document: PDFDocument): void {
  const acroForm = document.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const fields = acroForm?.lookupMaybe(PDFName.of('Fields'), PDFArray)
  if (!acroForm || !fields) return
  const liveWidgets = collectLiveWidgetObjects(document)
  for (let index = fields.size() - 1; index >= 0; index--) {
    if (!prunePdfFieldEntry(document, fields.get(index), liveWidgets)) fields.remove(index)
  }
  if (fields.size() === 0) document.catalog.delete(PDFName.of('AcroForm'))
}

async function extractPdfPageRangeBytes(
  bytes: Uint8Array | ArrayBuffer,
  firstPage: number,
  lastPage: number,
): Promise<Uint8Array> {
  let document = await PDFDocument.load(bytes, { updateMetadata: false })
  for (let pageIndex = document.getPageCount() - 1; pageIndex >= 0; pageIndex--) {
    if (pageIndex < firstPage || pageIndex > lastPage) document.removePage(pageIndex)
  }
  document = await PDFDocument.load(await document.save({ useObjectStreams: false }), {
    updateMetadata: false,
  })
  pruneOrphanedPdfFields(document)
  return document.save({ useObjectStreams: false })
}

async function compactPdfPageRangeBytes(
  source: PDFDocument,
  firstPage: number,
  lastPage: number,
): Promise<Uint8Array> {
  const output = await PDFDocument.create({ updateMetadata: false })
  await source.flush()
  const copier = PDFObjectCopier.for(source.context, output.context)
  for (let pageIndex = firstPage; pageIndex <= lastPage; pageIndex++) {
    const pageNode = copier.copy(source.getPage(pageIndex).node)
    const pageReference = output.context.register(pageNode)
    output.addPage(PDFPage.of(pageNode, pageReference, output))
  }
  const acroForm = source.catalog.get(PDFName.of('AcroForm'))
  if (acroForm) output.catalog.set(PDFName.of('AcroForm'), copier.copy(acroForm))
  copyMetadata(source, output)
  pruneOrphanedPdfFields(output)
  return output.save({ useObjectStreams: false, updateFieldAppearances: false })
}

async function splitPdfAtBoundariesBytes(
  bytes: Uint8Array | ArrayBuffer,
  boundaries: number[],
): Promise<Uint8Array[]> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = source.getPageCount()
  if (pageCount === 0) return []
  const outputs: Uint8Array[] = []
  let firstPage = 0
  for (const lastPage of boundaries) {
    if (lastPage < firstPage || lastPage >= pageCount) continue
    outputs.push(await extractPdfPageRangeBytes(bytes, firstPage, lastPage))
    firstPage = lastPage + 1
  }
  if (firstPage < pageCount) {
    outputs.push(await extractPdfPageRangeBytes(bytes, firstPage, pageCount - 1))
  }
  return outputs
}

export async function splitPdfByPageCountBytes(
  bytes: Uint8Array | ArrayBuffer,
  pagesPerDocument: number,
): Promise<Uint8Array[]> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  return splitPdfAtBoundariesBytes(
    bytes,
    splitRangeBoundariesByPageCount(document.getPageCount(), pagesPerDocument),
  )
}

export async function splitPdfByDocumentCountBytes(
  bytes: Uint8Array | ArrayBuffer,
  documentCount: number,
): Promise<Uint8Array[]> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  return splitPdfAtBoundariesBytes(
    bytes,
    splitRangeBoundariesByDocumentCount(document.getPageCount(), documentCount),
  )
}

export async function splitPdfBySizeBytes(
  bytes: Uint8Array | ArrayBuffer,
  maxBytes: number,
): Promise<Uint8Array[]> {
  if (!Number.isSafeInteger(maxBytes) || maxBytes <= 0) {
    throw new Error('maxBytes must be a positive safe integer')
  }
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = source.getPageCount()
  const outputs: Uint8Array[] = []
  let firstPage = 0
  while (firstPage < pageCount) {
    let bestLastPage = firstPage
    let bestBytes = await compactPdfPageRangeBytes(source, firstPage, firstPage)
    if (bestBytes.length <= maxBytes) {
      let low = firstPage + 1
      let high = pageCount - 1
      while (low <= high) {
        const trialLastPage = Math.floor((low + high) / 2)
        const trialBytes = await compactPdfPageRangeBytes(source, firstPage, trialLastPage)
        if (trialBytes.length <= maxBytes) {
          bestLastPage = trialLastPage
          bestBytes = trialBytes
          low = trialLastPage + 1
        } else {
          high = trialLastPage - 1
        }
      }
    }
    outputs.push(bestBytes)
    firstPage = bestLastPage + 1
  }
  return outputs
}

interface ChapterSplitPoint {
  title: string
  startPage: number
  sourceIndex: number
}

function collectChapterSplitPoints(
  bookmarks: PdfBookmark[],
  pageCount: number,
  maxLevel: number,
  depth = 0,
  output: ChapterSplitPoint[] = [],
): ChapterSplitPoint[] {
  if (depth > maxLevel) return output
  for (const bookmark of bookmarks) {
    output.push({
      title: bookmark.title.trim(),
      startPage: Math.min(Math.max(Math.trunc(bookmark.pageNumber) - 1, 0), pageCount - 1),
      sourceIndex: output.length,
    })
    collectChapterSplitPoints(bookmark.children, pageCount, maxLevel, depth + 1, output)
  }
  return output
}

function mergedChapterSplitPoints(points: ChapterSplitPoint[]): ChapterSplitPoint[] {
  const merged: ChapterSplitPoint[] = []
  for (const point of points) {
    const previous = merged.at(-1)
    if (previous?.startPage === point.startPage) {
      previous.title = [previous.title, point.title].filter(Boolean).join(' ').slice(0, 255)
    } else {
      merged.push({ ...point })
    }
  }
  return merged
}

export async function splitPdfByChaptersBytes(
  bytes: Uint8Array | ArrayBuffer,
  bookmarkLevel: number,
  allowDuplicates = false,
): Promise<PdfChapterSplitOutput[]> {
  if (!Number.isInteger(bookmarkLevel) || bookmarkLevel < 0 || bookmarkLevel > 20) {
    throw new Error('bookmarkLevel must be an integer from 0 to 20')
  }
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = document.getPageCount()
  if (pageCount === 0) return []
  const bookmarks = await listInternalPdfBookmarksBytes(bytes)
  if (bookmarks.length === 0) throw new Error('No PDF bookmarks were found')
  const sortedPoints = collectChapterSplitPoints(bookmarks, pageCount, bookmarkLevel).sort(
    (left, right) => left.startPage - right.startPage || left.sourceIndex - right.sourceIndex,
  )
  const points = allowDuplicates ? sortedPoints : mergedChapterSplitPoints(sortedPoints)
  const outputs: PdfChapterSplitOutput[] = []
  for (let index = 0; index < points.length; index++) {
    const point = points[index]!
    const next = points[index + 1]
    const lastPage = next
      ? next.startPage === point.startPage
        ? point.startPage
        : next.startPage - 1
      : pageCount - 1
    outputs.push({
      title: point.title || `Chapter ${index + 1}`,
      bytes: await extractPdfPageRangeBytes(bytes, point.startPage, lastPage),
    })
  }
  return outputs
}

export async function cropPdfMarginsBytes(
  bytes: Uint8Array | ArrayBuffer,
  margins: CropMargins,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const top = finiteNonNegative(margins.top, 'top')
  const right = finiteNonNegative(margins.right, 'right')
  const bottom = finiteNonNegative(margins.bottom, 'bottom')
  const left = finiteNonNegative(margins.left, 'left')
  for (const page of document.getPages()) {
    const box = page.getCropBox()
    const width = box.width - left - right
    const height = box.height - top - bottom
    if (width <= 0 || height <= 0) throw new Error('Crop margins exceed the page size')
    const x = box.x + left
    const y = box.y + bottom
    page.setMediaBox(x, y, width, height)
    page.setCropBox(x, y, width, height)
  }
  return document.save({ useObjectStreams: false })
}

export async function cropPdfPageBoxesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageBoxes: PdfPageCropBox[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const pages = document.getPages()
  if (pageBoxes.length !== pages.length) {
    throw new Error('Auto crop boxes must match the PDF page count')
  }
  pages.forEach((page, pageIndex) => {
    const source = page.getCropBox()
    const requested = pageBoxes[pageIndex]
    if (!requested) throw new Error(`Auto crop box is missing for page ${pageIndex + 1}`)
    const x = finiteNumber(requested.x, 'x')
    const y = finiteNumber(requested.y, 'y')
    const width = finitePositive(requested.width, 'width')
    const height = finitePositive(requested.height, 'height')
    const epsilon = 0.05
    if (
      x < source.x - epsilon ||
      y < source.y - epsilon ||
      x + width > source.x + source.width + epsilon ||
      y + height > source.y + source.height + epsilon
    ) {
      throw new Error(`Auto crop box is outside page ${pageIndex + 1}`)
    }
    const normalizedX = Math.max(source.x, x)
    const normalizedY = Math.max(source.y, y)
    const normalizedWidth = Math.min(source.x + source.width, x + width) - normalizedX
    const normalizedHeight = Math.min(source.y + source.height, y + height) - normalizedY
    if (normalizedWidth <= 0 || normalizedHeight <= 0) {
      throw new Error(`Auto crop box is empty for page ${pageIndex + 1}`)
    }
    page.setMediaBox(normalizedX, normalizedY, normalizedWidth, normalizedHeight)
    page.setCropBox(normalizedX, normalizedY, normalizedWidth, normalizedHeight)
  })
  return document.save({ useObjectStreams: false })
}

export async function scalePdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: Extract<PdfToolOperation, { kind: 'scale' }>,
): Promise<Uint8Array> {
  if (!Number.isFinite(options.scaleFactor) || options.scaleFactor <= 0) {
    throw new Error('scaleFactor must be positive')
  }
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  if (source.getPageCount() === 0) return output.save({ useObjectStreams: false })
  const embeddedPages = await embedCroppedPages(source, output)
  for (const embeddedPage of embeddedPages) {
    const targetSize =
      options.pageSize === 'KEEP'
        ? ([embeddedPage.width, embeddedPage.height] as const)
        : orientedSize(PAGE_SIZES[options.pageSize], options.orientation)
    const page = output.addPage([targetSize[0], targetSize[1]])
    const scale =
      Math.min(targetSize[0] / embeddedPage.width, targetSize[1] / embeddedPage.height) *
      options.scaleFactor
    const width = embeddedPage.width * scale
    const height = embeddedPage.height * scale
    page.drawPage(embeddedPage, {
      x: (targetSize[0] - width) / 2,
      y: (targetSize[1] - height) / 2,
      width,
      height,
    })
  }
  return output.save({ useObjectStreams: false })
}

export async function nUpPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: NUpOptions,
): Promise<Uint8Array> {
  const rows = positiveInteger(options.rows, 'rows')
  const columns = positiveInteger(options.columns, 'columns')
  if (rows > 30 || columns > 30) throw new Error('rows and columns must not exceed 30')
  const margins = {
    top: finiteNonNegative(options.topMargin ?? 0, 'topMargin'),
    right: finiteNonNegative(options.rightMargin ?? 0, 'rightMargin'),
    bottom: finiteNonNegative(options.bottomMargin ?? 0, 'bottomMargin'),
    left: finiteNonNegative(options.leftMargin ?? 0, 'leftMargin'),
    inner: finiteNonNegative(options.innerMargin ?? 0, 'innerMargin'),
  }
  const borderWidth = finiteNonNegative(options.borderWidth ?? 0, 'borderWidth')
  if (borderWidth > 72) throw new Error('borderWidth must not exceed 72')
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const targetSize = orientedSize(PAGE_SIZES.A4, options.orientation)
  const cellWidth = (targetSize[0] - margins.left - margins.right) / columns
  const cellHeight = (targetSize[1] - margins.top - margins.bottom) / rows
  const innerWidth = cellWidth - margins.inner * 2
  const innerHeight = cellHeight - margins.inner * 2
  if (innerWidth <= 0 || innerHeight <= 0) throw new Error('Margins leave no room for pages')

  const embeddedPages = await embedCroppedPages(source, output)
  const pagesPerSheet = rows * columns
  for (let offset = 0; offset < embeddedPages.length; offset += pagesPerSheet) {
    const sheet = output.addPage([targetSize[0], targetSize[1]])
    const group = embeddedPages.slice(offset, offset + pagesPerSheet)
    for (let index = 0; index < group.length; index++) {
      const embeddedPage = group[index]!
      const row = options.arrangement === 'columns' ? index % rows : Math.floor(index / columns)
      const logicalColumn =
        options.arrangement === 'columns' ? Math.floor(index / rows) : index % columns
      const column =
        options.readingDirection === 'rtl' ? columns - 1 - logicalColumn : logicalColumn
      drawFittedPage(
        sheet,
        embeddedPage,
        {
          x: margins.left + column * cellWidth + margins.inner,
          y: targetSize[1] - margins.top - (row + 1) * cellHeight + margins.inner,
          width: innerWidth,
          height: innerHeight,
        },
        borderWidth,
      )
    }
  }
  return output.save({ useObjectStreams: false })
}

export function bookletPagePairs(
  pageCount: number,
  duplexPass: PdfDuplexPass,
  flipOnShortEdge: boolean,
): Array<readonly [number, number]> {
  const paddedPageCount = Math.ceil(Math.max(0, pageCount) / 4) * 4
  const pairs: Array<readonly [number, number]> = []
  for (let sheet = 0; sheet < paddedPageCount / 4; sheet++) {
    const normalize = (pageIndex: number) => (pageIndex < pageCount ? pageIndex : -1)
    const front = [normalize(paddedPageCount - 1 - sheet * 2), normalize(sheet * 2)] as const
    const back = [normalize(sheet * 2 + 1), normalize(paddedPageCount - 2 - sheet * 2)] as const
    if (duplexPass === 'both' || duplexPass === 'front') pairs.push(front)
    if (duplexPass === 'both' || duplexPass === 'back') {
      pairs.push(flipOnShortEdge ? [back[1], back[0]] : back)
    }
  }
  return pairs
}

export async function bookletPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: BookletOptions,
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const sourcePages = source.getPages()
  const firstPage = sourcePages[0]
  if (!firstPage) throw new Error('Booklet requires at least one page')
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const embeddedPages = await embedCroppedPages(source, output)
  const firstBox = firstPage.getCropBox()
  const targetSize = orientedSize([firstBox.width, firstBox.height], 'landscape')
  const gutter = Math.min(
    finiteNonNegative(options.gutter, 'gutter'),
    Math.max(0, targetSize[0] - 2),
  )
  const cellWidth = (targetSize[0] - gutter) / 2
  const pairs = bookletPagePairs(embeddedPages.length, options.duplexPass, options.flipOnShortEdge)

  for (const pair of pairs) {
    const sheet = output.addPage([targetSize[0], targetSize[1]])
    for (let visualColumn = 0; visualColumn < 2; visualColumn++) {
      const pageIndex = pair[visualColumn]!
      const column = options.spine === 'right' ? 1 - visualColumn : visualColumn
      const box = {
        x: column === 0 ? 0 : cellWidth + gutter,
        y: 0,
        width: cellWidth,
        height: targetSize[1],
      }
      if (pageIndex >= 0) {
        drawFittedPage(sheet, embeddedPages[pageIndex]!, box, options.border ? 1 : 0)
      } else if (options.border) {
        sheet.drawRectangle({ ...box, borderColor: rgb(0, 0, 0), borderWidth: 1 })
      }
    }
  }
  return output.save({ useObjectStreams: false })
}

function sectionCoordinates(
  rows: number,
  columns: number,
  arrangement: PdfSectionArrangement,
): Array<readonly [number, number]> {
  const coordinates: Array<readonly [number, number]> = []
  if (arrangement === 'columns') {
    for (let column = 0; column < columns; column++) {
      for (let row = 0; row < rows; row++) coordinates.push([row, column])
    }
  } else {
    for (let row = 0; row < rows; row++) {
      for (let column = 0; column < columns; column++) coordinates.push([row, column])
    }
  }
  return coordinates
}

function drawPdfSection(
  output: PDFDocument,
  embeddedPage: Awaited<ReturnType<PDFDocument['embedPage']>>,
  rows: number,
  columns: number,
  row: number,
  column: number,
): void {
  const sectionWidth = embeddedPage.width / columns
  const sectionHeight = embeddedPage.height / rows
  const page = output.addPage([sectionWidth, sectionHeight])
  page.drawPage(embeddedPage, {
    x: -column * sectionWidth,
    y: -(rows - 1 - row) * sectionHeight,
    width: embeddedPage.width,
    height: embeddedPage.height,
  })
}

function drawFullEmbeddedPage(
  output: PDFDocument,
  embeddedPage: Awaited<ReturnType<PDFDocument['embedPage']>>,
): void {
  const page = output.addPage([embeddedPage.width, embeddedPage.height])
  page.drawPage(embeddedPage, {
    x: 0,
    y: 0,
    width: embeddedPage.width,
    height: embeddedPage.height,
  })
}

function pageCropBounds(page: ReturnType<PDFDocument['getPage']>) {
  const box = page.getCropBox()
  return {
    left: box.x,
    bottom: box.y,
    right: box.x + box.width,
    top: box.y + box.height,
  }
}

export async function splitPdfSectionsBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: SplitSectionsOptions,
): Promise<PdfSectionSplitOutput[]> {
  const rows = positiveInteger(options.rows, 'rows')
  const columns = positiveInteger(options.columns, 'columns')
  if (rows > 10 || columns > 10) throw new Error('rows and columns must not exceed 10')
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const sourcePages = source.getPages()
  if (sourcePages.length === 0) return []
  const selectedPages = new Set(checkedPageIndexes(sourcePages.length, options.pageIndexes))
  const coordinates = sectionCoordinates(rows, columns, options.arrangement ?? 'rows')
  ensureEmbeddablePages(source)

  if (options.merge) {
    const output = await PDFDocument.create()
    copyMetadata(source, output)
    const embeddedPages = await output.embedPages(sourcePages, sourcePages.map(pageCropBounds))
    for (let pageIndex = 0; pageIndex < embeddedPages.length; pageIndex++) {
      const embeddedPage = embeddedPages[pageIndex]!
      if (!selectedPages.has(pageIndex)) {
        drawFullEmbeddedPage(output, embeddedPage)
        continue
      }
      for (const [row, column] of coordinates) {
        drawPdfSection(output, embeddedPage, rows, columns, row, column)
      }
    }
    return [
      {
        sourcePageNumber: 0,
        sectionNumber: 0,
        bytes: await output.save({ useObjectStreams: false }),
      },
    ]
  }

  const outputs: PdfSectionSplitOutput[] = []
  for (let pageIndex = 0; pageIndex < sourcePages.length; pageIndex++) {
    const sourcePage = sourcePages[pageIndex]!
    const positions = selectedPages.has(pageIndex) ? coordinates : ([[0, 0]] as const)
    for (let sectionIndex = 0; sectionIndex < positions.length; sectionIndex++) {
      const [row, column] = positions[sectionIndex]!
      const output = await PDFDocument.create()
      copyMetadata(source, output)
      const embeddedPage = await output.embedPage(sourcePage, pageCropBounds(sourcePage))
      if (selectedPages.has(pageIndex)) {
        drawPdfSection(output, embeddedPage, rows, columns, row, column)
      } else {
        drawFullEmbeddedPage(output, embeddedPage)
      }
      outputs.push({
        sourcePageNumber: pageIndex + 1,
        sectionNumber: sectionIndex + 1,
        bytes: await output.save({ useObjectStreams: false }),
      })
    }
  }
  return outputs
}

export async function posterPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: PosterOptions,
): Promise<Uint8Array> {
  const rows = positiveInteger(options.rows, 'rows')
  const columns = positiveInteger(options.columns, 'columns')
  if (rows > 10 || columns > 10) throw new Error('rows and columns must not exceed 10')
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const embeddedPages = await embedCroppedPages(source, output)
  const targetSize = PAGE_SIZES[options.pageSize]

  for (const embeddedPage of embeddedPages) {
    const cellWidth = embeddedPage.width / columns
    const cellHeight = embeddedPage.height / rows
    const scale = Math.min(targetSize[0] / cellWidth, targetSize[1] / cellHeight)
    const scaledCellWidth = cellWidth * scale
    const scaledCellHeight = cellHeight * scale
    const offsetX = (targetSize[0] - scaledCellWidth) / 2
    const offsetY = (targetSize[1] - scaledCellHeight) / 2
    for (let row = 0; row < rows; row++) {
      for (let column = 0; column < columns; column++) {
        const sourceColumn = options.readingDirection === 'rtl' ? columns - 1 - column : column
        const cropX = sourceColumn * cellWidth
        const cropY = (rows - 1 - row) * cellHeight
        const page = output.addPage([targetSize[0], targetSize[1]])
        page.drawPage(embeddedPage, {
          x: offsetX - cropX * scale,
          y: offsetY - cropY * scale,
          width: embeddedPage.width * scale,
          height: embeddedPage.height * scale,
        })
      }
    }
  }
  return output.save({ useObjectStreams: false })
}

export async function singlePagePdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: SinglePageOptions,
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const embeddedPages = await embedCroppedPages(source, output)
  if (embeddedPages.length === 0) throw new Error('Single-page merge requires at least one page')
  const vertical = options.direction === 'vertical'
  const width = vertical
    ? Math.max(...embeddedPages.map((page) => page.width))
    : embeddedPages.reduce((total, page) => total + page.width, 0)
  const height = vertical
    ? embeddedPages.reduce((total, page) => total + page.height, 0)
    : Math.max(...embeddedPages.map((page) => page.height))
  const page = output.addPage([width, height])
  let cursor = vertical ? height : 0
  for (const embeddedPage of embeddedPages) {
    if (vertical) {
      cursor -= embeddedPage.height
      page.drawPage(embeddedPage, { x: 0, y: cursor })
    } else {
      page.drawPage(embeddedPage, { x: cursor, y: 0 })
      cursor += embeddedPage.width
    }
  }
  return output.save({ useObjectStreams: false })
}

export interface OverlayPageAssignment {
  documentIndex: number
  pageIndex: number
}

export function overlayPageAssignments(
  basePageCount: number,
  overlayPageCounts: number[],
  mode: PdfOverlayMode,
  repeatCounts?: number[],
): Array<OverlayPageAssignment | null> {
  if (!Number.isInteger(basePageCount) || basePageCount < 0) {
    throw new Error('basePageCount must be a non-negative integer')
  }
  if (overlayPageCounts.length === 0) throw new Error('At least one overlay PDF is required')
  for (const pageCount of overlayPageCounts) positiveInteger(pageCount, 'overlay page count')

  if (mode === 'sequential') {
    const sequence = overlayPageCounts.flatMap((pageCount, documentIndex) =>
      Array.from({ length: pageCount }, (_, pageIndex) => ({ documentIndex, pageIndex })),
    )
    return Array.from(
      { length: basePageCount },
      (_, pageIndex) => sequence[pageIndex % sequence.length]!,
    )
  }

  if (mode === 'interleaved') {
    return Array.from({ length: basePageCount }, (_, basePageIndex) => {
      const documentIndex = basePageIndex % overlayPageCounts.length
      const visitIndex = Math.floor(basePageIndex / overlayPageCounts.length)
      return {
        documentIndex,
        pageIndex: visitIndex % overlayPageCounts[documentIndex]!,
      }
    })
  }

  if (!repeatCounts || repeatCounts.length !== overlayPageCounts.length) {
    throw new Error('repeatCounts must match the number of overlay PDFs')
  }
  const sequence: OverlayPageAssignment[] = []
  for (let documentIndex = 0; documentIndex < overlayPageCounts.length; documentIndex++) {
    const repeatCount = positiveInteger(repeatCounts[documentIndex]!, 'repeat count')
    for (let repeat = 0; repeat < repeatCount; repeat++) {
      for (let pageIndex = 0; pageIndex < overlayPageCounts[documentIndex]!; pageIndex++) {
        sequence.push({ documentIndex, pageIndex })
      }
    }
  }
  return Array.from({ length: basePageCount }, (_, pageIndex) => sequence[pageIndex] ?? null)
}

export async function overlayPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: OverlayOptions,
): Promise<Uint8Array> {
  const opacity = unitInterval(options.opacity, 'opacity')
  if (options.overlayDocuments.length === 0) throw new Error('At least one overlay PDF is required')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const embeddedDocuments = await Promise.all(
    options.overlayDocuments.map(async (overlayBytes) => {
      if (overlayBytes.byteLength === 0) throw new Error('Overlay PDFs must not be empty')
      const overlayDocument = await PDFDocument.load(overlayBytes, { updateMetadata: false })
      if (overlayDocument.getPageCount() === 0) throw new Error('Overlay PDFs must contain a page')
      return embedCroppedPages(overlayDocument, document)
    }),
  )
  const assignments = overlayPageAssignments(
    document.getPageCount(),
    embeddedDocuments.map((pages) => pages.length),
    options.mode,
    options.repeatCounts,
  )

  for (let pageIndex = 0; pageIndex < assignments.length; pageIndex++) {
    const assignment = assignments[pageIndex]
    if (!assignment) continue
    const page = document.getPage(pageIndex)
    const overlayPage = embeddedDocuments[assignment.documentIndex]![assignment.pageIndex]!
    const cropBox = page.getCropBox()
    const scale = Math.min(cropBox.width / overlayPage.width, cropBox.height / overlayPage.height)
    const width = overlayPage.width * scale
    const height = overlayPage.height * scale
    page.drawPage(overlayPage, {
      x: cropBox.x + (cropBox.width - width) / 2,
      y: cropBox.y + (cropBox.height - height) / 2,
      width,
      height,
      opacity,
    })
    if (options.position === 'background') {
      const contents = page.node.Contents()
      if (contents instanceof PDFArray && contents.size() > 1) {
        const overlayStream = contents.get(contents.size() - 1)
        contents.remove(contents.size() - 1)
        contents.insert(0, overlayStream)
      }
    }
  }
  return document.save({ useObjectStreams: false })
}

function normalizedQuarterTurn(value: number): 0 | 90 | 180 | 270 {
  const normalized = ((value % 360) + 360) % 360
  if (normalized !== 0 && normalized !== 90 && normalized !== 180 && normalized !== 270) {
    throw new Error('PDF page rotation must be a multiple of 90 degrees')
  }
  return normalized
}

function imageOverlayVisualPosition(
  pageWidth: number,
  pageHeight: number,
  imageWidth: number,
  imageHeight: number,
  options: Pick<ImageOverlayOptions, 'position' | 'widthPercent' | 'margin' | 'x' | 'y'>,
): { left: number; top: number; width: number; height: number } {
  if (
    !Number.isFinite(pageWidth) ||
    pageWidth <= 0 ||
    !Number.isFinite(pageHeight) ||
    pageHeight <= 0
  ) {
    throw new Error('PDF page dimensions must be positive')
  }
  if (
    !Number.isFinite(imageWidth) ||
    imageWidth <= 0 ||
    !Number.isFinite(imageHeight) ||
    imageHeight <= 0
  ) {
    throw new Error('Image dimensions must be positive')
  }
  if (
    !Number.isFinite(options.widthPercent) ||
    options.widthPercent <= 0 ||
    options.widthPercent > 100
  ) {
    throw new Error('Image width must be greater than 0% and at most 100%')
  }
  const margin = finiteNonNegative(options.margin, 'margin')
  if (margin * 2 >= pageWidth || margin * 2 >= pageHeight) {
    throw new Error('Image margin is too large for the page')
  }
  const maxWidth = pageWidth - margin * 2
  const maxHeight = pageHeight - margin * 2
  const requestedWidth = (pageWidth * options.widthPercent) / 100
  const scale = Math.min(
    requestedWidth / imageWidth,
    maxWidth / imageWidth,
    maxHeight / imageHeight,
  )
  const width = imageWidth * scale
  const height = imageHeight * scale
  const horizontal = options.position.endsWith('Left')
    ? margin
    : options.position.endsWith('Right')
      ? pageWidth - margin - width
      : (pageWidth - width) / 2
  const vertical = options.position.startsWith('top')
    ? margin
    : options.position.startsWith('bottom')
      ? pageHeight - margin - height
      : (pageHeight - height) / 2
  const left = options.position === 'custom' ? finiteNonNegative(options.x ?? NaN, 'x') : horizontal
  const top = options.position === 'custom' ? finiteNonNegative(options.y ?? NaN, 'y') : vertical
  return {
    left: Math.min(left, pageWidth - width),
    top: Math.min(top, pageHeight - height),
    width,
    height,
  }
}

export function pdfImageOverlayPlacement(
  pageWidth: number,
  pageHeight: number,
  pageRotation: number,
  imageWidth: number,
  imageHeight: number,
  options: Pick<ImageOverlayOptions, 'position' | 'widthPercent' | 'margin' | 'x' | 'y'>,
  pageX = 0,
  pageY = 0,
): PdfImageOverlayPlacement {
  const rotation = normalizedQuarterTurn(pageRotation)
  const displayWidth = rotation % 180 === 0 ? pageWidth : pageHeight
  const displayHeight = rotation % 180 === 0 ? pageHeight : pageWidth
  const visual = imageOverlayVisualPosition(
    displayWidth,
    displayHeight,
    imageWidth,
    imageHeight,
    options,
  )
  if (rotation === 90) {
    return {
      x: pageX + visual.top + visual.height,
      y: pageY + visual.left,
      width: visual.width,
      height: visual.height,
      rotation,
    }
  }
  if (rotation === 180) {
    return {
      x: pageX + pageWidth - visual.left,
      y: pageY + visual.top + visual.height,
      width: visual.width,
      height: visual.height,
      rotation,
    }
  }
  if (rotation === 270) {
    return {
      x: pageX + pageWidth - visual.top - visual.height,
      y: pageY + pageHeight - visual.left,
      width: visual.width,
      height: visual.height,
      rotation,
    }
  }
  return {
    x: pageX + visual.left,
    y: pageY + pageHeight - visual.top - visual.height,
    width: visual.width,
    height: visual.height,
    rotation,
  }
}

function isPngImage(bytes: Uint8Array): boolean {
  return (
    bytes.length >= 8 &&
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a
  )
}

function isJpegImage(bytes: Uint8Array): boolean {
  return bytes.length >= 3 && bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff
}

function moveLastPageContentStreamToBackground(page: PDFPage): void {
  const contents = page.node.Contents()
  if (!(contents instanceof PDFArray) || contents.size() <= 1) return
  const stream = contents.get(contents.size() - 1)
  contents.remove(contents.size() - 1)
  contents.insert(0, stream)
}

export async function overlayImagePdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: ImageOverlayOptions,
): Promise<Uint8Array> {
  const opacity = unitInterval(options.opacity, 'opacity')
  if (options.image.length === 0) throw new Error('Choose an image to overlay')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageIndexes = checkedPageIndexes(document.getPageCount(), options.pageIndexes)
  const image = isPngImage(options.image)
    ? await document.embedPng(options.image)
    : isJpegImage(options.image)
      ? await document.embedJpg(options.image)
      : null
  if (!image) throw new Error('Only PNG and JPEG images are supported')

  for (const pageIndex of pageIndexes) {
    const page = document.getPage(pageIndex)
    const crop = page.getCropBox()
    const placement = pdfImageOverlayPlacement(
      crop.width,
      crop.height,
      page.getRotation().angle,
      image.width,
      image.height,
      options,
      crop.x,
      crop.y,
    )
    page.drawImage(image, {
      x: placement.x,
      y: placement.y,
      width: placement.width,
      height: placement.height,
      rotate: degrees(placement.rotation),
      opacity,
    })
    if (options.layer === 'background') moveLastPageContentStreamToBackground(page)
  }
  return document.save({ useObjectStreams: false })
}

function checkedRasterImage(image: PdfRasterPage, index: number): PdfRasterPage {
  if (!image.image || image.image.length === 0) {
    throw new Error(`Image ${index + 1} is empty`)
  }
  if (
    !Number.isFinite(image.width) ||
    !Number.isFinite(image.height) ||
    image.width <= 0 ||
    image.height <= 0
  ) {
    throw new Error(`Image ${index + 1} size is invalid`)
  }
  return image
}

const MAX_VECTOR_PAGE_PDF_BYTES = 50 * 1024 * 1024

function isVectorPdfPage(page: PdfImagesToPdfPage): page is PdfVectorPage {
  return 'kind' in page && page.kind === 'vectorPdf'
}

export async function imagesToPdfBytes(
  images: PdfImagesToPdfPage[],
  fitOption: PdfImagesToPdfFit,
  autoRotate: boolean,
): Promise<Uint8Array> {
  if (images.length === 0) throw new Error('Choose at least one image')
  if (!['fillPage', 'fitDocumentToImage', 'maintainAspectRatio'].includes(fitOption)) {
    throw new Error('Image fit option is invalid')
  }

  const output = await PDFDocument.create()
  for (let index = 0; index < images.length; index++) {
    const inputPage = images[index]!
    let sourceWidth: number
    let sourceHeight: number
    let drawContent: (
      page: PDFPage,
      placement: { x: number; y: number; width: number; height: number },
    ) => void

    if (isVectorPdfPage(inputPage)) {
      if (!inputPage.pdf.length) throw new Error(`Vector page ${index + 1} is empty`)
      if (inputPage.pdf.length > MAX_VECTOR_PAGE_PDF_BYTES) {
        throw new Error(`Vector page ${index + 1} must be 50 MB or smaller`)
      }
      let vectorDocument: PDFDocument
      try {
        vectorDocument = await PDFDocument.load(inputPage.pdf, { updateMetadata: false })
      } catch {
        throw new Error(`Vector page ${index + 1} must be a valid PDF`)
      }
      if (vectorDocument.getPageCount() !== 1) {
        throw new Error(`Vector page ${index + 1} must contain exactly one PDF page`)
      }
      const embeddedPage = await output.embedPage(vectorDocument.getPage(0))
      sourceWidth = embeddedPage.width
      sourceHeight = embeddedPage.height
      drawContent = (page, placement) => page.drawPage(embeddedPage, placement)
    } else {
      const rasterImage = checkedRasterImage(inputPage, index)
      const embeddedImage = isPngImage(rasterImage.image)
        ? await output.embedPng(rasterImage.image)
        : isJpegImage(rasterImage.image)
          ? await output.embedJpg(rasterImage.image)
          : null
      if (!embeddedImage) throw new Error(`Image ${index + 1} must be PNG or JPEG`)
      sourceWidth = rasterImage.width
      sourceHeight = rasterImage.height
      drawContent = (page, placement) => page.drawImage(embeddedImage, placement)
    }

    const [portraitWidth, portraitHeight] = PAGE_SIZES.A4
    const landscape = autoRotate && sourceWidth > sourceHeight
    const pageWidth =
      fitOption === 'fitDocumentToImage' ? sourceWidth : landscape ? portraitHeight : portraitWidth
    const pageHeight =
      fitOption === 'fitDocumentToImage' ? sourceHeight : landscape ? portraitWidth : portraitHeight
    const page = output.addPage([pageWidth, pageHeight])

    if (fitOption === 'maintainAspectRatio') {
      const scale = Math.min(pageWidth / sourceWidth, pageHeight / sourceHeight)
      const width = sourceWidth * scale
      const height = sourceHeight * scale
      drawContent(page, {
        x: (pageWidth - width) / 2,
        y: (pageHeight - height) / 2,
        width,
        height,
      })
    } else {
      drawContent(page, { x: 0, y: 0, width: pageWidth, height: pageHeight })
    }
  }

  return output.save({ useObjectStreams: false })
}

const MAX_CBZ_ARCHIVE_BYTES = 200 * 1024 * 1024
const MAX_CBZ_ENTRY_BYTES = 50 * 1024 * 1024
const MAX_CBZ_TOTAL_IMAGE_BYTES = 200 * 1024 * 1024
const MAX_CBZ_IMAGES = 200
const MAX_CBZ_ENTRIES = 2_000
const CBZ_IMAGE_EXTENSION = /\.(?:bmp|gif|jpe?g|png|webp)$/i
const CBZ_NATURAL_ORDER = new Intl.Collator('en', { numeric: true, sensitivity: 'base' })

type SizedZipEntry = JSZip.JSZipObject & {
  unsafeOriginalName?: string
  _data?: { uncompressedSize?: number }
}

function checkedCbzEntryName(entry: SizedZipEntry): string {
  const originalName = entry.unsafeOriginalName ?? entry.name
  const normalized = originalName.replace(/\\/g, '/')
  if (
    normalized.includes('\0') ||
    normalized.startsWith('/') ||
    /^[a-z]:\//i.test(normalized) ||
    normalized.split('/').includes('..')
  ) {
    throw new Error('CBZ archive contains an unsafe file path')
  }
  return normalized
}

function isVisibleCbzImage(name: string): boolean {
  if (!CBZ_IMAGE_EXTENSION.test(name)) return false
  const segments = name.split('/')
  return !segments.some((segment) => segment.startsWith('.') || segment === '__MACOSX')
}

export async function extractCbzImageEntries(
  archiveBytes: Uint8Array | ArrayBuffer,
): Promise<CbzImageEntry[]> {
  const bytes = archiveBytes instanceof Uint8Array ? archiveBytes : new Uint8Array(archiveBytes)
  if (bytes.length === 0) throw new Error('CBZ archive is empty')
  if (bytes.length > MAX_CBZ_ARCHIVE_BYTES) throw new Error('CBZ archive must be 200 MB or smaller')

  let archive: JSZip
  try {
    archive = await JSZip.loadAsync(bytes)
  } catch {
    throw new Error('CBZ archive is invalid or damaged')
  }

  const allEntries = Object.values(archive.files)
  if (allEntries.length > MAX_CBZ_ENTRIES) throw new Error('CBZ archive contains too many files')

  const imageEntries = allEntries
    .filter((entry) => !entry.dir)
    .map((entry) => ({ entry: entry as SizedZipEntry, name: checkedCbzEntryName(entry) }))
    .filter(({ name }) => isVisibleCbzImage(name))
    .sort((left, right) => CBZ_NATURAL_ORDER.compare(left.name, right.name))

  if (imageEntries.length === 0) throw new Error('CBZ archive contains no supported images')
  if (imageEntries.length > MAX_CBZ_IMAGES) {
    throw new Error(`CBZ archive may contain no more than ${MAX_CBZ_IMAGES} images`)
  }

  let totalBytes = 0
  for (const { entry } of imageEntries) {
    const declaredSize = entry._data?.uncompressedSize
    if (typeof declaredSize === 'number') {
      if (declaredSize > MAX_CBZ_ENTRY_BYTES)
        throw new Error('Each CBZ image must be 50 MB or smaller')
      totalBytes += declaredSize
    }
  }
  if (totalBytes > MAX_CBZ_TOTAL_IMAGE_BYTES) {
    throw new Error('CBZ images must total 200 MB or less')
  }

  const images: CbzImageEntry[] = []
  totalBytes = 0
  for (const { entry, name } of imageEntries) {
    const imageBytes = await entry.async('uint8array')
    if (imageBytes.length > MAX_CBZ_ENTRY_BYTES) {
      throw new Error(`${name}: image must be 50 MB or smaller`)
    }
    totalBytes += imageBytes.length
    if (totalBytes > MAX_CBZ_TOTAL_IMAGE_BYTES) {
      throw new Error('CBZ images must total 200 MB or less')
    }
    images.push({ name, bytes: imageBytes })
  }
  return images
}

export function cbzPdfOutputFileName(baseName?: string): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.(?:cbr|cbz|rar|zip)$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'comic'}_converted.pdf`
}

export function emailDocumentOutputFileName(
  baseName: string | undefined,
  format: EmailDocumentOutputFormat,
): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.eml$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'email'}_converted.${format}`
}

export function epubPdfOutputFileName(baseName?: string): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.epub$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'ebook'}_converted.pdf`
}

export function htmlPdfOutputFileName(baseName?: string): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.(?:html?|zip)$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'webpage'}_converted.pdf`
}

export function createdPdfOutputFileName(baseName?: string): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.pdf$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'document'}.pdf`
}

export function markdownPdfOutputFileName(baseName?: string): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.(?:md|markdown|zip)$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'markdown'}_converted.pdf`
}

async function epubPagesToPdfBytes(options: EpubToPdfOptions): Promise<Uint8Array> {
  const bytes = await imagesToPdfBytes(options.pages, 'fitDocumentToImage', false)
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const title = options.title?.trim()
  const author = options.author?.trim()
  if (title) document.setTitle(title.slice(0, 500))
  if (author) document.setAuthor(author.slice(0, 500))
  document.setCreator('GenOffice PDF')
  document.setProducer('GenOffice PDF')
  return document.save({ useObjectStreams: false })
}

async function htmlPagesToPdfBytes(options: HtmlToPdfOptions): Promise<Uint8Array> {
  const bytes = await imagesToPdfBytes(options.pages, 'fitDocumentToImage', false)
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const title = options.title?.trim()
  if (title) document.setTitle(title.slice(0, 500))
  document.setCreator('GenOffice PDF')
  document.setProducer('GenOffice PDF')
  return document.save({ useObjectStreams: false })
}

export async function appendImagesToPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  images: PdfImagesToPdfPage[],
  fitOption: PdfImagesToPdfFit,
  autoRotate: boolean,
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const imageDocument = await PDFDocument.load(
    await imagesToPdfBytes(images, fitOption, autoRotate),
    { updateMetadata: false },
  )
  const pages = await source.copyPages(imageDocument, imageDocument.getPageIndices())
  for (const page of pages) source.addPage(page)
  return source.save({ useObjectStreams: false })
}

const MAX_ATTACHMENT_SIZE = 50 * 1024 * 1024
const MAX_TOTAL_ATTACHMENT_SIZE = 200 * 1024 * 1024

interface PdfAttachmentRecord {
  key: string
  fileSpecObject: PDFObject
  fileSpec: PDFDict
  stream: PDFRawStream
  info: PdfAttachmentInfo
}

function decodePdfText(value: PDFObject | undefined): string | undefined {
  if (value instanceof PDFHexString || value instanceof PDFString || value instanceof PDFName) {
    return value.decodeText()
  }
  return undefined
}

function resolvedPdfObject(
  document: PDFDocument,
  value: PDFObject | undefined,
): PDFObject | undefined {
  if (!value) return undefined
  try {
    return document.context.lookup(value) as PDFObject
  } catch {
    return undefined
  }
}

function decodePdfJavaScriptBytes(bytes: Uint8Array): string {
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes.subarray(2))
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2))
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes)
  } catch {
    return new TextDecoder('windows-1252').decode(bytes)
  }
}

function decodePdfJavaScript(
  document: PDFDocument,
  value: PDFObject | undefined,
): string | undefined {
  const resolved = resolvedPdfObject(document, value)
  if (resolved instanceof PDFHexString || resolved instanceof PDFString) {
    try {
      return resolved.decodeText()
    } catch {
      return undefined
    }
  }
  if (!(resolved instanceof PDFRawStream)) return undefined
  try {
    return decodePdfJavaScriptBytes(decodePDFRawStream(resolved).decode())
  } catch {
    return undefined
  }
}

interface PdfJavaScriptLocation {
  source: PdfJavaScriptSource
  trigger: string
  name?: string
  pageNumber?: number
  annotationNumber?: number
  fieldName?: string
}

function collectPdfJavaScriptAction(
  document: PDFDocument,
  value: PDFObject | undefined,
  location: PdfJavaScriptLocation,
  actions: PdfJavaScriptAction[],
  signatures: Set<string>,
  visited: Set<PDFObject>,
  depth = 0,
): void {
  if (depth > 64 || actions.length >= 1000) return
  const resolved = resolvedPdfObject(document, value)
  if (!(resolved instanceof PDFDict) || visited.has(resolved)) return
  visited.add(resolved)

  if (nameIs(resolved.lookupMaybe(PDFName.of('S'), PDFName), 'JavaScript')) {
    const code = decodePdfJavaScript(document, resolved.get(PDFName.of('JS')))
    if (code?.trim()) {
      const action = { ...location, code }
      const signature = JSON.stringify(action)
      if (!signatures.has(signature)) {
        signatures.add(signature)
        actions.push(action)
      }
    }
  }

  const next = resolvedPdfObject(document, resolved.get(PDFName.of('Next')))
  if (next instanceof PDFArray) {
    for (let index = 0; index < next.size(); index++) {
      collectPdfJavaScriptAction(
        document,
        next.get(index),
        { ...location, trigger: `${location.trigger} / Next ${index + 1}` },
        actions,
        signatures,
        visited,
        depth + 1,
      )
    }
  } else if (next) {
    collectPdfJavaScriptAction(
      document,
      next,
      { ...location, trigger: `${location.trigger} / Next` },
      actions,
      signatures,
      visited,
      depth + 1,
    )
  }
  visited.delete(resolved)
}

function collectAdditionalPdfJavaScriptActions(
  document: PDFDocument,
  dictionary: PDFDict,
  location: Omit<PdfJavaScriptLocation, 'trigger'>,
  actions: PdfJavaScriptAction[],
  signatures: Set<string>,
): void {
  const additionalActions = dictionary.lookupMaybe(PDFName.of('AA'), PDFDict)
  if (!additionalActions) return
  for (const key of additionalActions.keys()) {
    collectPdfJavaScriptAction(
      document,
      additionalActions.get(key),
      { ...location, trigger: key.decodeText() },
      actions,
      signatures,
      new Set(),
    )
  }
}

function collectNamedPdfJavaScriptActions(
  document: PDFDocument,
  nodeValue: PDFObject | undefined,
  actions: PdfJavaScriptAction[],
  signatures: Set<string>,
  visited: Set<PDFObject>,
): void {
  const node = resolvedPdfObject(document, nodeValue)
  if (!(node instanceof PDFDict) || visited.has(node)) return
  visited.add(node)
  const names = node.lookupMaybe(PDFName.of('Names'), PDFArray)
  if (names) {
    for (let index = 0; index + 1 < names.size(); index += 2) {
      collectPdfJavaScriptAction(
        document,
        names.get(index + 1),
        {
          source: 'named',
          trigger: 'NamedJavaScript',
          name:
            decodePdfText(resolvedPdfObject(document, names.get(index))) ??
            `Script ${index / 2 + 1}`,
        },
        actions,
        signatures,
        new Set(),
      )
    }
  }
  const kids = node.lookupMaybe(PDFName.of('Kids'), PDFArray)
  if (!kids) return
  for (let index = 0; index < kids.size(); index++) {
    collectNamedPdfJavaScriptActions(document, kids.get(index), actions, signatures, visited)
  }
}

function collectFormPdfJavaScriptActions(
  document: PDFDocument,
  value: PDFObject | undefined,
  parentName: string,
  actions: PdfJavaScriptAction[],
  signatures: Set<string>,
  visited: Set<PDFObject>,
  formContainers: Set<PDFObject>,
  depth = 0,
): void {
  if (depth > 64) return
  const field = resolvedPdfObject(document, value)
  if (!(field instanceof PDFDict) || visited.has(field)) return
  visited.add(field)
  formContainers.add(field)
  const partialName = decodePdfText(resolvedPdfObject(document, field.get(PDFName.of('T'))))
  const fieldName = partialName
    ? parentName
      ? `${parentName}.${partialName}`
      : partialName
    : parentName
  const location = { source: 'form' as const, fieldName: fieldName || undefined }
  collectPdfJavaScriptAction(
    document,
    field.get(PDFName.of('A')),
    { ...location, trigger: 'A' },
    actions,
    signatures,
    new Set(),
  )
  collectAdditionalPdfJavaScriptActions(document, field, location, actions, signatures)
  const kids = field.lookupMaybe(PDFName.of('Kids'), PDFArray)
  if (!kids) return
  for (let index = 0; index < kids.size(); index++) {
    collectFormPdfJavaScriptActions(
      document,
      kids.get(index),
      fieldName,
      actions,
      signatures,
      visited,
      formContainers,
      depth + 1,
    )
  }
}

export async function auditPdfJavaScriptBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfJavaScriptAudit> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const actions: PdfJavaScriptAction[] = []
  const signatures = new Set<string>()
  const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  collectNamedPdfJavaScriptActions(
    document,
    names?.get(PDFName.of('JavaScript')),
    actions,
    signatures,
    new Set(),
  )
  collectPdfJavaScriptAction(
    document,
    document.catalog.get(PDFName.of('OpenAction')),
    { source: 'document', trigger: 'OpenAction' },
    actions,
    signatures,
    new Set(),
  )
  collectAdditionalPdfJavaScriptActions(
    document,
    document.catalog,
    { source: 'document' },
    actions,
    signatures,
  )

  const formContainers = new Set<PDFObject>()
  const acroForm = document.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const fields = acroForm?.lookupMaybe(PDFName.of('Fields'), PDFArray)
  if (fields) {
    const visitedFields = new Set<PDFObject>()
    for (let index = 0; index < fields.size(); index++) {
      collectFormPdfJavaScriptActions(
        document,
        fields.get(index),
        '',
        actions,
        signatures,
        visitedFields,
        formContainers,
      )
    }
  }

  document.getPages().forEach((page, pageIndex) => {
    const pageNumber = pageIndex + 1
    collectAdditionalPdfJavaScriptActions(
      document,
      page.node,
      { source: 'page', pageNumber },
      actions,
      signatures,
    )
    const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annotations) return
    for (let index = 0; index < annotations.size(); index++) {
      const annotation = resolvedPdfObject(document, annotations.get(index))
      if (!(annotation instanceof PDFDict) || formContainers.has(annotation)) continue
      const location = {
        source: 'annotation' as const,
        pageNumber,
        annotationNumber: index + 1,
      }
      collectPdfJavaScriptAction(
        document,
        annotation.get(PDFName.of('A')),
        { ...location, trigger: 'A' },
        actions,
        signatures,
        new Set(),
      )
      collectAdditionalPdfJavaScriptActions(document, annotation, location, actions, signatures)
    }
  })

  const encoder = new TextEncoder()
  return {
    actions,
    uniqueScriptCount: new Set(actions.map((action) => action.code)).size,
    totalCodeBytes: actions.reduce(
      (total, action) => total + encoder.encode(action.code).byteLength,
      0,
    ),
  }
}

function attachmentStream(document: PDFDocument, fileSpec: PDFDict): PDFRawStream | undefined {
  const embeddedFiles = fileSpec.lookupMaybe(PDFName.of('EF'), PDFDict)
  if (!embeddedFiles) return undefined
  for (const key of ['UF', 'F', 'DOS', 'Mac', 'Unix']) {
    const candidate = document.context.lookup(embeddedFiles.get(PDFName.of(key)))
    if (candidate instanceof PDFRawStream) return candidate
  }
  return undefined
}

function attachmentName(fileSpec: PDFDict, fallback: string): string {
  return (
    decodePdfText(fileSpec.get(PDFName.of('UF'))) ??
    decodePdfText(fileSpec.get(PDFName.of('F'))) ??
    fallback
  )
}

function collectAttachmentRecords(
  document: PDFDocument,
  node: PDFDict,
  records: PdfAttachmentRecord[],
  visited = new Set<PDFDict>(),
): void {
  if (visited.has(node)) return
  visited.add(node)
  const names = node.lookupMaybe(PDFName.of('Names'), PDFArray)
  if (names) {
    for (let index = 0; index + 1 < names.size(); index += 2) {
      const key = decodePdfText(names.get(index)) ?? `attachment_${records.length + 1}`
      const fileSpecObject = names.get(index + 1)
      const fileSpec = document.context.lookup(fileSpecObject)
      if (!(fileSpec instanceof PDFDict)) continue
      const stream = attachmentStream(document, fileSpec)
      if (!stream) continue
      const params = stream.dict.lookupMaybe(PDFName.of('Params'), PDFDict)
      const declaredSize = params?.lookupMaybe(PDFName.of('Size'), PDFNumber)?.asNumber()
      records.push({
        key,
        fileSpecObject,
        fileSpec,
        stream,
        info: {
          name: attachmentName(fileSpec, key),
          size: declaredSize ?? stream.getContents().byteLength,
          mimeType: decodePdfText(stream.dict.get(PDFName.of('Subtype'))),
          description: decodePdfText(fileSpec.get(PDFName.of('Desc'))),
        },
      })
    }
  }
  const kids = node.lookupMaybe(PDFName.of('Kids'), PDFArray)
  if (!kids) return
  for (let index = 0; index < kids.size(); index++) {
    const kid = kids.lookupMaybe(index, PDFDict)
    if (kid) collectAttachmentRecords(document, kid, records, visited)
  }
}

function attachmentRecords(document: PDFDocument): PdfAttachmentRecord[] {
  const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  const embeddedFiles = names?.lookupMaybe(PDFName.of('EmbeddedFiles'), PDFDict)
  if (!embeddedFiles) return []
  const records: PdfAttachmentRecord[] = []
  collectAttachmentRecords(document, embeddedFiles, records)
  return records
}

function samePdfObject(left: PDFObject, right: PDFObject): boolean {
  return left === right || left.toString() === right.toString()
}

function removeAssociatedFileReferences(document: PDFDocument, removedObjects: PDFObject[]): void {
  const associatedFiles = document.catalog.lookupMaybe(PDFName.of('AF'), PDFArray)
  if (!associatedFiles) return
  for (let index = associatedFiles.size() - 1; index >= 0; index--) {
    if (removedObjects.some((removed) => samePdfObject(removed, associatedFiles.get(index)))) {
      associatedFiles.remove(index)
    }
  }
  if (associatedFiles.size() === 0) document.catalog.delete(PDFName.of('AF'))
}

function rewriteAttachmentTree(
  document: PDFDocument,
  records: PdfAttachmentRecord[],
  removedObjects: PDFObject[] = [],
): void {
  const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  const embeddedFiles = names?.lookupMaybe(PDFName.of('EmbeddedFiles'), PDFDict)
  if (!names || !embeddedFiles) return
  if (records.length === 0) {
    names.delete(PDFName.of('EmbeddedFiles'))
    if (names.keys().length === 0) document.catalog.delete(PDFName.of('Names'))
  } else {
    const flattenedNames = document.context.obj([])
    for (const record of records) {
      flattenedNames.push(PDFHexString.fromText(record.key))
      flattenedNames.push(record.fileSpecObject)
    }
    embeddedFiles.set(PDFName.of('Names'), flattenedNames)
    embeddedFiles.delete(PDFName.of('Kids'))
  }
  removeAssociatedFileReferences(document, removedObjects)
}

function stripAsciiControlCharacters(value: string, includeDelete = true): string {
  return Array.from(value)
    .filter((character) => {
      const code = character.charCodeAt(0)
      return code > 0x1f && (!includeDelete || code !== 0x7f)
    })
    .join('')
}

function safeAttachmentName(name: string, fallback = 'attachment.bin'): string {
  const simpleName = name.replace(/\\/g, '/').split('/').pop()?.trim() ?? ''
  const sanitized = stripAsciiControlCharacters(simpleName).replace(/[:*?"<>|]/g, '_')
  return sanitized && sanitized !== '.' && sanitized !== '..' ? sanitized : fallback
}

function uniqueAttachmentName(name: string, usedNames: Set<string>): string {
  const dotIndex = name.lastIndexOf('.')
  const baseName = dotIndex > 0 ? name.slice(0, dotIndex) : name
  const extension = dotIndex > 0 ? name.slice(dotIndex) : ''
  let candidate = name
  let counter = 1
  while (usedNames.has(candidate)) candidate = `${baseName}_${counter++}${extension}`
  usedNames.add(candidate)
  return candidate
}

function decodedAttachmentBytes(record: PdfAttachmentRecord): Uint8Array {
  return decodePDFRawStream(record.stream).decode()
}

export async function listPdfAttachmentsBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfAttachmentInfo[]> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  return attachmentRecords(document).map((record) => record.info)
}

export async function addPdfAttachmentsBytes(
  bytes: Uint8Array | ArrayBuffer,
  attachments: PdfAttachmentInput[],
): Promise<Uint8Array> {
  if (attachments.length === 0) throw new Error('At least one attachment is required')
  let totalSize = 0
  for (const attachment of attachments) {
    if (attachment.bytes.byteLength === 0) throw new Error('Attachments must not be empty')
    if (attachment.bytes.byteLength > MAX_ATTACHMENT_SIZE) {
      throw new Error(`Attachment ${attachment.name} exceeds 50 MB`)
    }
    totalSize += attachment.bytes.byteLength
  }
  if (totalSize > MAX_TOTAL_ATTACHMENT_SIZE) throw new Error('Attachments exceed 200 MB in total')

  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const normalized = attachments.map((attachment, index) => ({
    ...attachment,
    name: safeAttachmentName(attachment.name, `attachment_${index + 1}.bin`),
  }))
  const replacedNames = new Set(normalized.map((attachment) => attachment.name))
  const existingRecords = attachmentRecords(document)
  const retainedRecords = existingRecords.filter((record) => !replacedNames.has(record.info.name))
  const removedRecords = existingRecords.filter((record) => replacedNames.has(record.info.name))
  if (removedRecords.length > 0) {
    rewriteAttachmentTree(
      document,
      retainedRecords,
      removedRecords.map((record) => record.fileSpecObject),
    )
  }
  for (const attachment of normalized) {
    await document.attach(attachment.bytes, attachment.name, {
      mimeType: attachment.mimeType,
      description: `Embedded attachment: ${attachment.name}`,
      creationDate: new Date(),
      modificationDate: new Date(),
    })
  }
  document.catalog.set(PDFName.of('PageMode'), PDFName.of('UseAttachments'))
  return document.save({ useObjectStreams: false })
}

export async function renamePdfAttachmentBytes(
  bytes: Uint8Array | ArrayBuffer,
  attachmentNameValue: string,
  newNameValue: string,
): Promise<Uint8Array> {
  const attachmentName = attachmentNameValue.trim()
  const newName = safeAttachmentName(newNameValue, '')
  if (!attachmentName) throw new Error('Attachment name is required')
  if (!newName) throw new Error('New attachment name is required')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const records = attachmentRecords(document)
  if (
    records.some((record) => record.info.name === newName && record.info.name !== attachmentName)
  ) {
    throw new Error(`Attachment ${newName} already exists`)
  }
  const record = records.find((candidate) => candidate.info.name === attachmentName)
  if (!record) throw new Error(`Attachment ${attachmentName} was not found`)
  record.key = newName
  record.info.name = newName
  record.fileSpec.set(PDFName.of('F'), PDFHexString.fromText(newName))
  record.fileSpec.set(PDFName.of('UF'), PDFHexString.fromText(newName))
  record.fileSpec.set(PDFName.of('Desc'), PDFHexString.fromText(`Embedded attachment: ${newName}`))
  rewriteAttachmentTree(document, records)
  return document.save({ useObjectStreams: false })
}

export async function deletePdfAttachmentBytes(
  bytes: Uint8Array | ArrayBuffer,
  attachmentNameValue: string,
): Promise<Uint8Array> {
  const attachmentName = attachmentNameValue.trim()
  if (!attachmentName) throw new Error('Attachment name is required')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const records = attachmentRecords(document)
  const removedRecords = records.filter((record) => record.info.name === attachmentName)
  if (removedRecords.length === 0) throw new Error(`Attachment ${attachmentName} was not found`)
  rewriteAttachmentTree(
    document,
    records.filter((record) => record.info.name !== attachmentName),
    removedRecords.map((record) => record.fileSpecObject),
  )
  return document.save({ useObjectStreams: false })
}

export async function extractPdfAttachmentsZipBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const records = attachmentRecords(document)
  if (records.length === 0) throw new Error('No embedded attachments were found')
  const archive = new JSZip()
  const usedNames = new Set<string>()
  let totalSize = 0
  for (let index = 0; index < records.length; index++) {
    const data = decodedAttachmentBytes(records[index]!)
    if (data.byteLength > MAX_ATTACHMENT_SIZE) continue
    if (totalSize + data.byteLength > MAX_TOTAL_ATTACHMENT_SIZE) continue
    totalSize += data.byteLength
    const name = uniqueAttachmentName(
      safeAttachmentName(records[index]!.info.name, `attachment_${index + 1}.bin`),
      usedNames,
    )
    archive.file(name, data)
  }
  if (usedNames.size === 0) throw new Error('No attachments met the extraction size limits')
  return archive.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

function safeExtractedImageBaseName(value: string): string {
  const withoutExtension = value.trim().replace(/\.pdf$/i, '')
  const sanitized = stripAsciiControlCharacters(withoutExtension)
    .replace(/[/\\:*?"<>|]/g, '_')
    .replace(/^\.+/, '')
    .trim()
  return sanitized || 'Document'
}

export async function extractPdfImagesZipBytes(
  images: PdfExtractedImage[],
  baseName: string,
  format: PdfExtractImageFormat,
): Promise<Uint8Array> {
  if (!['png', 'jpg', 'gif'].includes(format)) throw new Error('Unsupported image format')
  const archive = new JSZip()
  const safeBaseName = safeExtractedImageBaseName(baseName)
  for (const image of images) {
    if (!Number.isInteger(image.pageNumber) || image.pageNumber < 1) {
      throw new Error('Extracted image page number must be a positive integer')
    }
    if (!Number.isInteger(image.imageNumber) || image.imageNumber < 1) {
      throw new Error('Extracted image number must be a positive integer')
    }
    archive.file(
      `${safeBaseName}_page_${image.pageNumber}_${image.imageNumber}.${format}`,
      image.bytes,
    )
  }
  return archive.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 9 },
  })
}

function checkedPdfPageImages(
  pageCount: number,
  options: PdfToImagesOptions,
): PdfRenderedPageImage[] {
  const pageIndexes = checkedPageIndexes(pageCount, options.pageIndexes)
  if (!['png', 'jpg', 'gif', 'webp'].includes(options.format)) {
    throw new Error('PDF page image format is invalid')
  }
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 300) {
    throw new Error('PDF page image DPI must be a whole number from 72 to 300')
  }
  if (
    !Number.isInteger(options.imageQuality) ||
    options.imageQuality < 10 ||
    options.imageQuality > 100
  ) {
    throw new Error('PDF page image quality must be a whole percentage from 10 to 100')
  }
  if (!['color', 'greyscale', 'blackwhite'].includes(options.colorMode)) {
    throw new Error('PDF page image color mode is invalid')
  }
  if (!options.images) throw new Error('Rendered PDF page images are required')
  const expectedImageCount = options.outputMode === 'single' ? 1 : pageIndexes.length
  if (!['single', 'multiple'].includes(options.outputMode)) {
    throw new Error('PDF page image output mode is invalid')
  }
  if (options.images.length !== expectedImageCount) {
    throw new Error('Rendered PDF page images do not match the requested pages')
  }
  const expectedPages = pageIndexes.map((pageIndex) => pageIndex + 1)
  return options.images.map((image, index) => {
    const expectedPage = options.outputMode === 'single' ? expectedPages[0] : expectedPages[index]
    if (image.pageNumber !== expectedPage) {
      throw new Error('Rendered PDF page image order is invalid')
    }
    if (!(image.bytes instanceof Uint8Array) || image.bytes.length === 0) {
      throw new Error('Rendered PDF page image is empty')
    }
    return image
  })
}

export async function pdfPageImagesOutput(options: PdfToImagesOptions): Promise<PdfToolOutput> {
  if (!Number.isInteger(options.pageCount) || options.pageCount < 1) {
    throw new Error('PDF page count must be a positive integer')
  }
  const images = checkedPdfPageImages(options.pageCount, options)
  const extension = options.format === 'jpg' ? '.jpg' : `.${options.format}`
  const mimeType = options.format === 'jpg' ? 'image/jpeg' : `image/${options.format}`
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const pageDigits = Math.max(1, String(options.pageCount).length)
  const imageName = (image: PdfRenderedPageImage): string =>
    `${baseName}_page_${String(image.pageNumber).padStart(pageDigits, '0')}${extension}`

  if (options.outputMode === 'single') {
    return {
      suffix: `_long${extension}`,
      fileName: `${baseName}_long${extension}`,
      bytes: images[0]!.bytes,
      mimeType,
      extension,
    }
  }

  if (images.length === 1) {
    return {
      suffix: `_page_${images[0]!.pageNumber}${extension}`,
      fileName: imageName(images[0]!),
      bytes: images[0]!.bytes,
      mimeType,
      extension,
    }
  }

  const archive = new JSZip()
  for (const image of images) archive.file(imageName(image), image.bytes)
  return {
    suffix: '_images.zip',
    bytes: await archive.generateAsync({
      type: 'uint8array',
      compression: options.format === 'png' ? 'STORE' : 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
    mimeType: 'application/zip',
    extension: '.zip',
  }
}

export async function pdfToCbzBytes(options: PdfToCbzOptions): Promise<PdfToolOutput> {
  if (!['png', 'jpg', 'webp'].includes(options.format)) {
    throw new Error('PDF CBZ image format is invalid')
  }
  const images = checkedPdfPageImages(options.pageCount, {
    ...options,
    outputMode: 'multiple',
  })
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const pageDigits = Math.max(3, String(options.pageCount).length)
  const extension = options.format === 'jpg' ? 'jpg' : options.format
  const archive = new JSZip()
  for (const image of images) {
    archive.file(
      `page_${String(image.pageNumber).padStart(pageDigits, '0')}.${extension}`,
      image.bytes,
      { compression: 'STORE' },
    )
  }
  return {
    suffix: '_converted.cbz',
    fileName: `${baseName}_converted.cbz`,
    bytes: await archive.generateAsync({ type: 'uint8array', compression: 'STORE' }),
    mimeType: 'application/vnd.comicbook+zip',
    extension: '.cbz',
  }
}

function escapeHtml(value: string): string {
  return value.replace(
    /[&<>"']/g,
    (character) =>
      ({
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
      })[character]!,
  )
}

function checkedPdfHtmlPages(options: PdfToHtmlOptions): PdfHtmlPage[] {
  if (!Number.isInteger(options.pageCount) || options.pageCount < 1) {
    throw new Error('PDF page count must be a positive integer')
  }
  const pageIndexes = checkedPageIndexes(options.pageCount, options.pageIndexes)
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 300) {
    throw new Error('PDF HTML rendering DPI must be a whole number from 72 to 300')
  }
  if (!options.pages) throw new Error('Rendered PDF HTML pages are required')
  if (options.pages.length !== pageIndexes.length) {
    throw new Error('Rendered PDF HTML pages do not match the requested pages')
  }
  return options.pages.map((page, index) => {
    if (page.pageNumber !== pageIndexes[index]! + 1) {
      throw new Error('Rendered PDF HTML page order is invalid')
    }
    if (
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      !Number.isFinite(page.height) ||
      page.height <= 0
    ) {
      throw new Error('Rendered PDF HTML page size is invalid')
    }
    if (!(page.imageBytes instanceof Uint8Array) || page.imageBytes.length === 0) {
      throw new Error('Rendered PDF HTML page image is empty')
    }
    if (typeof page.text !== 'string' || !Array.isArray(page.textRuns)) {
      throw new Error('Rendered PDF HTML text is invalid')
    }
    for (const run of page.textRuns) {
      if (
        typeof run.text !== 'string' ||
        ![run.x, run.y, run.fontSize, run.angle].every(Number.isFinite) ||
        run.fontSize <= 0
      ) {
        throw new Error('Rendered PDF HTML text position is invalid')
      }
    }
    return page
  })
}

function pdfHtmlDocument(title: string, pages: PdfHtmlPage[], pageDigits: number): string {
  const pageMarkup = pages
    .map((page) => {
      const imageName = `pages/page-${String(page.pageNumber).padStart(pageDigits, '0')}.png`
      const textLayer = page.textRuns
        .filter((run) => run.text.length > 0)
        .map((run) => {
          const weight = run.bold ? ' font-weight="700"' : ''
          const style = run.italic ? ' font-style="italic"' : ''
          const family = run.fontFamily ? ` font-family="${escapeHtml(run.fontFamily)}"` : ''
          const rotation = run.angle ? ` transform="rotate(${run.angle} ${run.x} ${run.y})"` : ''
          return `<text x="${run.x}" y="${run.y}" font-size="${run.fontSize}"${family}${weight}${style}${rotation}>${escapeHtml(run.text)}</text>`
        })
        .join('')
      return `<section class="pdf-page" aria-label="Page ${page.pageNumber}">
<img src="${imageName}" alt="Page ${page.pageNumber}" width="${page.width}" height="${page.height}">
<svg class="text-layer" viewBox="0 0 ${page.width} ${page.height}" preserveAspectRatio="none" aria-hidden="true">${textLayer}</svg>
<p class="sr-only">${escapeHtml(page.text)}</p>
</section>`
    })
    .join('\n')

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta http-equiv="Content-Security-Policy" content="default-src 'none'; img-src 'self' data:; style-src 'unsafe-inline'">
<link rel="icon" href="data:,">
<title>${escapeHtml(title)}</title>
<style>
*{box-sizing:border-box}html{background:#e9edf1}body{margin:0;color:#18212b;font-family:Arial,sans-serif}header{position:sticky;top:0;z-index:2;padding:12px 20px;border-bottom:1px solid #d7dde3;background:rgba(255,255,255,.94);backdrop-filter:blur(10px)}header strong{font-size:14px}header span{margin-left:10px;color:#66717d;font-size:12px}main{display:grid;gap:24px;padding:24px}.pdf-page{position:relative;width:min(100%,900px);margin:0 auto;background:#fff;box-shadow:0 8px 26px rgba(20,30,40,.16)}.pdf-page img{display:block;width:100%;height:auto}.text-layer{position:absolute;inset:0;width:100%;height:100%;overflow:hidden}.text-layer text{fill:transparent;stroke:none;user-select:text}.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);white-space:pre-wrap;border:0}@media(max-width:600px){header{padding:10px 12px}main{gap:14px;padding:12px}}@media print{html{background:#fff}header{display:none}main{display:block;padding:0}.pdf-page{break-after:page;width:100%;box-shadow:none}}
</style>
</head>
<body>
<header><strong>${escapeHtml(title)}</strong><span>${pages.length} page${pages.length === 1 ? '' : 's'}</span></header>
<main>${pageMarkup}</main>
</body>
</html>`
}

export async function pdfToHtmlZipBytes(options: PdfToHtmlOptions): Promise<PdfToolOutput> {
  const pages = checkedPdfHtmlPages(options)
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const pageDigits = Math.max(1, String(options.pageCount).length)
  const archive = new JSZip()
  archive.file('index.html', pdfHtmlDocument(baseName, pages, pageDigits))
  for (const page of pages) {
    archive.file(
      `pages/page-${String(page.pageNumber).padStart(pageDigits, '0')}.png`,
      page.imageBytes,
      { compression: 'STORE' },
    )
  }
  return {
    suffix: 'ToHtml.zip',
    fileName: `${baseName}ToHtml.zip`,
    bytes: await archive.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
    mimeType: 'application/zip',
    extension: '.zip',
  }
}

function epubText(value: string): string {
  return escapeHtml(value.replace(/(?![\t\n\r])\p{Cc}/gu, ''))
}

function epubExternalUrl(value: string): string | undefined {
  const normalized = value.trim()
  if (!/^(?:https?:|mailto:)/i.test(normalized)) return undefined
  try {
    return new URL(normalized).toString()
  } catch {
    return undefined
  }
}

function epubPageLabel(page: PdfEpubPage): string {
  return (
    page.blocks.find((block) => block.kind === 'heading')?.text.trim() || `Page ${page.pageNumber}`
  )
}

function epubReflowableBody(page: PdfEpubPage): string {
  const content = page.blocks
    .map((block) => {
      const text = epubText(block.text)
      if (block.kind === 'heading') {
        const level = Math.max(1, Math.min(3, block.level ?? 2))
        return `<h${level}>${text}</h${level}>`
      }
      if (block.kind === 'listItem') {
        return `<ul><li>${epubText(block.text.replace(/^[•▪◦]\s*/, ''))}</li></ul>`
      }
      return `<p>${text}</p>`
    })
    .join('\n')
  const links = page.links
    .flatMap((link) => {
      const url = epubExternalUrl(link.url)
      if (!url) return []
      const label = link.label?.trim() || url
      return [`<li><a href="${epubText(url)}">${epubText(label)}</a></li>`]
    })
    .join('\n')
  return `${content}${links ? `\n<section class="links"><h2>Links</h2><ul>${links}</ul></section>` : ''}`
}

function epubPageDocument(page: PdfEpubPage, mode: PdfEpubMode, imageName?: string): string {
  const title = epubPageLabel(page)
  const body =
    mode === 'fixed'
      ? `<div class="fixed-page"><img src="../images/${imageName}" alt="${epubText(title)}"/></div>`
      : epubReflowableBody(page)
  const viewport =
    mode === 'fixed'
      ? `<meta name="viewport" content="width=${Math.round(page.width)},height=${Math.round(page.height)}"/>`
      : '<meta name="viewport" content="width=device-width,initial-scale=1"/>'
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head><title>${epubText(title)}</title>${viewport}<link rel="stylesheet" type="text/css" href="../styles/book.css"/></head>
<body${mode === 'fixed' ? ' class="fixed"' : ''}><section epub:type="chapter"><span epub:type="pagebreak" id="page-${page.pageNumber}" title="${page.pageNumber}"/>${body}</section></body>
</html>`
}

function epubNavigation(title: string, pages: PdfEpubPage[], pageDigits: number): string {
  const items = pages
    .map(
      (page) =>
        `<li><a href="text/page-${String(page.pageNumber).padStart(pageDigits, '0')}.xhtml">${epubText(epubPageLabel(page))}</a></li>`,
    )
    .join('')
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops">
<head><title>${epubText(title)}</title><meta name="viewport" content="width=device-width,initial-scale=1"/></head>
<body><nav epub:type="toc" id="toc"><h1>${epubText(title)}</h1><ol>${items}</ol></nav></body>
</html>`
}

function epubStyles(mode: PdfEpubMode): string {
  if (mode === 'fixed') {
    return 'html,body,section,.fixed-page{width:100%;height:100%;margin:0;padding:0;overflow:hidden}.fixed-page img{display:block;width:100%;height:100%;object-fit:contain}'
  }
  return 'html{color:#18212b;background:#fff}body{max-width:42em;margin:0 auto;padding:1.4em;font-family:serif;line-height:1.65}h1,h2,h3{line-height:1.25;margin:1.4em 0 .6em}p{margin:.8em 0}ul{margin:.5em 0;padding-left:1.4em}.links{margin-top:2em;padding-top:1em;border-top:1px solid #ccd3d9}.links a{overflow-wrap:anywhere}'
}

function epubModifiedDate(value?: string): string {
  const parsed = value ? new Date(value) : new Date()
  const date = Number.isNaN(parsed.getTime()) ? new Date() : parsed
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

function checkedPdfEpubPages(options: PdfToEpubOptions): PdfEpubPage[] {
  if (!Number.isInteger(options.pageCount) || options.pageCount < 1) {
    throw new Error('PDF page count must be a positive integer')
  }
  if (!['reflowable', 'fixed'].includes(options.mode)) throw new Error('PDF EPUB mode is invalid')
  const pageIndexes = checkedPageIndexes(options.pageCount, options.pageIndexes)
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 300) {
    throw new Error('PDF EPUB rendering DPI must be a whole number from 72 to 300')
  }
  if (!options.pages) throw new Error('Prepared PDF EPUB pages are required')
  if (options.pages.length !== pageIndexes.length) {
    throw new Error('Prepared PDF EPUB pages do not match the requested pages')
  }
  const pages = options.pages.map((page, index) => {
    if (page.pageNumber !== pageIndexes[index]! + 1) {
      throw new Error('Prepared PDF EPUB page order is invalid')
    }
    if (
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      typeof page.text !== 'string' ||
      !Array.isArray(page.blocks) ||
      !Array.isArray(page.links)
    ) {
      throw new Error('Prepared PDF EPUB page data is invalid')
    }
    if (
      options.mode === 'fixed' &&
      (!(page.imageBytes instanceof Uint8Array) || page.imageBytes.length === 0)
    ) {
      throw new Error('Fixed-layout EPUB page image is empty')
    }
    return page
  })
  if (options.mode === 'reflowable' && pages.every((page) => !page.text.trim())) {
    throw new Error('No extractable text was found; use fixed layout or run OCR first')
  }
  return pages
}

export async function pdfToEpubBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfToEpubOptions,
): Promise<PdfToolOutput> {
  const analysis = await analyzePdfBytes(bytes)
  if (analysis.pageCount !== options.pageCount) {
    throw new Error('PDF page count changed while preparing EPUB')
  }
  const pages = checkedPdfEpubPages(options)
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const title = analysis.properties.title?.trim() || baseName
  const author = analysis.properties.author?.trim()
  const pageDigits = Math.max(1, String(options.pageCount).length)
  const language = /[\u3400-\u9fff]/u.test(pages.map((page) => page.text).join('')) ? 'zh-CN' : 'en'
  const archive = new JSZip()
  archive.file('mimetype', 'application/epub+zip', { compression: 'STORE' })
  archive.file(
    'META-INF/container.xml',
    '<?xml version="1.0" encoding="UTF-8"?>\n<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="EPUB/package.opf" media-type="application/oebps-package+xml"/></rootfiles></container>',
  )
  archive.file('EPUB/nav.xhtml', epubNavigation(title, pages, pageDigits))
  archive.file('EPUB/styles/book.css', epubStyles(options.mode))

  const manifest: string[] = [
    '<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>',
    '<item id="style" href="styles/book.css" media-type="text/css"/>',
  ]
  const spine: string[] = []
  for (const page of pages) {
    const number = String(page.pageNumber).padStart(pageDigits, '0')
    const pageId = `page-${number}`
    const imageName = `${pageId}.png`
    archive.file(
      `EPUB/text/${pageId}.xhtml`,
      epubPageDocument(page, options.mode, options.mode === 'fixed' ? imageName : undefined),
    )
    manifest.push(
      `<item id="${pageId}" href="text/${pageId}.xhtml" media-type="application/xhtml+xml"/>`,
    )
    spine.push(`<itemref idref="${pageId}"/>`)
    if (options.mode === 'fixed') {
      archive.file(`EPUB/images/${imageName}`, page.imageBytes!, { compression: 'STORE' })
      manifest.push(
        `<item id="image-${number}" href="images/${imageName}" media-type="image/png"${page === pages[0] ? ' properties="cover-image"' : ''}/>`,
      )
    }
  }

  const modified = epubModifiedDate(
    analysis.properties.modificationDate ?? analysis.properties.creationDate,
  )
  archive.file(
    'EPUB/package.opf',
    `<?xml version="1.0" encoding="UTF-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="book-id" prefix="rendition: http://www.idpf.org/vocab/rendition/#">
<metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:identifier id="book-id">urn:genoffice:${encodeURIComponent(baseName)}</dc:identifier><dc:title>${epubText(title)}</dc:title><dc:language>${language}</dc:language>${author ? `<dc:creator>${epubText(author)}</dc:creator>` : ''}<meta property="dcterms:modified">${modified}</meta>${options.mode === 'fixed' ? '<meta property="rendition:layout">pre-paginated</meta><meta property="rendition:orientation">auto</meta><meta property="rendition:spread">none</meta>' : '<meta property="rendition:layout">reflowable</meta>'}</metadata>
<manifest>${manifest.join('')}</manifest><spine>${spine.join('')}</spine></package>`,
  )
  return {
    suffix: '_converted.epub',
    fileName: `${baseName}_converted.epub`,
    bytes: await archive.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
    mimeType: 'application/epub+zip',
    extension: '.epub',
  }
}

function checkedPdfPptxPages(options: PdfToPptxOptions): PdfPptxPage[] {
  if (!Number.isInteger(options.pageCount) || options.pageCount < 1) {
    throw new Error('PDF page count must be a positive integer')
  }
  if (!['editableText', 'fidelity'].includes(options.mode)) {
    throw new Error('PDF PowerPoint mode is invalid')
  }
  const pageIndexes = checkedPageIndexes(options.pageCount, options.pageIndexes)
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 300) {
    throw new Error('PDF PowerPoint rendering DPI must be a whole number from 72 to 300')
  }
  if (!options.pages) throw new Error('Prepared PDF PowerPoint pages are required')
  if (options.pages.length !== pageIndexes.length) {
    throw new Error('Prepared PDF PowerPoint pages do not match the requested pages')
  }
  let totalRuns = 0
  const pages = options.pages.map((page, index) => {
    if (page.pageNumber !== pageIndexes[index]! + 1) {
      throw new Error('Prepared PDF PowerPoint page order is invalid')
    }
    if (
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      !Array.isArray(page.textRuns)
    ) {
      throw new Error('Prepared PDF PowerPoint page data is invalid')
    }
    totalRuns += page.textRuns.length
    if (totalRuns > 100_000) throw new Error('PDF PowerPoint contains too many text elements')
    for (const run of page.textRuns) {
      if (
        typeof run.text !== 'string' ||
        run.text.length > 1_000_000 ||
        ![run.x, run.y, run.width, run.height, run.fontSize, run.angle].every(Number.isFinite) ||
        run.width < 0 ||
        run.height <= 0 ||
        run.fontSize <= 0
      ) {
        throw new Error('Prepared PDF PowerPoint text element is invalid')
      }
    }
    if (
      options.mode === 'fidelity' &&
      (!(page.imageBytes instanceof Uint8Array) || page.imageBytes.length === 0)
    ) {
      throw new Error('Fidelity PowerPoint page image is empty')
    }
    return page
  })
  if (
    options.mode === 'editableText' &&
    pages.every((page) => page.textRuns.every((run) => !run.text.trim()))
  ) {
    throw new Error('No editable text was found; use page fidelity or run OCR first')
  }
  return pages
}

function pptxPagePlacement(page: PdfPptxPage): {
  x: number
  y: number
  w: number
  h: number
  scale: number
} {
  const slideWidth = 13.333333
  const slideHeight = 7.5
  const scale = Math.min((slideWidth * 72) / page.width, (slideHeight * 72) / page.height)
  const w = (page.width * scale) / 72
  const h = (page.height * scale) / 72
  return { x: (slideWidth - w) / 2, y: (slideHeight - h) / 2, w, h, scale }
}

function pptxBytes(output: ArrayBuffer | Uint8Array | Blob): Promise<Uint8Array> | Uint8Array {
  if (output instanceof Uint8Array) return output
  if (output instanceof ArrayBuffer) return new Uint8Array(output)
  return output.arrayBuffer().then((bytes) => new Uint8Array(bytes))
}

export async function pdfToPptxBytes(options: PdfToPptxOptions): Promise<PdfToolOutput> {
  const pages = checkedPdfPptxPages(options)
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const presentation = new PptxGenJS()
  presentation.layout = 'LAYOUT_WIDE'
  presentation.author = 'GenOffice'
  presentation.company = 'GenOffice'
  presentation.subject =
    options.mode === 'editableText'
      ? 'PDF text converted to editable PowerPoint elements'
      : 'PDF pages converted to PowerPoint slides'
  presentation.title = baseName

  for (const page of pages) {
    const slide = presentation.addSlide()
    slide.background = { color: 'FFFFFF' }
    const placement = pptxPagePlacement(page)
    if (options.mode === 'fidelity') {
      slide.addImage({
        data: `image/png;base64,${Buffer.from(page.imageBytes!).toString('base64')}`,
        x: placement.x,
        y: placement.y,
        w: placement.w,
        h: placement.h,
        objectName: `PDF Page ${page.pageNumber}`,
        altText: `PDF page ${page.pageNumber}`,
      })
    } else {
      for (let index = 0; index < page.textRuns.length; index++) {
        const run = page.textRuns[index]!
        if (!run.text.trim()) continue
        slide.addText(run.text, {
          x: placement.x + (run.x * placement.scale) / 72,
          y: placement.y + (run.y * placement.scale) / 72,
          w: (Math.max(run.width, run.fontSize * 0.5) * placement.scale) / 72,
          h: (Math.max(run.height, run.fontSize * 1.15) * placement.scale) / 72,
          objectName: `PDF Text ${page.pageNumber}.${index + 1}`,
          fontSize: Math.max(1, run.fontSize * placement.scale),
          color: '111111',
          bold: run.bold,
          italic: run.italic,
          rotate: ((run.angle % 360) + 360) % 360,
          margin: 0,
          breakLine: false,
          fit: 'shrink',
          isTextBox: true,
          line: { type: 'none' },
          ...(run.fontFamily?.trim() ? { fontFace: run.fontFamily.trim() } : {}),
        })
      }
    }
  }

  const output = await presentation.write({ outputType: 'arraybuffer', compression: true })
  return {
    suffix: '_converted.pptx',
    fileName: `${baseName}_converted.pptx`,
    bytes: await pptxBytes(output as ArrayBuffer | Uint8Array | Blob),
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    extension: '.pptx',
  }
}

interface PdfDocxLine {
  x: number
  y: number
  runs: PdfDocxTextRun[]
}

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
const DOCX_MARGIN_TWIPS = 720
const DOCX_PORTRAIT_WIDTH_TWIPS = 11_906
const DOCX_PORTRAIT_HEIGHT_TWIPS = 16_838

function checkedPreparedPdfTextPages(
  options: Pick<PdfToRtfOptions, 'pageCount' | 'pageIndexes' | 'pages'>,
  format: string,
): PdfDocxPage[] {
  if (!Number.isInteger(options.pageCount) || options.pageCount < 1) {
    throw new Error('PDF page count must be a positive integer')
  }
  const pageIndexes = checkedPageIndexes(options.pageCount, options.pageIndexes)
  if (!options.pages) throw new Error(`Prepared PDF ${format} pages are required`)
  if (options.pages.length !== pageIndexes.length) {
    throw new Error(`Prepared PDF ${format} pages do not match the requested pages`)
  }
  let totalRuns = 0
  const pages = options.pages.map((page, index) => {
    if (page.pageNumber !== pageIndexes[index]! + 1) {
      throw new Error(`Prepared PDF ${format} page order is invalid`)
    }
    if (
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      !Array.isArray(page.textRuns)
    ) {
      throw new Error(`Prepared PDF ${format} page data is invalid`)
    }
    totalRuns += page.textRuns.length
    if (totalRuns > 100_000) throw new Error(`PDF ${format} contains too many text elements`)
    for (const run of page.textRuns) {
      if (
        typeof run.text !== 'string' ||
        run.text.length > 1_000_000 ||
        ![run.x, run.y, run.width, run.height, run.fontSize, run.angle].every(Number.isFinite) ||
        run.width < 0 ||
        run.height <= 0 ||
        run.fontSize <= 0
      ) {
        throw new Error(`Prepared PDF ${format} text element is invalid`)
      }
    }
    return page
  })
  return pages
}

function checkedPdfDocxPages(options: PdfToDocxOptions): PdfDocxPage[] {
  if (!['editableText', 'fidelity'].includes(options.mode)) {
    throw new Error('PDF Word mode is invalid')
  }
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 300) {
    throw new Error('PDF Word rendering DPI must be a whole number from 72 to 300')
  }
  const pages = checkedPreparedPdfTextPages(options, 'Word')
  for (const page of pages) {
    if (
      options.mode === 'fidelity' &&
      (!(page.imageBytes instanceof Uint8Array) ||
        page.imageBytes.length === 0 ||
        !Number.isFinite(page.imageWidth) ||
        page.imageWidth! <= 0 ||
        !Number.isFinite(page.imageHeight) ||
        page.imageHeight! <= 0)
    ) {
      throw new Error('Fidelity Word page image is empty or has invalid dimensions')
    }
  }
  if (
    options.mode === 'editableText' &&
    pages.every((page) => page.textRuns.every((run) => !run.text.trim()))
  ) {
    throw new Error('No editable text was found; use page fidelity or run OCR first')
  }
  return pages
}

function pdfDocxSection(page: PdfDocxPage): SectionSettings {
  const landscape = page.width > page.height
  return {
    pageWidth: landscape ? DOCX_PORTRAIT_HEIGHT_TWIPS : DOCX_PORTRAIT_WIDTH_TWIPS,
    pageHeight: landscape ? DOCX_PORTRAIT_WIDTH_TWIPS : DOCX_PORTRAIT_HEIGHT_TWIPS,
    orientation: landscape ? 'landscape' : 'portrait',
    marginTop: DOCX_MARGIN_TWIPS,
    marginRight: DOCX_MARGIN_TWIPS,
    marginBottom: DOCX_MARGIN_TWIPS,
    marginLeft: DOCX_MARGIN_TWIPS,
    pageBorder: false,
    columns: 1,
    headerDist: 360,
    footerDist: 360,
  }
}

function pdfDocxLines(page: PdfDocxPage): PdfDocxLine[] {
  const horizontal = page.textRuns
    .filter((run) => run.text.trim())
    .sort((left, right) => left.y - right.y || left.x - right.x)
  const rows: PdfDocxLine[] = []
  for (const run of horizontal) {
    let row: PdfDocxLine | undefined
    for (let index = rows.length - 1; index >= 0; index--) {
      const candidate = rows[index]!
      if (Math.abs(candidate.y - run.y) <= Math.max(2.5, run.fontSize * 0.45)) {
        row = candidate
        break
      }
    }
    if (row) {
      row.runs.push(run)
      row.x = Math.min(row.x, run.x)
      row.y = (row.y * (row.runs.length - 1) + run.y) / row.runs.length
    } else {
      rows.push({ x: run.x, y: run.y, runs: [run] })
    }
  }

  const lines: PdfDocxLine[] = []
  for (const row of rows) {
    const runs = row.runs.sort((left, right) => left.x - right.x)
    let current: PdfDocxLine | undefined
    for (const run of runs) {
      const previous = current?.runs.at(-1)
      const gap = previous ? run.x - (previous.x + previous.width) : 0
      if (!current || gap > Math.max(36, page.width * 0.12) || previous?.hasEol) {
        current = { x: run.x, y: row.y, runs: [run] }
        lines.push(current)
      } else {
        current.runs.push(run)
      }
    }
  }
  return lines.sort((left, right) => left.y - right.y || left.x - right.x)
}

function pdfDocxNeedsSpace(previous: PdfDocxTextRun, current: PdfDocxTextRun): boolean {
  if (/\s$/u.test(previous.text) || /^\s/u.test(current.text)) return false
  if (/[\u3400-\u9fff]$/u.test(previous.text) && /^[\u3400-\u9fff]/u.test(current.text))
    return false
  if (/^[,.;:!?%)}\]，。；：！？、）》】]/u.test(current.text)) return false
  if (/[({[（《【]$/u.test(previous.text)) return false
  return current.x - (previous.x + previous.width) > Math.max(1.5, previous.fontSize * 0.12)
}

function pdfDocxRun(run: PdfDocxTextRun, text = run.text): Run {
  const font = run.fontFamily?.trim().slice(0, 80)
  return {
    text,
    bold: run.bold,
    italic: run.italic,
    sizeHalfPoints: Math.min(400, Math.max(2, Math.round(run.fontSize * 2))),
    ...(font ? { font, fontAscii: font } : {}),
  }
}

function pdfDocxTextBlocks(pages: PdfDocxPage[]): SaveBlock[] {
  const blocks: SaveBlock[] = []
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    const page = pages[pageIndex]!
    const lines = pdfDocxLines(page)
    const minimumX = lines.length > 0 ? Math.min(...lines.map((line) => line.x)) : 0
    if (lines.length === 0 && pageIndex > 0) {
      blocks.push({
        kind: 'generated',
        block: { type: 'paragraph', runs: [], format: { pageBreakBefore: true } },
      })
      continue
    }
    for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
      const line = lines[lineIndex]!
      const runs: Run[] = []
      for (let runIndex = 0; runIndex < line.runs.length; runIndex++) {
        const run = line.runs[runIndex]!
        const previous = line.runs[runIndex - 1]
        if (previous && pdfDocxNeedsSpace(previous, run)) runs.push(pdfDocxRun(previous, ' '))
        runs.push(pdfDocxRun(run))
      }
      blocks.push({
        kind: 'generated',
        block: {
          type: 'paragraph',
          runs,
          format: {
            ...(pageIndex > 0 && lineIndex === 0 ? { pageBreakBefore: true } : {}),
            indentLeft: Math.max(
              0,
              Math.round(Math.min(line.x - minimumX, page.width * 0.75) * 20),
            ),
            spaceAfter: 0,
          },
        },
      })
    }
  }
  return blocks
}

function pdfDocxImageBlocks(pages: PdfDocxPage[], section: SectionSettings): SaveBlock[] {
  const maximumWidth = (section.pageWidth - section.marginLeft - section.marginRight) / 15
  const maximumHeight = (section.pageHeight - section.marginTop - section.marginBottom) / 15
  const blocks: SaveBlock[] = []
  for (let index = 0; index < pages.length; index++) {
    const page = pages[index]!
    if (index > 0) {
      blocks.push({
        kind: 'generated',
        block: { type: 'paragraph', runs: [], format: { pageBreakBefore: true } },
      })
    }
    const scale = Math.min(1, maximumWidth / page.imageWidth!, maximumHeight / page.imageHeight!)
    blocks.push({
      kind: 'image',
      image: {
        base64: Buffer.from(page.imageBytes!).toString('base64'),
        mime: 'image/png',
        widthPx: Math.max(1, Math.round(page.imageWidth! * scale)),
        heightPx: Math.max(1, Math.round(page.imageHeight! * scale)),
        align: 'center',
      },
    })
  }
  return blocks
}

export async function pdfToDocxBytes(options: PdfToDocxOptions): Promise<PdfToolOutput> {
  const pages = checkedPdfDocxPages(options)
  const section = pdfDocxSection(pages[0]!)
  const parsed = await parseDocx(await buildBlankDocx())
  const blocks =
    options.mode === 'editableText' ? pdfDocxTextBlocks(pages) : pdfDocxImageBlocks(pages, section)
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  return {
    suffix: '_converted.docx',
    fileName: `${baseName}_converted.docx`,
    bytes: await saveDocx(parsed, blocks, { section }),
    mimeType: DOCX_MIME,
    extension: '.docx',
  }
}

const ODT_MIME = 'application/vnd.oasis.opendocument.text'

interface OdtStyleRegistry {
  fonts: Map<string, string>
  paragraphs: Map<string, string>
  text: Map<string, string>
}

function checkedPdfOdtPages(options: PdfToOdtOptions): PdfDocxPage[] {
  if (!['editableText', 'fidelity'].includes(options.mode)) {
    throw new Error('PDF OpenDocument mode is invalid')
  }
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 300) {
    throw new Error('PDF OpenDocument rendering DPI must be a whole number from 72 to 300')
  }
  const pages = checkedPreparedPdfTextPages(options, 'OpenDocument')
  for (const page of pages) {
    if (
      options.mode === 'fidelity' &&
      (!(page.imageBytes instanceof Uint8Array) ||
        page.imageBytes.length === 0 ||
        !Number.isFinite(page.imageWidth) ||
        page.imageWidth! <= 0 ||
        !Number.isFinite(page.imageHeight) ||
        page.imageHeight! <= 0)
    ) {
      throw new Error('Fidelity OpenDocument page image is empty or has invalid dimensions')
    }
  }
  if (
    options.mode === 'editableText' &&
    pages.every((page) => page.textRuns.every((run) => !run.text.trim()))
  ) {
    throw new Error('No editable text was found; use page fidelity or run OCR first')
  }
  return pages
}

function odtLength(points: number): string {
  return `${Math.max(0, points / 72).toFixed(4)}in`
}

function odtTextContent(value: string): string {
  const output: string[] = []
  let offset = 0
  for (const match of value.matchAll(/\r\n|\r|\n|\t| +/gu)) {
    output.push(xmlSafeValue(value.slice(offset, match.index)))
    if (match[0] === '\t') output.push('<text:tab/>')
    else if (match[0][0] === '\r' || match[0][0] === '\n') output.push('<text:line-break/>')
    else output.push(match[0].length === 1 ? '<text:s/>' : `<text:s text:c="${match[0].length}"/>`)
    offset = match.index + match[0].length
  }
  output.push(xmlSafeValue(value.slice(offset)))
  return output.join('')
}

function odtParagraphStyle(
  registry: OdtStyleRegistry,
  indentPoints: number,
  pageBreak: boolean,
): string {
  const roundedIndent = Math.max(0, Math.round(indentPoints * 10) / 10)
  const key = `${roundedIndent}:${pageBreak ? 1 : 0}`
  const existing = registry.paragraphs.get(key)
  if (existing) return existing
  const name = `P${registry.paragraphs.size + 1}`
  registry.paragraphs.set(key, name)
  return name
}

function odtTextStyle(registry: OdtStyleRegistry, run: PdfDocxTextRun): string {
  const font = run.fontFamily?.trim().slice(0, 80) || 'Liberation Sans'
  if (!registry.fonts.has(font)) registry.fonts.set(font, `Font${registry.fonts.size + 1}`)
  const key = JSON.stringify([font, Math.round(run.fontSize * 10) / 10, run.bold, run.italic])
  const existing = registry.text.get(key)
  if (existing) return existing
  const name = `T${registry.text.size + 1}`
  registry.text.set(key, name)
  return name
}

function odtEditableBody(pages: PdfDocxPage[], registry: OdtStyleRegistry): string {
  const output: string[] = []
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    const page = pages[pageIndex]!
    const lines = pdfDocxLines(page)
    const minimumX = lines.length > 0 ? Math.min(...lines.map((line) => line.x)) : 0
    if (lines.length === 0) {
      output.push(`<text:p text:style-name="${odtParagraphStyle(registry, 0, pageIndex > 0)}"/>`)
      continue
    }
    for (let lineIndex = 0; lineIndex < lines.length; lineIndex++) {
      const line = lines[lineIndex]!
      const paragraphStyle = odtParagraphStyle(
        registry,
        Math.min(Math.max(0, line.x - minimumX), page.width * 0.75),
        pageIndex > 0 && lineIndex === 0,
      )
      const spans: string[] = []
      for (let runIndex = 0; runIndex < line.runs.length; runIndex++) {
        const run = line.runs[runIndex]!
        const previous = line.runs[runIndex - 1]
        if (previous && pdfDocxNeedsSpace(previous, run)) spans.push('<text:s/>')
        spans.push(
          `<text:span text:style-name="${odtTextStyle(registry, run)}">${odtTextContent(run.text)}</text:span>`,
        )
      }
      output.push(`<text:p text:style-name="${paragraphStyle}">${spans.join('')}</text:p>`)
    }
  }
  return output.join('')
}

function odtFidelityBody(
  pages: PdfDocxPage[],
  registry: OdtStyleRegistry,
): { body: string; images: { path: string; bytes: Uint8Array }[] } {
  const body: string[] = []
  const images: { path: string; bytes: Uint8Array }[] = []
  for (let index = 0; index < pages.length; index++) {
    const page = pages[index]!
    const path = `Pictures/page-${index + 1}.png`
    const maximumWidth = Math.max(72, page.width - 72)
    const maximumHeight = Math.max(72, page.height - 72)
    const scale = Math.min(1, maximumWidth / page.imageWidth!, maximumHeight / page.imageHeight!)
    body.push(
      `<text:p text:style-name="${odtParagraphStyle(registry, 0, index > 0)}"><draw:frame draw:style-name="PageImage" draw:name="Page ${page.pageNumber}" text:anchor-type="paragraph" svg:width="${odtLength(page.imageWidth! * scale)}" svg:height="${odtLength(page.imageHeight! * scale)}"><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p>`,
    )
    images.push({ path, bytes: page.imageBytes! })
  }
  return { body: body.join(''), images }
}

function odtAutomaticStyles(registry: OdtStyleRegistry): string {
  const paragraphStyles = [...registry.paragraphs.entries()].map(([key, name]) => {
    const [indent, pageBreak] = key.split(':')
    return `<style:style style:name="${name}" style:family="paragraph" style:parent-style-name="Standard"><style:paragraph-properties fo:margin-left="${odtLength(Number(indent))}" fo:margin-top="0in" fo:margin-bottom="0in"${pageBreak === '1' ? ' fo:break-before="page"' : ''}/></style:style>`
  })
  const textStyles = [...registry.text.entries()].map(([key, name]) => {
    const [font, size, bold, italic] = JSON.parse(key) as [string, number, boolean, boolean]
    const fontName = registry.fonts.get(font)!
    return `<style:style style:name="${name}" style:family="text"><style:text-properties style:font-name="${fontName}" style:font-name-asian="${fontName}" fo:font-size="${size.toFixed(1)}pt" style:font-size-asian="${size.toFixed(1)}pt"${bold ? ' fo:font-weight="bold" style:font-weight-asian="bold"' : ''}${italic ? ' fo:font-style="italic" style:font-style-asian="italic"' : ''}/></style:style>`
  })
  return `${paragraphStyles.join('')}${textStyles.join('')}<style:style style:name="PageImage" style:family="graphic" style:parent-style-name="Graphics"><style:graphic-properties style:horizontal-pos="center" style:horizontal-rel="paragraph" style:vertical-pos="top" style:vertical-rel="paragraph" draw:stroke="none" draw:fill="none"/></style:style>`
}

function odtContentXml(body: string, registry: OdtStyleRegistry): string {
  const fontFaces = [...registry.fonts.entries()]
    .map(
      ([font, name]) =>
        `<style:font-face style:name="${name}" svg:font-family="${xmlSafeValue(font)}"/>`,
    )
    .join('')
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.3"><office:font-face-decls>${fontFaces}</office:font-face-decls><office:automatic-styles>${odtAutomaticStyles(registry)}</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`
}

function odtStylesXml(page: PdfDocxPage): string {
  const orientation = page.width > page.height ? 'landscape' : 'portrait'
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" office:version="1.3"><office:styles><style:default-style style:family="paragraph"><style:paragraph-properties fo:margin-top="0in" fo:margin-bottom="0in"/><style:text-properties fo:font-size="11pt"/></style:default-style><style:style style:name="Standard" style:family="paragraph" style:class="text"/><style:style style:name="Graphics" style:family="graphic"/></office:styles><office:automatic-styles><style:page-layout style:name="PageLayout"><style:page-layout-properties fo:page-width="${odtLength(page.width)}" fo:page-height="${odtLength(page.height)}" style:print-orientation="${orientation}" fo:margin="0.5in"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PageLayout"/></office:master-styles></office:document-styles>`
}

function odtManifestXml(images: { path: string }[]): string {
  const imageEntries = images
    .map(
      ({ path }) =>
        `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="image/png"/>`,
    )
    .join('')
  return `<?xml version="1.0" encoding="UTF-8"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODT_MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/>${imageEntries}</manifest:manifest>`
}

export async function pdfToOdtBytes(options: PdfToOdtOptions): Promise<PdfToolOutput> {
  const pages = checkedPdfOdtPages(options)
  const registry: OdtStyleRegistry = {
    fonts: new Map(),
    paragraphs: new Map(),
    text: new Map(),
  }
  const fidelity =
    options.mode === 'fidelity'
      ? odtFidelityBody(pages, registry)
      : { body: odtEditableBody(pages, registry), images: [] }
  const archive = new JSZip()
  archive.file('mimetype', ODT_MIME, { compression: 'STORE' })
  archive.file('content.xml', odtContentXml(fidelity.body, registry))
  archive.file('styles.xml', odtStylesXml(pages[0]!))
  archive.file(
    'meta.xml',
    '<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.3"><office:meta><meta:generator>GenOffice</meta:generator></office:meta></office:document-meta>',
  )
  archive.file('META-INF/manifest.xml', odtManifestXml(fidelity.images))
  for (const image of fidelity.images) archive.file(image.path, image.bytes)
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  return {
    suffix: '_converted.odt',
    fileName: `${baseName}_converted.odt`,
    bytes: await archive.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      platform: 'UNIX',
    }),
    mimeType: ODT_MIME,
    extension: '.odt',
  }
}

const RTF_MIME = 'application/rtf'

function pdfRtfEscape(value: string): string {
  let result = ''
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    if (code === 0x5c || code === 0x7b || code === 0x7d) {
      result += `\\${value[index]}`
    } else if (code === 0x09) {
      result += '\\tab '
    } else if (code === 0x0a) {
      result += '\\line '
    } else if (code === 0x0d) {
      if (value.charCodeAt(index + 1) !== 0x0a) result += '\\line '
    } else if (code >= 0x20 && code <= 0x7e) {
      result += value[index]
    } else if (code >= 0x20) {
      result += `\\u${code > 0x7fff ? code - 0x1_0000 : code}?`
    }
  }
  return result
}

function pdfRtfFontName(value: string | undefined): string {
  const cleaned = value
    ?.trim()
    .replace(/[\\{};]/g, ' ')
    .replace(/\s+/g, ' ')
    .slice(0, 80)
  return cleaned || 'Arial'
}

function pdfRtfBytes(pages: PdfDocxPage[]): Uint8Array {
  const fonts = new Map<string, number>()
  for (const page of pages) {
    for (const run of page.textRuns) {
      const font = pdfRtfFontName(run.fontFamily)
      if (!fonts.has(font)) fonts.set(font, fonts.size)
    }
  }
  if (fonts.size === 0) fonts.set('Arial', 0)
  const fontTable = Array.from(
    fonts,
    ([font, index]) => `{\\f${index}\\fnil ${pdfRtfEscape(font)};}`,
  ).join('')
  const body: string[] = []
  for (let pageIndex = 0; pageIndex < pages.length; pageIndex++) {
    if (pageIndex > 0) body.push('\\page\n')
    const page = pages[pageIndex]!
    const lines = pdfDocxLines(page)
    const minimumX = lines.length > 0 ? Math.min(...lines.map((line) => line.x)) : 0
    for (const line of lines) {
      const indent = Math.max(0, Math.round(Math.min(line.x - minimumX, page.width * 0.75) * 20))
      body.push(`\\pard\\ql\\li${indent}\\sa0\\sb0 `)
      for (let runIndex = 0; runIndex < line.runs.length; runIndex++) {
        const run = line.runs[runIndex]!
        const previous = line.runs[runIndex - 1]
        if (previous && pdfDocxNeedsSpace(previous, run)) body.push(' ')
        const font = fonts.get(pdfRtfFontName(run.fontFamily)) ?? 0
        const size = Math.min(400, Math.max(2, Math.round(run.fontSize * 2)))
        body.push(
          `{\\f${font}\\fs${size}${run.bold ? '\\b' : ''}${run.italic ? '\\i' : ''} ${pdfRtfEscape(run.text)}}`,
        )
      }
      body.push('\\par\n')
    }
  }
  return new TextEncoder().encode(
    `{\\rtf1\\ansi\\ansicpg1252\\uc1\\deff0{\\fonttbl${fontTable}}\n${body.join('')}}`,
  )
}

export function pdfToRtfBytes(options: PdfToRtfOptions): PdfToolOutput {
  const pages = checkedPreparedPdfTextPages(options, 'RTF')
  if (pages.every((page) => page.textRuns.every((run) => !run.text.trim()))) {
    throw new Error('No editable text was found; run OCR first')
  }
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  return {
    suffix: '_converted.rtf',
    fileName: `${baseName}_converted.rtf`,
    bytes: pdfRtfBytes(pages),
    mimeType: RTF_MIME,
    extension: '.rtf',
  }
}

function namedDestinationFromTree(
  document: PDFDocument,
  node: PDFDict,
  name: string,
  visited = new Set<PDFDict>(),
): PDFObject | undefined {
  if (visited.has(node)) return undefined
  visited.add(node)
  const names = node.lookupMaybe(PDFName.of('Names'), PDFArray)
  if (names) {
    for (let index = 0; index + 1 < names.size(); index += 2) {
      if (decodePdfText(names.get(index)) === name) return names.get(index + 1)
    }
  }
  const kids = node.lookupMaybe(PDFName.of('Kids'), PDFArray)
  if (!kids) return undefined
  for (let index = 0; index < kids.size(); index++) {
    const child = kids.lookupMaybe(index, PDFDict)
    if (!child) continue
    const found = namedDestinationFromTree(document, child, name, visited)
    if (found) return found
  }
  return undefined
}

function namedBookmarkDestination(document: PDFDocument, name: string): PDFObject | undefined {
  const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  const destinations = names?.lookupMaybe(PDFName.of('Dests'), PDFDict)
  const fromTree = destinations ? namedDestinationFromTree(document, destinations, name) : undefined
  if (fromTree) return fromTree
  const legacy = document.catalog.lookupMaybe(PDFName.of('Dests'), PDFDict)
  if (!legacy) return undefined
  for (const [key, value] of legacy.entries()) {
    if (key.decodeText() === name) return value
  }
  return undefined
}

function bookmarkDestinationArray(
  document: PDFDocument,
  object: PDFObject | undefined,
  visitedNames = new Set<string>(),
): PDFArray | undefined {
  const resolved = document.context.lookup(object)
  if (resolved instanceof PDFArray) return resolved
  if (resolved instanceof PDFDict) return resolved.lookupMaybe(PDFName.of('D'), PDFArray)
  const name = decodePdfText(resolved)
  if (!name || visitedNames.has(name)) return undefined
  visitedNames.add(name)
  return bookmarkDestinationArray(document, namedBookmarkDestination(document, name), visitedNames)
}

function bookmarkPageNumber(document: PDFDocument, item: PDFDict): number | undefined {
  let destination = bookmarkDestinationArray(document, item.get(PDFName.of('Dest')))
  if (!destination) {
    const action = item.lookupMaybe(PDFName.of('A'), PDFDict)
    const actionType = action?.lookupMaybe(PDFName.of('S'), PDFName)?.decodeText()
    if (!actionType || actionType === 'GoTo') {
      destination = bookmarkDestinationArray(document, action?.get(PDFName.of('D')))
    }
  }
  if (!destination || destination.size() === 0) return undefined
  const pageObject = destination.get(0)
  const resolvedPage = document.context.lookup(pageObject)
  const pageIndex = document
    .getPages()
    .findIndex((page) => samePdfObject(pageObject, page.ref) || resolvedPage === page.node)
  return pageIndex < 0 ? undefined : pageIndex + 1
}

function readBookmarkLevel(
  document: PDFDocument,
  firstEntry: PDFObject | undefined,
  visited: Set<string>,
  counter: { count: number },
  depth: number,
  internalOnly = false,
): PdfBookmark[] {
  if (depth > 20) return []
  const bookmarks: PdfBookmark[] = []
  let current = firstEntry
  while (current && counter.count < 1000) {
    const key = current.toString()
    if (visited.has(key)) break
    visited.add(key)
    const item = document.context.lookupMaybe(current, PDFDict)
    if (!item) break
    counter.count++
    const children = readBookmarkLevel(
      document,
      item.get(PDFName.of('First')),
      visited,
      counter,
      depth + 1,
      internalOnly,
    )
    const pageNumber = bookmarkPageNumber(document, item)
    if (pageNumber !== undefined || !internalOnly) {
      bookmarks.push({
        title: decodePdfText(item.get(PDFName.of('Title'))) ?? '',
        pageNumber: pageNumber ?? 1,
        children,
      })
    } else {
      bookmarks.push(...children)
    }
    current = item.get(PDFName.of('Next'))
  }
  return bookmarks
}

export async function listPdfBookmarksBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfBookmark[]> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const outlines = document.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (!outlines) return []
  return readBookmarkLevel(document, outlines.get(PDFName.of('First')), new Set(), { count: 0 }, 0)
}

async function listInternalPdfBookmarksBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfBookmark[]> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const outlines = document.catalog.lookupMaybe(PDFName.of('Outlines'), PDFDict)
  if (!outlines) return []
  return readBookmarkLevel(
    document,
    outlines.get(PDFName.of('First')),
    new Set(),
    { count: 0 },
    0,
    true,
  )
}

function normalizedBookmarks(
  bookmarks: PdfBookmark[],
  pageCount: number,
  counter = { count: 0 },
  depth = 0,
): PdfBookmark[] {
  if (!Array.isArray(bookmarks)) throw new Error('bookmarks must be an array')
  if (depth > 20) throw new Error('bookmark nesting must not exceed 20 levels')
  return bookmarks.map((bookmark) => {
    counter.count++
    if (counter.count > 1000) throw new Error('bookmarks must not exceed 1000 items')
    const title = bookmark.title?.trim()
    if (!title) throw new Error('bookmark titles must not be empty')
    const requestedPage = Number.isFinite(bookmark.pageNumber) ? Math.trunc(bookmark.pageNumber) : 1
    return {
      title,
      pageNumber: Math.min(Math.max(requestedPage, 1), pageCount),
      children: normalizedBookmarks(bookmark.children ?? [], pageCount, counter, depth + 1),
    }
  })
}

interface OutlineLevelResult {
  first?: PDFRef
  last?: PDFRef
  count: number
}

function writeBookmarkLevel(
  document: PDFDocument,
  bookmarks: PdfBookmark[],
  parent: PDFRef,
): OutlineLevelResult {
  const items = bookmarks.map((bookmark) => {
    const dictionary = document.context.obj({})
    dictionary.set(PDFName.of('Title'), PDFHexString.fromText(bookmark.title))
    dictionary.set(PDFName.of('Parent'), parent)
    dictionary.set(
      PDFName.of('Dest'),
      document.context.obj([document.getPage(bookmark.pageNumber - 1).ref, PDFName.of('Fit')]),
    )
    return { bookmark, dictionary, ref: document.context.register(dictionary) }
  })
  let count = items.length
  for (let index = 0; index < items.length; index++) {
    const item = items[index]!
    const previous = items[index - 1]
    const next = items[index + 1]
    if (previous) item.dictionary.set(PDFName.of('Prev'), previous.ref)
    if (next) item.dictionary.set(PDFName.of('Next'), next.ref)
    const children = writeBookmarkLevel(document, item.bookmark.children, item.ref)
    count += children.count
    if (children.first && children.last) {
      item.dictionary.set(PDFName.of('First'), children.first)
      item.dictionary.set(PDFName.of('Last'), children.last)
      item.dictionary.set(PDFName.of('Count'), PDFNumber.of(children.count))
    }
  }
  return { first: items[0]?.ref, last: items.at(-1)?.ref, count }
}

export async function setPdfBookmarksBytes(
  bytes: Uint8Array | ArrayBuffer,
  bookmarks: PdfBookmark[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = document.getPageCount()
  if (pageCount === 0) throw new Error('PDF must contain at least one page')
  const normalized = normalizedBookmarks(bookmarks, pageCount)
  if (normalized.length === 0) {
    document.catalog.delete(PDFName.of('Outlines'))
    if (
      document.catalog.lookupMaybe(PDFName.of('PageMode'), PDFName)?.decodeText() === 'UseOutlines'
    ) {
      document.catalog.delete(PDFName.of('PageMode'))
    }
    return document.save({ useObjectStreams: false })
  }
  const root = document.context.obj({})
  root.set(PDFName.of('Type'), PDFName.of('Outlines'))
  const rootRef = document.context.register(root)
  const level = writeBookmarkLevel(document, normalized, rootRef)
  root.set(PDFName.of('First'), level.first!)
  root.set(PDFName.of('Last'), level.last!)
  root.set(PDFName.of('Count'), PDFNumber.of(level.count))
  document.catalog.set(PDFName.of('Outlines'), rootRef)
  document.catalog.set(PDFName.of('PageMode'), PDFName.of('UseOutlines'))
  return document.save({ useObjectStreams: false })
}

function pdfHeaderVersion(bytes: Uint8Array): string | undefined {
  let header = ''
  for (let index = 0; index < Math.min(bytes.length, 32); index++) {
    header += String.fromCharCode(bytes[index]!)
  }
  return header.match(/%PDF-(\d\.\d)/)?.[1]
}

function isoDate(value: Date | undefined): string | undefined {
  if (!value) return undefined
  try {
    return value.toISOString()
  } catch {
    return undefined
  }
}

const STANDARD_METADATA_KEYS = new Set([
  'title',
  'author',
  'subject',
  'keywords',
  'creator',
  'producer',
  'creationdate',
  'moddate',
  'trapped',
])

function pdfInfoDictionary(document: PDFDocument, create: boolean): PDFDict | undefined {
  const existing = document.context.lookup(document.context.trailerInfo.Info)
  if (existing instanceof PDFDict) return existing
  if (!create) return undefined
  const dictionary = document.context.obj({})
  document.context.trailerInfo.Info = document.context.register(dictionary)
  return dictionary
}

function pdfMetadataCustomFields(document: PDFDocument): PdfMetadataCustomField[] {
  const info = pdfInfoDictionary(document, false)
  if (!info) return []
  return info
    .keys()
    .filter(
      (key) =>
        !STANDARD_METADATA_KEYS.has(key.decodeText().toLowerCase()) &&
        key.decodeText().toLowerCase() !== PDF_CLASSIFICATION_METADATA_KEY_NORMALIZED,
    )
    .map((key) => ({
      key: key.decodeText(),
      value: decodePdfText(info.get(key)) ?? info.get(key)?.toString() ?? '',
    }))
    .sort((left, right) => left.key.localeCompare(right.key))
}

const PDF_CLASSIFICATION_SENSITIVITIES = new Set<PdfClassificationSensitivity>([
  'standard',
  'internal',
  'confidential',
  'restricted',
])
const PDF_CLASSIFICATION_METADATA_KEY_NORMALIZED = PDF_CLASSIFICATION_METADATA_KEY.toLowerCase()

function checkedPdfClassificationMetadata(
  metadata: PdfClassificationMetadata,
): PdfClassificationMetadata {
  if (!PDF_CLASSIFICATION_SENSITIVITIES.has(metadata.sensitivity)) {
    throw new Error('PDF classification sensitivity is invalid')
  }
  if (!Array.isArray(metadata.labels) || metadata.labels.length > 5) {
    throw new Error('PDF classification may contain at most 5 labels')
  }
  const ids = new Set<string>()
  const labels = metadata.labels.map((label, index) => {
    const id = typeof label?.id === 'string' ? label.id.trim() : ''
    const name = typeof label?.name === 'string' ? label.name.trim() : ''
    if (!id || id.length > 64 || !/^[a-z0-9]+(?:-[a-z0-9]+)*$/.test(id)) {
      throw new Error(`PDF classification label ${index + 1} has an invalid id`)
    }
    if (!name || name.length > 128 || /\p{Cc}/u.test(name)) {
      throw new Error(`PDF classification label ${index + 1} has an invalid name`)
    }
    if (ids.has(id)) throw new Error(`PDF classification label ${id} is duplicated`)
    ids.add(id)
    return { id, name }
  })
  return { labels, sensitivity: metadata.sensitivity }
}

function parsePdfClassificationMetadata(
  value: string | undefined,
): PdfClassificationMetadata | null {
  if (!value) return null
  try {
    const parsed = JSON.parse(value) as Partial<PdfClassificationMetadata> & { version?: unknown }
    if (
      parsed.version !== 1 ||
      !Array.isArray(parsed.labels) ||
      typeof parsed.sensitivity !== 'string'
    ) {
      return null
    }
    return checkedPdfClassificationMetadata({
      labels: parsed.labels as PdfClassificationLabel[],
      sensitivity: parsed.sensitivity as PdfClassificationSensitivity,
    })
  } catch {
    return null
  }
}

export function applyPdfClassificationMetadata(
  document: PDFDocument,
  metadata: PdfClassificationMetadata,
): void {
  const checked = checkedPdfClassificationMetadata(metadata)
  const info = pdfInfoDictionary(document, true)!
  info.set(
    PDFName.of(PDF_CLASSIFICATION_METADATA_KEY),
    PDFHexString.fromText(JSON.stringify({ version: 1, ...checked })),
  )
  document.setModificationDate(new Date())
}

export function readPdfClassificationMetadata(
  document: PDFDocument,
): PdfClassificationMetadata | null {
  const info = pdfInfoDictionary(document, false)
  return parsePdfClassificationMetadata(
    info ? decodePdfText(info.get(PDFName.of(PDF_CLASSIFICATION_METADATA_KEY))) : undefined,
  )
}

export async function readPdfClassificationMetadataBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfClassificationMetadata | null> {
  try {
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    return readPdfClassificationMetadata(document)
  } catch {
    return null
  }
}

interface PdfResourceAnalysis {
  fonts: Set<string>
  images: Set<PDFRawStream>
}

interface MutablePdfFontInfo extends Omit<PdfFontInfo, 'pages'> {
  pages: Set<number>
}

function pdfFontDescriptor(document: PDFDocument, font: PDFDict): PDFDict | undefined {
  const direct = font.lookupMaybe(PDFName.of('FontDescriptor'), PDFDict)
  if (direct) return direct
  const descendants = font.lookupMaybe(PDFName.of('DescendantFonts'), PDFArray)
  return descendants?.lookupMaybe(0, PDFDict)?.lookupMaybe(PDFName.of('FontDescriptor'), PDFDict)
}

function pdfFontName(document: PDFDocument, font: PDFDict, resourceName: PDFName): string {
  const descendant = font
    .lookupMaybe(PDFName.of('DescendantFonts'), PDFArray)
    ?.lookupMaybe(0, PDFDict)
  return (
    decodePdfText(font.get(PDFName.of('BaseFont'))) ??
    decodePdfText(descendant?.get(PDFName.of('BaseFont'))) ??
    resourceName.decodeText() ??
    'Unknown'
  )
}

function pdfFontEncoding(font: PDFDict): string | undefined {
  const encoding = font.get(PDFName.of('Encoding'))
  if (encoding instanceof PDFName) return encoding.decodeText()
  const dictionary = font.lookupMaybe(PDFName.of('Encoding'), PDFDict)
  return dictionary?.lookupMaybe(PDFName.of('BaseEncoding'), PDFName)?.decodeText()
}

function pdfFontEmbedded(document: PDFDocument, font: PDFDict, subtype: string): boolean {
  if (subtype === 'Type3') return true
  const descriptor = pdfFontDescriptor(document, font)
  return Boolean(
    descriptor?.has(PDFName.of('FontFile')) ||
    descriptor?.has(PDFName.of('FontFile2')) ||
    descriptor?.has(PDFName.of('FontFile3')),
  )
}

function collectPdfFonts(
  document: PDFDocument,
  resources: PDFDict | undefined,
  pageNumber: number,
  fonts: Map<string, MutablePdfFontInfo>,
  visited = new Set<PDFDict>(),
): void {
  if (!resources || visited.has(resources)) return
  visited.add(resources)
  const fontResources = resources.lookupMaybe(PDFName.of('Font'), PDFDict)
  if (fontResources) {
    for (const resourceName of fontResources.keys()) {
      const font = document.context.lookupMaybe(fontResources.get(resourceName), PDFDict)
      if (!font) continue
      const rawName = pdfFontName(document, font, resourceName)
      const subset = /^[A-Z]{6}\+/.test(rawName)
      const name = rawName.replace(/^[A-Z]{6}\+/, '')
      const subtype = font.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() ?? 'Unknown'
      const encoding = pdfFontEncoding(font)
      const embedded = pdfFontEmbedded(document, font, subtype)
      const hasToUnicode = font.has(PDFName.of('ToUnicode'))
      const key = [name, subtype, embedded, subset, encoding ?? '', hasToUnicode].join('\u0000')
      const existing = fonts.get(key)
      if (existing) {
        existing.pages.add(pageNumber)
      } else {
        fonts.set(key, {
          name,
          subtype,
          embedded,
          subset,
          ...(encoding ? { encoding } : {}),
          hasToUnicode,
          pages: new Set([pageNumber]),
        })
      }
    }
  }
  const xObjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xObjects) return
  for (const resourceName of xObjects.keys()) {
    const stream = document.context.lookup(xObjects.get(resourceName))
    if (!(stream instanceof PDFRawStream)) continue
    if (stream.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() !== 'Form') continue
    collectPdfFonts(
      document,
      stream.dict.lookupMaybe(PDFName.of('Resources'), PDFDict),
      pageNumber,
      fonts,
      visited,
    )
  }
}

function analyzePdfFonts(document: PDFDocument): PdfFontReport {
  const fonts = new Map<string, MutablePdfFontInfo>()
  document.getPages().forEach((page, pageIndex) => {
    collectPdfFonts(document, page.node.Resources(), pageIndex + 1, fonts)
  })
  const items = [...fonts.values()]
    .map((font): PdfFontInfo => ({ ...font, pages: [...font.pages].sort((a, b) => a - b) }))
    .sort(
      (left, right) =>
        left.name.localeCompare(right.name) || left.subtype.localeCompare(right.subtype),
    )
  return {
    fontCount: items.length,
    embeddedCount: items.filter((font) => font.embedded).length,
    subsetCount: items.filter((font) => font.subset).length,
    fonts: items,
  }
}

export async function analyzePdfFontsBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfFontReport> {
  return analyzePdfFonts(await PDFDocument.load(bytes, { updateMetadata: false }))
}

function pdfaPreservationReport(document: PDFDocument): PdfaPreservationReport {
  const fonts = analyzePdfFonts(document)
  const unembeddedFonts = fonts.fonts
    .filter((font) => !font.embedded)
    .map((font) => font.name)
    .filter((name, index, names) => names.indexOf(name) === index)
  return {
    eligible: unembeddedFonts.length === 0,
    fontCount: fonts.fontCount,
    embeddedFontCount: fonts.embeddedCount,
    unembeddedFonts,
  }
}

export async function pdfaPreservationReportBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfaPreservationReport> {
  return pdfaPreservationReport(await PDFDocument.load(bytes, { updateMetadata: false }))
}

function pdfAnnotationDate(
  document: PDFDocument,
  value: PDFObject | undefined,
): string | undefined {
  const resolved = resolvedPdfObject(document, value)
  if (!(resolved instanceof PDFString || resolved instanceof PDFHexString)) return undefined
  try {
    return isoDate(resolved.decodeDate())
  } catch {
    return undefined
  }
}

function pdfAnnotationRectangle(annotation: PDFDict): PdfAnnotationInfo['rectangle'] {
  const rectangle = annotation.lookupMaybe(PDFName.of('Rect'), PDFArray)
  if (!rectangle || rectangle.size() < 4) return undefined
  const values = Array.from({ length: 4 }, (_, index) =>
    rectangle.lookupMaybe(index, PDFNumber)?.asNumber(),
  )
  if (values.some((value) => value === undefined || !Number.isFinite(value))) return undefined
  const [left, bottom, right, top] = values as [number, number, number, number]
  return {
    x: Math.min(left, right),
    y: Math.min(bottom, top),
    width: Math.abs(right - left),
    height: Math.abs(top - bottom),
  }
}

function analyzePdfAnnotations(document: PDFDocument): PdfAnnotationReport {
  const typeBreakdown: Record<string, number> = {}
  const annotations: PdfAnnotationInfo[] = []
  document.getPages().forEach((page, pageIndex) => {
    const pageAnnotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!pageAnnotations) return
    for (let index = 0; index < pageAnnotations.size(); index++) {
      const annotation = pageAnnotations.lookupMaybe(index, PDFDict)
      if (!annotation) continue
      const subtype =
        annotation.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText() ?? 'Unknown'
      typeBreakdown[subtype] = (typeBreakdown[subtype] ?? 0) + 1
      const author = decodePdfText(resolvedPdfObject(document, annotation.get(PDFName.of('T'))))
      const subject = decodePdfText(resolvedPdfObject(document, annotation.get(PDFName.of('Subj'))))
      const contents = decodePdfText(
        resolvedPdfObject(document, annotation.get(PDFName.of('Contents'))),
      )
      const name = decodePdfText(resolvedPdfObject(document, annotation.get(PDFName.of('NM'))))
      const flags = annotation.lookupMaybe(PDFName.of('F'), PDFNumber)?.asNumber()
      const rectangle = pdfAnnotationRectangle(annotation)
      const modifiedAt = pdfAnnotationDate(document, annotation.get(PDFName.of('M')))
      annotations.push({
        pageNumber: pageIndex + 1,
        annotationNumber: index + 1,
        subtype,
        ...(author ? { author } : {}),
        ...(subject ? { subject } : {}),
        ...(contents ? { contents } : {}),
        ...(modifiedAt ? { modifiedAt } : {}),
        ...(name ? { name } : {}),
        ...(flags !== undefined ? { flags } : {}),
        ...(rectangle ? { rectangle } : {}),
      })
    }
  })
  return { totalCount: annotations.length, typeBreakdown, annotations }
}

export async function analyzePdfAnnotationsBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfAnnotationReport> {
  return analyzePdfAnnotations(await PDFDocument.load(bytes, { updateMetadata: false }))
}

function analyzePdfResources(
  document: PDFDocument,
  resources: PDFDict | undefined,
  analysis: PdfResourceAnalysis,
  visited = new Set<PDFDict>(),
): void {
  if (!resources || visited.has(resources)) return
  visited.add(resources)
  const fonts = resources.lookupMaybe(PDFName.of('Font'), PDFDict)
  if (fonts) {
    for (const resourceName of fonts.keys()) {
      const font = document.context.lookupMaybe(fonts.get(resourceName), PDFDict)
      const name =
        decodePdfText(font?.get(PDFName.of('BaseFont'))) ?? resourceName.decodeText() ?? 'Unknown'
      analysis.fonts.add(name.replace(/^[A-Z]{6}\+/, ''))
    }
  }
  const xObjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xObjects) return
  for (const resourceName of xObjects.keys()) {
    const stream = document.context.lookup(xObjects.get(resourceName))
    if (!(stream instanceof PDFRawStream)) continue
    const subtype = stream.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
    if (subtype === 'Image') analysis.images.add(stream)
    if (subtype === 'Form') {
      analyzePdfResources(
        document,
        stream.dict.lookupMaybe(PDFName.of('Resources'), PDFDict),
        analysis,
        visited,
      )
    }
  }
}

function bookmarkCount(bookmarks: PdfBookmark[]): number {
  return bookmarks.reduce((total, bookmark) => total + 1 + bookmarkCount(bookmark.children), 0)
}

function analyzeFormFieldDictionary(
  document: PDFDocument,
  dictionary: PDFDict,
  inheritedType: string | undefined,
  visited: Set<PDFDict>,
): { fieldCount: number; signatureCount: number } {
  if (visited.has(dictionary)) return { fieldCount: 0, signatureCount: 0 }
  visited.add(dictionary)
  const fieldType = dictionary.lookupMaybe(PDFName.of('FT'), PDFName)?.decodeText() ?? inheritedType
  const kids = dictionary.lookupMaybe(PDFName.of('Kids'), PDFArray)
  const childFields: PDFDict[] = []
  if (kids) {
    for (let index = 0; index < kids.size(); index++) {
      const child = kids.lookupMaybe(index, PDFDict)
      if (!child) continue
      const subtype = child.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
      if (subtype !== 'Widget') childFields.push(child)
    }
  }
  if (childFields.length === 0) {
    return { fieldCount: 1, signatureCount: fieldType === 'Sig' ? 1 : 0 }
  }
  return childFields.reduce(
    (total, child) => {
      const result = analyzeFormFieldDictionary(document, child, fieldType, visited)
      total.fieldCount += result.fieldCount
      total.signatureCount += result.signatureCount
      return total
    },
    { fieldCount: 0, signatureCount: 0 },
  )
}

function analyzeFormFields(
  document: PDFDocument,
  acroForm: PDFDict | undefined,
): { fieldCount: number; signatureCount: number } {
  const fields = acroForm?.lookupMaybe(PDFName.of('Fields'), PDFArray)
  if (!fields) return { fieldCount: 0, signatureCount: 0 }
  const visited = new Set<PDFDict>()
  const result = { fieldCount: 0, signatureCount: 0 }
  for (let index = 0; index < fields.size(); index++) {
    const field = fields.lookupMaybe(index, PDFDict)
    if (!field) continue
    const fieldResult = analyzeFormFieldDictionary(document, field, undefined, visited)
    result.fieldCount += fieldResult.fieldCount
    result.signatureCount += fieldResult.signatureCount
  }
  return result
}

export async function analyzePdfBytes(bytes: Uint8Array | ArrayBuffer): Promise<PdfAnalysis> {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const document = await PDFDocument.load(input, { updateMetadata: false })
  const acroForm = document.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const formFields = analyzeFormFields(document, acroForm)

  const annotationReport = analyzePdfAnnotations(document)
  const resources: PdfResourceAnalysis = { fonts: new Set(), images: new Set() }
  const visitedResources = new Set<PDFDict>()
  const pages = document.getPages().map((page, pageIndex) => {
    analyzePdfResources(document, page.node.Resources(), resources, visitedResources)
    const box = page.getCropBox()
    return {
      pageNumber: pageIndex + 1,
      width: box.width,
      height: box.height,
      rotation: page.getRotation().angle,
    }
  })
  const bookmarks = await listPdfBookmarksBytes(input)
  const attachments = await listPdfAttachmentsBytes(input)

  return {
    pageCount: document.getPageCount(),
    pdfVersion: pdfHeaderVersion(input),
    fileSize: input.byteLength,
    isEncrypted: document.isEncrypted,
    properties: {
      title: document.getTitle(),
      author: document.getAuthor(),
      subject: document.getSubject(),
      keywords: document.getKeywords(),
      creator: document.getCreator(),
      producer: document.getProducer(),
      creationDate: isoDate(document.getCreationDate()),
      modificationDate: isoDate(document.getModificationDate()),
      trapped: (() => {
        const trapped = pdfInfoDictionary(document, false)
          ?.lookupMaybe(PDFName.of('Trapped'), PDFName)
          ?.decodeText()
        return trapped === 'True' || trapped === 'False' || trapped === 'Unknown'
          ? trapped
          : undefined
      })(),
      custom: pdfMetadataCustomFields(document),
    },
    pages,
    form: {
      fieldCount: formFields.fieldCount,
      hasXfa: acroForm?.has(PDFName.of('XFA')) ?? false,
      signatureCount: formFields.signatureCount,
    },
    annotations: {
      totalCount: annotationReport.totalCount,
      typeBreakdown: annotationReport.typeBreakdown,
    },
    fonts: [...resources.fonts].sort((left, right) => left.localeCompare(right)),
    imageCount: resources.images.size,
    attachmentCount: attachments.length,
    bookmarkCount: bookmarkCount(bookmarks),
  }
}

function lastAsciiIndex(bytes: Uint8Array, value: string): number {
  const pattern = new TextEncoder().encode(value)
  outer: for (let index = bytes.length - pattern.length; index >= 0; index--) {
    for (let offset = 0; offset < pattern.length; offset++) {
      if (bytes[index + offset] !== pattern[offset]) continue outer
    }
    return index
  }
  return -1
}

function asciiOccurrences(bytes: Uint8Array, value: string): number[] {
  const pattern = new TextEncoder().encode(value)
  const indexes: number[] = []
  outer: for (let index = 0; index <= bytes.length - pattern.length; index++) {
    for (let offset = 0; offset < pattern.length; offset++) {
      if (bytes[index + offset] !== pattern[offset]) continue outer
    }
    indexes.push(index)
  }
  return indexes
}

function asciiSlice(bytes: Uint8Array, start: number, length: number): string {
  let result = ''
  for (let index = start; index < Math.min(bytes.length, start + length); index++) {
    result += String.fromCharCode(bytes[index]!)
  }
  return result
}

function validStartXrefTarget(bytes: Uint8Array, offset: number | undefined): boolean {
  if (offset === undefined || offset < 0 || offset >= bytes.byteLength) return false
  const target = asciiSlice(bytes, offset, 512).trimStart()
  if (/^xref\b/.test(target)) return true
  return /^\d+\s+\d+\s+obj\b/.test(target) && /\/Type\s*\/XRef\b/.test(target)
}

function decodedMetadataStream(document: PDFDocument): string | undefined {
  const metadata = document.context.lookup(document.catalog.get(PDFName.of('Metadata')))
  if (!(metadata instanceof PDFRawStream)) return undefined
  try {
    return new TextDecoder().decode(decodePDFRawStream(metadata).decode())
  } catch {
    return undefined
  }
}

function xmlScalarValues(
  value: unknown,
  localName: string,
  output: string[],
  prefix?: string,
): void {
  if (Array.isArray(value)) {
    for (const item of value) xmlScalarValues(item, localName, output, prefix)
    return
  }
  if (!value || typeof value !== 'object') return
  for (const [key, child] of Object.entries(value)) {
    const normalized = key.startsWith('@_') ? key.slice(2) : key
    const [keyPrefix, keyName] = normalized.includes(':')
      ? (normalized.split(':', 2) as [string, string])
      : ['', normalized]
    if (
      keyName.toLowerCase() === localName.toLowerCase() &&
      (!prefix || keyPrefix.toLowerCase() === prefix.toLowerCase())
    ) {
      if (typeof child === 'string' || typeof child === 'number') output.push(String(child))
    }
    xmlScalarValues(child, localName, output, prefix)
  }
}

function xmpScalar(value: unknown, prefix: string, localName: string): string | undefined {
  const values: string[] = []
  xmlScalarValues(value, localName, values, prefix)
  return values.map((item) => item.trim()).find(Boolean)
}

function standardDeclarations(xmp: string | undefined): {
  declarations: PdfStandardDeclaration[]
  valid: boolean
} {
  if (!xmp) return { declarations: [], valid: false }
  const validation = XMLValidator.validate(xmp)
  if (validation !== true) return { declarations: [], valid: false }
  const parsed = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    parseAttributeValue: false,
    parseTagValue: false,
  }).parse(xmp) as unknown
  const declarations: PdfStandardDeclaration[] = []
  const pdfaPart = xmpScalar(parsed, 'pdfaid', 'part')
  const pdfaConformance = xmpScalar(parsed, 'pdfaid', 'conformance')?.toUpperCase()
  if (/^\d+$/.test(pdfaPart ?? '') && /pdfaid:/i.test(xmp)) {
    declarations.push({
      family: 'PDF/A',
      part: pdfaPart!,
      ...(pdfaConformance ? { conformance: pdfaConformance } : {}),
      label: `PDF/A-${pdfaPart}${pdfaConformance ? pdfaConformance.toLowerCase() : ''}`,
    })
  }
  const pdfuaPart = xmpScalar(parsed, 'pdfuaid', 'part')
  const pdfuaRevision = xmpScalar(parsed, 'pdfuaid', 'rev')
  if (/^\d+$/.test(pdfuaPart ?? '') && /pdfuaid:/i.test(xmp)) {
    declarations.push({
      family: 'PDF/UA',
      part: pdfuaPart!,
      ...(pdfuaRevision ? { revision: pdfuaRevision } : {}),
      label: `PDF/UA-${pdfuaPart}${pdfuaRevision ? `:${pdfuaRevision}` : ''}`,
    })
  }
  return { declarations, valid: true }
}

function preflightStatus(findings: PdfPreflightFinding[]): PdfPreflightReport['status'] {
  if (findings.some((finding) => finding.severity === 'error')) return 'error'
  if (findings.some((finding) => finding.severity === 'warning')) return 'warning'
  return 'pass'
}

export async function preflightPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfPreflightReport> {
  const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const findings: PdfPreflightFinding[] = []
  const headerOffset = asciiSlice(input, 0, 1024).indexOf('%PDF-')
  const eofOffsets = asciiOccurrences(input, '%%EOF')
  const lastEofOffset = eofOffsets.at(-1) ?? -1
  const trailingBytes =
    lastEofOffset < 0 ? input.byteLength : input.byteLength - (lastEofOffset + 5)
  const startXrefMarker = lastAsciiIndex(input, 'startxref')
  const startXrefValue =
    startXrefMarker < 0
      ? undefined
      : Number(/startxref\s+(\d+)/.exec(asciiSlice(input, startXrefMarker, 80))?.[1])
  const startXrefOffset = Number.isSafeInteger(startXrefValue) ? startXrefValue : undefined
  const startXrefInRange =
    startXrefOffset !== undefined && startXrefOffset >= 0 && startXrefOffset < input.byteLength
  const startXrefTargetValid = validStartXrefTarget(input, startXrefOffset)

  if (headerOffset < 0) findings.push({ code: 'missingPdfHeader', severity: 'error' })
  else if (headerOffset > 0) {
    findings.push({ code: 'headerNotAtStart', severity: 'warning', detail: String(headerOffset) })
  }
  if (eofOffsets.length === 0) findings.push({ code: 'missingEofMarker', severity: 'error' })
  else if (trailingBytes > 1024) {
    findings.push({ code: 'largeTrailingData', severity: 'warning', detail: String(trailingBytes) })
  }
  if (startXrefMarker < 0 || startXrefOffset === undefined) {
    findings.push({ code: 'missingStartXref', severity: 'error' })
  } else if (!startXrefInRange) {
    findings.push({
      code: 'startXrefOutOfRange',
      severity: 'error',
      detail: String(startXrefOffset),
    })
  } else if (!startXrefTargetValid) {
    findings.push({
      code: 'startXrefTargetInvalid',
      severity: 'error',
      detail: String(startXrefOffset),
    })
  }

  let document: PDFDocument
  try {
    document = await PDFDocument.load(input, { throwOnInvalidObject: true, updateMetadata: false })
  } catch (loadError) {
    findings.push({
      code: 'strictParseFailed',
      severity: 'error',
      detail: loadError instanceof Error ? loadError.message : String(loadError),
    })
    return {
      schema: 'genoffice.pdf.preflight',
      version: 1,
      status: 'error',
      fileSize: input.byteLength,
      pdfVersion: pdfHeaderVersion(input),
      parseable: false,
      strictParsing: false,
      pageCount: 0,
      structure: {
        headerOffset,
        eofMarkerCount: eofOffsets.length,
        trailingBytes,
        ...(startXrefOffset === undefined ? {} : { startXrefOffset }),
        startXrefInRange,
        startXrefTargetValid,
        incrementalUpdates: Math.max(0, eofOffsets.length - 1),
      },
      standards: [],
      features: {
        hasXmpMetadata: false,
        xmpValid: false,
        tagged: false,
        marked: false,
        outputIntentCount: 0,
        javaScriptActionCount: 0,
        attachmentCount: 0,
        formFieldCount: 0,
        hasXfa: false,
        signatureCount: 0,
        encrypted: false,
      },
      findings,
      disclaimer: 'local-structural-preflight',
    }
  }

  const pageCount = document.getPageCount()
  if (pageCount === 0) findings.push({ code: 'emptyPageTree', severity: 'error' })
  const pages = document.getPages()
  for (let index = 0; index < pages.length; index++) {
    const box = pages[index]!.getCropBox()
    if (
      ![box.x, box.y, box.width, box.height].every(Number.isFinite) ||
      box.width <= 0 ||
      box.height <= 0
    ) {
      findings.push({ code: 'invalidPageBox', severity: 'error', pageNumber: index + 1 })
    }
  }

  const xmp = decodedMetadataStream(document)
  const hasXmpMetadata = document.catalog.has(PDFName.of('Metadata'))
  const declared = standardDeclarations(xmp)
  const structureRoot = document.catalog.lookupMaybe(PDFName.of('StructTreeRoot'), PDFDict)
  const markInfo = document.catalog.lookupMaybe(PDFName.of('MarkInfo'), PDFDict)
  const marked = markInfo?.lookupMaybe(PDFName.of('Marked'), PDFBool)?.asBoolean() ?? false
  const tagged = Boolean(structureRoot) && marked
  const language = decodePdfText(document.catalog.get(PDFName.of('Lang')))?.trim() || undefined
  const outputIntents = document.catalog.lookupMaybe(PDFName.of('OutputIntents'), PDFArray)
  const outputIntentCount = outputIntents?.size() ?? 0
  const acroForm = document.catalog.lookupMaybe(PDFName.of('AcroForm'), PDFDict)
  const form = analyzeFormFields(document, acroForm)
  const attachments = await listPdfAttachmentsBytes(input)
  const javaScriptAudit = await auditPdfJavaScriptBytes(input)

  if (hasXmpMetadata && !declared.valid) {
    findings.push({ code: 'invalidXmpMetadata', severity: 'warning' })
  }
  for (const declaration of declared.declarations) {
    findings.push({ code: 'standardDeclaredOnly', severity: 'info', detail: declaration.label })
    if (declaration.family === 'PDF/A') {
      if (outputIntentCount === 0) {
        findings.push({ code: 'pdfaMissingOutputIntent', severity: 'warning' })
      }
      if (attachments.length > 0 && Number(declaration.part) < 3) {
        findings.push({ code: 'pdfaAttachmentsRisk', severity: 'warning' })
      }
      if (javaScriptAudit.actions.length > 0) {
        findings.push({ code: 'pdfaJavaScriptRisk', severity: 'warning' })
      }
    }
    if (declaration.family === 'PDF/UA') {
      if (!tagged) findings.push({ code: 'pdfuaNotTagged', severity: 'warning' })
      if (!language) findings.push({ code: 'pdfuaMissingLanguage', severity: 'warning' })
    }
  }
  if (javaScriptAudit.actions.length > 0) {
    findings.push({
      code: 'javascriptPresent',
      severity: 'warning',
      detail: String(javaScriptAudit.actions.length),
    })
  }
  if (attachments.length > 0) {
    findings.push({
      code: 'attachmentsPresent',
      severity: 'info',
      detail: String(attachments.length),
    })
  }
  if (acroForm?.has(PDFName.of('XFA'))) findings.push({ code: 'xfaPresent', severity: 'warning' })
  if (form.signatureCount > 0) {
    findings.push({
      code: 'signaturesPresent',
      severity: 'info',
      detail: String(form.signatureCount),
    })
  }

  return {
    schema: 'genoffice.pdf.preflight',
    version: 1,
    status: preflightStatus(findings),
    fileSize: input.byteLength,
    pdfVersion: pdfHeaderVersion(input),
    parseable: true,
    strictParsing: true,
    pageCount,
    structure: {
      headerOffset,
      eofMarkerCount: eofOffsets.length,
      trailingBytes,
      ...(startXrefOffset === undefined ? {} : { startXrefOffset }),
      startXrefInRange,
      startXrefTargetValid,
      incrementalUpdates: Math.max(0, eofOffsets.length - 1),
    },
    standards: declared.declarations,
    features: {
      hasXmpMetadata,
      xmpValid: hasXmpMetadata && declared.valid,
      tagged,
      marked,
      ...(language ? { language } : {}),
      outputIntentCount,
      javaScriptActionCount: javaScriptAudit.actions.length,
      attachmentCount: attachments.length,
      formFieldCount: form.fieldCount,
      hasXfa: acroForm?.has(PDFName.of('XFA')) ?? false,
      signatureCount: form.signatureCount,
      encrypted: document.isEncrypted,
    },
    findings,
    disclaimer: 'local-structural-preflight',
  }
}

function checkedMetadataDate(value: string, name: string): Date | undefined {
  if (!value.trim()) return undefined
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) throw new Error(`${name} is invalid`)
  return date
}

function setInfoText(dictionary: PDFDict, key: string, value: string): void {
  const name = PDFName.of(key)
  if (value.trim()) dictionary.set(name, PDFHexString.fromText(value.trim()))
  else dictionary.delete(name)
}

function checkedCustomMetadata(fields: PdfMetadataCustomField[]): PdfMetadataCustomField[] {
  const keys = new Set<string>()
  return fields.map((field, index) => {
    const key = field.key.trim()
    if (!key || key.length > 128 || /\p{Cc}/u.test(key)) {
      throw new Error(`Custom metadata key ${index + 1} is invalid`)
    }
    const normalized = key.toLowerCase()
    if (STANDARD_METADATA_KEYS.has(normalized)) {
      throw new Error(`Custom metadata key ${key} is reserved`)
    }
    if (normalized === PDF_CLASSIFICATION_METADATA_KEY_NORMALIZED) {
      throw new Error(`Custom metadata key ${key} is reserved`)
    }
    if (keys.has(normalized)) throw new Error(`Custom metadata key ${key} is duplicated`)
    keys.add(normalized)
    return { key, value: field.value }
  })
}

export async function updatePdfMetadataBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfMetadataOptions,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  if (options.deleteAll) {
    document.context.trailerInfo.Info = document.context.register(document.context.obj({}))
    document.catalog.delete(PDFName.of('Metadata'))
    document.catalog.delete(PDFName.of('PieceInfo'))
    return document.save({ useObjectStreams: false })
  }

  const custom = checkedCustomMetadata(options.metadata.custom)
  const creationDate = checkedMetadataDate(options.metadata.creationDate, 'Creation date')
  const modificationDate = checkedMetadataDate(
    options.metadata.modificationDate,
    'Modification date',
  )
  if (!['', 'True', 'False', 'Unknown'].includes(options.metadata.trapped)) {
    throw new Error('Trapped status is invalid')
  }

  const info = pdfInfoDictionary(document, true)!
  for (const key of info.keys()) {
    if (
      !STANDARD_METADATA_KEYS.has(key.decodeText().toLowerCase()) &&
      key.decodeText().toLowerCase() !== PDF_CLASSIFICATION_METADATA_KEY_NORMALIZED
    ) {
      info.delete(key)
    }
  }
  setInfoText(info, 'Title', options.metadata.title)
  setInfoText(info, 'Author', options.metadata.author)
  setInfoText(info, 'Subject', options.metadata.subject)
  setInfoText(info, 'Keywords', options.metadata.keywords)
  setInfoText(info, 'Creator', options.metadata.creator)
  setInfoText(info, 'Producer', options.metadata.producer)
  if (creationDate) info.set(PDFName.of('CreationDate'), PDFString.fromDate(creationDate))
  else info.delete(PDFName.of('CreationDate'))
  if (modificationDate) info.set(PDFName.of('ModDate'), PDFString.fromDate(modificationDate))
  else info.delete(PDFName.of('ModDate'))
  if (options.metadata.trapped) {
    info.set(PDFName.of('Trapped'), PDFName.of(options.metadata.trapped))
  } else {
    info.delete(PDFName.of('Trapped'))
  }
  for (const field of custom) {
    info.set(PDFName.of(field.key), PDFHexString.fromText(field.value))
  }
  return document.save({ useObjectStreams: false })
}

export async function rotatePdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
  angle: PdfPageRotation,
): Promise<Uint8Array> {
  if (![90, 180, 270].includes(angle)) throw new Error('angle must be 90, 180, or 270')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(document.getPageCount(), pageIndexes)
  for (const pageIndex of indexes) {
    const page = document.getPage(pageIndex)
    page.setRotation(degrees((page.getRotation().angle + angle) % 360))
  }
  return document.save({ useObjectStreams: false })
}

export async function autoRotatePdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageRotations: PdfPageRotationCorrection[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const seen = new Set<number>()
  for (const correction of pageRotations) {
    if (
      !Number.isInteger(correction.pageIndex) ||
      correction.pageIndex < 0 ||
      correction.pageIndex >= document.getPageCount()
    ) {
      throw new Error('Auto-rotation page index is out of range')
    }
    if (seen.has(correction.pageIndex)) {
      throw new Error('Auto-rotation page indexes must be unique')
    }
    seen.add(correction.pageIndex)
    if (![90, 180, 270].includes(correction.angle)) {
      throw new Error('Auto-rotation angle must be 90, 180, or 270')
    }
    const page = document.getPage(correction.pageIndex)
    page.setRotation(degrees((page.getRotation().angle + correction.angle) % 360))
  }
  return document.save({ useObjectStreams: false })
}

function removeImagesFromResources(
  document: PDFDocument,
  resources: PDFDict | undefined,
  visited: Set<PDFDict>,
): number {
  if (!resources || visited.has(resources)) return 0
  visited.add(resources)
  const xObjects = resources.lookupMaybe(PDFName.of('XObject'), PDFDict)
  if (!xObjects) return 0
  let removed = 0
  for (const name of [...xObjects.keys()]) {
    const stream = document.context.lookup(xObjects.get(name))
    if (!(stream instanceof PDFRawStream)) continue
    const subtype = stream.dict.lookupMaybe(PDFName.of('Subtype'), PDFName)?.decodeText()
    if (subtype === 'Image') {
      xObjects.delete(name)
      removed++
    } else if (subtype === 'Form') {
      removed += removeImagesFromResources(
        document,
        stream.dict.lookupMaybe(PDFName.of('Resources'), PDFDict),
        visited,
      )
    }
  }
  return removed
}

export async function removePdfImagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(document.getPageCount(), pageIndexes)
  const visited = new Set<PDFDict>()
  for (const pageIndex of indexes) {
    removeImagesFromResources(document, document.getPage(pageIndex).node.Resources(), visited)
  }
  return document.save({ useObjectStreams: false })
}

function copyPdfDict(document: PDFDocument, source: PDFDict | undefined): PDFDict {
  const copy = document.context.obj({})
  if (!source) return copy
  for (const key of source.keys()) {
    const value = source.get(key)
    if (value) copy.set(key, value)
  }
  return copy
}

function pdfNumber(value: number): string {
  return PDFNumber.of(value).toString()
}

function pdfRgbColor(value: string, name: string): readonly [number, number, number] {
  const match = /^#?([\da-f]{2})([\da-f]{2})([\da-f]{2})$/i.exec(value.trim())
  if (!match) throw new Error(`${name} must be a six-digit hex color`)
  return [
    Number.parseInt(match[1]!, 16),
    Number.parseInt(match[2]!, 16),
    Number.parseInt(match[3]!, 16),
  ]
}

function pdfRgbOperator(color: readonly [number, number, number]): string {
  return color.map((component) => pdfNumber(component / 255)).join(' ')
}

export async function invertPdfColorsBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(document.getPageCount(), pageIndexes)
  const graphicsState = document.context.obj({
    Type: PDFName.of('ExtGState'),
    BM: PDFName.of('Difference'),
    ca: PDFNumber.of(1),
    CA: PDFNumber.of(1),
  })
  const graphicsStateRef = document.context.register(graphicsState)

  for (const pageIndex of indexes) {
    const page = document.getPage(pageIndex)
    const resources = copyPdfDict(document, page.node.Resources())
    const extGState = copyPdfDict(document, resources.lookupMaybe(PDFName.of('ExtGState'), PDFDict))
    resources.set(PDFName.of('ExtGState'), extGState)
    page.node.set(PDFName.of('Resources'), resources)

    let resourceIndex = 1
    let resourceName = PDFName.of('GOInvert')
    while (extGState.has(resourceName)) {
      resourceIndex++
      resourceName = PDFName.of(`GOInvert${resourceIndex}`)
    }
    extGState.set(resourceName, graphicsStateRef)

    const cropBox = page.getCropBox()
    const content = [
      'q',
      `${resourceName.toString()} gs`,
      '1 1 1 rg',
      `${pdfNumber(cropBox.x)} ${pdfNumber(cropBox.y)} ${pdfNumber(cropBox.width)} ${pdfNumber(cropBox.height)} re`,
      'f',
      'Q',
    ].join('\n')
    const contentStream = document.context.flateStream(content)
    page.node.addContentStream(document.context.register(contentStream))
  }

  return document.save({ useObjectStreams: false })
}

export async function replacePdfColorsBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
  textColor: string,
  backgroundColor: string,
): Promise<Uint8Array> {
  const textRgb = pdfRgbColor(textColor, 'textColor')
  const backgroundRgb = pdfRgbColor(backgroundColor, 'backgroundColor')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(document.getPageCount(), pageIndexes)

  for (const pageIndex of indexes) {
    const page = document.getPage(pageIndex)
    const cropBox = page.getCropBox()
    const mediaBox = page.getMediaBox()
    if (!page.node.Contents()) {
      page.drawRectangle({ x: cropBox.x, y: cropBox.y, width: 0, height: 0, opacity: 0 })
    }
    const maskPage = await document.embedPage(page, {
      left: mediaBox.x,
      bottom: mediaBox.y,
      right: mediaBox.x + mediaBox.width,
      top: mediaBox.y + mediaBox.height,
    })
    await maskPage.embed()
    const maskStream = document.context.lookup(maskPage.ref)
    if (!(maskStream instanceof PDFRawStream)) throw new Error('Could not create color mask')
    maskStream.dict.set(
      PDFName.of('Group'),
      document.context.obj({
        S: PDFName.of('Transparency'),
        CS: PDFName.of('DeviceRGB'),
        I: PDFBool.True,
      }),
    )

    const softMask = document.context.obj({
      S: PDFName.of('Luminosity'),
      G: maskPage.ref,
      BC: [PDFNumber.of(1), PDFNumber.of(1), PDFNumber.of(1)],
    })
    const graphicsState = document.context.obj({
      Type: PDFName.of('ExtGState'),
      SMask: softMask,
      AIS: PDFBool.False,
    })
    const graphicsStateRef = document.context.register(graphicsState)

    const resources = copyPdfDict(document, page.node.Resources())
    const extGState = copyPdfDict(document, resources.lookupMaybe(PDFName.of('ExtGState'), PDFDict))
    resources.set(PDFName.of('ExtGState'), extGState)
    page.node.set(PDFName.of('Resources'), resources)
    let resourceIndex = 1
    let resourceName = PDFName.of('GORecolor')
    while (extGState.has(resourceName)) {
      resourceIndex++
      resourceName = PDFName.of(`GORecolor${resourceIndex}`)
    }
    extGState.set(resourceName, graphicsStateRef)

    const rectangle = `${pdfNumber(cropBox.x)} ${pdfNumber(cropBox.y)} ${pdfNumber(cropBox.width)} ${pdfNumber(cropBox.height)} re`
    const content = [
      'q',
      `${pdfRgbOperator(textRgb)} rg`,
      rectangle,
      'f',
      `${resourceName.toString()} gs`,
      `${pdfRgbOperator(backgroundRgb)} rg`,
      rectangle,
      'f',
      'Q',
    ].join('\n')
    page.node.addContentStream(document.context.register(document.context.flateStream(content)))
  }

  return document.save({ useObjectStreams: false })
}

export async function overlayAdjustedPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
  pageImages: Uint8Array[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(document.getPageCount(), pageIndexes)
  if (pageImages.length !== indexes.length) {
    throw new Error('Adjusted page images must match the selected pages')
  }

  for (let imageIndex = 0; imageIndex < indexes.length; imageIndex++) {
    const imageBytes = pageImages[imageIndex]
    if (!imageBytes || imageBytes.length === 0) throw new Error('Adjusted page image is empty')
    const page = document.getPage(indexes[imageIndex]!)
    const cropBox = page.getCropBox()
    const image = await document.embedPng(imageBytes)
    page.drawImage(image, {
      x: cropBox.x,
      y: cropBox.y,
      width: cropBox.width,
      height: cropBox.height,
    })
  }

  return document.save({ useObjectStreams: false })
}

export async function flattenPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageImages: Uint8Array[],
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  if (pageImages.length !== source.getPageCount()) {
    throw new Error('Flattened page images must match every PDF page')
  }
  const output = await PDFDocument.create()
  copyMetadata(source, output)

  for (let pageIndex = 0; pageIndex < pageImages.length; pageIndex++) {
    const imageBytes = pageImages[pageIndex]
    if (!imageBytes || imageBytes.length === 0) throw new Error('Flattened page image is empty')
    const sourcePage = source.getPage(pageIndex)
    const box = sourcePage.getCropBox()
    const rotation = ((sourcePage.getRotation().angle % 360) + 360) % 360
    const width = rotation === 90 || rotation === 270 ? box.height : box.width
    const height = rotation === 90 || rotation === 270 ? box.width : box.height
    const page = output.addPage([width, height])
    const isPng =
      imageBytes.length >= 8 &&
      imageBytes[0] === 0x89 &&
      imageBytes[1] === 0x50 &&
      imageBytes[2] === 0x4e &&
      imageBytes[3] === 0x47
    const image = isPng ? await output.embedPng(imageBytes) : await output.embedJpg(imageBytes)
    page.drawImage(image, { x: 0, y: 0, width, height })
    page.node.delete(PDFName.of('Annots'))
  }

  return output.save({ useObjectStreams: false })
}

function pdfaTimestamp(date: Date): string {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z')
}

function pdfaXmpMetadata(document: PDFDocument): string {
  const now = new Date()
  const title = document.getTitle()?.trim()
  const author = document.getAuthor()?.trim()
  const subject = document.getSubject()?.trim()
  const keywords = document.getKeywords()?.trim()
  const creator = document.getCreator()?.trim()
  const xml = (value: string) =>
    value
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&apos;')
  return `<?xpacket begin="\uFEFF" id="W5M0MpCehiHzreSzNTczkc9d"?>
<x:xmpmeta xmlns:x="adobe:ns:meta/">
  <rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">
    <rdf:Description rdf:about=""
      xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/"
      xmlns:xmp="http://ns.adobe.com/xap/1.0/"
      xmlns:pdf="http://ns.adobe.com/pdf/1.3/"
      xmlns:dc="http://purl.org/dc/elements/1.1/"
      pdfaid:part="2"
      pdfaid:conformance="B"
      xmp:CreateDate="${pdfaTimestamp(document.getCreationDate() ?? now)}"
      xmp:ModifyDate="${pdfaTimestamp(now)}"
      xmp:MetadataDate="${pdfaTimestamp(now)}"
      ${creator ? `xmp:CreatorTool="${xml(creator)}"` : ''}
      ${keywords ? `pdf:Keywords="${xml(keywords)}"` : ''}
      pdf:Producer="GenOffice PDF/A local converter">
      ${title ? `<dc:title><rdf:Alt><rdf:li xml:lang="x-default">${xml(title)}</rdf:li></rdf:Alt></dc:title>` : ''}
      ${author ? `<dc:creator><rdf:Seq><rdf:li>${xml(author)}</rdf:li></rdf:Seq></dc:creator>` : ''}
      ${subject ? `<dc:description><rdf:Alt><rdf:li xml:lang="x-default">${xml(subject)}</rdf:li></rdf:Alt></dc:description>` : ''}
    </rdf:Description>
  </rdf:RDF>
</x:xmpmeta>
<?xpacket end="w"?>`
}

export async function pdfToPdfaBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfToPdfaOptions,
): Promise<Uint8Array> {
  if (options.format !== 'PDF/A-2b') throw new Error('PDF archive format is invalid')
  if (options.archiveMode !== 'auto' && options.archiveMode !== 'raster') {
    throw new Error('PDF/A archive mode is invalid')
  }
  if (!Number.isInteger(options.renderDpi) || options.renderDpi < 72 || options.renderDpi > 600) {
    throw new Error('PDF/A rendering DPI must be a whole number from 72 to 600')
  }
  if (
    !Number.isInteger(options.imageQuality) ||
    options.imageQuality < 10 ||
    options.imageQuality > 100
  ) {
    throw new Error('PDF/A image quality must be a whole percentage from 10 to 100')
  }
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const preservation = pdfaPreservationReport(source)
  const preserveContent = options.archiveMode === 'auto' && preservation.eligible
  if (!preserveContent && !options.pageImages) {
    const detail = preservation.unembeddedFonts.length
      ? ` Unembedded fonts: ${preservation.unembeddedFonts.join(', ')}.`
      : ''
    throw new Error(`PDF/A image fallback requires a rendered image for every page.${detail}`)
  }
  const document = preserveContent
    ? source
    : await PDFDocument.load(await flattenPdfPagesBytes(bytes, options.pageImages!), {
        updateMetadata: false,
      })

  sanitizeJavaScript(document)
  const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  names?.delete(PDFName.of('EmbeddedFiles'))
  names?.delete(PDFName.of('JavaScript'))
  if (names?.keys().length === 0) document.catalog.delete(PDFName.of('Names'))
  document.catalog.delete(PDFName.of('AcroForm'))
  document.catalog.delete(PDFName.of('OpenAction'))
  document.catalog.delete(PDFName.of('AA'))
  document.catalog.delete(PDFName.of('AF'))
  document.catalog.delete(PDFName.of('Perms'))
  for (const page of document.getPages()) {
    page.node.delete(PDFName.of('Annots'))
    page.node.delete(PDFName.of('AA'))
    page.node.delete(PDFName.of('AF'))
  }
  for (const [, object] of document.context.enumerateIndirectObjects()) {
    if (object instanceof PDFDict) object.delete(PDFName.of('AF'))
  }

  document.setProducer('GenOffice PDF/A local converter')
  document.setModificationDate(new Date())

  const outputProfile = document.context.flateStream(SRGB_2014_ICC_BYTES, {
    N: 3,
    Alternate: 'DeviceRGB',
  })
  const outputProfileRef = document.context.register(outputProfile)
  const outputIntent = document.context.obj({
    Type: 'OutputIntent',
    S: 'GTS_PDFA1',
    OutputConditionIdentifier: PDFString.of('sRGB2014'),
    Info: PDFString.of('sRGB2014'),
    RegistryName: PDFString.of('http://www.color.org'),
    DestOutputProfile: outputProfileRef,
  })
  document.catalog.set(PDFName.of('OutputIntents'), document.context.obj([outputIntent]))

  const metadata = document.context.flateStream(
    new TextEncoder().encode(pdfaXmpMetadata(document)),
    {
      Type: 'Metadata',
      Subtype: 'XML',
    },
  )
  document.catalog.set(PDFName.of('Metadata'), document.context.register(metadata))

  return document.save({ useObjectStreams: false, updateFieldAppearances: false })
}

export async function scannerEffectPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageImages: Uint8Array[],
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  if (pageImages.length !== source.getPageCount()) {
    throw new Error('Scanner page images must match every PDF page')
  }
  const output = await PDFDocument.create()
  copyMetadata(source, output)

  for (let pageIndex = 0; pageIndex < pageImages.length; pageIndex++) {
    const imageBytes = pageImages[pageIndex]
    if (!imageBytes || imageBytes.length === 0) throw new Error('Scanner page image is empty')
    const sourcePage = source.getPage(pageIndex)
    const box = sourcePage.getCropBox()
    const rotation = ((sourcePage.getRotation().angle % 360) + 360) % 360
    const width = rotation === 90 || rotation === 270 ? box.height : box.width
    const height = rotation === 90 || rotation === 270 ? box.width : box.height
    const page = output.addPage([width, height])
    const isPng =
      imageBytes.length >= 8 &&
      imageBytes[0] === 0x89 &&
      imageBytes[1] === 0x50 &&
      imageBytes[2] === 0x4e &&
      imageBytes[3] === 0x47
    const image = isPng ? await output.embedPng(imageBytes) : await output.embedJpg(imageBytes)
    const scale = Math.max(width / image.width, height / image.height)
    const drawWidth = image.width * scale
    const drawHeight = image.height * scale
    page.drawImage(image, {
      x: (width - drawWidth) / 2,
      y: (height - drawHeight) / 2,
      width: drawWidth,
      height: drawHeight,
    })
  }

  return output.save({ useObjectStreams: false })
}

export async function deskewPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
  maxAngle: number,
  pages: PdfDeskewPage[],
): Promise<Uint8Array> {
  const sourceBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
  const source = await PDFDocument.load(sourceBytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(source.getPageCount(), pageIndexes)
  if (!Number.isFinite(maxAngle) || maxAngle < 0.5 || maxAngle > 15) {
    throw new Error('Deskew maximum angle must be from 0.5 to 15 degrees')
  }
  if (pages.length > indexes.length) throw new Error('Too many deskewed page images')
  const selected = new Set(indexes)
  const prepared = new Map<number, PdfDeskewPage>()
  for (const page of pages) {
    if (!Number.isInteger(page.pageIndex) || !selected.has(page.pageIndex)) {
      throw new Error('Deskewed page is outside the selected pages')
    }
    if (prepared.has(page.pageIndex)) throw new Error('Deskewed pages must be unique')
    if (
      !Number.isFinite(page.angle) ||
      Math.abs(page.angle) < 0.05 ||
      Math.abs(page.angle) > maxAngle
    ) {
      throw new Error('Deskewed page angle is invalid')
    }
    if (!(page.image instanceof Uint8Array) || page.image.length === 0) {
      throw new Error('Deskewed page image is empty')
    }
    prepared.set(page.pageIndex, page)
  }
  if (prepared.size === 0) return new Uint8Array(sourceBytes)

  const output = await PDFDocument.create()
  copyMetadata(source, output)
  for (let pageIndex = 0; pageIndex < source.getPageCount(); pageIndex++) {
    const corrected = prepared.get(pageIndex)
    if (!corrected) {
      const [copiedPage] = await output.copyPages(source, [pageIndex])
      output.addPage(copiedPage)
      continue
    }
    const sourcePage = source.getPage(pageIndex)
    const box = sourcePage.getCropBox()
    const rotation = ((sourcePage.getRotation().angle % 360) + 360) % 360
    const width = rotation === 90 || rotation === 270 ? box.height : box.width
    const height = rotation === 90 || rotation === 270 ? box.width : box.height
    const page = output.addPage([width, height])
    const image = await output.embedPng(corrected.image)
    page.drawImage(image, { x: 0, y: 0, width, height })
  }
  return output.save({ useObjectStreams: false })
}

async function rasterPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pages: PdfRasterPage[],
  label: string,
): Promise<Uint8Array> {
  if (pages.length === 0) throw new Error(`At least one ${label} page is required`)
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const output = await PDFDocument.create()
  copyMetadata(source, output)

  for (const rasterPage of pages) {
    if (!rasterPage.image || rasterPage.image.length === 0) {
      throw new Error(`${label} page image is empty`)
    }
    if (
      !Number.isFinite(rasterPage.width) ||
      !Number.isFinite(rasterPage.height) ||
      rasterPage.width <= 0 ||
      rasterPage.height <= 0
    ) {
      throw new Error(`${label} page size is invalid`)
    }
    const page = output.addPage([rasterPage.width, rasterPage.height])
    const image = await output.embedPng(rasterPage.image)
    page.drawImage(image, {
      x: 0,
      y: 0,
      width: rasterPage.width,
      height: rasterPage.height,
    })
  }

  return output.save({ useObjectStreams: false })
}

export async function comparisonPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pages: PdfComparisonPage[],
): Promise<Uint8Array> {
  return rasterPdfPagesBytes(bytes, pages, 'comparison')
}

export async function scannerImageSplitPdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  pages: PdfScannedImagePage[],
): Promise<Uint8Array> {
  return rasterPdfPagesBytes(bytes, pages, 'scanned image')
}

export async function compressPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageImages: Uint8Array[],
  options: { forceRasterized?: boolean } = {},
): Promise<Uint8Array> {
  const sourceBytes =
    bytes instanceof Uint8Array ? new Uint8Array(bytes) : new Uint8Array(bytes.slice(0))
  const compressed = await flattenPdfPagesBytes(sourceBytes, pageImages)
  return options.forceRasterized || compressed.length < sourceBytes.length
    ? compressed
    : new Uint8Array(sourceBytes)
}

export async function redactPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageImages: Uint8Array[],
): Promise<Uint8Array> {
  return flattenPdfPagesBytes(bytes, pageImages)
}

export async function redactSelectedPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pages: PdfRedactedPage[],
): Promise<Uint8Array> {
  if (pages.length === 0) throw new Error('At least one redacted page is required')
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const replacements = new Map<number, Uint8Array>()
  for (const page of pages) {
    if (
      !Number.isInteger(page.pageIndex) ||
      page.pageIndex < 0 ||
      page.pageIndex >= source.getPageCount()
    ) {
      throw new Error('Redacted page is outside the PDF')
    }
    if (replacements.has(page.pageIndex)) throw new Error('Redacted pages must be unique')
    if (!(page.image instanceof Uint8Array) || page.image.length === 0) {
      throw new Error('Redacted page image is empty')
    }
    replacements.set(page.pageIndex, page.image)
  }

  const output = await PDFDocument.create()
  copyMetadata(source, output)
  for (let pageIndex = 0; pageIndex < source.getPageCount(); pageIndex++) {
    const imageBytes = replacements.get(pageIndex)
    if (!imageBytes) {
      const [copiedPage] = await output.copyPages(source, [pageIndex])
      output.addPage(copiedPage)
      continue
    }
    const sourcePage = source.getPage(pageIndex)
    const box = sourcePage.getCropBox()
    const rotation = ((sourcePage.getRotation().angle % 360) + 360) % 360
    const width = rotation === 90 || rotation === 270 ? box.height : box.width
    const height = rotation === 90 || rotation === 270 ? box.width : box.height
    const page = output.addPage([width, height])
    const isPng =
      imageBytes.length >= 8 &&
      imageBytes[0] === 0x89 &&
      imageBytes[1] === 0x50 &&
      imageBytes[2] === 0x4e &&
      imageBytes[3] === 0x47
    const image = isPng ? await output.embedPng(imageBytes) : await output.embedJpg(imageBytes)
    page.drawImage(image, { x: 0, y: 0, width, height })
  }
  return output.save({ useObjectStreams: false })
}

export async function addPdfCommentsBytes(
  bytes: Uint8Array | ArrayBuffer,
  comments: PdfCommentInput[],
): Promise<Uint8Array> {
  if (comments.length === 0) throw new Error('At least one comment is required')
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const appliedComments = comments.filter((comment) => {
    return (
      Number.isInteger(comment.pageIndex) &&
      comment.pageIndex >= 0 &&
      comment.pageIndex < document.getPageCount() &&
      Number.isFinite(comment.x) &&
      Number.isFinite(comment.y) &&
      Number.isFinite(comment.width) &&
      comment.width > 0 &&
      Number.isFinite(comment.height) &&
      comment.height > 0 &&
      comment.text.trim().length > 0 &&
      comment.text.length <= 100_000
    )
  })
  if (appliedComments.length === 0) throw new Error('No valid comments were provided')

  const createdAt = PDFString.fromDate(new Date())
  for (let commentIndex = 0; commentIndex < appliedComments.length; commentIndex++) {
    const comment = appliedComments[commentIndex]!
    const page = document.getPage(comment.pageIndex)
    const annotation = document.context.obj({
      Type: 'Annot',
      Subtype: 'Text',
      Rect: [comment.x, comment.y, comment.x + comment.width, comment.y + comment.height],
      Name: 'Comment',
      C: [1, 0.95, 0.4],
      CA: 0.9,
      F: 4,
      Open: false,
      P: page.ref,
    })
    annotation.set(PDFName.of('Contents'), PDFHexString.fromText(comment.text))
    annotation.set(PDFName.of('T'), PDFHexString.fromText(comment.author?.trim() || 'GenOffice AI'))
    annotation.set(
      PDFName.of('Subj'),
      PDFHexString.fromText(comment.subject?.trim() || 'GenOffice AI Comment'),
    )
    annotation.set(PDFName.of('CreationDate'), createdAt)
    annotation.set(PDFName.of('M'), createdAt)
    annotation.set(
      PDFName.of('NM'),
      PDFHexString.fromText(`genoffice-comment-${Date.now()}-${commentIndex + 1}`),
    )
    const annotationReference = document.context.register(annotation)
    let annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annotations) {
      annotations = document.context.obj([])
      page.node.set(PDFName.of('Annots'), annotations)
    }
    annotations.push(annotationReference)
  }
  return document.save({ useObjectStreams: false })
}

export async function removePdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const removed = new Set(checkedPageIndexes(source.getPageCount(), pageIndexes))
  if (removed.size >= source.getPageCount()) throw new Error('At least one page must remain')
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const retainedIndexes = source.getPageIndices().filter((pageIndex) => !removed.has(pageIndex))
  const pages = await output.copyPages(source, retainedIndexes)
  for (const page of pages) output.addPage(page)
  return output.save({ useObjectStreams: false })
}

export function rearrangePageIndexes(pageCount: number, options: RearrangeOptions): number[] {
  if (!Number.isInteger(pageCount) || pageCount <= 0) {
    throw new Error('pageCount must be a positive integer')
  }
  const allPages = Array.from({ length: pageCount }, (_, pageIndex) => pageIndex)
  if (options.mode === 'custom') {
    return checkedPageIndexes(pageCount, options.customOrder ?? [], {
      allowDuplicates: true,
    })
  }
  if (options.mode === 'reverse') return allPages.reverse()
  if (options.mode === 'oddEven') {
    return [
      ...allPages.filter((pageIndex) => pageIndex % 2 === 0),
      ...allPages.filter((pageIndex) => pageIndex % 2 === 1),
    ]
  }
  if (options.mode === 'duplex') {
    const order: number[] = []
    for (let left = 0, right = pageCount - 1; left <= right; left++, right--) {
      order.push(left)
      if (left !== right) order.push(right)
    }
    return order
  }
  if (options.mode === 'removeFirst') return allPages.slice(1)
  if (options.mode === 'removeLast') return allPages.slice(0, -1)
  if (options.mode === 'removeFirstAndLast') return allPages.slice(1, -1)
  const duplicateCount = positiveInteger(options.duplicateCount ?? 2, 'duplicateCount')
  if (duplicateCount > 100) throw new Error('duplicateCount must not exceed 100')
  return allPages.flatMap((pageIndex) => Array.from({ length: duplicateCount }, () => pageIndex))
}

export async function rearrangePdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: RearrangeOptions,
): Promise<Uint8Array> {
  const source = await PDFDocument.load(bytes, { updateMetadata: false })
  const order = rearrangePageIndexes(source.getPageCount(), options)
  if (order.length === 0) throw new Error('At least one page must remain')
  const output = await PDFDocument.create()
  copyMetadata(source, output)
  const pages = await output.copyPages(source, order)
  for (const page of pages) output.addPage(page)
  return output.save({ useObjectStreams: false })
}

export async function processPdfFormBytes(
  bytes: Uint8Array | ArrayBuffer,
  action: PdfFormAction,
  fieldValues: PdfFormFieldValue[] = [],
  fieldNames: string[] = [],
  modifications: PdfFormFieldModification[] = [],
  creations: PdfFormFieldCreation[] = [],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const acroForm = document.catalog.getAcroForm()
  if (!acroForm && action !== 'create') return document.save({ useObjectStreams: false })
  const xfa = acroForm?.dict.get(PDFName.of('XFA'))
  if (xfa) acroForm?.dict.delete(PDFName.of('XFA'))
  const form = document.getForm()
  if (action === 'flatten') {
    form.flatten({ updateFieldAppearances: false })
    for (const page of document.getPages()) {
      const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
      if (!annotations) continue
      for (let index = annotations.size() - 1; index >= 0; index--) {
        if (!document.context.lookup(annotations.get(index))) annotations.remove(index)
      }
      if (annotations.size() === 0) page.node.delete(PDFName.of('Annots'))
    }
  } else if (action === 'unlock') {
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    for (const field of form.getFields()) {
      field.disableReadOnly()
      field.acroField.dict.delete(PDFName.of('Lock'))
    }
    if (xfa) {
      form.acroForm.dict.set(PDFName.of('XFA'), unlockXfaFormData(document, xfa))
    }
  } else if (action === 'delete') {
    const names = new Set(fieldNames.map((name) => name.trim()).filter(Boolean))
    if (names.size === 0) throw new Error('At least one PDF form field is required')
    const matched = form.getFields().filter((field) => names.has(field.getName()))
    if (matched.length === 0) throw new Error('No matching PDF form fields were found')
    for (const field of matched) {
      for (const widget of field.acroField.getWidgets()) {
        const widgetReference = document.context.getObjectRef(widget.dict)
        if (widgetReference) pageForWidget(document, widget)?.node.removeAnnot(widgetReference)
      }
      for (const page of document.getPages()) page.node.removeAnnot(field.ref)
      form.removeField(field)
    }
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    if (xfa) form.acroForm.dict.set(PDFName.of('XFA'), xfa)
    if (form.getFields().length === 0 && !xfa) document.catalog.delete(PDFName.of('AcroForm'))
  } else if (action === 'create') {
    if (creations.length === 0) throw new Error('At least one PDF form field creation is required')
    if (creations.length > 500)
      throw new Error('No more than 500 PDF form fields can be created at once')

    const existingNames = form.getFields().map((field) => field.getName())
    const requestedNames: string[] = []
    const resolved = creations.map((creation) => {
      const name = creation.name.trim()
      if (!['text', 'checkbox', 'radio', 'dropdown', 'optionList'].includes(creation.type)) {
        throw new Error(`Unsupported PDF form field type: ${String(creation.type)}`)
      }
      const nameParts = name.split('.')
      if (!name || name.length > 500 || nameParts.some((part) => !part.trim())) {
        throw new Error(`Invalid PDF form field name: ${creation.name}`)
      }
      const conflictingName = [...existingNames, ...requestedNames].find(
        (otherName) =>
          otherName === name ||
          otherName.startsWith(`${name}.`) ||
          name.startsWith(`${otherName}.`),
      )
      if (conflictingName) {
        throw new Error(`A PDF form field name conflicts with an existing field: ${name}`)
      }
      requestedNames.push(name)

      if (!Number.isInteger(creation.pageIndex) || creation.pageIndex < 0) {
        throw new Error(`Invalid PDF form field page: ${name}`)
      }
      const page = document.getPages()[creation.pageIndex]
      if (!page) throw new Error(`Invalid PDF form field page: ${name}`)
      const coordinates = [creation.x, creation.y, creation.width, creation.height]
      if (
        coordinates.some((value) => !Number.isFinite(value)) ||
        creation.x < 0 ||
        creation.y < 0 ||
        creation.width <= 0 ||
        creation.height <= 0
      ) {
        throw new Error(`Invalid PDF form field rectangle: ${name}`)
      }
      const { width: pageWidth, height: pageHeight } = page.getSize()
      if (creation.x + creation.width > pageWidth || creation.y + creation.height > pageHeight) {
        throw new Error(`PDF form field rectangle is outside the page: ${name}`)
      }

      const options = [
        ...new Set((creation.options ?? []).map((option) => option.trim()).filter(Boolean)),
      ]
      const choiceField = ['radio', 'dropdown', 'optionList'].includes(creation.type)
      if (choiceField && options.length === 0) {
        throw new Error(`A PDF choice field requires at least one option: ${name}`)
      }
      if (options.length > 100 || options.some((option) => option.length > 500)) {
        throw new Error(`PDF form field options are too large: ${name}`)
      }

      const defaultValues = Array.isArray(creation.defaultValue)
        ? [...new Set(creation.defaultValue.map((value) => value.trim()).filter(Boolean))]
        : typeof creation.defaultValue === 'string' && creation.defaultValue.trim()
          ? [creation.defaultValue.trim()]
          : []
      if (choiceField && defaultValues.some((value) => !options.includes(value))) {
        throw new Error(`PDF form field default values must match its options: ${name}`)
      }
      if ((creation.type === 'radio' || !creation.multiselect) && defaultValues.length > 1) {
        throw new Error(`PDF form field only supports one default value: ${name}`)
      }
      if (creation.type === 'text' && creation.defaultValue !== undefined) {
        if (typeof creation.defaultValue !== 'string') {
          throw new Error(`PDF text field default value must be text: ${name}`)
        }
      }
      if (creation.type === 'checkbox' && creation.defaultValue !== undefined) {
        if (typeof creation.defaultValue !== 'boolean') {
          throw new Error(`PDF checkbox default value must be boolean: ${name}`)
        }
      }
      if (!choiceField && creation.options && creation.options.length > 0) {
        throw new Error(`Options can only be used with PDF choice fields: ${name}`)
      }

      const optionSpacing = creation.optionSpacing ?? creation.height + 6
      if (
        creation.type === 'radio' &&
        (!Number.isFinite(optionSpacing) ||
          optionSpacing < creation.height ||
          creation.y - optionSpacing * (options.length - 1) < 0)
      ) {
        throw new Error(`Invalid PDF radio group spacing: ${name}`)
      }

      return {
        ...creation,
        name,
        label: creation.label?.trim() ?? '',
        page,
        options,
        defaultValues,
        optionSpacing,
      }
    })

    for (const creation of resolved) {
      const appearance = {
        x: creation.x,
        y: creation.y,
        width: creation.width,
        height: creation.height,
      }
      let field: PDFField
      if (creation.type === 'text') {
        const textField = form.createTextField(creation.name)
        if (creation.multiline) textField.enableMultiline()
        textField.addToPage(creation.page, appearance)
        if (typeof creation.defaultValue === 'string') textField.setText(creation.defaultValue)
        field = textField
      } else if (creation.type === 'checkbox') {
        const checkbox = form.createCheckBox(creation.name)
        checkbox.addToPage(creation.page, appearance)
        if (creation.defaultValue === true) checkbox.check()
        field = checkbox
      } else if (creation.type === 'dropdown') {
        const dropdown = form.createDropdown(creation.name)
        dropdown.addToPage(creation.page, appearance)
        dropdown.setOptions(creation.options)
        if (creation.multiselect) dropdown.enableMultiselect()
        if (creation.defaultValues.length > 0) dropdown.select(creation.defaultValues)
        field = dropdown
      } else if (creation.type === 'optionList') {
        const optionList = form.createOptionList(creation.name)
        optionList.addToPage(creation.page, appearance)
        optionList.setOptions(creation.options)
        if (creation.multiselect) optionList.enableMultiselect()
        if (creation.defaultValues.length > 0) optionList.select(creation.defaultValues)
        field = optionList
      } else {
        const radioGroup = form.createRadioGroup(creation.name)
        creation.options.forEach((option, index) => {
          radioGroup.addOptionToPage(option, creation.page, {
            ...appearance,
            y: creation.y - creation.optionSpacing * index,
          })
        })
        const [defaultValue] = creation.defaultValues
        if (defaultValue) radioGroup.select(defaultValue)
        field = radioGroup
      }

      if (creation.label) {
        field.acroField.dict.set(PDFName.of('TU'), PDFHexString.fromText(creation.label))
      }
      if (creation.readOnly) field.enableReadOnly()
      if (creation.required) field.enableRequired()
    }
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    if (xfa) form.acroForm.dict.set(PDFName.of('XFA'), xfa)
  } else if (action === 'modify') {
    const requested = modifications.filter((modification) => modification.name.trim())
    if (requested.length === 0)
      throw new Error('At least one PDF form field modification is required')

    const fields = form.getFields()
    const fieldsByName = new Map(fields.map((field) => [field.getName(), field]))
    const requestedNames = new Set<string>()
    const resolved = requested.map((modification) => {
      const name = modification.name.trim()
      if (requestedNames.has(name))
        throw new Error(`Duplicate PDF form field modification: ${name}`)
      requestedNames.add(name)
      const field = fieldsByName.get(name)
      if (!field) throw new Error(`No matching PDF form field was found: ${name}`)

      const oldParts = name.split('.')
      const oldParent = oldParts.slice(0, -1).join('.')
      const requestedNewName = modification.newName?.trim()
      let newName = name
      if (requestedNewName) {
        const newParts = requestedNewName.split('.')
        if (newParts.some((part) => !part.trim())) {
          throw new Error(
            `PDF form field names cannot contain empty path segments: ${requestedNewName}`,
          )
        }
        if (newParts.length === 1) {
          newName = oldParent ? `${oldParent}.${requestedNewName}` : requestedNewName
        } else {
          const newParent = newParts.slice(0, -1).join('.')
          if (newParent !== oldParent) {
            throw new Error(`Moving a PDF form field to another group is not supported: ${name}`)
          }
          newName = requestedNewName
        }
      }
      return { field, modification, name, newName, partialName: newName.split('.').at(-1)! }
    })

    const finalNames = new Set(fields.map((field) => field.getName()))
    for (const item of resolved) finalNames.delete(item.name)
    for (const item of resolved) {
      if (finalNames.has(item.newName)) {
        throw new Error(`A PDF form field already uses this name: ${item.newName}`)
      }
      finalNames.add(item.newName)
    }

    for (const { field, modification, name, newName, partialName } of resolved) {
      if (newName !== name) field.acroField.setPartialName(partialName)
      if (modification.label !== undefined) {
        const label = modification.label.trim()
        if (label) field.acroField.dict.set(PDFName.of('TU'), PDFHexString.fromText(label))
        else field.acroField.dict.delete(PDFName.of('TU'))
      }
      if (modification.readOnly === true) field.enableReadOnly()
      else if (modification.readOnly === false) field.disableReadOnly()
      if (modification.required === true) field.enableRequired()
      else if (modification.required === false) field.disableRequired()
      if (modification.options !== undefined || modification.multiselect !== undefined) {
        if (!(field instanceof PDFDropdown || field instanceof PDFOptionList)) {
          throw new Error(`Options can only be modified on PDF choice fields: ${name}`)
        }
        const options =
          modification.options === undefined
            ? field.getOptions()
            : [...new Set(modification.options.map((option) => option.trim()).filter(Boolean))]
        if (options.length === 0)
          throw new Error(`A PDF choice field requires at least one option: ${name}`)
        const multiselect = modification.multiselect ?? field.isMultiselect()
        const selected = field.getSelected().filter((option) => options.includes(option))
        field.setOptions(options)
        field.clear()
        const retained = multiselect ? selected : selected.slice(0, 1)
        if (retained.length > 0) field.select(retained)
        if (multiselect) field.enableMultiselect()
        else field.disableMultiselect()
      }
    }
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    if (xfa) form.acroForm.dict.set(PDFName.of('XFA'), xfa)
  } else {
    fillPdfFormFields(form.getFields(), fieldValues)
    form.acroForm.dict.set(PDFName.of('NeedAppearances'), PDFBool.True)
    if (xfa) form.acroForm.dict.set(PDFName.of('XFA'), xfa)
  }
  return document.save({ useObjectStreams: false, updateFieldAppearances: false })
}

function decodedXfaXml(bytes: Uint8Array): {
  text: string
  encode: (value: string) => Uint8Array
} {
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    return {
      text: new TextDecoder('utf-16be').decode(bytes),
      encode: (value) => encodeUtf16(value, false),
    }
  }
  if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    return {
      text: new TextDecoder('utf-16le').decode(bytes),
      encode: (value) => encodeUtf16(value, true),
    }
  }
  const hasBom = bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf
  return {
    text: new TextDecoder().decode(bytes),
    encode: (value) => {
      const encoded = new TextEncoder().encode(value)
      if (!hasBom) return encoded
      const output = new Uint8Array(encoded.length + 3)
      output.set([0xef, 0xbb, 0xbf])
      output.set(encoded, 3)
      return output
    },
  }
}

function encodeUtf16(value: string, littleEndian: boolean): Uint8Array {
  const output = new Uint8Array(value.length * 2 + 2)
  output.set(littleEndian ? [0xff, 0xfe] : [0xfe, 0xff])
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index)
    const offset = index * 2 + 2
    output[offset + (littleEndian ? 0 : 1)] = code & 0xff
    output[offset + (littleEndian ? 1 : 0)] = code >> 8
  }
  return output
}

function unlockedXfaStream(document: PDFDocument, stream: PDFRawStream): PDFRawStream {
  try {
    const decoded = decodePDFRawStream(stream).decode()
    const xml = decodedXfaXml(decoded)
    const unlocked = xml.text.replace(
      /\baccess\s*=\s*(["'])readOnly\1/gi,
      (_match, quote: string) => `access=${quote}open${quote}`,
    )
    if (unlocked === xml.text) return stream
    const dictionary = stream.dict.clone(document.context)
    dictionary.delete(PDFName.of('Filter'))
    dictionary.delete(PDFName.of('DecodeParms'))
    dictionary.delete(PDFName.of('D'))
    return PDFRawStream.of(dictionary, xml.encode(unlocked))
  } catch {
    return stream
  }
}

function unlockXfaObject(document: PDFDocument, object: PDFObject): PDFObject {
  const resolved = document.context.lookup(object)
  if (!(resolved instanceof PDFRawStream)) return object
  const unlocked = unlockedXfaStream(document, resolved)
  if (unlocked === resolved) return object
  if (object instanceof PDFRef) {
    document.context.assign(object, unlocked)
    return object
  }
  return unlocked
}

function unlockXfaFormData(document: PDFDocument, xfa: PDFObject): PDFObject {
  const resolved = document.context.lookup(xfa)
  if (resolved instanceof PDFArray) {
    for (let index = 1; index < resolved.size(); index += 2) {
      resolved.set(index, unlockXfaObject(document, resolved.get(index)))
    }
    return xfa
  }
  return unlockXfaObject(document, xfa)
}

function pdfFormFieldInfo(field: PDFField): PdfFormFieldInfo {
  const label = field.acroField.dict
    .lookupMaybe(PDFName.of('TU'), PDFString, PDFHexString)
    ?.decodeText()
  const shared = {
    name: field.getName(),
    ...(label ? { label } : {}),
    readOnly: field.isReadOnly(),
    required: field.isRequired(),
  }
  if (field instanceof PDFTextField) {
    let value = ''
    try {
      value = field.getText() ?? ''
    } catch {
      // Rich text fields can still be replaced with a normal text value.
    }
    return { ...shared, type: 'text', value, multiline: field.isMultiline() }
  }
  if (field instanceof PDFCheckBox) {
    return { ...shared, type: 'checkbox', value: field.isChecked() }
  }
  if (field instanceof PDFRadioGroup) {
    return {
      ...shared,
      type: 'radio',
      value: field.getSelected() ?? '',
      options: field.getOptions(),
    }
  }
  if (field instanceof PDFDropdown) {
    return {
      ...shared,
      type: 'dropdown',
      value: field.getSelected(),
      options: field.getOptions(),
      editable: field.isEditable(),
      multiselect: field.isMultiselect(),
    }
  }
  if (field instanceof PDFOptionList) {
    return {
      ...shared,
      type: 'optionList',
      value: field.getSelected(),
      options: field.getOptions(),
      multiselect: field.isMultiselect(),
    }
  }
  if (field instanceof PDFButton) return { ...shared, type: 'button' }
  if (field instanceof PDFSignature) return { ...shared, type: 'signature' }
  return { ...shared, type: 'unknown' }
}

export async function listPdfFormFieldsBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<PdfFormFieldInfo[]> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  if (!document.catalog.getAcroForm()) return []
  return document.getForm().getFields().map(pdfFormFieldInfo)
}

function formFieldExportValue(field: PdfFormFieldInfo): string | boolean | string[] | null {
  return field.value ?? null
}

function csvCell(value: string): string {
  return `"${value.replace(/"/g, '""')}"`
}

export function pdfFormFieldsJsonBytes(fields: PdfFormFieldInfo[]): Uint8Array {
  const values = Object.fromEntries(
    fields.map((field) => [field.name, formFieldExportValue(field)]),
  )
  return new TextEncoder().encode(`${JSON.stringify(values, null, 2)}\n`)
}

export function pdfFormFieldsCsvBytes(fields: PdfFormFieldInfo[]): Uint8Array {
  const rows = [
    ['Field Name', 'Value'],
    ...fields.map((field) => {
      const value = formFieldExportValue(field)
      return [
        field.name,
        Array.isArray(value) ? value.join('; ') : value === null ? '' : String(value),
      ]
    }),
  ]
  return new TextEncoder().encode(
    `\ufeff${rows.map((row) => row.map(csvCell).join(',')).join('\r\n')}\r\n`,
  )
}

export async function pdfFormFieldsXlsxBytes(fields: PdfFormFieldInfo[]): Promise<Uint8Array> {
  if (fields.length === 0) throw new Error('No PDF form fields were provided')
  return xlsxWorkbookBytes([
    {
      name: 'Form Fields',
      rows: [
        ['Field Name', 'Value'],
        ...fields.map((field) => {
          const value = formFieldExportValue(field)
          return [
            field.name,
            Array.isArray(value) ? value.join('; ') : value === null ? '' : String(value),
          ]
        }),
      ],
    },
  ])
}

function cleanExportText(value: string): string {
  return value
    .replace(/\r\n?/g, '\n')
    .replace(/[\t ]+\n/g, '\n')
    .trim()
}

function checkedExtractedTextPages(pages: PdfExtractedTextPage[]): PdfExtractedTextPage[] {
  const seen = new Set<number>()
  return pages.map((page) => {
    if (!Number.isInteger(page.pageNumber) || page.pageNumber < 1 || seen.has(page.pageNumber)) {
      throw new Error('Extracted text pages must have unique positive page numbers')
    }
    seen.add(page.pageNumber)
    return {
      ...page,
      text: cleanExportText(page.text),
      blocks: page.blocks
        .map((block) => ({ ...block, text: cleanExportText(block.text) }))
        .filter((block) => block.text.length > 0),
      links: page.links
        .map((link) => ({
          url: link.url.replace(/[\r\n]/g, '').trim(),
          label: link.label ? cleanExportText(link.label) : undefined,
        }))
        .filter((link) => link.url.length > 0),
    }
  })
}

function escapeMarkdownText(value: string): string {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/([`*_{}[\]<>|~])/g, '\\$1')
    .replace(/^([#+-]|\d+[.)])\s/gm, '\\$1 ')
}

function markdownLink(link: PdfExtractedTextLink): string {
  const label = escapeMarkdownText(link.label?.trim() || link.url)
  const url = link.url.replace(/[<>\\]/g, (character) => encodeURIComponent(character))
  return `[${label}](<${url}>)`
}

export function pdfTextPagesTxtBytes(pages: PdfExtractedTextPage[]): Uint8Array {
  const checked = checkedExtractedTextPages(pages)
  const output = checked.map((page) => page.text).join('\n\n\f\n\n')
  return new TextEncoder().encode(output ? `${output}\n` : '')
}

export function pdfTextPagesMarkdownBytes(
  pages: PdfExtractedTextPage[],
  includePageBreaks = true,
): Uint8Array {
  const checked = checkedExtractedTextPages(pages)
  const sections = checked.map((page) => {
    const content = page.blocks.map((block) => {
      const value = escapeMarkdownText(block.text)
      if (block.kind === 'heading') {
        const offset = includePageBreaks ? 2 : 0
        return `${'#'.repeat(Math.min(6, (block.level ?? 2) + offset))} ${value}`
      }
      if (block.kind === 'listItem') return `- ${value.replace(/^[•▪◦]\s*/, '')}`
      return value
    })
    const links = [...new Map(page.links.map((link) => [link.url, link])).values()]
    if (links.length > 0) {
      content.push(
        `${'#'.repeat(includePageBreaks ? 3 : 2)} Links`,
        ...links.map((link) => `- ${markdownLink(link)}`),
      )
    }
    return includePageBreaks
      ? [`## Page ${page.pageNumber}`, ...content].join('\n\n')
      : content.join('\n\n')
  })
  const output = sections.join(includePageBreaks ? '\n\n---\n\n' : '\n\n')
  return new TextEncoder().encode(output ? `${output}\n` : '')
}

export function pdfMarkdownOutputFileName(baseName?: string): string {
  const stem = (baseName ?? '')
    .replace(/^.*[/\\]/, '')
    .replace(/\.pdf$/i, '')
    .replace(/[<>:"/\\|?*\p{Cc}]/gu, '')
    .trim()
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 120)
  return `${stem || 'document'}_converted.md`
}

function checkedPdfJsonPages(
  pages: PdfJsonPage[],
  pageIndexes: number[],
  lightweight: boolean,
): PdfJsonPage[] {
  if (pages.length !== pageIndexes.length) {
    throw new Error('Structured PDF pages do not match the requested pages')
  }
  return pages.map((page, index) => {
    if (page.pageNumber !== pageIndexes[index]! + 1) {
      throw new Error('Structured PDF page order is invalid')
    }
    if (
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      !Number.isFinite(page.rotation)
    ) {
      throw new Error('Structured PDF page geometry is invalid')
    }
    if (!lightweight && !page.textRuns) {
      throw new Error('Full structured PDF export requires text layout data')
    }
    for (const run of page.textRuns ?? []) {
      if (
        typeof run.text !== 'string' ||
        ![run.x, run.y, run.width, run.height, run.fontSize].every(Number.isFinite) ||
        run.width < 0 ||
        run.height <= 0 ||
        run.fontSize <= 0
      ) {
        throw new Error('Structured PDF text layout is invalid')
      }
    }
    return page
  })
}

export async function pdfToJsonOutput(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfToJsonOptions,
): Promise<PdfToolOutput> {
  const analysis = await analyzePdfBytes(bytes)
  const pageIndexes = checkedPageIndexes(analysis.pageCount, options.pageIndexes)
  if (!options.pages) throw new Error('Structured PDF pages are required')
  const pages = checkedPdfJsonPages(options.pages, pageIndexes, options.lightweight).map((page) =>
    options.lightweight ? { ...page, textRuns: undefined } : page,
  )
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const document = {
    schema: 'genoffice.pdf.json',
    version: 1,
    mode: options.lightweight ? 'semantic' : 'layout',
    source: {
      fileName: baseName,
      pageCount: analysis.pageCount,
      selectedPages: pages.map((page) => page.pageNumber),
      pdfVersion: analysis.pdfVersion,
      fileSize: analysis.fileSize,
    },
    metadata: analysis.properties,
    pages,
    bookmarks: await listPdfBookmarksBytes(bytes),
    formFields: await listPdfFormFieldsBytes(bytes),
  }
  return {
    suffix: '_structured.json',
    fileName: `${baseName}_structured.json`,
    bytes: new TextEncoder().encode(`${JSON.stringify(document, null, 2)}\n`),
    mimeType: 'application/json;charset=utf-8',
    extension: '.json',
  }
}

const MAX_STRUCTURED_PDF_JSON_BYTES = 25 * 1024 * 1024
const MAX_STRUCTURED_PDF_PAGES = 2_000
const MAX_STRUCTURED_PDF_TEXT_RUNS = 100_000
const MAX_STRUCTURED_PDF_TEXT_LENGTH = 10_000_000

interface GenOfficePdfJsonDocument {
  schema: 'genoffice.pdf.json'
  version: 1
  mode: 'semantic' | 'layout'
  source?: { fileName?: string }
  metadata?: Partial<PdfAnalysis['properties']>
  pages: PdfJsonPage[]
  bookmarks: PdfBookmark[]
}

interface EmbeddedPdfJsonFonts {
  regular: PDFFont
  bold: PDFFont
  italic: PDFFont
  boldItalic: PDFFont
  unicode: PDFFont
}

function structuredPdfJsonObject(value: unknown, name: string): Record<string, unknown> {
  if (!value || typeof value !== 'object' || Array.isArray(value)) {
    throw new Error(`${name} must be an object`)
  }
  return value as Record<string, unknown>
}

function structuredPdfJsonText(value: unknown, name: string, maximum = 1_000_000): string {
  if (typeof value !== 'string' || value.length > maximum) {
    throw new Error(`${name} is invalid`)
  }
  return value
}

function structuredPdfJsonNumber(
  value: unknown,
  name: string,
  minimum: number,
  maximum: number,
): number {
  if (typeof value !== 'number' || !Number.isFinite(value) || value < minimum || value > maximum) {
    throw new Error(`${name} is invalid`)
  }
  return value
}

function parseStructuredPdfJson(jsonBytes: Uint8Array): GenOfficePdfJsonDocument {
  if (!(jsonBytes instanceof Uint8Array) || jsonBytes.length === 0) {
    throw new Error('Choose a non-empty structured PDF JSON file')
  }
  if (jsonBytes.length > MAX_STRUCTURED_PDF_JSON_BYTES) {
    throw new Error('Structured PDF JSON is too large')
  }
  let parsed: unknown
  try {
    parsed = JSON.parse(new TextDecoder('utf-8', { fatal: true }).decode(jsonBytes))
  } catch {
    throw new Error('Structured PDF JSON is malformed')
  }
  const root = structuredPdfJsonObject(parsed, 'Structured PDF JSON')
  if (root.schema !== 'genoffice.pdf.json' || root.version !== 1) {
    throw new Error('Only GenOffice structured PDF JSON version 1 is supported')
  }
  if (root.mode !== 'semantic' && root.mode !== 'layout') {
    throw new Error('Structured PDF JSON mode is invalid')
  }
  if (!Array.isArray(root.pages) || root.pages.length === 0) {
    throw new Error('Structured PDF JSON must contain at least one page')
  }
  if (root.pages.length > MAX_STRUCTURED_PDF_PAGES) {
    throw new Error('Structured PDF JSON contains too many pages')
  }
  let totalRuns = 0
  let totalTextLength = 0
  const pageNumbers = new Set<number>()
  const pages = root.pages.map((rawPage, pageIndex): PdfJsonPage => {
    const page = structuredPdfJsonObject(rawPage, `Page ${pageIndex + 1}`)
    const pageNumber = structuredPdfJsonNumber(
      page.pageNumber,
      `Page ${pageIndex + 1} number`,
      1,
      1_000_000,
    )
    if (!Number.isInteger(pageNumber) || pageNumbers.has(pageNumber)) {
      throw new Error(`Page ${pageIndex + 1} number is invalid or duplicated`)
    }
    pageNumbers.add(pageNumber)
    const width = structuredPdfJsonNumber(page.width, `Page ${pageNumber} width`, 1, 20_000)
    const height = structuredPdfJsonNumber(page.height, `Page ${pageNumber} height`, 1, 20_000)
    const rotation = structuredPdfJsonNumber(page.rotation, `Page ${pageNumber} rotation`, 0, 270)
    if (![0, 90, 180, 270].includes(rotation)) {
      throw new Error(`Page ${pageNumber} rotation is invalid`)
    }
    const text = structuredPdfJsonText(page.text ?? '', `Page ${pageNumber} text`)
    totalTextLength += text.length
    const blocks = Array.isArray(page.blocks)
      ? page.blocks.map((rawBlock, blockIndex): PdfExtractedTextBlock => {
          const block = structuredPdfJsonObject(
            rawBlock,
            `Page ${pageNumber} block ${blockIndex + 1}`,
          )
          if (!['heading', 'paragraph', 'listItem'].includes(String(block.kind))) {
            throw new Error(`Page ${pageNumber} block ${blockIndex + 1} kind is invalid`)
          }
          const blockText = structuredPdfJsonText(
            block.text,
            `Page ${pageNumber} block ${blockIndex + 1} text`,
          )
          totalTextLength += blockText.length
          const level = block.level
          if (level !== undefined && ![1, 2, 3].includes(Number(level))) {
            throw new Error(`Page ${pageNumber} block ${blockIndex + 1} level is invalid`)
          }
          return {
            kind: block.kind as PdfTextBlockKind,
            text: blockText,
            ...(level === undefined ? {} : { level: Number(level) as 1 | 2 | 3 }),
          }
        })
      : []
    const links = Array.isArray(page.links)
      ? page.links.map((rawLink, linkIndex): PdfExtractedTextLink => {
          const link = structuredPdfJsonObject(rawLink, `Page ${pageNumber} link ${linkIndex + 1}`)
          return {
            url: structuredPdfJsonText(
              link.url,
              `Page ${pageNumber} link ${linkIndex + 1} URL`,
              8_192,
            ),
            ...(link.label === undefined
              ? {}
              : {
                  label: structuredPdfJsonText(
                    link.label,
                    `Page ${pageNumber} link ${linkIndex + 1} label`,
                    16_384,
                  ),
                }),
          }
        })
      : []
    let textRuns: PdfJsonTextRun[] | undefined
    if (root.mode === 'layout') {
      if (!Array.isArray(page.textRuns)) {
        throw new Error(`Layout page ${pageNumber} requires text runs`)
      }
      textRuns = page.textRuns.map((rawRun, runIndex): PdfJsonTextRun => {
        totalRuns++
        if (totalRuns > MAX_STRUCTURED_PDF_TEXT_RUNS) {
          throw new Error('Structured PDF JSON contains too many text runs')
        }
        const run = structuredPdfJsonObject(rawRun, `Page ${pageNumber} text run ${runIndex + 1}`)
        const runText = structuredPdfJsonText(
          run.text,
          `Page ${pageNumber} text run ${runIndex + 1} text`,
        )
        totalTextLength += runText.length
        return {
          text: runText,
          x: structuredPdfJsonNumber(run.x, 'Text run x', -100_000, 100_000),
          y: structuredPdfJsonNumber(run.y, 'Text run y', -100_000, 100_000),
          width: structuredPdfJsonNumber(run.width, 'Text run width', 0, 100_000),
          height: structuredPdfJsonNumber(run.height, 'Text run height', 0.1, 10_000),
          fontSize: structuredPdfJsonNumber(run.fontSize, 'Text run font size', 0.1, 1_000),
          ...(run.fontFamily === undefined
            ? {}
            : { fontFamily: structuredPdfJsonText(run.fontFamily, 'Text run font family', 256) }),
          bold: run.bold === true,
          italic: run.italic === true,
        }
      })
    }
    if (totalTextLength > MAX_STRUCTURED_PDF_TEXT_LENGTH) {
      throw new Error('Structured PDF JSON contains too much text')
    }
    return {
      pageNumber,
      width,
      height,
      rotation,
      text,
      blocks,
      links,
      ...(textRuns ? { textRuns } : {}),
    }
  })
  const source =
    root.source && typeof root.source === 'object' && !Array.isArray(root.source)
      ? (root.source as Record<string, unknown>)
      : undefined
  return {
    schema: 'genoffice.pdf.json',
    version: 1,
    mode: root.mode,
    ...(source?.fileName === undefined
      ? {}
      : { source: { fileName: structuredPdfJsonText(source.fileName, 'Source file name', 512) } }),
    ...(root.metadata && typeof root.metadata === 'object' && !Array.isArray(root.metadata)
      ? { metadata: root.metadata as Partial<PdfAnalysis['properties']> }
      : {}),
    pages,
    bookmarks: Array.isArray(root.bookmarks) ? (root.bookmarks as PdfBookmark[]) : [],
  }
}

function checkedStructuredPdfFontBytes(fonts: PdfJsonImportFonts): PdfJsonImportFonts {
  const entries = Object.entries(fonts) as [keyof PdfJsonImportFonts, Uint8Array][]
  for (const [name, bytes] of entries) {
    if (!(bytes instanceof Uint8Array) || bytes.length === 0 || bytes.length > 25 * 1024 * 1024) {
      throw new Error(`Structured PDF ${name} font is invalid`)
    }
  }
  return fonts
}

async function embedStructuredPdfFonts(
  document: PDFDocument,
  fontBytes: PdfJsonImportFonts | undefined,
): Promise<EmbeddedPdfJsonFonts> {
  if (!fontBytes) {
    const [regular, bold, italic, boldItalic] = await Promise.all([
      document.embedFont(StandardFonts.Helvetica),
      document.embedFont(StandardFonts.HelveticaBold),
      document.embedFont(StandardFonts.HelveticaOblique),
      document.embedFont(StandardFonts.HelveticaBoldOblique),
    ])
    return { regular, bold, italic, boldItalic, unicode: regular }
  }
  document.registerFontkit(pdfLibFontkit)
  const checked = checkedStructuredPdfFontBytes(fontBytes)
  const [regular, bold, italic, boldItalic, unicode] = await Promise.all([
    document.embedFont(checked.regular, { subset: true }),
    document.embedFont(checked.bold, { subset: true }),
    document.embedFont(checked.italic, { subset: true }),
    document.embedFont(checked.boldItalic, { subset: true }),
    document.embedFont(checked.unicode, { subset: true }),
  ])
  return { regular, bold, italic, boldItalic, unicode }
}

function structuredPdfStyledFont(
  fonts: EmbeddedPdfJsonFonts,
  bold: boolean,
  italic: boolean,
): PDFFont {
  if (bold && italic) return fonts.boldItalic
  if (bold) return fonts.bold
  if (italic) return fonts.italic
  return fonts.regular
}

const structuredPdfFontCharacterSets = new WeakMap<PDFFont, Set<number>>()

function structuredPdfFontCharacterSet(font: PDFFont): Set<number> {
  let characters = structuredPdfFontCharacterSets.get(font)
  if (!characters) {
    characters = new Set(font.getCharacterSet())
    structuredPdfFontCharacterSets.set(font, characters)
  }
  return characters
}

function structuredPdfFontSupports(font: PDFFont, text: string): boolean {
  const characters = structuredPdfFontCharacterSet(font)
  return Array.from(text).every(
    (character) =>
      character === '\r' || character === '\n' || characters.has(character.codePointAt(0)!),
  )
}

function structuredPdfDrawableText(font: PDFFont, text: string): string {
  const characters = structuredPdfFontCharacterSet(font)
  const replacement = characters.has(0xfffd) ? '\ufffd' : characters.has(0x3f) ? '?' : ''
  return Array.from(text, (character) =>
    character === '\r' || character === '\n' || characters.has(character.codePointAt(0)!)
      ? character
      : replacement,
  ).join('')
}

function structuredPdfFontForText(
  fonts: EmbeddedPdfJsonFonts,
  text: string,
  bold = false,
  italic = false,
): { font: PDFFont; text: string } {
  const styled = structuredPdfStyledFont(fonts, bold, italic)
  const normalized = text.replace(/(?![\t\n\r])\p{Cc}/gu, '\ufffd')
  if (structuredPdfFontSupports(styled, normalized)) return { font: styled, text: normalized }
  return { font: fonts.unicode, text: structuredPdfDrawableText(fonts.unicode, normalized) }
}

function structuredPdfWrapText(
  text: string,
  font: PDFFont,
  size: number,
  maximumWidth: number,
): string[] {
  const lines: string[] = []
  for (const paragraph of text.split(/\r\n|\r|\n/u)) {
    let line = ''
    for (const character of Array.from(paragraph)) {
      const candidate = line + character
      if (!line || font.widthOfTextAtSize(candidate, size) <= maximumWidth) {
        line = candidate
      } else {
        lines.push(line.trimEnd())
        line = character.trimStart()
      }
    }
    lines.push(line.trimEnd())
  }
  return lines.length > 0 ? lines : ['']
}

function drawStructuredPdfSemanticPage(
  page: PDFPage,
  source: PdfJsonPage,
  fonts: EmbeddedPdfJsonFonts,
): void {
  const blocks =
    source.blocks.length > 0 ? source.blocks : [{ kind: 'paragraph' as const, text: source.text }]
  const margin = Math.min(48, Math.max(18, Math.min(page.getWidth(), page.getHeight()) * 0.08))
  const maximumWidth = Math.max(1, page.getWidth() - margin * 2)
  let y = page.getHeight() - margin
  for (const block of blocks) {
    const headingLevel = block.kind === 'heading' ? (block.level ?? 2) : undefined
    const fontSize =
      headingLevel === 1 ? 24 : headingLevel === 2 ? 18 : headingLevel === 3 ? 14 : 11
    const bold = block.kind === 'heading'
    const prefix = block.kind === 'listItem' && !/^[•*-]\s/u.test(block.text) ? '• ' : ''
    const selected = structuredPdfFontForText(fonts, `${prefix}${block.text}`, bold, false)
    const lineHeight = fontSize * 1.35
    const lines = structuredPdfWrapText(selected.text, selected.font, fontSize, maximumWidth)
    for (const line of lines) {
      if (y - lineHeight < margin) return
      y -= lineHeight
      if (line) page.drawText(line, { x: margin, y, size: fontSize, font: selected.font })
    }
    y -= fontSize * 0.45
  }
}

function remapStructuredPdfBookmarks(
  bookmarks: PdfBookmark[],
  pageNumbers: Map<number, number>,
): PdfBookmark[] {
  if (!Array.isArray(bookmarks)) throw new Error('Structured PDF bookmarks must be an array')
  const output: PdfBookmark[] = []
  for (const bookmark of bookmarks) {
    const record = structuredPdfJsonObject(bookmark, 'Structured PDF bookmark')
    const title = structuredPdfJsonText(record.title, 'Structured PDF bookmark title', 4_096)
    const children = remapStructuredPdfBookmarks(
      Array.isArray(record.children) ? (record.children as PdfBookmark[]) : [],
      pageNumbers,
    )
    const mappedPage = pageNumbers.get(Number(record.pageNumber))
    if (mappedPage) output.push({ title, pageNumber: mappedPage, children })
    else output.push(...children)
  }
  return output
}

function structuredPdfMetadata(
  metadata: Partial<PdfAnalysis['properties']> | undefined,
): PdfMetadataValues {
  const value = metadata ?? {}
  const stringValue = (input: unknown): string => (typeof input === 'string' ? input : '')
  return {
    title: stringValue(value.title),
    author: stringValue(value.author),
    subject: stringValue(value.subject),
    keywords: stringValue(value.keywords),
    creator: stringValue(value.creator),
    producer: stringValue(value.producer),
    creationDate: stringValue(value.creationDate),
    modificationDate: stringValue(value.modificationDate),
    trapped:
      value.trapped === 'True' || value.trapped === 'False' || value.trapped === 'Unknown'
        ? value.trapped
        : '',
    custom: Array.isArray(value.custom)
      ? value.custom.map((field) => ({
          key: structuredPdfJsonText(field?.key, 'Custom metadata key', 128),
          value: structuredPdfJsonText(field?.value, 'Custom metadata value'),
        }))
      : [],
  }
}

export async function jsonToPdfBytes(options: JsonToPdfOptions): Promise<PdfToolOutput> {
  const source = parseStructuredPdfJson(options.jsonBytes)
  const document = await PDFDocument.create()
  const fonts = await embedStructuredPdfFonts(document, options.fonts)
  const pageNumbers = new Map<number, number>()
  for (const [index, sourcePage] of source.pages.entries()) {
    const rotated =
      source.mode === 'layout' && (sourcePage.rotation === 90 || sourcePage.rotation === 270)
    const page = document.addPage(
      rotated ? [sourcePage.height, sourcePage.width] : [sourcePage.width, sourcePage.height],
    )
    if (source.mode === 'layout' && sourcePage.rotation) {
      page.setRotation(degrees(sourcePage.rotation))
    }
    pageNumbers.set(sourcePage.pageNumber, index + 1)
    if (source.mode === 'layout') {
      for (const run of sourcePage.textRuns ?? []) {
        if (!run.text) continue
        const selected = structuredPdfFontForText(
          fonts,
          run.text.replace(/\r\n|\r|\n/gu, ' '),
          run.bold,
          run.italic,
        )
        if (!selected.text) continue
        page.drawText(selected.text, {
          x: run.x,
          y: run.y,
          size: Math.min(1_000, Math.max(0.1, run.fontSize)),
          font: selected.font,
        })
      }
    } else {
      drawStructuredPdfSemanticPage(page, sourcePage, fonts)
    }
  }
  let bytes = await document.save({ useObjectStreams: false })
  bytes = await updatePdfMetadataBytes(bytes, {
    deleteAll: false,
    metadata: structuredPdfMetadata(source.metadata),
  })
  const bookmarks = remapStructuredPdfBookmarks(source.bookmarks, pageNumbers)
  if (bookmarks.length > 0) bytes = await setPdfBookmarksBytes(bytes, bookmarks)
  const baseName = safeExtractedImageBaseName(
    (options.baseName || source.source?.fileName || 'Document').replace(/\.json$/i, ''),
  )
  return {
    suffix: '_restored.pdf',
    fileName: `${baseName}_restored.pdf`,
    bytes,
    mimeType: 'application/pdf',
    extension: '.pdf',
  }
}

function xmlSafeValue(value: unknown): string {
  const safe = Array.from(String(value), (character) => {
    const codePoint = character.codePointAt(0)!
    return codePoint === 0x09 ||
      codePoint === 0x0a ||
      codePoint === 0x0d ||
      (codePoint >= 0x20 && codePoint <= 0xd7ff) ||
      (codePoint >= 0xe000 && codePoint <= 0xfffd) ||
      (codePoint >= 0x10000 && codePoint <= 0x10ffff)
      ? character
      : '\ufffd'
  }).join('')
  return escapeXml(safe)
}

function xmlAttributes(values: Record<string, string | number | boolean | undefined>): string {
  return Object.entries(values)
    .filter((entry): entry is [string, string | number | boolean] => entry[1] !== undefined)
    .map(([name, value]) => ` ${name}="${xmlSafeValue(value)}"`)
    .join('')
}

function pdfBookmarksXml(bookmarks: PdfBookmark[], indent = '    '): string[] {
  return bookmarks.flatMap((bookmark) => [
    `${indent}<bookmark${xmlAttributes({ title: bookmark.title, pageNumber: bookmark.pageNumber })}>`,
    ...pdfBookmarksXml(bookmark.children, `${indent}  `),
    `${indent}</bookmark>`,
  ])
}

function pdfFormFieldXml(field: PdfFormFieldInfo): string[] {
  const lines = [
    `    <formField${xmlAttributes({ name: field.name, type: field.type, readOnly: field.readOnly, required: field.required, multiline: field.multiline, multiselect: field.multiselect, editable: field.editable })}>`,
  ]
  if (field.label !== undefined) lines.push(`      <label>${xmlSafeValue(field.label)}</label>`)
  if (field.value !== undefined) {
    const values = Array.isArray(field.value) ? field.value : [field.value]
    lines.push('      <values>')
    lines.push(...values.map((value) => `        <value>${xmlSafeValue(value)}</value>`))
    lines.push('      </values>')
  }
  if (field.options) {
    lines.push('      <options>')
    lines.push(...field.options.map((option) => `        <option>${xmlSafeValue(option)}</option>`))
    lines.push('      </options>')
  }
  lines.push('    </formField>')
  return lines
}

export async function pdfToXmlOutput(
  bytes: Uint8Array | ArrayBuffer,
  options: PdfToXmlOptions,
): Promise<PdfToolOutput> {
  const analysis = await analyzePdfBytes(bytes)
  const pageIndexes = checkedPageIndexes(analysis.pageCount, options.pageIndexes)
  if (!options.pages) throw new Error('Structured PDF pages are required')
  const pages = checkedPdfJsonPages(options.pages, pageIndexes, options.lightweight)
  const bookmarks = await listPdfBookmarksBytes(bytes)
  const formFields = await listPdfFormFieldsBytes(bytes)
  const baseName = safeExtractedImageBaseName(options.baseName ?? 'Document')
  const mode = options.lightweight ? 'semantic' : 'layout'
  const lines = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    `<pdfDocument${xmlAttributes({ schema: 'genoffice.pdf.xml', version: 1, mode })}>`,
    `  <source${xmlAttributes({ fileName: baseName, pageCount: analysis.pageCount, pdfVersion: analysis.pdfVersion, fileSize: analysis.fileSize })}>`,
    '    <selectedPages>',
    ...pages.map((page) => `      <pageNumber>${page.pageNumber}</pageNumber>`),
    '    </selectedPages>',
    '  </source>',
    '  <metadata>',
  ]
  for (const [name, value] of Object.entries(analysis.properties)) {
    if (name === 'custom' || value === undefined) continue
    lines.push(`    <${name}>${xmlSafeValue(value)}</${name}>`)
  }
  if (analysis.properties.custom.length > 0) {
    lines.push('    <custom>')
    lines.push(
      ...analysis.properties.custom.map(
        (field) =>
          `      <field${xmlAttributes({ key: field.key })}>${xmlSafeValue(field.value)}</field>`,
      ),
    )
    lines.push('    </custom>')
  }
  lines.push('  </metadata>', '  <pages>')
  for (const page of pages) {
    lines.push(
      `    <page${xmlAttributes({ number: page.pageNumber, width: page.width, height: page.height, rotation: page.rotation })}>`,
      `      <text>${xmlSafeValue(page.text)}</text>`,
      '      <blocks>',
      ...page.blocks.map(
        (block) =>
          `        <block${xmlAttributes({ kind: block.kind, level: block.level })}>${xmlSafeValue(block.text)}</block>`,
      ),
      '      </blocks>',
      '      <links>',
      ...page.links.map(
        (link) => `        <link${xmlAttributes({ url: link.url, label: link.label })} />`,
      ),
      '      </links>',
    )
    if (!options.lightweight) {
      lines.push(
        '      <textRuns>',
        ...(page.textRuns ?? []).map(
          (run) =>
            `        <run${xmlAttributes({ x: run.x, y: run.y, width: run.width, height: run.height, fontSize: run.fontSize, fontFamily: run.fontFamily, bold: run.bold, italic: run.italic })}>${xmlSafeValue(run.text)}</run>`,
        ),
        '      </textRuns>',
      )
    }
    lines.push('    </page>')
  }
  lines.push(
    '  </pages>',
    '  <bookmarks>',
    ...pdfBookmarksXml(bookmarks),
    '  </bookmarks>',
    '  <formFields>',
    ...formFields.flatMap(pdfFormFieldXml),
    '  </formFields>',
    '</pdfDocument>',
    '',
  )
  return {
    suffix: '_structured.xml',
    fileName: `${baseName}_structured.xml`,
    bytes: new TextEncoder().encode(lines.join('\n')),
    mimeType: 'application/xml;charset=utf-8',
    extension: '.xml',
  }
}

function checkedExtractedTables(tables: PdfExtractedTable[]): PdfExtractedTable[] {
  if (tables.length > 1000) throw new Error('Too many extracted PDF tables')
  let totalCells = 0
  const keys = new Set<string>()
  return tables.map((table) => {
    const key = `${table.pageNumber}:${table.tableNumber}`
    if (
      !Number.isInteger(table.pageNumber) ||
      table.pageNumber < 1 ||
      !Number.isInteger(table.tableNumber) ||
      table.tableNumber < 1 ||
      keys.has(key) ||
      table.rows.length < 2 ||
      table.rows.length > 10_000
    ) {
      throw new Error('Extracted PDF table metadata is invalid')
    }
    keys.add(key)
    const columnCount = Math.max(0, ...table.rows.map((row) => row.length))
    if (columnCount < 2 || columnCount > 1000) {
      throw new Error('Extracted PDF table dimensions are invalid')
    }
    totalCells += table.rows.length * columnCount
    if (totalCells > 1_000_000) throw new Error('Extracted PDF tables contain too many cells')
    return {
      ...table,
      rows: table.rows.map((row) =>
        Array.from({ length: columnCount }, (_, index) =>
          cleanExportText(String(row[index] ?? '')).slice(0, 1_000_000),
        ),
      ),
    }
  })
}

function pdfTableCsvBytes(table: PdfExtractedTable): Uint8Array {
  return new TextEncoder().encode(
    `\ufeff${table.rows.map((row) => row.map(csvCell).join(',')).join('\r\n')}\r\n`,
  )
}

function safeArchiveStem(value: string): string {
  const stem = Array.from(
    value.replace(/\.pdf$/i, '').replace(/[<>:"/\\|?*]/g, '_'),
    (character) => (character.charCodeAt(0) < 32 ? '_' : character),
  )
    .join('')
    .trim()
  return stem.slice(0, 80) || 'Document'
}

export async function pdfTablesCsvOutput(
  tables: PdfExtractedTable[],
  baseName: string,
): Promise<PdfToolOutput> {
  if (tables.length === 0) throw new Error('No extracted PDF tables were provided')
  const checked = checkedExtractedTables(tables)
  if (checked.length === 1) {
    return {
      suffix: '_extracted.csv',
      bytes: pdfTableCsvBytes(checked[0]!),
      mimeType: 'text/csv;charset=utf-8',
      extension: '.csv',
    }
  }
  const archive = new JSZip()
  const stem = safeArchiveStem(baseName)
  for (const table of checked) {
    archive.file(`${stem}_p${table.pageNumber}_t${table.tableNumber}.csv`, pdfTableCsvBytes(table))
  }
  return {
    suffix: '_extracted_csv.zip',
    bytes: await archive.generateAsync({
      type: 'uint8array',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    }),
    mimeType: 'application/zip',
    extension: '.zip',
  }
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function xlsxColumnName(index: number): string {
  let value = index + 1
  let result = ''
  while (value > 0) {
    value--
    result = String.fromCharCode(65 + (value % 26)) + result
    value = Math.floor(value / 26)
  }
  return result
}

function xlsxSheetName(table: PdfExtractedTable, tables: PdfExtractedTable[]): string {
  const count = tables.filter((item) => item.pageNumber === table.pageNumber).length
  return count === 1
    ? `Page ${table.pageNumber}`
    : `Page ${table.pageNumber} Table ${table.tableNumber}`
}

function uniqueXlsxSheetNames(tables: PdfExtractedTable[]): string[] {
  const used = new Set<string>()
  return tables.map((table) => {
    const base =
      xlsxSheetName(table, tables)
        .replace(/[\\/?*:[\]]/g, '_')
        .slice(0, 31) || 'Table'
    let name = base
    for (let sequence = 2; used.has(name.toLowerCase()); sequence++) {
      const suffix = ` (${sequence})`
      name = `${base.slice(0, 31 - suffix.length)}${suffix}`
    }
    used.add(name.toLowerCase())
    return name
  })
}

function xlsxWorksheetXml(rows: string[][]): string {
  const body = rows
    .map((row, rowIndex) => {
      const cells = row
        .map((value, columnIndex) => {
          const reference = `${xlsxColumnName(columnIndex)}${rowIndex + 1}`
          return `<c r="${reference}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(value)}</t></is></c>`
        })
        .join('')
      return `<row r="${rowIndex + 1}">${cells}</row>`
    })
    .join('')
  const lastColumn = xlsxColumnName(Math.max(0, ...rows.map((row) => row.length)) - 1)
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    `<dimension ref="A1:${lastColumn}${rows.length}"/><sheetData>${body}</sheetData></worksheet>`
  )
}

interface XlsxWorksheet {
  name: string
  rows: string[][]
}

async function xlsxWorkbookBytes(sheets: XlsxWorksheet[]): Promise<Uint8Array> {
  const zip = new JSZip()
  const overrides = sheets
    .map(
      (_, index) =>
        `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    )
    .join('')
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      `${overrides}</Types>`,
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<sheets>${sheets.map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`).join('')}</sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets
        .map(
          (_, index) =>
            `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`,
        )
        .join('') +
      '</Relationships>',
  )
  sheets.forEach((sheet, index) =>
    zip.file(`xl/worksheets/sheet${index + 1}.xml`, xlsxWorksheetXml(sheet.rows)),
  )
  return zip.generateAsync({
    type: 'uint8array',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  })
}

export async function pdfTablesXlsxBytes(tables: PdfExtractedTable[]): Promise<Uint8Array> {
  if (tables.length === 0) throw new Error('No extracted PDF tables were provided')
  const checked = checkedExtractedTables(tables)
  const names = uniqueXlsxSheetNames(checked)
  return xlsxWorkbookBytes(
    checked.map((table, index) => ({ name: names[index]!, rows: table.rows })),
  )
}

function checkedFormSelection(
  field: PDFDropdown | PDFOptionList,
  value: string | boolean | string[],
): string[] {
  if (!Array.isArray(value) || value.some((option) => typeof option !== 'string')) {
    throw new Error(`${field.getName()} must be a string selection`)
  }
  if (!field.isMultiselect() && value.length > 1) {
    throw new Error(`${field.getName()} does not allow multiple selections`)
  }
  const available = new Set(field.getOptions())
  const editable = field instanceof PDFDropdown && field.isEditable()
  if (!editable && value.some((option) => !available.has(option))) {
    throw new Error(`${field.getName()} contains an unknown option`)
  }
  return [...new Set(value)]
}

function fillPdfFormFields(fields: PDFField[], fieldValues: PdfFormFieldValue[]): void {
  if (!Array.isArray(fieldValues) || fieldValues.length > 1000) {
    throw new Error('fields must contain at most 1000 items')
  }
  const fieldsByName = new Map(fields.map((field) => [field.getName(), field]))
  const seen = new Set<string>()
  for (const entry of fieldValues) {
    if (!entry || typeof entry.name !== 'string' || !entry.name || seen.has(entry.name)) {
      throw new Error('Each form field name must be unique and non-empty')
    }
    seen.add(entry.name)
    const field = fieldsByName.get(entry.name)
    if (!field) throw new Error(`Unknown form field: ${entry.name}`)
    if (field.isReadOnly()) throw new Error(`${entry.name} is read-only`)
    if (field instanceof PDFTextField) {
      if (typeof entry.value !== 'string') throw new Error(`${entry.name} must be text`)
      if (entry.value.length > 1_000_000) throw new Error(`${entry.name} is too long`)
      field.setText(entry.value)
    } else if (field instanceof PDFCheckBox) {
      if (typeof entry.value !== 'boolean') throw new Error(`${entry.name} must be true or false`)
      if (entry.value) field.check()
      else field.uncheck()
    } else if (field instanceof PDFRadioGroup) {
      if (typeof entry.value !== 'string') throw new Error(`${entry.name} must be one option`)
      if (!entry.value) field.clear()
      else if (!field.getOptions().includes(entry.value)) {
        throw new Error(`${entry.name} contains an unknown option`)
      } else field.select(entry.value)
    } else if (field instanceof PDFDropdown || field instanceof PDFOptionList) {
      const selection = checkedFormSelection(field, entry.value)
      if (selection.length === 0) field.clear()
      else field.select(selection)
    } else {
      throw new Error(`${entry.name} cannot be filled`)
    }
  }
}

export async function removePdfAnnotationsBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  for (const page of document.getPages()) page.node.delete(PDFName.of('Annots'))
  document.catalog.delete(PDFName.of('AcroForm'))
  return document.save({ useObjectStreams: false })
}

async function retainPdfPagesBytes(
  bytes: Uint8Array | ArrayBuffer,
  pageIndexes: number[],
): Promise<Uint8Array> {
  let document = await PDFDocument.load(bytes, { updateMetadata: false })
  const indexes = checkedPageIndexes(document.getPageCount(), pageIndexes)
  const retained = new Set(indexes)
  for (let pageIndex = document.getPageCount() - 1; pageIndex >= 0; pageIndex--) {
    if (!retained.has(pageIndex)) document.removePage(pageIndex)
  }
  document = await PDFDocument.load(await document.save({ useObjectStreams: false }), {
    updateMetadata: false,
  })
  pruneOrphanedPdfFields(document)
  return document.save({ useObjectStreams: false })
}

export function contentFilterOutputPageIndexes(
  pageCount: number,
  scannedPageIndexes: number[],
  matchedPageIndexes: number[],
  action: PdfContentFilterAction,
): number[] {
  if (!Number.isInteger(pageCount) || pageCount <= 0) {
    throw new Error('pageCount must be a positive integer')
  }
  const scanned = checkedPageIndexes(pageCount, scannedPageIndexes)
  const matched = checkedPageIndexes(pageCount, matchedPageIndexes, { allowEmpty: true })
  const scannedSet = new Set(scanned)
  if (matched.some((pageIndex) => !scannedSet.has(pageIndex))) {
    throw new Error('Matched pages must be inside the scanned page range')
  }
  if (matched.length === 0) throw new Error('No pages matched the selected content filter')
  if (action === 'keep') return matched
  const removed = new Set(matched)
  const retained = Array.from({ length: pageCount }, (_, pageIndex) => pageIndex).filter(
    (pageIndex) => !removed.has(pageIndex),
  )
  if (retained.length === 0) throw new Error('The content filter would remove every page')
  return retained
}

function comparePageFilterValue(
  actual: number,
  expected: number,
  comparator: PdfPageFilterComparator,
): boolean {
  if (comparator === 'less') return actual < expected
  if (comparator === 'greater') return actual > expected
  if (comparator === 'equal') return Math.abs(actual - expected) <= Math.max(1, expected * 0.01)
  throw new Error('Page filter comparator is invalid')
}

function compareDocumentFilterValue(
  actual: number,
  expected: number,
  comparator: PdfPageFilterComparator,
  tolerance = 0,
): boolean {
  if (comparator === 'less') return actual < expected
  if (comparator === 'greater') return actual > expected
  if (comparator === 'equal') return Math.abs(actual - expected) <= tolerance
  throw new Error('Document filter comparator is invalid')
}

export async function documentMatchesFilterBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: Pick<
    Extract<PdfToolOperation, { kind: 'filterDocuments' }>,
    'criterion' | 'comparator' | 'pageCount' | 'fileSizeBytes' | 'pageSize' | 'rotation'
  >,
): Promise<boolean> {
  if (!['pageCount', 'fileSize', 'pageSize', 'rotation'].includes(options.criterion)) {
    throw new Error('Document filter criterion is invalid')
  }
  if (!['less', 'equal', 'greater'].includes(options.comparator)) {
    throw new Error('Document filter comparator is invalid')
  }
  const byteLength = bytes.byteLength
  if (options.criterion === 'fileSize') {
    if (!Number.isInteger(options.fileSizeBytes) || (options.fileSizeBytes ?? 0) < 1) {
      throw new Error('Document filter file size must be a positive whole number')
    }
    return compareDocumentFilterValue(byteLength, options.fileSizeBytes!, options.comparator)
  }
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  if (document.getPageCount() === 0) throw new Error('The PDF must contain at least one page')
  if (options.criterion === 'pageCount') {
    if (!Number.isInteger(options.pageCount) || (options.pageCount ?? 0) < 1) {
      throw new Error('Document filter page count must be a positive whole number')
    }
    return compareDocumentFilterValue(
      document.getPageCount(),
      options.pageCount!,
      options.comparator,
    )
  }
  const firstPage = document.getPage(0)
  if (options.criterion === 'rotation') {
    if (![0, 90, 180, 270].includes(options.rotation ?? -1)) {
      throw new Error('Choose a valid document rotation')
    }
    const actualRotation = ((firstPage.getRotation().angle % 360) + 360) % 360
    return compareDocumentFilterValue(actualRotation, options.rotation!, options.comparator)
  }
  if (!options.pageSize || !(options.pageSize in PAGE_SIZES)) {
    throw new Error('Choose a valid document page size')
  }
  const mediaBox = firstPage.getMediaBox()
  const [standardWidth, standardHeight] = PAGE_SIZES[options.pageSize]
  const expectedArea = standardWidth * standardHeight
  return compareDocumentFilterValue(
    mediaBox.width * mediaBox.height,
    expectedArea,
    options.comparator,
    Math.max(1, expectedArea * 0.01),
  )
}

function uniquePdfOutputName(fileName: string, index: number, usedNames: Set<string>): string {
  const leafName = fileName.split(/[/\\]/).at(-1)?.trim() || `document-${index + 1}.pdf`
  const safeName = leafName.replace(/[\p{Cc}<>:"|?*]/gu, '_')
  const withExtension = /\.pdf$/i.test(safeName) ? safeName : `${safeName}.pdf`
  const stem = withExtension.replace(/\.pdf$/i, '') || `document-${index + 1}`
  let candidate = `${stem}.pdf`
  let duplicate = 2
  while (usedNames.has(candidate.toLocaleLowerCase())) {
    candidate = `${stem} (${duplicate++}).pdf`
  }
  usedNames.add(candidate.toLocaleLowerCase())
  return candidate
}

export function geometricFilterMatchedPageIndexes(
  document: PDFDocument,
  options: Pick<
    Extract<PdfToolOperation, { kind: 'filterPages' }>,
    'criterion' | 'pageIndexes' | 'pageSize' | 'orientation' | 'rotation' | 'comparator'
  >,
): number[] {
  if (!['pageSize', 'orientation', 'rotation'].includes(options.criterion)) {
    throw new Error('Geometric page filter criterion is invalid')
  }
  const pageIndexes = checkedPageIndexes(document.getPageCount(), options.pageIndexes)
  return pageIndexes.filter((pageIndex) => {
    const page = document.getPage(pageIndex)
    const box = page.getCropBox()
    const mediaBox = page.getMediaBox()
    const rotation = ((page.getRotation().angle % 360) + 360) % 360
    const visualWidth = rotation === 90 || rotation === 270 ? box.height : box.width
    const visualHeight = rotation === 90 || rotation === 270 ? box.width : box.height
    if (options.criterion === 'orientation') {
      if (!options.orientation) throw new Error('Choose a page orientation')
      const actual: PdfOrientation = visualWidth > visualHeight ? 'landscape' : 'portrait'
      return actual === options.orientation
    }
    if (options.criterion === 'rotation') {
      if (![0, 90, 180, 270].includes(options.rotation ?? -1)) {
        throw new Error('Choose a valid page rotation')
      }
      return rotation === options.rotation
    }
    if (!options.pageSize || !(options.pageSize in PAGE_SIZES)) {
      throw new Error('Choose a valid page size')
    }
    if (!options.comparator) throw new Error('Choose a page size comparison')
    const [standardWidth, standardHeight] = PAGE_SIZES[options.pageSize]
    return comparePageFilterValue(
      mediaBox.width * mediaBox.height,
      standardWidth * standardHeight,
      options.comparator,
    )
  })
}

function nameIs(value: PDFName | undefined, expected: string): boolean {
  return value?.asString() === `/${expected}`
}

function actionIs(dict: PDFDict, key: string, actionTypes: readonly string[]): boolean {
  const action = dict.lookupMaybe(PDFName.of(key), PDFDict)
  const type = action?.lookupMaybe(PDFName.of('S'), PDFName)
  return actionTypes.some((actionType) => nameIs(type, actionType))
}

function removeJavaScriptActionAt(
  document: PDFDocument,
  dictionary: PDFDict,
  key: PDFName,
  visited = new Set<PDFObject>(),
): void {
  const action = resolvedPdfObject(document, dictionary.get(key))
  if (!(action instanceof PDFDict) || visited.has(action)) return
  if (nameIs(action.lookupMaybe(PDFName.of('S'), PDFName), 'JavaScript')) {
    dictionary.delete(key)
    return
  }
  visited.add(action)
  removeJavaScriptNextActions(document, action, visited)
}

function removeJavaScriptNextActions(
  document: PDFDocument,
  action: PDFDict,
  visited: Set<PDFObject>,
): void {
  const nextKey = PDFName.of('Next')
  const next = resolvedPdfObject(document, action.get(nextKey))
  if (next instanceof PDFArray) {
    for (let index = next.size() - 1; index >= 0; index--) {
      const nextAction = resolvedPdfObject(document, next.get(index))
      if (!(nextAction instanceof PDFDict) || visited.has(nextAction)) continue
      if (nameIs(nextAction.lookupMaybe(PDFName.of('S'), PDFName), 'JavaScript')) {
        next.remove(index)
      } else {
        visited.add(nextAction)
        removeJavaScriptNextActions(document, nextAction, visited)
      }
    }
    if (next.size() === 0) action.delete(nextKey)
  } else if (next instanceof PDFDict) {
    if (nameIs(next.lookupMaybe(PDFName.of('S'), PDFName), 'JavaScript')) {
      action.delete(nextKey)
    } else if (!visited.has(next)) {
      visited.add(next)
      removeJavaScriptNextActions(document, next, visited)
    }
  }
}

function removeAdditionalJavaScriptActions(document: PDFDocument, dict: PDFDict): void {
  const additionalActions = dict.lookupMaybe(PDFName.of('AA'), PDFDict)
  if (!additionalActions) return
  for (const key of additionalActions.keys()) {
    removeJavaScriptActionAt(document, additionalActions, key)
  }
}

function sanitizeJavaScript(document: PDFDocument): void {
  const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
  names?.delete(PDFName.of('JavaScript'))
  removeJavaScriptActionAt(document, document.catalog, PDFName.of('OpenAction'))
  removeAdditionalJavaScriptActions(document, document.catalog)
  for (const [, object] of document.context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFDict)) continue
    removeJavaScriptActionAt(document, object, PDFName.of('A'))
    removeAdditionalJavaScriptActions(document, object)
  }
}

function sanitizePageAnnotations(
  document: PDFDocument,
  options: Pick<SanitizeOptions, 'removeEmbeddedFiles' | 'removeLinks'>,
): void {
  for (const page of document.getPages()) {
    const annotations = page.node.lookupMaybe(PDFName.of('Annots'), PDFArray)
    if (!annotations) continue
    for (let index = annotations.size() - 1; index >= 0; index--) {
      const annotation = annotations.lookupMaybe(index, PDFDict)
      if (!annotation) continue
      const subtype = annotation.lookupMaybe(PDFName.of('Subtype'), PDFName)
      if (options.removeEmbeddedFiles && nameIs(subtype, 'FileAttachment')) {
        annotations.remove(index)
      } else if (
        options.removeLinks &&
        nameIs(subtype, 'Link') &&
        actionIs(annotation, 'A', ['URI', 'Launch'])
      ) {
        annotation.delete(PDFName.of('A'))
      }
    }
  }
}

export async function sanitizePdfBytes(
  bytes: Uint8Array | ArrayBuffer,
  options: SanitizeOptions,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  if (options.removeJavaScript) sanitizeJavaScript(document)
  if (options.removeEmbeddedFiles) {
    const names = document.catalog.lookupMaybe(PDFName.of('Names'), PDFDict)
    names?.delete(PDFName.of('EmbeddedFiles'))
  }
  if (options.removeXmpMetadata) document.catalog.delete(PDFName.of('Metadata'))
  if (options.removeMetadata) {
    document.context.trailerInfo.Info = document.context.register(document.context.obj({}))
  }
  if (options.removeEmbeddedFiles || options.removeLinks) {
    sanitizePageAnnotations(document, options)
  }
  return document.save({ useObjectStreams: false })
}

export async function repairPdfBytes(bytes: Uint8Array | ArrayBuffer): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, {
    parseSpeed: ParseSpeeds.Fastest,
    throwOnInvalidObject: false,
    updateMetadata: false,
  })
  return document.save({
    useObjectStreams: false,
    addDefaultPage: false,
    updateFieldAppearances: false,
  })
}

export async function decompressPdfBytes(bytes: Uint8Array | ArrayBuffer): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  for (const [reference, object] of document.context.enumerateIndirectObjects()) {
    if (!(object instanceof PDFRawStream)) continue
    const filter = object.dict.get(PDFName.of('Filter'))
    const decodeParms = object.dict.get(PDFName.of('DecodeParms'))
    const abbreviatedDecodeParms = object.dict.get(PDFName.of('D'))
    if (!filter && !decodeParms && !abbreviatedDecodeParms) continue
    try {
      const decoded = decodePDFRawStream(object).decode()
      const dictionary = object.dict.clone(document.context)
      dictionary.delete(PDFName.of('Filter'))
      dictionary.delete(PDFName.of('DecodeParms'))
      dictionary.delete(PDFName.of('D'))
      document.context.assign(reference, PDFRawStream.of(dictionary, decoded))
    } catch {
      // Preserve streams whose codecs are not available in the browser runtime.
    }
  }
  return document.save({
    useObjectStreams: false,
    addDefaultPage: false,
    updateFieldAppearances: false,
  })
}

function pageForWidget(document: PDFDocument, widget: PDFWidgetAnnotation): PDFPage | undefined {
  const pageReference = widget.P()
  if (pageReference) return document.getPages().find((page) => page.ref === pageReference)
  const widgetReference = document.context.getObjectRef(widget.dict)
  return widgetReference ? document.findPageForAnnotationRef(widgetReference) : undefined
}

export function removeEmptyPdfSignatureFields(
  document: PDFDocument,
  names: readonly string[],
): string[] {
  if (names.length === 0 || !document.catalog.getAcroForm()) return []
  const requested = new Set(names)
  const form = document.getForm()
  const removed: string[] = []
  for (const field of form.getFields()) {
    if (
      !(field instanceof PDFSignature) ||
      !requested.has(field.getName()) ||
      field.acroField.V() !== undefined
    ) {
      continue
    }
    for (const widget of field.acroField.getWidgets()) {
      const normalAppearance = widget.AP()?.get(PDFName.of('N'))
      if (normalAppearance instanceof PDFRef) continue
      const { width, height } = widget.getRectangle()
      widget.setNormalAppearance(
        document.context.register(
          document.context.formXObject([], {
            BBox: [0, 0, Math.max(0, width), Math.max(0, height)],
            Resources: {},
          }),
        ),
      )
    }
    form.removeField(field)
    removed.push(field.getName())
  }
  return removed
}

export async function removePdfSignaturesBytes(
  bytes: Uint8Array | ArrayBuffer,
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  if (!document.catalog.getAcroForm()) return document.save({ useObjectStreams: false })
  const form = document.getForm()
  for (const signature of form.getFields().filter((field) => field instanceof PDFSignature)) {
    for (const widget of signature.acroField.getWidgets()) {
      const page = pageForWidget(document, widget)
      if (!page) continue
      const appearance = widget.AP()?.get(PDFName.of('N'))
      if (appearance instanceof PDFRef) {
        const rectangle = widget.getRectangle()
        const xObjectKey = page.node.newXObject('UnsignedWidget', appearance)
        page.pushOperators(
          pushGraphicsState(),
          translate(rectangle.x, rectangle.y),
          ...rotateInPlace({ ...rectangle, rotation: 0 }),
          drawObject(xObjectKey),
          popGraphicsState(),
        )
      }
      const widgetReference = document.context.getObjectRef(widget.dict)
      if (widgetReference) page.node.removeAnnot(widgetReference)
    }
    form.removeField(signature)
  }
  if (form.getFields().length === 0) document.catalog.delete(PDFName.of('AcroForm'))
  return document.save({ useObjectStreams: false, updateFieldAppearances: false })
}

function chapterOutputSuffix(title: string, index: number): string {
  const safeTitle = stripAsciiControlCharacters(title, false)
    .replace(/[<>:"/\\|?*]/g, '')
    .replace(/\s+/g, '_')
    .replace(/^\.+|\.+$/g, '')
    .slice(0, 80)
  return `_chapter_${index + 1}_${safeTitle || `Chapter_${index + 1}`}.pdf`
}

export async function addPdfOcrTextLayersBytes(
  bytes: Uint8Array | ArrayBuffer,
  textLayers: PdfOcrTextLayer[],
): Promise<Uint8Array> {
  const document = await PDFDocument.load(bytes, { updateMetadata: false })
  const pageCount = document.getPageCount()
  const seen = new Set<number>()
  for (const layer of textLayers) {
    if (!Number.isInteger(layer.pageIndex) || layer.pageIndex < 0 || layer.pageIndex >= pageCount) {
      throw new Error('OCR text layer page is invalid')
    }
    if (seen.has(layer.pageIndex)) throw new Error('OCR text layer page is duplicated')
    if (!(layer.bytes instanceof Uint8Array) || layer.bytes.length === 0) {
      throw new Error('OCR text layer is empty')
    }
    seen.add(layer.pageIndex)
    const [embeddedLayer] = await document.embedPdf(layer.bytes, [0])
    if (!embeddedLayer) throw new Error('OCR text layer has no page')
    const page = document.getPage(layer.pageIndex)
    const cropBox = page.getCropBox()
    if (layer.replacePage) {
      page.node.set(PDFName.of('Contents'), document.context.obj([]))
      page.node.set(PDFName.of('Resources'), document.context.obj({}))
    }
    page.drawPage(embeddedLayer, cropBox)
  }
  return document.save({ useObjectStreams: false, updateFieldAppearances: false })
}

export function pdfOcrSidecarText(pageCount: number, pageTexts: PdfOcrPageText[]): string {
  if (!Number.isInteger(pageCount) || pageCount <= 0) {
    throw new Error('OCR sidecar page count must be a positive whole number')
  }
  const texts = new Map<number, string>()
  let totalLength = 0
  for (const page of pageTexts) {
    if (!Number.isInteger(page.pageIndex) || page.pageIndex < 0 || page.pageIndex >= pageCount) {
      throw new Error('OCR sidecar text page is invalid')
    }
    if (page.source !== 'existing' && page.source !== 'ocr') {
      throw new Error('OCR sidecar text source is invalid')
    }
    if (texts.has(page.pageIndex)) throw new Error('OCR sidecar text page is duplicated')
    const text = page.text.replace(/\r\n?/g, '\n').trim()
    if (text.length > 5_000_000) throw new Error('OCR sidecar page text is too large')
    totalLength += text.length
    if (totalLength > 50_000_000) throw new Error('OCR sidecar text is too large')
    texts.set(page.pageIndex, text)
  }
  return `${Array.from({ length: pageCount }, (_, pageIndex) => {
    const text = texts.get(pageIndex) || '[No text recognized]'
    return `===== Page ${pageIndex + 1} =====\n${text}`
  }).join('\n\n\f\n\n')}\n`
}

export async function pdfOcrSidecarZipBytes(
  pdfBytes: Uint8Array,
  pageCount: number,
  pageTexts: PdfOcrPageText[],
  baseName: string,
): Promise<Uint8Array> {
  if (!(pdfBytes instanceof Uint8Array) || pdfBytes.length === 0) {
    throw new Error('OCR PDF is required for sidecar output')
  }
  const stem = safeArchiveStem(baseName)
  const archive = new JSZip()
  archive.file(`${stem}_OCR.pdf`, pdfBytes, { compression: 'STORE' })
  archive.file(`${stem}_OCR.txt`, new TextEncoder().encode(pdfOcrSidecarText(pageCount, pageTexts)))
  return archive.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

export async function runPdfToolBytes(
  bytes: Uint8Array | ArrayBuffer,
  operation: PdfToolOperation,
): Promise<PdfToolOutput[]> {
  if (operation.kind === 'pipeline') {
    if (operation.steps.length === 0 || operation.steps.length > 12) {
      throw new Error('PDF pipeline must contain from 1 to 12 steps')
    }
    let current: Uint8Array<ArrayBufferLike> =
      bytes instanceof Uint8Array ? new Uint8Array(bytes) : new Uint8Array(bytes.slice(0))
    for (const step of operation.steps) {
      if (step.kind === 'sanitize') {
        if (
          !step.removeJavaScript &&
          !step.removeEmbeddedFiles &&
          !step.removeXmpMetadata &&
          !step.removeMetadata &&
          !step.removeLinks
        ) {
          throw new Error('Pipeline sanitize step must select at least one cleanup action')
        }
        current = await sanitizePdfBytes(current, step)
      } else if (step.kind === 'removeAnnotations') {
        current = await removePdfAnnotationsBytes(current)
      } else if (step.kind === 'removeSignatures') {
        current = await removePdfSignaturesBytes(current)
      } else if (step.kind === 'flattenForms') {
        current = await processPdfFormBytes(current, 'flatten')
      } else if (step.kind === 'repair') {
        current = await repairPdfBytes(current)
      } else if (step.kind === 'decompress') {
        current = await decompressPdfBytes(current)
      } else {
        throw new Error('Unsupported PDF pipeline step')
      }
    }
    return [{ suffix: '_processed.pdf', bytes: current }]
  }
  if (operation.kind === 'split') {
    if (operation.mode === 'chapters') {
      const chapters = await splitPdfByChaptersBytes(
        bytes,
        operation.bookmarkLevel,
        operation.allowDuplicates,
      )
      return chapters.map((chapter, index) => ({
        suffix: chapterOutputSuffix(chapter.title, index),
        bytes: chapter.bytes,
      }))
    }
    const outputs =
      operation.mode === 'afterPages'
        ? await splitPdfAtBoundariesBytes(bytes, operation.splitAfterPages)
        : operation.mode === 'fileSize'
          ? await splitPdfBySizeBytes(bytes, operation.maxBytes)
          : operation.mode === 'pagesPerDocument'
            ? await splitPdfByPageCountBytes(bytes, operation.value)
            : await splitPdfByDocumentCountBytes(bytes, operation.value)
    return outputs.map((output, index) => ({ suffix: `_split_${index + 1}.pdf`, bytes: output }))
  }
  if (operation.kind === 'password') {
    const input = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    const encryption = await pdfEncryptionInfoBytes(input)
    if (operation.action === 'unlock') {
      if (!encryption.encrypted) throw new Error('This PDF is not password protected')
      return [{ suffix: '_unlocked.pdf', bytes: await decryptPDF(input, operation.password) }]
    }
    if (encryption.encrypted) throw new Error('Unlock the PDF before applying new protection')
    if (!['AES-256', 'RC4'].includes(operation.algorithm)) {
      throw new Error('PDF encryption algorithm is invalid')
    }
    const userPassword = operation.userPassword
    const ownerPassword = operation.ownerPassword
    if (!userPassword && !ownerPassword) {
      throw new Error('Enter an open password or an owner password')
    }
    const permissions = operation.permissions
    const permissionValues = Object.values(permissions)
    if (
      permissionValues.length !== 8 ||
      permissionValues.some((value) => typeof value !== 'boolean')
    ) {
      throw new Error('PDF permissions are invalid')
    }
    const hasRestrictions = permissionValues.some((allowed) => !allowed)
    if (hasRestrictions && !ownerPassword) {
      throw new Error('An owner password is required when permissions are restricted')
    }
    if (hasRestrictions && userPassword && ownerPassword === userPassword) {
      throw new Error('Use a different owner password when permissions are restricted')
    }
    return [
      {
        suffix: '_protected.pdf',
        bytes: await encryptPDF(input, userPassword, {
          ownerPassword: ownerPassword || userPassword,
          algorithm: operation.algorithm,
          ...permissions,
        }),
      },
    ]
  }
  if (operation.kind === 'ocr') {
    if (!['skipText', 'force', 'strict'].includes(operation.mode)) {
      throw new Error('OCR mode is invalid')
    }
    if (
      operation.languages.length === 0 ||
      operation.languages.some((language) => !['eng', 'chi_sim'].includes(language))
    ) {
      throw new Error('Choose at least one supported OCR language')
    }
    if (
      !Number.isInteger(operation.renderDpi) ||
      operation.renderDpi < 100 ||
      operation.renderDpi > 300
    ) {
      throw new Error('OCR rendering DPI must be a whole number from 100 to 300')
    }
    if (!operation.textLayers) throw new Error('OCR text layer analysis is required')
    if (operation.textLayers.length === 0) {
      throw new Error('No pages require OCR')
    }
    const outputPdf = await addPdfOcrTextLayersBytes(bytes, operation.textLayers)
    if (!operation.sidecar) return [{ suffix: '_ocr.pdf', bytes: outputPdf }]
    if (operation.baseName === undefined || !operation.pageTexts) {
      throw new Error('OCR sidecar text analysis is required')
    }
    const document = await PDFDocument.load(outputPdf, { updateMetadata: false })
    const stem = safeArchiveStem(operation.baseName)
    return [
      {
        suffix: '_ocr.zip',
        fileName: `${stem}_OCR.zip`,
        bytes: await pdfOcrSidecarZipBytes(
          outputPdf,
          document.getPageCount(),
          operation.pageTexts,
          stem,
        ),
        mimeType: 'application/zip',
        extension: '.zip',
      },
    ]
  }
  if (operation.kind === 'merge') {
    if (operation.documents.length === 0) throw new Error('Choose at least one PDF to merge')
    if (
      !Number.isInteger(operation.currentDocumentIndex) ||
      operation.currentDocumentIndex < 0 ||
      operation.currentDocumentIndex > operation.documents.length
    ) {
      throw new Error('Current PDF position is invalid')
    }
    if (operation.documents.some((document) => document.length === 0)) {
      throw new Error('Merged PDF is empty')
    }
    const sourceBytes = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)
    const documents = [...operation.documents]
    documents.splice(operation.currentDocumentIndex, 0, sourceBytes)
    return [{ suffix: '_merged.pdf', bytes: await mergePdfBytes(documents) }]
  }
  if (operation.kind === 'compare') {
    if (
      !Number.isInteger(operation.renderDpi) ||
      operation.renderDpi < 72 ||
      operation.renderDpi > 300
    ) {
      throw new Error('Comparison rendering DPI must be a whole number from 72 to 300')
    }
    if (
      !Number.isFinite(operation.threshold) ||
      operation.threshold < 0 ||
      operation.threshold > 1
    ) {
      throw new Error('Comparison threshold must be from 0 to 1')
    }
    if (!operation.comparisonDocument || operation.comparisonDocument.length === 0) {
      throw new Error('Comparison PDF is empty')
    }
    if (!operation.pages) throw new Error('Rendered comparison pages are required')
    return [
      {
        suffix: '_comparison.pdf',
        bytes: await comparisonPdfPagesBytes(bytes, operation.pages),
      },
    ]
  }
  if (operation.kind === 'autoRename') {
    if (!['largestHeading', 'firstText'].includes(operation.strategy)) {
      throw new Error('Auto rename strategy is invalid')
    }
    if (!operation.suggestedName) throw new Error('An inferred PDF title is required')
    return [
      {
        suffix: '',
        fileName: pdfAutoRenameFileName(operation.suggestedName),
        bytes: bytes instanceof Uint8Array ? new Uint8Array(bytes) : new Uint8Array(bytes.slice(0)),
      },
    ]
  }
  if (operation.kind === 'pageNumbers') {
    return [
      {
        suffix: '_page_numbers_added.pdf',
        bytes: await addPdfPageNumbersBytes(bytes, operation),
      },
    ]
  }
  if (operation.kind === 'extractPages') {
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const pageIndexes = checkedPageIndexes(document.getPageCount(), operation.pageIndexes, {
      allowDuplicates: true,
    })
    return [
      {
        suffix: '_extracted_pages.pdf',
        bytes: await extractPagesBytes(bytes, pageIndexes),
      },
    ]
  }
  if (operation.kind === 'splitSections') {
    const outputs = await splitPdfSectionsBytes(bytes, operation)
    if (operation.merge) {
      return outputs.map((output) => ({ suffix: '_sections.pdf', bytes: output.bytes }))
    }
    return outputs.map((output) => ({
      suffix: `_page_${output.sourcePageNumber}_section_${output.sectionNumber}.pdf`,
      bytes: output.bytes,
    }))
  }
  if (operation.kind === 'crop') {
    if (operation.mode === 'auto') {
      if (!operation.pageBoxes) throw new Error('Auto crop boxes are required')
      return [
        {
          suffix: '_auto_cropped.pdf',
          bytes: await cropPdfPageBoxesBytes(bytes, operation.pageBoxes),
        },
      ]
    }
    return [{ suffix: '_cropped.pdf', bytes: await cropPdfMarginsBytes(bytes, operation.margins) }]
  }
  if (operation.kind === 'scale') {
    return [{ suffix: '_scaled.pdf', bytes: await scalePdfPagesBytes(bytes, operation) }]
  }
  if (operation.kind === 'nup') {
    return [{ suffix: '_nup.pdf', bytes: await nUpPdfBytes(bytes, operation) }]
  }
  if (operation.kind === 'booklet') {
    return [{ suffix: '_booklet.pdf', bytes: await bookletPdfBytes(bytes, operation) }]
  }
  if (operation.kind === 'poster') {
    return [{ suffix: '_poster.pdf', bytes: await posterPdfBytes(bytes, operation) }]
  }
  if (operation.kind === 'singlePage') {
    return [{ suffix: '_single_page.pdf', bytes: await singlePagePdfBytes(bytes, operation) }]
  }
  if (operation.kind === 'rotatePages') {
    return [
      {
        suffix: '_rotated.pdf',
        bytes: await rotatePdfPagesBytes(bytes, operation.pageIndexes, operation.angle),
      },
    ]
  }
  if (operation.kind === 'autoRotate') {
    if (!operation.pageRotations) throw new Error('Auto-rotation analysis is required')
    return [
      {
        suffix: '_auto_rotated.pdf',
        bytes: await autoRotatePdfPagesBytes(bytes, operation.pageRotations),
      },
    ]
  }
  if (operation.kind === 'deskew') {
    if (
      !Number.isInteger(operation.renderDpi) ||
      operation.renderDpi < 72 ||
      operation.renderDpi > 300
    ) {
      throw new Error('Deskew rendering DPI must be a whole number from 72 to 300')
    }
    if (typeof operation.includeAnnotations !== 'boolean') {
      throw new Error('Deskew annotation setting is invalid')
    }
    if (!operation.pages) throw new Error('Deskew page analysis is required')
    return [
      {
        suffix: '_deskewed.pdf',
        bytes: await deskewPdfPagesBytes(
          bytes,
          operation.pageIndexes,
          operation.maxAngle,
          operation.pages,
        ),
      },
    ]
  }
  if (operation.kind === 'scannerEffect') {
    if (!operation.pageImages) throw new Error('Scanner page images are required')
    return [
      {
        suffix: '_scanner_effect.pdf',
        bytes: await scannerEffectPdfPagesBytes(bytes, operation.pageImages),
      },
    ]
  }
  if (operation.kind === 'scannerImageSplit') {
    const checks = [
      ['angle threshold', operation.angleThreshold, 0, 45],
      ['tolerance', operation.tolerance, 0, 255],
      ['minimum area', operation.minArea, 1, 100_000_000],
      ['minimum contour area', operation.minContourArea, 1, 100_000_000],
      ['border size', operation.borderSize, 0, 200],
      ['rendering DPI', operation.renderDpi, 72, 300],
    ] as const
    for (const [name, value, minimum, maximum] of checks) {
      if (!Number.isInteger(value) || value < minimum || value > maximum) {
        throw new Error(`${name} must be a whole number from ${minimum} to ${maximum}`)
      }
    }
    if (!operation.pages) throw new Error('Detected scanned image pages are required')
    return [
      {
        suffix: '_scanned_images.pdf',
        bytes: await scannerImageSplitPdfBytes(bytes, operation.pages),
      },
    ]
  }
  if (operation.kind === 'autoSplit') {
    if (operation.action === 'divider') {
      return [
        {
          suffix: '_auto_split_divider.pdf',
          bytes: await createPdfAutoSplitDividerBytes(),
        },
      ]
    }
    if (!operation.dividerPageIndexes || operation.baseName === undefined) {
      throw new Error('Detected QR divider pages are required')
    }
    return [
      {
        suffix: '_auto_split.zip',
        bytes: await autoSplitPdfZipBytes(
          bytes,
          operation.dividerPageIndexes,
          operation.duplexMode,
          operation.baseName,
        ),
        mimeType: 'application/zip',
        extension: '.zip',
      },
    ]
  }
  if (operation.kind === 'removePages') {
    return [
      {
        suffix: '_removed_pages.pdf',
        bytes: await removePdfPagesBytes(bytes, operation.pageIndexes),
      },
    ]
  }
  if (operation.kind === 'removeImages') {
    return [
      {
        suffix: '_images_removed.pdf',
        bytes: await removePdfImagesBytes(bytes, operation.pageIndexes),
      },
    ]
  }
  if (operation.kind === 'extractText') {
    if (!operation.pages) throw new Error('Extracted PDF text pages are required')
    if (!['txt', 'markdown', 'both'].includes(operation.format)) {
      throw new Error('Text export format is invalid')
    }
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const pageIndexes = checkedPageIndexes(document.getPageCount(), operation.pageIndexes)
    const pages = checkedExtractedTextPages(operation.pages)
    if (
      pages.length !== pageIndexes.length ||
      pages.some((page, index) => page.pageNumber !== pageIndexes[index]! + 1)
    ) {
      throw new Error('Extracted text pages do not match the requested pages')
    }
    if (
      pages.every(
        (page) => page.text.length === 0 && page.blocks.length === 0 && page.links.length === 0,
      )
    ) {
      throw new Error('No extractable text was found; scanned PDFs require OCR')
    }
    const outputs: PdfToolOutput[] = []
    if (operation.format === 'txt' || operation.format === 'both') {
      outputs.push({
        suffix: '_text.txt',
        bytes: pdfTextPagesTxtBytes(pages),
        mimeType: 'text/plain;charset=utf-8',
        extension: '.txt',
      })
    }
    if (operation.format === 'markdown' || operation.format === 'both') {
      outputs.push({
        suffix: '_text.md',
        bytes: pdfTextPagesMarkdownBytes(pages),
        mimeType: 'text/markdown;charset=utf-8',
        extension: '.md',
      })
    }
    return outputs
  }
  if (operation.kind === 'pdfToMarkdown') {
    if (!operation.pages) throw new Error('Extracted PDF text pages are required')
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const pageIndexes = checkedPageIndexes(document.getPageCount(), operation.pageIndexes)
    const pages = checkedExtractedTextPages(operation.pages)
    if (
      pages.length !== pageIndexes.length ||
      pages.some((page, index) => page.pageNumber !== pageIndexes[index]! + 1)
    ) {
      throw new Error('Extracted text pages do not match the requested pages')
    }
    if (
      pages.every(
        (page) => page.text.length === 0 && page.blocks.length === 0 && page.links.length === 0,
      )
    ) {
      throw new Error('No extractable text was found; scanned PDFs require OCR')
    }
    return [
      {
        suffix: '_converted.md',
        fileName: pdfMarkdownOutputFileName(operation.baseName),
        bytes: pdfTextPagesMarkdownBytes(pages, operation.includePageBreaks),
        mimeType: 'text/markdown;charset=utf-8',
        extension: '.md',
      },
    ]
  }
  if (operation.kind === 'pdfToJson') {
    return [await pdfToJsonOutput(bytes, operation)]
  }
  if (operation.kind === 'jsonToPdf') {
    return [await jsonToPdfBytes(operation)]
  }
  if (operation.kind === 'pdfToXml') {
    return [await pdfToXmlOutput(bytes, operation)]
  }
  if (operation.kind === 'pdfToVideo') {
    const analysis = await analyzePdfBytes(bytes)
    const pageIndexes = checkedPageIndexes(analysis.pageCount, operation.pageIndexes)
    if (
      !Number.isInteger(operation.secondsPerPage) ||
      operation.secondsPerPage < 1 ||
      operation.secondsPerPage > 10
    ) {
      throw new Error('Video page duration must be a whole number from 1 to 10 seconds')
    }
    if (!['480p', '720p', '1080p'].includes(operation.resolution)) {
      throw new Error('PDF video resolution is invalid')
    }
    if (
      !Number.isFinite(operation.transitionSeconds) ||
      operation.transitionSeconds < 0 ||
      operation.transitionSeconds > 1
    ) {
      throw new Error('Video transition must be from 0 to 1 second')
    }
    if (pageIndexes.length > 100 || pageIndexes.length * operation.secondsPerPage > 300) {
      throw new Error('PDF video is limited to 100 pages or 5 minutes')
    }
    return [pdfToVideoOutput(operation)]
  }
  if (operation.kind === 'extractTables') {
    if (!operation.tables || operation.baseName === undefined) {
      throw new Error('Extracted PDF tables are required')
    }
    if (!['csv', 'xlsx', 'both'].includes(operation.format)) {
      throw new Error('Table export format is invalid')
    }
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const pageIndexes = checkedPageIndexes(document.getPageCount(), operation.pageIndexes)
    const tables = checkedExtractedTables(operation.tables)
    const selectedPages = new Set(pageIndexes.map((pageIndex) => pageIndex + 1))
    if (tables.length === 0) throw new Error('No tables were detected on the selected pages')
    if (tables.some((table) => !selectedPages.has(table.pageNumber))) {
      throw new Error('Extracted tables do not match the requested pages')
    }
    const outputs: PdfToolOutput[] = []
    if (operation.format === 'csv' || operation.format === 'both') {
      outputs.push(await pdfTablesCsvOutput(tables, operation.baseName))
    }
    if (operation.format === 'xlsx' || operation.format === 'both') {
      outputs.push({
        suffix: '_tables.xlsx',
        bytes: await pdfTablesXlsxBytes(tables),
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        extension: '.xlsx',
      })
    }
    return outputs
  }
  if (operation.kind === 'pdfToXlsx') {
    if (!operation.tables || operation.baseName === undefined) {
      throw new Error('Extracted PDF tables are required')
    }
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const pageIndexes = checkedPageIndexes(document.getPageCount(), operation.pageIndexes)
    const tables = checkedExtractedTables(operation.tables)
    const selectedPages = new Set(pageIndexes.map((pageIndex) => pageIndex + 1))
    if (tables.length === 0) throw new Error('No tables were detected on the selected pages')
    if (tables.some((table) => !selectedPages.has(table.pageNumber))) {
      throw new Error('Extracted tables do not match the requested pages')
    }
    const baseName = safeExtractedImageBaseName(operation.baseName)
    return [
      {
        suffix: '_converted.xlsx',
        fileName: `${baseName}_converted.xlsx`,
        bytes: await pdfTablesXlsxBytes(tables),
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        extension: '.xlsx',
      },
    ]
  }
  if (operation.kind === 'extractImages') {
    if (!operation.images || operation.baseName === undefined) {
      throw new Error('Extracted image data is required')
    }
    return [
      {
        suffix: '_extracted-images.zip',
        bytes: await extractPdfImagesZipBytes(
          operation.images,
          operation.baseName,
          operation.format,
        ),
        mimeType: 'application/zip',
        extension: '.zip',
      },
    ]
  }
  if (operation.kind === 'pdfToImages') {
    return [await pdfPageImagesOutput(operation)]
  }
  if (operation.kind === 'pdfToCbz') {
    return [await pdfToCbzBytes(operation)]
  }
  if (operation.kind === 'pdfToHtml') {
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    if (document.getPageCount() !== operation.pageCount) {
      throw new Error('PDF page count changed while preparing HTML')
    }
    return [await pdfToHtmlZipBytes(operation)]
  }
  if (operation.kind === 'pdfToEpub') {
    return [await pdfToEpubBytes(bytes, operation)]
  }
  if (operation.kind === 'pdfToPptx') {
    return [await pdfToPptxBytes(operation)]
  }
  if (operation.kind === 'pdfToDocx') {
    return [await pdfToDocxBytes(operation)]
  }
  if (operation.kind === 'pdfToOdt') {
    return [await pdfToOdtBytes(operation)]
  }
  if (operation.kind === 'pdfToPdfa') {
    return [
      {
        suffix: '_PDFA-2b.pdf',
        bytes: await pdfToPdfaBytes(bytes, {
          format: operation.format,
          archiveMode: operation.archiveMode,
          renderDpi: operation.renderDpi,
          imageQuality: operation.imageQuality,
          pageImages: operation.pageImages,
        }),
      },
    ]
  }
  if (operation.kind === 'pdfToRtf') {
    return [pdfToRtfBytes(operation)]
  }
  if (operation.kind === 'removeAnnotations') {
    return [
      {
        suffix: '_annotations_removed.pdf',
        bytes: await removePdfAnnotationsBytes(bytes),
      },
    ]
  }
  if (operation.kind === 'removeBlanks') {
    if (
      !Number.isFinite(operation.threshold) ||
      operation.threshold < 0 ||
      operation.threshold > 255
    ) {
      throw new Error('threshold must be from 0 to 255')
    }
    if (
      !Number.isFinite(operation.whitePercent) ||
      operation.whitePercent <= 0 ||
      operation.whitePercent > 100
    ) {
      throw new Error('whitePercent must be greater than 0 and at most 100')
    }
    if (!operation.blankPageIndexes) throw new Error('Detected blank page indexes are required')
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const blankPageIndexes = checkedPageIndexes(
      document.getPageCount(),
      operation.blankPageIndexes,
      { allowEmpty: true },
    ).sort((left, right) => left - right)
    const blankPages = new Set(blankPageIndexes)
    const nonBlankPageIndexes = document
      .getPageIndices()
      .filter((pageIndex) => !blankPages.has(pageIndex))
    if (nonBlankPageIndexes.length === 0) {
      return [
        {
          suffix: '_allBlankPages.pdf',
          bytes: await retainPdfPagesBytes(bytes, blankPageIndexes),
        },
      ]
    }
    const outputs: PdfToolOutput[] = [
      {
        suffix: '_nonBlankPages.pdf',
        bytes: await retainPdfPagesBytes(bytes, nonBlankPageIndexes),
      },
    ]
    if (operation.includeBlankPages && blankPageIndexes.length > 0) {
      outputs.push({
        suffix: '_blankPages.pdf',
        bytes: await retainPdfPagesBytes(bytes, blankPageIndexes),
      })
    }
    return outputs
  }
  if (operation.kind === 'invertColors') {
    return [
      {
        suffix: '_inverted.pdf',
        bytes: await invertPdfColorsBytes(bytes, operation.pageIndexes),
      },
    ]
  }
  if (operation.kind === 'replaceColors') {
    return [
      {
        suffix: '_recolored.pdf',
        bytes: await replacePdfColorsBytes(
          bytes,
          operation.pageIndexes,
          operation.textColor,
          operation.backgroundColor,
        ),
      },
    ]
  }
  if (operation.kind === 'adjustColors') {
    if (!operation.pageImages) throw new Error('Adjusted page images are required')
    return [
      {
        suffix: '_adjusted.pdf',
        bytes: await overlayAdjustedPdfPagesBytes(
          bytes,
          operation.pageIndexes,
          operation.pageImages,
        ),
      },
    ]
  }
  if (operation.kind === 'rearrange') {
    return [{ suffix: '_rearranged.pdf', bytes: await rearrangePdfPagesBytes(bytes, operation) }]
  }
  if (operation.kind === 'redact') {
    if (!operation.pages && !operation.pageImages) {
      throw new Error('Redacted page images are required')
    }
    return [
      {
        suffix: '_redacted.pdf',
        bytes: operation.pages
          ? await redactSelectedPdfPagesBytes(bytes, operation.pages)
          : await redactPdfPagesBytes(bytes, operation.pageImages!),
      },
    ]
  }
  if (operation.kind === 'comments') {
    return [
      {
        suffix: '_commented.pdf',
        bytes: await addPdfCommentsBytes(bytes, operation.comments),
      },
    ]
  }
  if (operation.kind === 'compress') {
    if (!operation.pageImages) throw new Error('Compressed page images are required')
    if (
      operation.lineArt &&
      (!Number.isFinite(operation.lineArtThreshold ?? 55) ||
        (operation.lineArtThreshold ?? 55) < 0 ||
        (operation.lineArtThreshold ?? 55) > 100)
    ) {
      throw new Error('Line-art threshold must be from 0 to 100')
    }
    if (
      operation.lineArt &&
      (!Number.isInteger(operation.lineArtEdgeLevel ?? 1) ||
        (operation.lineArtEdgeLevel ?? 1) < 1 ||
        (operation.lineArtEdgeLevel ?? 1) > 3)
    ) {
      throw new Error('Line-art edge level must be a whole number from 1 to 3')
    }
    return [
      {
        suffix: '_compressed.pdf',
        bytes: await compressPdfPagesBytes(bytes, operation.pageImages, {
          forceRasterized: operation.lineArt === true,
        }),
      },
    ]
  }
  if (operation.kind === 'flatten') {
    if (operation.mode === 'forms') {
      return [{ suffix: '_flattened.pdf', bytes: await processPdfFormBytes(bytes, 'flatten') }]
    }
    if (!operation.pageImages) throw new Error('Flattened page images are required')
    return [
      { suffix: '_flattened.pdf', bytes: await flattenPdfPagesBytes(bytes, operation.pageImages) },
    ]
  }
  if (operation.kind === 'forms') {
    if (operation.action === 'export') {
      const fields = await listPdfFormFieldsBytes(bytes)
      if (fields.length === 0) throw new Error('No PDF form fields were found')
      return [
        {
          suffix: '_form_data.json',
          bytes: pdfFormFieldsJsonBytes(fields),
          mimeType: 'application/json',
          extension: '.json',
        },
        {
          suffix: '_form_data.csv',
          bytes: pdfFormFieldsCsvBytes(fields),
          mimeType: 'text/csv;charset=utf-8',
          extension: '.csv',
        },
        {
          suffix: '_form_data.xlsx',
          bytes: await pdfFormFieldsXlsxBytes(fields),
          mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          extension: '.xlsx',
        },
      ]
    }
    return [
      {
        suffix:
          operation.action === 'flatten'
            ? '_flattened.pdf'
            : operation.action === 'fill'
              ? '_filled.pdf'
              : operation.action === 'delete'
                ? '_updated.pdf'
                : operation.action === 'modify' || operation.action === 'create'
                  ? '_updated.pdf'
                  : '_unlocked_forms.pdf',
        bytes: await processPdfFormBytes(
          bytes,
          operation.action,
          operation.fields,
          operation.fieldNames,
          operation.modifications,
          operation.creations,
        ),
      },
    ]
  }
  if (operation.kind === 'repair') {
    return [{ suffix: '_repaired.pdf', bytes: await repairPdfBytes(bytes) }]
  }
  if (operation.kind === 'decompress') {
    return [{ suffix: '_decompressed.pdf', bytes: await decompressPdfBytes(bytes) }]
  }
  if (operation.kind === 'removeSignatures') {
    return [{ suffix: '_unsigned.pdf', bytes: await removePdfSignaturesBytes(bytes) }]
  }
  if (operation.kind === 'timestamp') {
    if (!operation.timestampedBytes?.length) throw new Error('Prepared timestamped PDF is required')
    return [{ suffix: '_timestamped.pdf', bytes: operation.timestampedBytes }]
  }
  if (operation.kind === 'certificateSign') {
    const result = await signPdfWithP12Bytes(bytes, operation)
    return [{ suffix: '_signed.pdf', bytes: result.bytes }]
  }
  if (operation.kind === 'sanitize') {
    return [{ suffix: '_sanitized.pdf', bytes: await sanitizePdfBytes(bytes, operation) }]
  }
  if (operation.kind === 'overlay') {
    return [{ suffix: '_overlayed.pdf', bytes: await overlayPdfBytes(bytes, operation) }]
  }
  if (operation.kind === 'overlayImage') {
    return [
      {
        suffix: '_image_overlayed.pdf',
        bytes: await overlayImagePdfBytes(bytes, operation),
      },
    ]
  }
  if (operation.kind === 'imagesToPdf') {
    return [
      {
        suffix: operation.appendToCurrent ? '_with_images.pdf' : '_from_images.pdf',
        bytes: operation.appendToCurrent
          ? await appendImagesToPdfBytes(
              bytes,
              operation.images,
              operation.fitOption,
              operation.autoRotate,
            )
          : await imagesToPdfBytes(operation.images, operation.fitOption, operation.autoRotate),
      },
    ]
  }
  if (operation.kind === 'cbzToPdf') {
    return [
      {
        suffix: '_converted.pdf',
        fileName: cbzPdfOutputFileName(operation.baseName),
        bytes: await imagesToPdfBytes(operation.images, operation.fitOption, operation.autoRotate),
      },
    ]
  }
  if (operation.kind === 'emailToPdf') {
    if (operation.outputFormat === 'html') {
      if (!operation.htmlBytes?.length) throw new Error('Prepared email HTML is required')
      return [
        {
          suffix: '_converted.html',
          fileName: emailDocumentOutputFileName(operation.baseName, 'html'),
          bytes: operation.htmlBytes,
          mimeType: 'text/html',
          extension: '.html',
        },
      ]
    }
    if (!operation.pages?.length) throw new Error('Prepared email pages are required')
    let output = await imagesToPdfBytes(operation.pages, 'fitDocumentToImage', false)
    if (operation.attachments?.length) {
      output = await addPdfAttachmentsBytes(output, operation.attachments)
    }
    return [
      {
        suffix: '_converted.pdf',
        fileName: emailDocumentOutputFileName(operation.baseName, 'pdf'),
        bytes: output,
      },
    ]
  }
  if (operation.kind === 'epubToPdf') {
    if (!operation.pages.length) throw new Error('Prepared EPUB pages are required')
    return [
      {
        suffix: '_converted.pdf',
        fileName: epubPdfOutputFileName(operation.baseName),
        bytes: await epubPagesToPdfBytes(operation),
      },
    ]
  }
  if (operation.kind === 'htmlToPdf') {
    if (!operation.pages.length) throw new Error('Prepared HTML pages are required')
    return [
      {
        suffix: '_converted.pdf',
        fileName: htmlPdfOutputFileName(operation.baseName),
        bytes: await htmlPagesToPdfBytes(operation),
      },
    ]
  }
  if (operation.kind === 'createPdf') {
    if (!operation.pages.length) throw new Error('Prepared PDF pages are required')
    return [
      {
        suffix: '.pdf',
        fileName: createdPdfOutputFileName(operation.baseName),
        bytes: await htmlPagesToPdfBytes(operation),
      },
    ]
  }
  if (operation.kind === 'markdownToPdf') {
    if (!operation.pages.length) throw new Error('Prepared Markdown pages are required')
    return [
      {
        suffix: '_converted.pdf',
        fileName: markdownPdfOutputFileName(operation.baseName),
        bytes: await htmlPagesToPdfBytes(operation),
      },
    ]
  }
  if (operation.kind === 'metadata') {
    return [{ suffix: '_metadata.pdf', bytes: await updatePdfMetadataBytes(bytes, operation) }]
  }
  if (operation.kind === 'filterPages') {
    if (!['text', 'image', 'pageSize', 'orientation', 'rotation'].includes(operation.criterion)) {
      throw new Error('Page filter criterion is invalid')
    }
    if (!['keep', 'remove'].includes(operation.action)) {
      throw new Error('Page filter action is invalid')
    }
    if (operation.criterion === 'text' && !operation.text?.trim()) {
      throw new Error('Enter text to find')
    }
    const document = await PDFDocument.load(bytes, { updateMetadata: false })
    const matchedPageIndexes =
      operation.criterion === 'text' || operation.criterion === 'image'
        ? operation.matchedPageIndexes
        : geometricFilterMatchedPageIndexes(document, operation)
    if (!matchedPageIndexes) throw new Error('Content filter analysis is required')
    const outputPageIndexes = contentFilterOutputPageIndexes(
      document.getPageCount(),
      operation.pageIndexes,
      matchedPageIndexes,
      operation.action,
    )
    return [
      {
        suffix: operation.action === 'keep' ? '_filtered_pages.pdf' : '_content_removed.pdf',
        bytes: await retainPdfPagesBytes(bytes, outputPageIndexes),
      },
    ]
  }
  if (operation.kind === 'filterDocuments') {
    if (!operation.currentFileName.trim()) throw new Error('Current PDF file name is required')
    if (operation.documents.length > 100) {
      throw new Error('Document filter supports at most 100 additional PDF files')
    }
    if (
      operation.documents.some(
        (document) => !document.fileName.trim() || !(document.bytes instanceof Uint8Array),
      )
    ) {
      throw new Error('Document filter inputs are invalid')
    }
    const documents: PdfDocumentFilterInput[] = [
      {
        fileName: operation.currentFileName,
        bytes: bytes instanceof Uint8Array ? new Uint8Array(bytes) : new Uint8Array(bytes.slice(0)),
        contentMatched: operation.currentContentMatched,
      },
      ...operation.documents,
    ]
    if (operation.criterion === 'text' && !operation.text?.trim()) {
      throw new Error('Enter text to find in PDF documents')
    }
    const contentCriterion = operation.criterion === 'text' || operation.criterion === 'image'
    if (contentCriterion && documents.some((document) => document.contentMatched === undefined)) {
      throw new Error('Document content filter analysis is required')
    }
    const matches = await Promise.all(
      documents.map(async (document) => ({
        ...document,
        matched: contentCriterion
          ? document.contentMatched === true
          : await documentMatchesFilterBytes(document.bytes, operation),
      })),
    )
    const filtered = matches.filter((document) => document.matched)
    if (filtered.length === 0) throw new Error('No PDF documents matched the selected filter')
    const usedNames = new Set<string>()
    return filtered.map((document, index) => ({
      suffix: '_filtered.pdf',
      fileName: uniquePdfOutputName(document.fileName, index, usedNames),
      bytes: document.bytes,
    }))
  }
  if (operation.kind === 'bookmarks') {
    return [
      {
        suffix: '_with_toc.pdf',
        bytes: await setPdfBookmarksBytes(bytes, operation.bookmarks),
      },
    ]
  }
  if (operation.action === 'add') {
    return [
      {
        suffix: '_with_attachments.pdf',
        bytes: await addPdfAttachmentsBytes(bytes, operation.attachments),
      },
    ]
  }
  if (operation.action === 'extract') {
    return [
      {
        suffix: '_attachments.zip',
        bytes: await extractPdfAttachmentsZipBytes(bytes),
        mimeType: 'application/zip',
        extension: '.zip',
      },
    ]
  }
  if (operation.action === 'rename') {
    return [
      {
        suffix: '_attachment_renamed.pdf',
        bytes: await renamePdfAttachmentBytes(bytes, operation.attachmentName, operation.newName),
      },
    ]
  }
  return [
    {
      suffix: '_attachment_deleted.pdf',
      bytes: await deletePdfAttachmentBytes(bytes, operation.attachmentName),
    },
  ]
}
