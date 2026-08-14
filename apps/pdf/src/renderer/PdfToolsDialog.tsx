import { useEffect, useMemo, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { createPortal } from 'react-dom'
import type { Lang } from '@genoffice/i18n'
import {
  analyzePdfAnnotationsBytes,
  analyzePdfBytes,
  analyzePdfFontsBytes,
  analyzePdfSecurityBytes,
  auditPdfJavaScriptBytes,
  auditPdfSignaturesBytes,
  PDF_TIMESTAMP_TSA_PRESETS,
  preflightPdfBytes,
  listPdfAttachmentsBytes,
  listPdfBookmarksBytes,
  listPdfFormFieldsBytes,
  type PdfAttachmentAction,
  type PdfAttachmentInfo,
  type PdfAnnotationReport,
  type PdfAnalysis,
  type PdfArchiveMode,
  type PdfAutoRenameStrategy,
  type PdfBookmark,
  type PdfColorAdjustments,
  type PdfComicImageFormat,
  type PdfCommentInput,
  type PdfContentFilterAction,
  type PdfContentFilterCriterion,
  type PdfDocumentFilterCriterion,
  type PdfExtractImageFormat,
  type PdfFlattenMode,
  type PdfFontReport,
  type PdfOrientation,
  type PdfCreatableFormFieldType,
  type PdfFormAction,
  type PdfFormFieldCreation,
  type PdfFormFieldInfo,
  type PdfFormFieldModification,
  type PdfFormFieldValue,
  type PdfImageOverlayPosition,
  type PdfImagesToPdfFit,
  type PdfOverlayMode,
  type PdfOverlayPosition,
  type PdfPageRotation,
  type PdfPageNumberFont,
  type PdfPageNumberMargin,
  type PdfPageNumberPosition,
  type PdfPageImageColorMode,
  type PdfPageImageFormat,
  type PdfPageImageOutputMode,
  type PdfPageFilterComparator,
  type PdfPageFilterRotation,
  type PdfEpubMode,
  type PdfDocxMode,
  type PdfPageSize,
  type PdfPasswordAlgorithm,
  type PdfPasswordPermissions,
  type PdfPptxMode,
  type PdfPreflightFinding,
  type PdfPreflightReport,
  type PdfPipelineStep,
  type PdfRearrangeMode,
  type PdfRedactionArea,
  type PdfScannerColorspace,
  type PdfScannerQuality,
  type PdfScannerRotation,
  type PdfSecurityReport,
  type PdfSectionArrangement,
  type PdfSplitMode,
  type PdfTableExportFormat,
  type PdfTextExportFormat,
  type PdfVideoResolution,
  type PdfJavaScriptAction,
  type PdfJavaScriptAudit,
  type PdfSignatureAudit,
  type PdfSignatureAuditItem,
  type PdfMetadataCustomField,
  type PdfMetadataTrapped,
  type PdfMetadataValues,
  type PdfOcrLanguage,
  type PdfOcrMode,
  type PdfToolOperation,
} from '@genoffice/pdf-tools'
import type { TFunc } from './i18n/locale'
import { prepareComicArchiveForPdf } from './cbz-to-pdf'
import {
  prepareEmailDocumentForPdf,
  prepareEmailMarkup,
  type EmailDocumentLabels,
} from './email-to-pdf'
import { prepareEpubDocumentForPdf } from './epub-to-pdf'
import { prepareLocalWebDocumentForPdf, prepareRemoteWebDocumentForPdf } from './html-to-pdf'
import { prepareImagesForPdf, type PdfImageColorMode } from './images-to-pdf'
import { prepareMarkdownDocumentForPdf } from './markdown-to-pdf'
import {
  MAX_BOOKMARK_COUNT,
  MAX_BOOKMARK_DEPTH,
  MAX_BOOKMARK_JSON_BYTES,
  MAX_BOOKMARK_TITLE_LENGTH,
  bookmarkJsonText,
  bookmarkTreeCount,
  parseBookmarkJson,
} from './bookmark-json'
import type { PdfOcrProgress } from './ocr'
import type { PdfVideoProgress } from './pdf-to-video'
import { CameraCapture } from './CameraCapture'
import { MobileScannerCapture } from './MobileScannerCapture'
import { RedactionAreaPreview } from './RedactionAreaPreview'

export type PdfToolKind =
  | Exclude<PdfToolOperation['kind'], 'createPdf'>
  | 'bulkReplaceText'
  | 'info'
  | 'fontInfo'
  | 'annotationInfo'
  | 'securityInfo'
  | 'preflight'
  | 'javascriptAudit'
  | 'signatureAudit'

export interface PdfBulkTextReplacementRequest {
  rules: { find: string; replace: string }[]
  pageIndexes: number[]
  caseSensitive: boolean
  wholeWord: boolean
}

const EN = {
  tab: 'Tools',
  title: 'PDF tools',
  split: 'Split',
  merge: 'Merge PDFs',
  imagesToPdf: 'Images / SVG to PDF',
  cbzToPdf: 'CBZ / CBR to PDF',
  emailToPdf: 'Email to PDF',
  epubToPdf: 'EPUB to PDF',
  htmlToPdf: 'Website / HTML to PDF',
  markdownToPdf: 'Markdown / ZIP to PDF',
  pdfToImages: 'PDF to images',
  pdfToVideo: 'PDF to video',
  pdfToCbz: 'PDF to CBZ',
  pdfToHtml: 'PDF to HTML',
  pdfToEpub: 'PDF to EPUB',
  pdfToPptx: 'PDF to PowerPoint',
  pdfToDocx: 'PDF to Word',
  pdfToOdt: 'PDF to ODT',
  pdfToRtf: 'PDF to RTF',
  pdfToPdfa: 'Archive PDF (PDF/A)',
  pdfToMarkdown: 'PDF to Markdown',
  pdfToXlsx: 'PDF to Excel',
  pdfToXml: 'PDF to XML',
  compare: 'Compare PDFs',
  extractPages: 'Extract pages',
  splitSections: 'Split sections',
  crop: 'Crop',
  scale: 'Scale',
  nup: 'N-up',
  booklet: 'Booklet',
  poster: 'Poster',
  singlePage: 'Long page',
  rotatePages: 'Rotate pages',
  autoRotate: 'Auto rotate',
  deskew: 'Deskew scans',
  autoRename: 'Auto rename',
  pageNumbers: 'Page numbers',
  scannerEffect: 'Scanner effect',
  scannerImageSplit: 'Split scanned photos',
  autoSplit: 'QR auto split',
  removePages: 'Remove pages',
  removeImages: 'Remove images',
  bulkReplaceText: 'Replace text',
  extractText: 'Export text',
  pdfToJson: 'PDF to JSON',
  jsonToPdf: 'JSON to PDF',
  extractTables: 'Extract tables',
  extractImages: 'Extract images',
  removeAnnotations: 'Remove annotations',
  removeBlanks: 'Remove blank pages',
  invertColors: 'Invert colors',
  replaceColors: 'Recolor / contrast',
  adjustColors: 'Adjust colors',
  rearrange: 'Rearrange',
  redact: 'Secure redaction',
  comments: 'Add comments',
  compress: 'Compress PDF',
  flatten: 'Flatten PDF',
  forms: 'Forms',
  repair: 'Repair PDF',
  decompress: 'Decompress PDF',
  certificateSign: 'Certificate signing',
  timestamp: 'Document timestamp',
  removeSignatures: 'Remove digital signatures',
  password: 'Password protection',
  ocr: 'OCR',
  sanitize: 'Sanitize',
  pipeline: 'Processing pipeline',
  overlay: 'Overlay',
  overlayImage: 'Add image',
  filterPages: 'Filter pages',
  filterDocuments: 'Filter PDFs',
  attachments: 'Attachments',
  bookmarks: 'Bookmarks',
  metadata: 'Edit metadata',
  info: 'PDF info',
  preflight: 'Local preflight',
  javascriptAudit: 'JavaScript audit',
  signatureAudit: 'Validate signatures',
  fontInfo: 'Font report',
  annotationInfo: 'Annotation report',
  securityInfo: 'Security & permissions',
  protectPdf: 'Protect PDF',
  unlockPdf: 'Unlock PDF',
  openPassword: 'Open password',
  ownerPassword: 'Owner password',
  currentPassword: 'Current password',
  confirmPassword: 'Confirm password',
  encryption: 'Encryption',
  aes256: 'AES-256',
  rc4: 'RC4-128',
  aesRecommended: 'Recommended for modern PDF readers. Supports Unicode passwords.',
  rc4Compatibility: 'Legacy compatibility only. Password characters are limited.',
  pdfPermissions: 'Allowed actions',
  allowPrinting: 'Print',
  allowHighQualityPrint: 'High-quality print',
  allowModifying: 'Modify content',
  allowCopying: 'Copy text and images',
  allowAnnotating: 'Add annotations',
  allowFillingForms: 'Fill forms and sign',
  allowExtraction: 'Accessibility extraction',
  allowAssembly: 'Assemble pages',
  permissionAdvisory:
    'PDF permissions are advisory and depend on the reader. Use an open password for confidentiality.',
  passwordRewritesPdf:
    'Protecting or unlocking rewrites the PDF and invalidates existing signatures.',
  unlockHint: 'Exports an unencrypted copy. The current document remains unchanged.',
  passwordRequired: 'Enter at least an open password or owner password.',
  passwordMismatch: 'The password confirmation does not match.',
  ownerPasswordRequired: 'A separate owner password is required when actions are restricted.',
  ocrMode: 'OCR mode',
  ocrSkipText: 'Auto · skip text pages',
  ocrForce: 'Force · replace every page',
  ocrStrict: 'Strict · stop if text exists',
  ocrLanguages: 'Recognition languages',
  ocrEnglish: 'English',
  ocrChinese: 'Simplified Chinese',
  ocrRenderDpi: 'Rendering DPI',
  ocrClean: 'Clean scan before recognition',
  ocrCleanHint: 'Improves contrast and converts pages to grayscale before OCR.',
  ocrSidecar: 'Include sidecar text',
  ocrSidecarHint: 'Downloads the searchable PDF and page-by-page plain text together as a ZIP.',
  ocrLocalHint: 'Recognition runs locally. Language data is cached after the first load.',
  ocrRewritesPdf: 'OCR rewrites processed pages and invalidates existing signatures.',
  ocrNoLanguage: 'Choose at least one recognition language.',
  invalidOcrDpi: 'OCR rendering DPI must be a whole number from 100 to 300.',
  runOcr: 'Run OCR',
  ocrAnalyzing: 'Checking existing text layers…',
  ocrLoading: 'Loading local OCR engine…',
  ocrRecognizing: 'Recognizing page {page} of {total}…',
  ocrFinishing: 'Building searchable PDF…',
  certificateFile: 'P12 / PFX certificate',
  timestampAuthority: 'Time Stamp Authority',
  timestampPrivacyHint:
    'Only a SHA-256 digest is sent to the selected authority. The PDF never leaves this device.',
  timestampSecurityHint:
    'The output contains a digital signature. Editing it later invalidates the timestamp.',
  runTimestamp: 'Add timestamp',
  chooseCertificate: 'Choose certificate',
  noCertificate: 'No certificate selected',
  certificatePassword: 'Certificate password',
  signerNameInput: 'Signer name',
  contactInfo: 'Contact information',
  certificateSignerPlaceholder: 'Optional PDF display name',
  certificateReasonPlaceholder: 'Document approval',
  certificateLocationPlaceholder: 'City or organization',
  certificateContactPlaceholder: 'Email or phone (optional)',
  certificateRequired: 'Choose a P12 or PFX certificate file.',
  certificateTooLarge: 'The certificate file must be 16 MB or smaller.',
  certificateLocalHint:
    'The certificate and password are used only for this local signing operation and are not stored.',
  certificateSecurityHint:
    'A signed copy is created. Any later edit to that copy will invalidate its digital signature.',
  runCertificateSign: 'Sign PDF',
  pdfaFormat: 'Archive standard',
  pdfaMode: 'Archive mode',
  pdfaAuto: 'Auto preserve',
  pdfaRaster: 'Image archive',
  pdfaRenderDpi: 'Archive rendering DPI',
  pdfaImageQuality: 'Archive image quality',
  pdfaAutoHint:
    'Preserves selectable text and vector graphics when all used fonts are embedded. Otherwise it automatically uses the image archive fallback. Forms, links, attachments, and scripts are removed locally.',
  pdfaRasterHint:
    'Always creates an image-based PDF/A-2b copy. This maximizes visual consistency but selectable text and vector graphics are flattened.',
  invalidPdfa: 'Choose a whole-number DPI from 72 to 600 and image quality from 10% to 100%.',
  splitMethod: 'Split method',
  afterPages: 'After specified pages',
  byFileSize: 'By file size',
  pagesPerDocument: 'Pages per document',
  documentCount: 'Number of documents',
  chapters: 'By bookmarks / chapters',
  splitAfter: 'Split after',
  splitPlaceholder: '2, 5 or all',
  splitValue: 'Amount',
  targetFileSize: 'Target size',
  chapterLevels: 'Bookmark levels',
  allowDuplicateChapters: 'Keep same-page bookmarks as separate files',
  sectionArrangement: 'Section order',
  rowMajor: 'Across rows',
  columnMajor: 'Down columns',
  mergeSections: 'Merge sections into one PDF',
  cropMode: 'Crop mode',
  cropAuto: 'Auto detect',
  cropManual: 'Manual margins',
  whiteThreshold: 'White threshold',
  cropPadding: 'Keep padding',
  autoCropHint: 'Detects white page margins locally and keeps the original vector content.',
  top: 'Top',
  right: 'Right',
  bottom: 'Bottom',
  left: 'Left',
  points: 'pt',
  pageSize: 'Page size',
  contentScale: 'Content scale',
  pagesPerSheet: 'Pages/sheet',
  customLayout: 'Custom layout',
  nupPageOrder: 'Page order',
  readingDirection: 'Reading direction',
  leftToRight: 'Left to right',
  outerMargins: 'Outer margins',
  innerMargin: 'Inner margin',
  borderWidth: 'Border width',
  orientation: 'Orientation',
  portrait: 'Portrait',
  landscape: 'Landscape',
  border: 'Border',
  spine: 'Spine',
  spineLeft: 'Left binding',
  spineRight: 'Right binding',
  gutter: 'Gutter',
  duplexPass: 'Print sides',
  both: 'Both',
  front: 'Front',
  back: 'Back',
  shortEdge: 'Short-edge flip',
  columns: 'Columns',
  rows: 'Rows',
  rightToLeft: 'Right to left',
  joinDirection: 'Join direction',
  vertical: 'Vertical',
  horizontal: 'Horizontal',
  pages: 'Pages',
  pagesPlaceholder: 'all or 1, 3-5',
  colorMode: 'Color mode',
  highContrast: 'High contrast',
  customColors: 'Custom colors',
  colorPreset: 'Color preset',
  whiteOnBlack: 'White on black',
  blackOnWhite: 'Black on white',
  yellowOnBlack: 'Yellow on black',
  greenOnBlack: 'Green on black',
  textColor: 'Text color',
  backgroundColor: 'Background',
  contrast: 'Contrast',
  brightness: 'Brightness',
  saturation: 'Saturation',
  rgbChannels: 'RGB channels',
  redChannel: 'Red',
  greenChannel: 'Green',
  blueChannel: 'Blue',
  whitenessThreshold: 'Whiteness tolerance',
  whitePercent: 'Minimum white area',
  includeBlankPages: 'Export detected blank pages',
  outputFormat: 'Output format',
  outputMode: 'Output mode',
  separateImages: 'Separate images',
  singleLongImage: 'Single long image',
  includeAnnotations: 'Include comments and form values',
  pdfToImagesHint: 'Pages are rendered locally. Multiple images are downloaded as one ZIP file.',
  videoDuration: 'Seconds per page',
  videoResolution: 'Video resolution',
  videoTransition: 'Fade transition',
  videoEncoding: 'Encoding page {page} of {total}…',
  pdfToVideoHint:
    'Creates a local WebM slideshow without FFmpeg or a server. Videos are limited to 100 pages or 5 minutes.',
  pdfToCbzHint:
    'Creates a comic-book archive from locally rendered pages in natural reading order.',
  cbzFile: 'Comic archive',
  chooseCbz: 'Choose CBZ / CBR',
  replaceCbz: 'Replace archive',
  selectCbz: 'Choose a CBZ, CBR, ZIP, or RAR comic archive.',
  cbzToPdfHint:
    'ZIP and RAR comics are unpacked and naturally sorted locally. Hidden and non-image entries are ignored.',
  emailFile: 'Email file',
  chooseEmail: 'Choose EML',
  replaceEmail: 'Replace email',
  selectEmail: 'Choose an EML email file.',
  emailOutputFormat: 'Email output',
  emailPdf: 'PDF document',
  emailHtml: 'Offline HTML',
  emailIncludeAttachments: 'Embed original attachments',
  emailMaxAttachmentSize: 'Attachment size limit',
  emailIncludeAllRecipients: 'Include CC and BCC',
  emailLocalHint:
    'The email is parsed locally. Remote images and active content are removed before rendering.',
  emailFromLabel: 'From',
  emailToLabel: 'To',
  emailCcLabel: 'Cc',
  emailBccLabel: 'Bcc',
  emailDateLabel: 'Date',
  emailAttachmentsLabel: 'Attachments',
  emailEmptyBody: 'Email content could not be parsed.',
  emailRemoteImageRemoved: 'remote image removed',
  emailUntitled: 'Untitled email',
  invalidEmailAttachmentSize: 'Attachment size limit must be a whole number from 1 to 50 MB.',
  epubFile: 'EPUB file',
  chooseEpub: 'Choose EPUB',
  replaceEpub: 'Replace EPUB',
  selectEpub: 'Choose an EPUB ebook file.',
  epubEmbedFonts: 'Embed book fonts',
  epubIncludeToc: 'Add table of contents',
  epubIncludePageNumbers: 'Add page numbers',
  epubOptimize: 'Optimize for ebook readers',
  epubLocalHint:
    'The reading order, chapters and local images are parsed on this device. Remote and active content are removed.',
  epubTocLabel: 'Table of contents',
  epubUntitled: 'Untitled chapter',
  jsonFile: 'Structured PDF JSON',
  chooseJson: 'Choose JSON',
  replaceJson: 'Replace JSON',
  selectJson: 'Choose a GenOffice structured PDF JSON version 1 file.',
  jsonToPdfHint:
    'Rebuilds GenOffice structured PDF JSON version 1 locally. Layout mode restores text positions; semantic mode reflows text. Images, annotations, links, and form geometry are not present in this JSON schema.',
  htmlFile: 'Web document',
  htmlSource: 'Source',
  htmlSourceFile: 'Local file',
  htmlSourceWebsite: 'Website URL',
  htmlWebsiteUrl: 'Website address',
  htmlWebsitePlaceholder: 'https://example.com/article',
  chooseHtml: 'Choose HTML / ZIP',
  replaceHtml: 'Replace file',
  selectHtml: 'Choose an HTML, HTM, or ZIP web document.',
  selectWebsiteUrl: 'Enter a valid website URL.',
  htmlZoom: 'Page zoom',
  htmlIncludePageNumbers: 'Add page numbers',
  htmlLocalHint:
    'HTML, CSS and local images are processed on this device. Remote and active content are removed.',
  htmlRemoteHint:
    'The public page, styles and images are fetched securely. Active and private-network content is removed.',
  htmlUntitled: 'Untitled web document',
  invalidHtmlZoom: 'Page zoom must be a whole number from 50% to 200%.',
  markdownFile: 'Markdown document',
  chooseMarkdown: 'Choose Markdown / ZIP',
  replaceMarkdown: 'Replace file',
  selectMarkdown: 'Choose an MD, Markdown, or ZIP document.',
  markdownZoom: 'Page zoom',
  markdownIncludePageNumbers: 'Add page numbers',
  markdownLocalHint:
    'GFM tables, lists, code blocks and local ZIP images are processed on this device. Raw active content is removed.',
  markdownUntitled: 'Untitled Markdown document',
  invalidMarkdownZoom: 'Page zoom must be a whole number from 50% to 200%.',
  pdfToHtmlHint:
    'Creates an offline ZIP with page images and a searchable text layer. No file is uploaded.',
  epubMode: 'Book layout',
  epubReflowable: 'Reading',
  epubFixed: 'Page fidelity',
  pdfToEpubHint:
    'Reading mode reflows selectable text; page fidelity mode preserves appearance with local page images.',
  pptxMode: 'PowerPoint mode',
  pptxEditableText: 'Editable text',
  pptxFidelity: 'Page fidelity',
  pdfToPptxHint:
    'Editable text creates positioned text boxes; images and complex graphics are omitted. Page fidelity preserves the full appearance as a page image.',
  docxMode: 'Word mode',
  odtMode: 'OpenDocument mode',
  docxEditableText: 'Editable text',
  docxFidelity: 'Page fidelity',
  pdfToDocxHint:
    'Editable text rebuilds locally editable paragraphs in reading order. Page fidelity preserves the full appearance as one image per page.',
  pdfToOdtHint:
    'Creates a standards-based ODT locally. Editable text preserves reading order and common text styles; page fidelity embeds one image per page.',
  pdfToRtfHint:
    'Exports editable rich text locally in reading order, preserving common fonts, sizes, bold, italic, Unicode text, and page breaks.',
  pdfToMarkdownPageBreaks: 'Preserve page boundaries',
  pdfToMarkdownHint:
    'Rebuilds headings, paragraphs, lists and links in reading order. Scanned pages require OCR first.',
  pdfToXlsxHint:
    'Detects aligned tables locally and creates one editable Excel worksheet per table.',
  pdfToXmlHint:
    'Exports a standards-based XML document locally with metadata, bookmarks, forms, links, and selected pages.',
  pageOrder: 'Page order',
  pageOrderPlaceholder: '3, 1, 2, 2',
  angle: 'Angle',
  inferUndetected: 'Infer sparse pages from document',
  deskewMaxAngle: 'Maximum skew angle',
  deskewHint:
    'Locally detects small scan angles. Only corrected pages are rasterized; unchanged pages keep their original PDF content.',
  renameStrategy: 'Title source',
  largestHeading: 'Largest heading',
  firstText: 'First text',
  pageNumberPosition: 'Position',
  pageNumberMargin: 'Margin',
  small: 'Small',
  medium: 'Medium',
  large: 'Large',
  xLarge: 'Extra large',
  pageNumberFont: 'Font',
  helvetica: 'Helvetica',
  times: 'Times',
  courier: 'Courier',
  fontSize: 'Font size',
  fontColor: 'Text color',
  startingNumber: 'Starting number',
  zeroPad: 'Zero-padding',
  textPattern: 'Text format',
  textPatternPlaceholder: 'Page {n} of {total}',
  pageNumbersHint: 'Variables: {n}, {total}, and {filename}.',
  scanQuality: 'Scan quality',
  scanHigh: 'High',
  scanMedium: 'Medium',
  scanLow: 'Low',
  scanCustom: 'Custom',
  scanRotation: 'Page rotation',
  scanNone: 'None',
  scanSlight: 'Slight',
  scanModerate: 'Moderate',
  scanSevere: 'Severe',
  scanColorspace: 'Color space',
  grayscale: 'Grayscale',
  color: 'Color',
  paperBorder: 'Paper border',
  baseRotation: 'Base rotation',
  rotationVariance: 'Rotation variance',
  blur: 'Blur',
  noise: 'Noise',
  yellowedPaper: 'Yellowed paper',
  textureSeed: 'Texture seed',
  scanAngleThreshold: 'Rotation threshold',
  scanTolerance: 'Background tolerance',
  minimumPhotoArea: 'Minimum photo area',
  minimumForegroundArea: 'Minimum foreground area',
  detectionBorder: 'Detection border',
  mode: 'Mode',
  custom: 'Custom order',
  reverse: 'Reverse',
  oddEven: 'Odd pages, then even',
  duplex: 'Duplex order',
  removeFirst: 'Remove first',
  removeLast: 'Remove last',
  removeFirstAndLast: 'Remove first & last',
  duplicate: 'Duplicate pages',
  copies: 'Copies per page',
  flattenMode: 'Flatten mode',
  flattenPages: 'Entire document',
  flattenFormsOnly: 'Forms only',
  renderDpi: 'Rendering DPI',
  imageQuality: 'Image quality',
  lineArt: 'Convert to line art',
  lineArtThreshold: 'Line sensitivity',
  lineArtEdgeLevel: 'Edge strength',
  redactPatterns: 'Text or patterns',
  redactPatternsPlaceholder: 'One term per line',
  redactTextMode: 'Text matching',
  redactAreaMode: 'Page areas',
  redactionMode: 'Redaction mode',
  addRedactionArea: 'Add area',
  redactionArea: 'Area',
  deleteRedactionArea: 'Delete area',
  redactionPreview: 'Redaction area preview',
  resizeRedactionArea: 'Resize redaction area',
  redactLeft: 'Left',
  redactTop: 'Top',
  redactWidth: 'Width',
  redactHeight: 'Height',
  redactionAreaHint:
    'Coordinates are percentages from the top-left. Affected pages are securely flattened; other pages keep their original PDF content.',
  useRegex: 'Use regular expressions',
  wholeWord: 'Match whole words only',
  redactColor: 'Redaction color',
  redactPadding: 'Extra padding',
  pipelineStep: 'Pipeline step',
  addPipelineStep: 'Add step',
  deletePipelineStep: 'Delete step',
  flattenForms: 'Flatten forms',
  pipelineHint:
    'Runs each step locally in the listed order and exports one PDF. Reorder steps to control the result.',
  addComment: 'Add comment',
  comment: 'Comment',
  commentText: 'Comment text',
  commentTextPlaceholder: 'Enter the note shown in PDF readers',
  anchorText: 'Anchor text (optional)',
  anchorTextPlaceholder: 'Place beside the first match on this page',
  fallbackPosition: 'Fallback position',
  positionX: 'X',
  positionY: 'Y',
  deleteComment: 'Delete comment',
  formAction: 'Form action',
  exportFormData: 'Export data',
  fillFields: 'Fill fields',
  createFields: 'Create fields',
  flattenFields: 'Flatten fields',
  unlock: 'Unlock fields',
  modifyFields: 'Edit properties',
  deleteFields: 'Delete fields',
  selectAllFields: 'Select all',
  clearFieldSelection: 'Clear',
  selectedFields: 'Selected fields',
  modifiedFields: 'Modified fields',
  createdFields: 'New fields',
  addField: 'Add field',
  removeField: 'Remove field',
  fieldName: 'Field name',
  fieldLabel: 'Field label',
  fieldType: 'Field type',
  fieldPage: 'Page',
  fieldWidth: 'Width',
  fieldHeight: 'Height',
  defaultValue: 'Default value',
  defaultValues: 'Default values (one per line)',
  defaultChecked: 'Checked by default',
  multilineField: 'Multiple lines',
  optionSpacing: 'Option spacing',
  choiceOptions: 'Options (one per line)',
  multiselectField: 'Allow multiple selections',
  formCreateHint: 'Coordinates use PDF points from the page’s bottom-left corner.',
  loadingFormFields: 'Reading form fields…',
  formFieldsLoadFailed: 'Could not read PDF form fields.',
  noFormFields: 'No AcroForm fields were found.',
  readOnlyField: 'Read-only',
  requiredField: 'Required',
  selectEmpty: 'No selection',
  unsupportedFormField: 'This field type cannot be filled here.',
  formText: 'Text',
  formCheckbox: 'Checkbox',
  formRadio: 'Radio group',
  formDropdown: 'Dropdown',
  formOptionList: 'Option list',
  formButton: 'Button',
  formSignature: 'Signature',
  formUnknown: 'Unknown field',
  textExportFormat: 'Output format',
  plainText: 'Plain text',
  markdown: 'Markdown',
  exportBoth: 'Both',
  textExportHint: 'Exports selected pages locally. Scanned pages require OCR first.',
  structuredExportMode: 'Structure detail',
  semanticJson: 'Compact semantics',
  layoutJson: 'Full layout',
  pdfToJsonHint:
    'Exports metadata, bookmarks, forms, links, and selected pages locally. Full layout also includes text coordinates and fonts.',
  tableExportFormat: 'Output format',
  csv: 'CSV',
  excel: 'Excel',
  includeTwoColumnTables: 'Detect text-only two-column tables',
  tableExportHint:
    'Detects aligned tables locally. Complex, rotated, or scanned tables may need review.',
  downloadDivider: 'Download divider page',
  duplexMode: 'Duplex scanning',
  autoSplitHint:
    'Scans pages locally for compatible QR dividers, removes them, and exports the documents as a ZIP.',
  removeJavaScript: 'Remove JavaScript',
  removeEmbeddedFiles: 'Remove embedded files',
  removeXmpMetadata: 'Remove XMP metadata',
  removeMetadata: 'Remove document metadata',
  removeLinks: 'Remove external link actions',
  overlayFiles: 'Overlay PDFs',
  chooseFiles: 'Choose PDFs',
  noFiles: 'No files selected',
  mergeFiles: 'PDF order',
  currentPdf: 'Current PDF',
  addMergeFiles: 'Add PDFs',
  deleteFile: 'Remove file',
  comparisonFile: 'Comparison PDF',
  chooseComparisonFile: 'Choose PDF',
  comparisonTolerance: 'Difference tolerance',
  overlayMode: 'Page assignment',
  sequential: 'Sequential pages',
  interleaved: 'Alternate files',
  fixedRepeat: 'Fixed repeats',
  repeatCounts: 'Repeat counts',
  repeatCountsPlaceholder: '1, 2, 1',
  layer: 'Layer',
  foreground: 'Foreground',
  background: 'Background',
  opacity: 'Opacity',
  overlayImageFile: 'Image',
  chooseImage: 'Choose image',
  overlayImagePosition: 'Position',
  topLeft: 'Top left',
  topCenter: 'Top center',
  topRight: 'Top right',
  middleLeft: 'Middle left',
  center: 'Center',
  middleRight: 'Middle right',
  bottomLeft: 'Bottom left',
  bottomCenter: 'Bottom center',
  bottomRight: 'Bottom right',
  customPosition: 'Custom',
  overlayImageWidth: 'Page width',
  overlayImageMargin: 'Margin',
  customPositionHint: 'Coordinates start at the visible page’s top-left corner.',
  overlayImageHint: 'Adds a PNG or JPEG locally to the selected pages.',
  imagePdfFiles: 'Image / SVG order',
  addImages: 'Add images / SVG',
  svgVectorHint:
    'Color SVG files stay vector-based for sharp scaling. Grayscale modes rasterize them locally.',
  cameraStart: 'Open camera',
  cameraStop: 'Close camera',
  cameraSwitch: 'Switch camera',
  cameraCapture: 'Capture page',
  cameraReadyHint: 'Camera preview',
  cameraPrivacyHint: 'Scan pages locally. Camera frames never leave this device.',
  cameraDenied:
    'Camera access was denied. Allow it in your browser or system settings, then try again.',
  cameraUnavailable: 'No usable camera was found, or another application is using it.',
  cameraUnsupported: 'Camera scanning requires a secure browser or supported desktop system.',
  cameraCaptureFailed: 'The camera frame is not ready. Try again.',
  mobileScannerStart: 'Scan with phone',
  mobileScannerStop: 'End session',
  mobileScannerTitle: 'Phone upload',
  mobileScannerHint: 'Scan the QR code on a phone connected to this service.',
  mobileScannerWaiting: 'Waiting for phone images…',
  mobileScannerReceived: 'Images received:',
  mobileScannerCopyLink: 'Copy link',
  mobileScannerCopied: 'Copied',
  mobileScannerExpiresIn: 'Expires in',
  mobileScannerExpired: 'The phone upload session has expired.',
  mobileScannerUnavailable: 'Phone scanning is unavailable.',
  imageFit: 'Page fit',
  fillPage: 'Fill A4 page',
  fitDocumentToImage: 'Use image size',
  maintainAspectRatio: 'Fit inside A4',
  imageColorMode: 'Color mode',
  blackWhite: 'Black & white',
  autoRotateImages: 'Use landscape A4 for landscape images',
  imagePdfDestination: 'Output',
  newPdf: 'New PDF',
  appendCurrent: 'Append to current PDF',
  filterCriterion: 'Criterion',
  filterText: 'Text',
  filterImage: 'Images',
  filterPageSize: 'Page size',
  filterOrientation: 'Orientation',
  filterRotation: 'Rotation',
  filterComparator: 'Comparison',
  filterLess: 'Smaller than',
  filterEqual: 'Equal to',
  filterGreater: 'Larger than',
  filterQuery: 'Text to find',
  filterQueryPlaceholder: 'Enter literal text',
  caseSensitive: 'Match case',
  filterAction: 'Result',
  keepMatches: 'Keep matching pages',
  removeMatches: 'Remove matching pages',
  filterPagesHint: 'Checks selected pages locally and exports a new PDF.',
  filterDocumentsHint:
    'Checks the current PDF and added files locally. Matching files are exported unchanged.',
  filterDocumentsFiles: 'PDF candidates',
  addFilterDocuments: 'Add PDFs',
  currentDocument: 'Current document',
  removeFilterDocument: 'Remove PDF',
  filterPageCount: 'Page count',
  filterFileSize: 'File size',
  filterContainsText: 'Contains text',
  filterContainsImage: 'Contains image',
  filterFirstPageSize: 'First-page size',
  filterFirstPageRotation: 'First-page rotation',
  filterExpectedValue: 'Expected value',
  filterSizeUnit: 'Unit',
  replacementRules: 'Replacement rules',
  findText: 'Find',
  replaceWith: 'Replace with',
  addReplacement: 'Add rule',
  deleteReplacement: 'Delete rule',
  emptyReplacementDeletes: 'Leave replacement empty to delete matching text.',
  bulkReplaceHint:
    'Rules run from top to bottom. Later rules see earlier results. A new PDF copy is exported locally.',
  attachmentAction: 'Action',
  addAttachments: 'Add files',
  extractAttachments: 'Export ZIP',
  renameAttachment: 'Rename',
  deleteAttachment: 'Delete',
  attachmentFiles: 'Files',
  chooseAttachments: 'Choose files',
  existingAttachments: 'Embedded files',
  loadingAttachments: 'Loading…',
  noAttachments: 'No embedded files',
  newAttachmentName: 'New name',
  addBookmark: 'Add bookmark',
  addChildBookmark: 'Add child',
  bookmarkImportMode: 'Import mode',
  bookmarkReplace: 'Replace',
  bookmarkAppend: 'Append',
  importBookmarkJson: 'Import JSON',
  pasteBookmarkJson: 'Paste JSON',
  downloadBookmarkJson: 'Download JSON',
  copyBookmarkJson: 'Copy JSON',
  bookmarkJsonHint:
    'JSON uses title, pageNumber, and nested children. Maximum 1,000 items / 20 levels.',
  bookmarkJsonImported: 'Bookmark JSON imported.',
  bookmarkJsonCopied: 'Bookmark JSON copied.',
  bookmarkTitle: 'Title',
  bookmarkPage: 'Page',
  loadingAnalysis: 'Analyzing…',
  fontInfoLoadFailed: 'Could not read PDF font information.',
  fontSummary: 'Font summary',
  embeddedFonts: 'Embedded fonts',
  subsetFonts: 'Subset fonts',
  fontName: 'Font name',
  fontSubtype: 'Type',
  fontEncoding: 'Encoding',
  fontEmbedded: 'Embedded',
  fontSubset: 'Subset',
  fontToUnicode: 'ToUnicode',
  fontPages: 'Pages',
  noFonts: 'No PDF font resources were found.',
  annotationInfoLoadFailed: 'Could not read PDF annotation information.',
  annotationSummary: 'Annotation summary',
  annotationIndex: 'No.',
  annotationSubtype: 'Type',
  annotationContents: 'Contents',
  annotationModified: 'Modified',
  annotationPosition: 'Position',
  annotationFlags: 'Flags',
  noAnnotations: 'No PDF annotations were found.',
  securityInfoLoadFailed: 'Could not read PDF security information.',
  securitySummary: 'Security summary',
  encryptionAlgorithm: 'Algorithm',
  keyLength: 'Key length',
  securityVersion: 'Handler version / revision',
  restrictedActions: 'Restricted actions',
  permissionAllowed: 'Allowed',
  permissionRestricted: 'Restricted',
  permissionUnavailable: 'The encryption permissions dictionary could not be read.',
  analysisLoadFailed: 'Could not analyze this PDF.',
  preflightLoadFailed: 'Could not complete the local PDF preflight.',
  preflightStatus: 'Preflight status',
  preflightPass: 'Structure passed',
  preflightWarning: 'Review warnings',
  preflightError: 'Structural errors found',
  downloadPreflightJson: 'Download JSON',
  reviewRepair: 'Open repair tool',
  preflightDisclaimer:
    'This is a local structural preflight and declaration check. It is not formal veraPDF or accessibility compliance certification.',
  structure: 'File structure',
  strictParsing: 'Strict parsing',
  headerOffset: 'PDF header offset',
  eofMarkers: 'EOF markers',
  trailingBytes: 'Trailing bytes',
  startXref: 'startxref offset',
  incrementalUpdates: 'Incremental updates',
  standardDeclarations: 'Declared standards',
  noStandardDeclarations: 'No PDF/A or PDF/UA declaration found',
  xmpMetadata: 'XMP metadata',
  xmpValid: 'XMP parseable',
  taggedPdf: 'Tagged PDF',
  documentLanguage: 'Document language',
  outputIntents: 'Output intents',
  findings: 'Findings',
  noPreflightFindings: 'No structural findings.',
  findingMissingPdfHeader: 'The PDF header is missing.',
  findingHeaderNotAtStart: 'The PDF header is not at the start of the file.',
  findingMissingEofMarker: 'The final EOF marker is missing.',
  findingLargeTrailingData: 'A large amount of data follows the final EOF marker.',
  findingMissingStartXref: 'The startxref pointer is missing.',
  findingStartXrefOutOfRange: 'The startxref pointer is outside the file.',
  findingStartXrefTargetInvalid: 'The startxref pointer does not target cross-reference data.',
  findingStrictParseFailed: 'Strict PDF parsing failed.',
  findingEmptyPageTree: 'The PDF has no pages.',
  findingInvalidPageBox: 'A page has invalid dimensions.',
  findingInvalidXmpMetadata: 'XMP metadata is present but could not be parsed.',
  findingStandardDeclaredOnly:
    'The document declares a standard; formal compliance was not certified.',
  findingPdfaMissingOutputIntent: 'PDF/A is declared but no output intent was found.',
  findingPdfaAttachmentsRisk: 'Attachments may be incompatible with the declared PDF/A part.',
  findingPdfaJavaScriptRisk: 'JavaScript is incompatible with PDF/A expectations.',
  findingPdfuaNotTagged: 'PDF/UA is declared but the document is not marked as tagged.',
  findingPdfuaMissingLanguage: 'PDF/UA is declared but the document language is missing.',
  findingJavascriptPresent: 'JavaScript actions are present.',
  findingAttachmentsPresent: 'Embedded attachments are present.',
  findingXfaPresent: 'An XFA form is present and may have limited reader support.',
  findingSignaturesPresent: 'Digital signature fields are present.',
  metadataLoadFailed: 'Could not read PDF metadata.',
  overview: 'Overview',
  pageCount: 'Pages',
  pdfVersion: 'PDF version',
  fileSize: 'File size',
  encrypted: 'Encrypted',
  yes: 'Yes',
  no: 'No',
  documentProperties: 'Document properties',
  documentTitle: 'Title',
  author: 'Author',
  subject: 'Subject',
  keywords: 'Keywords',
  creator: 'Creator',
  producer: 'Producer',
  creationDate: 'Created',
  modificationDate: 'Modified',
  trapped: 'Trapping status',
  trappedTrue: 'True',
  trappedFalse: 'False',
  trappedUnknown: 'Unknown',
  customMetadata: 'Custom metadata',
  customMetadataKey: 'Key',
  customMetadataValue: 'Value',
  addCustomMetadata: 'Add field',
  deleteAllMetadata: 'Clear all document metadata',
  resources: 'Resources',
  fonts: 'Fonts',
  images: 'Images',
  embeddedFiles: 'Attachments',
  bookmarkCount: 'Bookmarks',
  javaScriptActions: 'Script actions',
  uniqueScripts: 'Unique scripts',
  scriptCodeSize: 'Code size',
  loadingJavaScript: 'Auditing JavaScript…',
  javaScriptLoadFailed: 'Could not audit PDF JavaScript.',
  noJavaScript: 'No JavaScript actions were found.',
  loadingSignatures: 'Validating digital signatures…',
  signatureLoadFailed: 'Could not validate PDF signatures.',
  noSignatures: 'No signed signature values were found.',
  validSignatures: 'Cryptographically valid',
  invalidSignatures: 'Invalid or unreadable',
  unsupportedSignatures: 'Unsupported',
  signatureValid: 'Valid',
  signatureInvalid: 'Invalid',
  signatureUnsupported: 'Unsupported',
  signatureError: 'Could not verify',
  signatureNotChecked: 'Not checked',
  signatureField: 'Signature field',
  signatureSigner: 'Signer',
  reason: 'Reason',
  location: 'Location',
  signatureType: 'Type',
  approvalSignature: 'Approval signature',
  certificationSignature: 'Certification signature',
  documentTimestamp: 'Document timestamp',
  unknownSignature: 'Unknown signature',
  signedAt: 'Claimed signing time',
  timestampAt: 'Timestamp time',
  signatureCoverage: 'Coverage',
  entireDocument: 'Entire file',
  earlierRevision: 'Earlier revision',
  unsignedTail: 'Unsigned trailing data',
  digestAlgorithm: 'Digest algorithm',
  signatureAlgorithm: 'Signature algorithm',
  certificateSubject: 'Certificate subject',
  certificateIssuer: 'Certificate issuer',
  certificateSerial: 'Certificate serial',
  certificateValidFrom: 'Certificate valid from',
  certificateValidUntil: 'Certificate valid until',
  certificateValidAtSigning: 'Valid at signing time',
  certificateSelfSigned: 'Self-signed certificate',
  trustNotEvaluated:
    'Certificate trust and revocation are not evaluated locally. Cryptographic validity only confirms the signed bytes have not changed.',
  copy: 'Copy',
  copyAll: 'Copy all',
  copied: 'Copied',
  reviewRemoval: 'Review removal',
  sourceNamed: 'Named script',
  sourceDocument: 'Document',
  sourcePage: 'Page',
  sourceAnnotation: 'Annotation',
  sourceForm: 'Form field',
  trigger: 'Trigger',
  scriptName: 'Script',
  field: 'Field',
  annotationNumber: 'Annotation',
  interactiveContent: 'Interactive content',
  formFields: 'Form fields',
  signatures: 'Signatures',
  xfa: 'XFA form',
  annotations: 'Annotations',
  annotationTypes: 'Types',
  pageDetails: 'Page details',
  dimensions: 'Dimensions',
  rotation: 'Rotation',
  none: 'None',
  close: 'Close',
  loadingBookmarks: 'Loading…',
  noBookmarks: 'No bookmarks',
  moveUp: 'Move up',
  moveDown: 'Move down',
  indentBookmark: 'Indent',
  outdentBookmark: 'Outdent',
  deleteBookmark: 'Delete bookmark',
  invalidSplit: 'Enter valid page numbers or a positive whole number.',
  invalidFileSize: 'Enter a valid positive file size.',
  invalidChapterLevels: 'Bookmark levels must be an integer from 1 to 21.',
  invalidMargins: 'Margins must be non-negative numbers.',
  invalidAutoCrop: 'White threshold must be 0–255 and padding must be 0–144 pt.',
  invalidGrid: 'Rows and columns must be integers from 1 to 10.',
  invalidNup: 'Check the grid, margins, and border width.',
  invalidGutter: 'Gutter must be a non-negative number.',
  invalidPages: 'Enter valid pages from this document.',
  invalidExtractPages: 'Enter at least one valid page in the requested order.',
  invalidColors: 'Choose valid text and background colors.',
  invalidBlankDetection:
    'Whiteness tolerance must be 0–255 and white area must be greater than 0% and at most 100%.',
  invalidRemoval: 'Choose pages to remove, but keep at least one page.',
  invalidOrder: 'Enter a valid page order that leaves at least one page.',
  invalidCopies: 'Copies must be an integer from 1 to 100.',
  invalidRenderDpi: 'Rendering DPI must be an integer from 72 to 600.',
  invalidScannerEffect: 'Enter valid scanner effect settings.',
  invalidDeskew: 'Check the page range, maximum angle, and rendering DPI.',
  invalidScannerImageSplit: 'Enter valid scanned-photo detection settings.',
  invalidImageQuality: 'Image quality must be an integer from 10% to 100%.',
  invalidLineArt: 'Line sensitivity must be 0%–100% and edge strength must be 1–3.',
  invalidPdfToImages: 'Check the page range, DPI, and image quality.',
  invalidPdfToVideo: 'Check the page range, timing, transition, and video duration.',
  invalidPdfToCbz: 'Check the page range, format, DPI, and image quality.',
  invalidPdfToHtml: 'Check the page range and rendering DPI.',
  invalidPdfToEpub: 'Check the page range and rendering DPI.',
  invalidPdfToPptx: 'Check the page range and rendering DPI.',
  invalidPdfToDocx: 'Check the page range and rendering DPI.',
  invalidPdfToRtf: 'Check the page range.',
  invalidRedactionPatterns: 'Enter at least one valid text pattern.',
  invalidRedactionAreas: 'Enter valid page areas that stay inside each page.',
  invalidRedactionPadding: 'Padding must be from 0 to 72 points.',
  invalidPipeline: 'Add from 1 to 12 valid processing steps.',
  invalidComments: 'Each comment needs valid page, position, and comment text.',
  invalidFormFields: 'Complete all required form fields.',
  invalidCreatedFields: 'Check field names, pages, coordinates, sizes, options, and defaults.',
  modifyAtLeastOneField: 'Change at least one form field property.',
  invalidModifiedFieldNames: 'Field names must be unique and cannot contain empty groups.',
  invalidChoiceOptions: 'Each choice field needs at least one option.',
  selectFieldsToDelete: 'Select at least one form field to delete.',
  selectSanitize: 'Select at least one item to remove.',
  selectOverlayFiles: 'Choose at least one overlay PDF.',
  selectOverlayImage: 'Choose a PNG or JPEG image.',
  selectImagesToPdf: 'Choose at least one PNG, JPEG, or WebP image.',
  selectMergeFiles: 'Choose at least one additional PDF.',
  selectComparisonFile: 'Choose a PDF to compare.',
  invalidComparison: 'Rendering DPI must be 72–300 and tolerance must be 0%–100%.',
  invalidRepeatCounts: 'Enter one positive repeat count for each overlay PDF.',
  invalidOpacity: 'Opacity must be from 0% to 100%.',
  invalidImageOverlay: 'Check the pages, size, margin, and custom coordinates.',
  invalidContentFilter: 'Check the page range and filter settings.',
  invalidDocumentFilter: 'Check the candidate files and document filter settings.',
  invalidBulkReplace: 'Choose valid pages and provide at least one non-empty find value.',
  invalidPageNumbers: 'Check the pages, numbering, font, color, and text format.',
  invalidMetadata: 'Check metadata dates and custom field names.',
  selectAttachmentFiles: 'Choose at least one file to attach.',
  selectExistingAttachment: 'Choose an embedded file.',
  invalidAttachmentName: 'Enter a valid new attachment name.',
  attachmentLoadFailed: 'Could not read embedded files.',
  bookmarkLoadFailed: 'Could not read PDF bookmarks.',
  invalidBookmarks: 'Each bookmark needs a title and a valid page number.',
  invalidBookmarkJson: 'Choose valid bookmark JSON for pages in this document.',
  bookmarkLimitReached: 'Bookmarks are limited to 1,000 items and 20 levels.',
  completed: 'PDF tool completed',
  outputs: 'files',
}

const ZH: typeof EN = {
  tab: '工具',
  title: 'PDF 工具',
  split: '拆分',
  merge: '合并 PDF',
  imagesToPdf: '图片 / SVG 转 PDF',
  cbzToPdf: 'CBZ / CBR 转 PDF',
  emailToPdf: '邮件转 PDF',
  epubToPdf: 'EPUB 转 PDF',
  htmlToPdf: '网页 / HTML 转 PDF',
  markdownToPdf: 'Markdown / ZIP 转 PDF',
  pdfToImages: 'PDF 转图片',
  pdfToVideo: 'PDF 转视频',
  pdfToCbz: 'PDF 转 CBZ',
  pdfToHtml: 'PDF 转 HTML',
  pdfToEpub: 'PDF 转 EPUB',
  pdfToPptx: 'PDF 转 PowerPoint',
  pdfToDocx: 'PDF 转 Word',
  pdfToOdt: 'PDF 转 ODT',
  pdfToRtf: 'PDF 转 RTF',
  pdfToPdfa: '归档 PDF（PDF/A）',
  pdfToMarkdown: 'PDF 转 Markdown',
  pdfToXlsx: 'PDF 转 Excel',
  pdfToXml: 'PDF 转 XML',
  compare: 'PDF 对比',
  extractPages: '提取页面',
  splitSections: '页面分区',
  crop: '裁边',
  scale: '缩放',
  nup: '多页合一',
  booklet: '小册子',
  poster: '海报分页',
  singlePage: '长页合并',
  rotatePages: '批量旋转',
  autoRotate: '自动旋转',
  deskew: '扫描页纠偏',
  autoRename: '自动重命名',
  pageNumbers: '添加页码',
  scannerEffect: '扫描件效果',
  scannerImageSplit: '扫描照片分割',
  autoSplit: '二维码自动拆分',
  removePages: '批量删除',
  removeImages: '移除图片',
  bulkReplaceText: '批量替换文字',
  extractText: '导出文本',
  pdfToJson: 'PDF 转 JSON',
  jsonToPdf: 'JSON 转 PDF',
  extractTables: '提取表格',
  extractImages: '提取图片',
  removeAnnotations: '移除注释',
  removeBlanks: '移除空白页',
  invertColors: '反转颜色',
  replaceColors: '换色与高对比度',
  adjustColors: '色彩调节',
  rearrange: '页面重排',
  redact: '安全涂黑',
  comments: '添加便签评论',
  compress: '压缩 PDF',
  flatten: 'PDF 扁平化',
  forms: '表单处理',
  repair: '修复 PDF',
  decompress: '解压缩 PDF',
  certificateSign: '证书签名',
  timestamp: '文档时间戳',
  removeSignatures: '移除数字签名',
  password: '密码保护',
  ocr: 'OCR 文字识别',
  sanitize: '文档清理',
  pipeline: '处理流水线',
  overlay: 'PDF 叠加',
  overlayImage: '添加图片',
  filterPages: '筛选页面',
  filterDocuments: '批量筛选 PDF',
  attachments: '附件管理',
  bookmarks: '目录书签',
  metadata: '编辑元数据',
  info: 'PDF 信息',
  preflight: '本地预检',
  javascriptAudit: 'JavaScript 审计',
  signatureAudit: '验证数字签名',
  fontInfo: '字体清单',
  annotationInfo: '批注清单',
  securityInfo: '安全与权限',
  protectPdf: '保护 PDF',
  unlockPdf: '解锁 PDF',
  openPassword: '打开密码',
  ownerPassword: '所有者密码',
  currentPassword: '当前密码',
  confirmPassword: '确认密码',
  encryption: '加密算法',
  aes256: 'AES-256',
  rc4: 'RC4-128',
  aesRecommended: '推荐用于现代 PDF 阅读器，支持中文等 Unicode 密码。',
  rc4Compatibility: '仅用于兼容旧阅读器，密码字符范围有限。',
  pdfPermissions: '允许的操作',
  allowPrinting: '打印',
  allowHighQualityPrint: '高质量打印',
  allowModifying: '修改内容',
  allowCopying: '复制文本和图片',
  allowAnnotating: '添加批注',
  allowFillingForms: '填写表单和签名',
  allowExtraction: '辅助功能提取',
  allowAssembly: '页面组合',
  permissionAdvisory: 'PDF 权限依赖阅读器遵循。如需内容保密，请设置打开密码。',
  passwordRewritesPdf: '保护或解锁会重写 PDF，并使已有数字签名失效。',
  unlockHint: '导出不加密的副本，当前文档保持不变。',
  passwordRequired: '请至少输入打开密码或所有者密码。',
  passwordMismatch: '两次输入的密码不一致。',
  ownerPasswordRequired: '限制操作时必须设置不同的所有者密码。',
  ocrMode: '识别模式',
  ocrSkipText: '自动 · 跳过已有文字页',
  ocrForce: '强制 · 替换全部页面',
  ocrStrict: '严格 · 发现文字即停止',
  ocrLanguages: '识别语言',
  ocrEnglish: '英语',
  ocrChinese: '简体中文',
  ocrRenderDpi: '渲染 DPI',
  ocrClean: '识别前清理扫描图像',
  ocrCleanHint: '识别前增强对比度并转为灰度，适合背景发灰的扫描件。',
  ocrSidecar: '同时导出旁挂文本',
  ocrSidecarHint: '将可搜索 PDF 与按页整理的纯文本一起打包为 ZIP。',
  ocrLocalHint: '识别完全在本地执行，语言模型首次加载后会由浏览器缓存。',
  ocrRewritesPdf: 'OCR 会重写处理页面，并使已有数字签名失效。',
  ocrNoLanguage: '请至少选择一种识别语言。',
  invalidOcrDpi: 'OCR 渲染 DPI 必须为 100 到 300 的整数。',
  runOcr: '开始 OCR',
  ocrAnalyzing: '正在检查已有文字层…',
  ocrLoading: '正在加载本地 OCR 引擎…',
  ocrRecognizing: '正在识别第 {page}/{total} 页…',
  ocrFinishing: '正在生成可搜索 PDF…',
  certificateFile: 'P12 / PFX 证书',
  timestampAuthority: '时间戳服务机构',
  timestampPrivacyHint: '仅向所选机构发送 SHA-256 摘要，PDF 文件不会离开当前设备。',
  timestampSecurityHint: '输出文件包含数字签名，后续编辑会使时间戳失效。',
  runTimestamp: '添加时间戳',
  chooseCertificate: '选择证书',
  noCertificate: '尚未选择证书',
  certificatePassword: '证书密码',
  signerNameInput: '签署人名称',
  contactInfo: '联系方式',
  certificateSignerPlaceholder: '可选，用于 PDF 显示',
  certificateReasonPlaceholder: '文档审批',
  certificateLocationPlaceholder: '城市或组织',
  certificateContactPlaceholder: '邮箱或电话（可选）',
  certificateRequired: '请选择 P12 或 PFX 证书文件。',
  certificateTooLarge: '证书文件不能超过 16 MB。',
  certificateLocalHint: '证书和密码仅用于本次本地签名，不会被保存或上传。',
  certificateSecurityHint: '系统会生成已签名副本；之后对该副本的任何编辑都会使数字签名失效。',
  runCertificateSign: '签署 PDF',
  pdfaFormat: '归档标准',
  pdfaMode: '归档方式',
  pdfaAuto: '自动保真',
  pdfaRaster: '图像归档',
  pdfaRenderDpi: '归档渲染 DPI',
  pdfaImageQuality: '归档图像质量',
  pdfaAutoHint:
    '所有已用字体均已嵌入时保留可选文字和矢量图形，否则自动回退为图像归档。交互表单、链接、附件和脚本均在本地移除。',
  pdfaRasterHint:
    '始终生成图像化的 PDF/A-2b 副本，优先保持视觉一致性，但会扁平化可选文字和矢量图形。',
  invalidPdfa: '渲染 DPI 必须为 72 到 600 的整数，图像质量必须为 10% 到 100%。',
  splitMethod: '拆分方式',
  afterPages: '在指定页后拆分',
  byFileSize: '按文件大小',
  pagesPerDocument: '每份固定页数',
  documentCount: '平均拆成固定份数',
  chapters: '按目录章节拆分',
  splitAfter: '拆分位置',
  splitPlaceholder: '2, 5 或 all',
  splitValue: '数量',
  targetFileSize: '目标大小',
  chapterLevels: '包含目录层级',
  allowDuplicateChapters: '同页书签分别输出',
  sectionArrangement: '分区输出顺序',
  rowMajor: '按行从左到右',
  columnMajor: '按列从上到下',
  mergeSections: '合并为一个 PDF',
  cropMode: '裁边方式',
  cropAuto: '自动识别',
  cropManual: '手动边距',
  whiteThreshold: '白色阈值',
  cropPadding: '保留边距',
  autoCropHint: '在本地识别页面白边，并保留原始矢量内容。',
  top: '上边距',
  right: '右边距',
  bottom: '下边距',
  left: '左边距',
  points: '点',
  pageSize: '页面尺寸',
  contentScale: '内容比例',
  pagesPerSheet: '每张页数',
  customLayout: '自定义布局',
  nupPageOrder: '页面排列',
  readingDirection: '阅读方向',
  leftToRight: '从左向右',
  outerMargins: '外边距',
  innerMargin: '页间边距',
  borderWidth: '边框宽度',
  orientation: '页面方向',
  portrait: '纵向',
  landscape: '横向',
  border: '显示边框',
  spine: '装订方向',
  spineLeft: '左侧装订',
  spineRight: '右侧装订',
  gutter: '装订间距',
  duplexPass: '输出页面',
  both: '正反面',
  front: '仅正面',
  back: '仅反面',
  shortEdge: '短边翻转',
  columns: '横向份数',
  rows: '纵向份数',
  rightToLeft: '从右向左',
  joinDirection: '合并方向',
  vertical: '纵向',
  horizontal: '横向',
  pages: '页面范围',
  pagesPlaceholder: 'all 或 1, 3-5',
  colorMode: '换色方式',
  highContrast: '高对比度',
  customColors: '自定义颜色',
  colorPreset: '配色方案',
  whiteOnBlack: '黑底白字',
  blackOnWhite: '白底黑字',
  yellowOnBlack: '黑底黄字',
  greenOnBlack: '黑底绿字',
  textColor: '文字颜色',
  backgroundColor: '背景颜色',
  contrast: '对比度',
  brightness: '亮度',
  saturation: '饱和度',
  rgbChannels: 'RGB 通道',
  redChannel: '红色',
  greenChannel: '绿色',
  blueChannel: '蓝色',
  whitenessThreshold: '白色容差',
  whitePercent: '最小白色比例',
  includeBlankPages: '另存检测到的空白页',
  outputFormat: '输出格式',
  outputMode: '输出方式',
  separateImages: '分页图片',
  singleLongImage: '单张长图',
  includeAnnotations: '包含评论与表单值',
  pdfToImagesHint: '页面仅在本地渲染；多张图片会打包为一个 ZIP 文件下载。',
  videoDuration: '每页时长（秒）',
  videoResolution: '视频分辨率',
  videoTransition: '淡入过渡（秒）',
  videoEncoding: '正在编码第 {page}/{total} 页…',
  pdfToVideoHint: '无需 FFmpeg 或服务器，在本地生成 WebM 幻灯片；最多 100 页或 5 分钟。',
  pdfToCbzHint: '在本地渲染页面并按自然阅读顺序生成漫画书归档，不会上传文件。',
  cbzFile: '漫画压缩包',
  chooseCbz: '选择 CBZ / CBR',
  replaceCbz: '更换压缩包',
  selectCbz: '请选择一个 CBZ、CBR、ZIP 或 RAR 漫画压缩包。',
  cbzToPdfHint: 'ZIP 和 RAR 漫画包会在本地解包并自然排序；隐藏文件和非图片条目会被忽略。',
  emailFile: '邮件文件',
  chooseEmail: '选择 EML',
  replaceEmail: '更换邮件',
  selectEmail: '请选择一个 EML 邮件文件。',
  emailOutputFormat: '邮件输出',
  emailPdf: 'PDF 文档',
  emailHtml: '离线 HTML',
  emailIncludeAttachments: '嵌入原始附件',
  emailMaxAttachmentSize: '单个附件上限',
  emailIncludeAllRecipients: '包含抄送和密送',
  emailLocalHint: '邮件仅在本地解析；渲染前会移除远程图片和主动内容。',
  emailFromLabel: '发件人',
  emailToLabel: '收件人',
  emailCcLabel: '抄送',
  emailBccLabel: '密送',
  emailDateLabel: '日期',
  emailAttachmentsLabel: '附件',
  emailEmptyBody: '无法解析邮件正文。',
  emailRemoteImageRemoved: '远程图片已移除',
  emailUntitled: '无主题邮件',
  invalidEmailAttachmentSize: '单个附件上限必须是 1 到 50 MB 的整数。',
  epubFile: 'EPUB 电子书',
  chooseEpub: '选择 EPUB',
  replaceEpub: '更换 EPUB',
  selectEpub: '请选择一个 EPUB 电子书文件。',
  epubEmbedFonts: '嵌入书籍字体',
  epubIncludeToc: '添加目录页',
  epubIncludePageNumbers: '添加页码',
  epubOptimize: '针对电子书阅读器优化',
  epubLocalHint: '阅读顺序、章节和本地图片仅在当前设备解析；远程和主动内容会被移除。',
  epubTocLabel: '目录',
  epubUntitled: '未命名章节',
  jsonFile: '结构化 PDF JSON',
  chooseJson: '选择 JSON',
  replaceJson: '更换 JSON',
  selectJson: '请选择一个 GenOffice 结构化 PDF JSON v1 文件。',
  jsonToPdfHint:
    '仅在本地还原 GenOffice 结构化 PDF JSON v1；布局模式恢复文字坐标，语义模式重排文字。该格式当前不包含图片、批注、链接位置和表单位置。',
  htmlFile: '网页文档',
  htmlSource: '来源',
  htmlSourceFile: '本地文件',
  htmlSourceWebsite: '网页地址',
  htmlWebsiteUrl: '网页地址',
  htmlWebsitePlaceholder: 'https://example.com/article',
  chooseHtml: '选择 HTML / ZIP',
  replaceHtml: '更换文件',
  selectHtml: '请选择一个 HTML、HTM 或 ZIP 网页文档。',
  selectWebsiteUrl: '请输入有效的网页地址。',
  htmlZoom: '页面缩放',
  htmlIncludePageNumbers: '添加页码',
  htmlLocalHint: 'HTML、CSS 和本地图片仅在当前设备处理；远程和主动内容会被移除。',
  htmlRemoteHint: '仅安全获取公网网页、样式和图片；主动内容及内网资源会被移除。',
  htmlUntitled: '未命名网页文档',
  invalidHtmlZoom: '页面缩放必须是 50% 到 200% 的整数。',
  markdownFile: 'Markdown 文档',
  chooseMarkdown: '选择 Markdown / ZIP',
  replaceMarkdown: '更换文件',
  selectMarkdown: '请选择一个 MD、Markdown 或 ZIP 文档。',
  markdownZoom: '页面缩放',
  markdownIncludePageNumbers: '添加页码',
  markdownLocalHint:
    'GFM 表格、列表、代码块和 ZIP 内本地图片仅在当前设备处理；原始主动内容会被移除。',
  markdownUntitled: '未命名 Markdown 文档',
  invalidMarkdownZoom: '页面缩放必须是 50% 到 200% 的整数。',
  pdfToHtmlHint: '在本地生成包含页面图片和可搜索文字层的离线 ZIP，不会上传文件。',
  epubMode: '电子书版式',
  epubReflowable: '阅读版',
  epubFixed: '保真版',
  pdfToEpubHint: '阅读版会重排可选择文字；保真版使用本地页面图像保留原始外观。',
  pptxMode: 'PowerPoint 模式',
  pptxEditableText: '可编辑文字',
  pptxFidelity: '页面保真',
  pdfToPptxHint:
    '可编辑文字会生成带坐标的文本框，不会伪转换图片和复杂图形；页面保真会用整页图像保留完整外观。',
  docxMode: 'Word 模式',
  odtMode: 'OpenDocument 模式',
  docxEditableText: '可编辑文字',
  docxFidelity: '页面保真',
  pdfToDocxHint: '可编辑文字会在本地按阅读顺序重建段落；页面保真会用每页一张图像保留完整外观。',
  pdfToOdtHint:
    '在本地生成标准 ODT；可编辑文字保留阅读顺序和常用文字样式，页面保真会嵌入每页图像。',
  pdfToRtfHint:
    '在本地按阅读顺序导出可编辑富文本，保留常用字体、字号、粗体、斜体、Unicode 文字和分页。',
  pdfToMarkdownPageBreaks: '保留页面边界',
  pdfToMarkdownHint: '按阅读顺序重建标题、段落、列表和链接；扫描件请先执行 OCR。',
  pdfToXlsxHint: '在本地识别对齐表格，并为每个表格生成一张可编辑的 Excel 工作表。',
  pdfToXmlHint: '在本地导出标准 XML，包含元数据、书签、表单、链接和所选页面。',
  pageOrder: '页面顺序',
  pageOrderPlaceholder: '3, 1, 2, 2',
  angle: '旋转角度',
  inferUndetected: '根据文档推断文字不足的页面',
  deskewMaxAngle: '最大倾斜角度',
  deskewHint: '在本地检测扫描页的小角度倾斜；仅纠偏页面会栅格化，未变化页面保留原始 PDF 内容。',
  renameStrategy: '标题来源',
  largestHeading: '最大字号标题',
  firstText: '首段文字',
  pageNumberPosition: '页码位置',
  pageNumberMargin: '页边距',
  small: '小',
  medium: '中',
  large: '大',
  xLarge: '超大',
  pageNumberFont: '字体',
  helvetica: 'Helvetica',
  times: 'Times',
  courier: 'Courier',
  fontSize: '字号',
  fontColor: '文字颜色',
  startingNumber: '起始编号',
  zeroPad: '补零位数',
  textPattern: '页码格式',
  textPatternPlaceholder: 'Page {n} of {total}',
  pageNumbersHint: '可用变量：{n}、{total}、{filename}。',
  scanQuality: '扫描质量',
  scanHigh: '高',
  scanMedium: '中',
  scanLow: '低',
  scanCustom: '自定义',
  scanRotation: '页面倾斜',
  scanNone: '无',
  scanSlight: '轻微',
  scanModerate: '中等',
  scanSevere: '明显',
  scanColorspace: '色彩空间',
  grayscale: '灰度',
  color: '彩色',
  paperBorder: '纸张边缘',
  baseRotation: '基础角度',
  rotationVariance: '随机偏移',
  blur: '模糊',
  noise: '噪点',
  yellowedPaper: '泛黄纸张',
  textureSeed: '纹理种子',
  scanAngleThreshold: '旋转纠偏阈值',
  scanTolerance: '背景色容差',
  minimumPhotoArea: '最小照片面积',
  minimumForegroundArea: '最小前景面积',
  detectionBorder: '检测边缘',
  mode: '重排方式',
  custom: '自定义顺序',
  reverse: '倒序',
  oddEven: '奇数页后偶数页',
  duplex: '双面打印顺序',
  removeFirst: '移除首页',
  removeLast: '移除末页',
  removeFirstAndLast: '移除首尾页',
  duplicate: '重复每一页',
  copies: '每页份数',
  flattenMode: '扁平化方式',
  flattenPages: '整份文档',
  flattenFormsOnly: '仅表单',
  renderDpi: '渲染 DPI',
  imageQuality: '图片质量',
  lineArt: '转换为线稿',
  lineArtThreshold: '线条灵敏度',
  lineArtEdgeLevel: '边缘强度',
  redactPatterns: '文字或匹配规则',
  redactPatternsPlaceholder: '每行输入一个关键词',
  redactTextMode: '文字匹配',
  redactAreaMode: '页面区域',
  redactionMode: '涂黑方式',
  addRedactionArea: '添加区域',
  redactionArea: '区域',
  deleteRedactionArea: '删除区域',
  redactionPreview: '涂黑区域预览',
  resizeRedactionArea: '调整涂黑区域大小',
  redactLeft: '左侧',
  redactTop: '顶部',
  redactWidth: '宽度',
  redactHeight: '高度',
  redactionAreaHint:
    '坐标为从页面左上角开始的百分比。命中页面会安全栅格化，其他页面保留原始 PDF 内容。',
  useRegex: '使用正则表达式',
  wholeWord: '仅匹配完整单词',
  redactColor: '涂黑颜色',
  redactPadding: '额外留白',
  pipelineStep: '流水线步骤',
  addPipelineStep: '添加步骤',
  deletePipelineStep: '删除步骤',
  flattenForms: '扁平化表单',
  pipelineHint: '按列表顺序在本地逐步处理，并只导出一份 PDF。调整步骤顺序可控制最终结果。',
  addComment: '添加评论',
  comment: '评论',
  commentText: '评论内容',
  commentTextPlaceholder: '输入在 PDF 阅读器中显示的便签内容',
  anchorText: '锚点文字（可选）',
  anchorTextPlaceholder: '放在本页首个匹配文字旁',
  fallbackPosition: '回退坐标',
  positionX: 'X',
  positionY: 'Y',
  deleteComment: '删除评论',
  formAction: '处理方式',
  exportFormData: '导出数据',
  fillFields: '填写表单',
  createFields: '新建字段',
  flattenFields: '扁平化表单',
  unlock: '解除只读',
  modifyFields: '修改属性',
  deleteFields: '删除字段',
  selectAllFields: '全选',
  clearFieldSelection: '清空',
  selectedFields: '已选字段',
  modifiedFields: '已修改字段',
  createdFields: '新建字段',
  addField: '添加字段',
  removeField: '移除字段',
  fieldName: '字段名称',
  fieldLabel: '字段标签',
  fieldType: '字段类型',
  fieldPage: '页码',
  fieldWidth: '宽度',
  fieldHeight: '高度',
  defaultValue: '默认值',
  defaultValues: '默认值（每行一项）',
  defaultChecked: '默认选中',
  multilineField: '允许多行',
  optionSpacing: '选项间距',
  choiceOptions: '选项（每行一项）',
  multiselectField: '允许多选',
  formCreateHint: '坐标使用 PDF 点，原点位于页面左下角。',
  loadingFormFields: '正在读取表单字段…',
  formFieldsLoadFailed: '无法读取 PDF 表单字段。',
  noFormFields: '未发现 AcroForm 表单字段。',
  readOnlyField: '只读',
  requiredField: '必填',
  selectEmpty: '不选择',
  unsupportedFormField: '该字段类型不能在这里填写。',
  formText: '文本',
  formCheckbox: '复选框',
  formRadio: '单选组',
  formDropdown: '下拉列表',
  formOptionList: '多选列表',
  formButton: '按钮',
  formSignature: '签名',
  formUnknown: '未知字段',
  textExportFormat: '输出格式',
  plainText: '纯文本',
  markdown: 'Markdown',
  exportBoth: '两种格式',
  textExportHint: '在本地导出所选页面；扫描件请先执行 OCR。',
  structuredExportMode: '结构详细程度',
  semanticJson: '紧凑语义',
  layoutJson: '完整布局',
  pdfToJsonHint: '在本地导出元数据、书签、表单、链接和所选页面；完整布局还包含文字坐标与字体。',
  tableExportFormat: '输出格式',
  csv: 'CSV',
  excel: 'Excel',
  includeTwoColumnTables: '识别纯文本双列表格',
  tableExportHint: '在本地识别对齐表格；复杂、旋转或扫描表格建议导出后复核。',
  downloadDivider: '下载分隔页',
  duplexMode: '双面扫描模式',
  autoSplitHint: '在本地扫描兼容二维码，移除分隔页，并将各份文档打包为 ZIP。',
  removeJavaScript: '移除 JavaScript',
  removeEmbeddedFiles: '移除嵌入文件',
  removeXmpMetadata: '移除 XMP 元数据',
  removeMetadata: '移除文档属性',
  removeLinks: '移除外部链接动作',
  overlayFiles: '叠加 PDF',
  chooseFiles: '选择 PDF',
  noFiles: '未选择文件',
  mergeFiles: 'PDF 顺序',
  currentPdf: '当前 PDF',
  addMergeFiles: '添加 PDF',
  deleteFile: '移除文件',
  comparisonFile: '对比 PDF',
  chooseComparisonFile: '选择 PDF',
  comparisonTolerance: '差异容差',
  overlayMode: '页面分配',
  sequential: '顺序循环',
  interleaved: '文件交错',
  fixedRepeat: '固定次数',
  repeatCounts: '重复次数',
  repeatCountsPlaceholder: '1, 2, 1',
  layer: '叠加层级',
  foreground: '前景',
  background: '背景',
  opacity: '透明度',
  overlayImageFile: '图片',
  chooseImage: '选择图片',
  overlayImagePosition: '位置',
  topLeft: '左上',
  topCenter: '上中',
  topRight: '右上',
  middleLeft: '左中',
  center: '居中',
  middleRight: '右中',
  bottomLeft: '左下',
  bottomCenter: '下中',
  bottomRight: '右下',
  customPosition: '自定义',
  overlayImageWidth: '页面宽度占比',
  overlayImageMargin: '边距',
  customPositionHint: '坐标以可见页面左上角为起点。',
  overlayImageHint: '在本地将 PNG 或 JPEG 添加到所选页面。',
  imagePdfFiles: '图片 / SVG 顺序',
  addImages: '添加图片 / SVG',
  svgVectorHint: '彩色 SVG 会保留矢量内容，缩放依然清晰；灰度模式会在本地栅格化。',
  cameraStart: '打开相机',
  cameraStop: '关闭相机',
  cameraSwitch: '切换相机',
  cameraCapture: '拍摄此页',
  cameraReadyHint: '相机预览',
  cameraPrivacyHint: '在本地扫描页面，相机画面不会离开此设备。',
  cameraDenied: '相机权限被拒绝，请在浏览器或系统设置中允许后重试。',
  cameraUnavailable: '未找到可用相机，或相机正被其他应用占用。',
  cameraUnsupported: '相机扫描需要安全浏览器环境或受支持的桌面系统。',
  cameraCaptureFailed: '相机画面尚未就绪，请重试。',
  mobileScannerStart: '手机扫码',
  mobileScannerStop: '结束会话',
  mobileScannerTitle: '手机上传',
  mobileScannerHint: '用可访问当前服务的手机扫描二维码。',
  mobileScannerWaiting: '等待手机发送图片…',
  mobileScannerReceived: '已接收图片：',
  mobileScannerCopyLink: '复制链接',
  mobileScannerCopied: '已复制',
  mobileScannerExpiresIn: '剩余时间',
  mobileScannerExpired: '手机上传会话已过期。',
  mobileScannerUnavailable: '手机扫码功能暂不可用。',
  imageFit: '页面适配',
  fillPage: '铺满 A4 页面',
  fitDocumentToImage: '使用图片尺寸',
  maintainAspectRatio: '等比放入 A4',
  imageColorMode: '颜色模式',
  blackWhite: '黑白',
  autoRotateImages: '横图使用 A4 横向页面',
  imagePdfDestination: '输出位置',
  newPdf: '新建 PDF',
  appendCurrent: '追加到当前 PDF',
  filterCriterion: '筛选条件',
  filterText: '文字',
  filterImage: '图片',
  filterPageSize: '页面尺寸',
  filterOrientation: '页面方向',
  filterRotation: '页面旋转',
  filterComparator: '比较方式',
  filterLess: '小于',
  filterEqual: '等于',
  filterGreater: '大于',
  filterQuery: '查找文字',
  filterQueryPlaceholder: '输入要查找的原文',
  caseSensitive: '区分大小写',
  filterAction: '输出结果',
  keepMatches: '保留匹配页面',
  removeMatches: '移除匹配页面',
  filterPagesHint: '在本地检查所选页面并导出新 PDF。',
  filterDocumentsHint: '在本地检查当前 PDF 和添加的文件，仅原样导出符合条件的文档。',
  filterDocumentsFiles: '候选 PDF',
  addFilterDocuments: '添加 PDF',
  currentDocument: '当前文档',
  removeFilterDocument: '移除 PDF',
  filterPageCount: '页数',
  filterFileSize: '文件大小',
  filterContainsText: '包含文字',
  filterContainsImage: '包含图片',
  filterFirstPageSize: '首页尺寸',
  filterFirstPageRotation: '首页旋转',
  filterExpectedValue: '目标值',
  filterSizeUnit: '单位',
  replacementRules: '替换规则',
  findText: '查找',
  replaceWith: '替换为',
  addReplacement: '添加规则',
  deleteReplacement: '删除规则',
  emptyReplacementDeletes: '替换内容留空即可删除匹配文字。',
  bulkReplaceHint: '规则从上到下依次执行，后续规则会看到前序结果，并在本地导出新的 PDF 副本。',
  attachmentAction: '操作',
  addAttachments: '添加文件',
  extractAttachments: '导出 ZIP',
  renameAttachment: '重命名',
  deleteAttachment: '删除',
  attachmentFiles: '附件文件',
  chooseAttachments: '选择文件',
  existingAttachments: '已有附件',
  loadingAttachments: '正在读取…',
  noAttachments: '没有嵌入附件',
  newAttachmentName: '新名称',
  addBookmark: '添加书签',
  addChildBookmark: '添加子级',
  bookmarkImportMode: '导入方式',
  bookmarkReplace: '替换',
  bookmarkAppend: '追加',
  importBookmarkJson: '导入 JSON',
  pasteBookmarkJson: '粘贴 JSON',
  downloadBookmarkJson: '下载 JSON',
  copyBookmarkJson: '复制 JSON',
  bookmarkJsonHint: 'JSON 使用 title、pageNumber 和嵌套 children，最多 1000 项 / 20 层。',
  bookmarkJsonImported: '目录 JSON 已导入。',
  bookmarkJsonCopied: '目录 JSON 已复制。',
  bookmarkTitle: '标题',
  bookmarkPage: '页码',
  loadingAnalysis: '正在分析…',
  fontInfoLoadFailed: '无法读取 PDF 字体信息。',
  fontSummary: '字体概览',
  embeddedFonts: '已嵌入字体',
  subsetFonts: '子集字体',
  fontName: '字体名称',
  fontSubtype: '类型',
  fontEncoding: '编码',
  fontEmbedded: '嵌入',
  fontSubset: '子集',
  fontToUnicode: 'ToUnicode',
  fontPages: '使用页码',
  noFonts: '未发现 PDF 字体资源。',
  annotationInfoLoadFailed: '无法读取 PDF 批注信息。',
  annotationSummary: '批注概览',
  annotationIndex: '序号',
  annotationSubtype: '类型',
  annotationContents: '内容',
  annotationModified: '修改时间',
  annotationPosition: '位置',
  annotationFlags: '标志',
  noAnnotations: '未发现 PDF 批注。',
  securityInfoLoadFailed: '无法读取 PDF 安全信息。',
  securitySummary: '安全概览',
  encryptionAlgorithm: '加密算法',
  keyLength: '密钥长度',
  securityVersion: '处理器版本 / 修订版',
  restrictedActions: '受限操作',
  permissionAllowed: '允许',
  permissionRestricted: '受限',
  permissionUnavailable: '无法读取加密权限字典。',
  analysisLoadFailed: '无法分析此 PDF。',
  preflightLoadFailed: '无法完成 PDF 本地预检。',
  preflightStatus: '预检状态',
  preflightPass: '结构检查通过',
  preflightWarning: '存在待检查警告',
  preflightError: '发现结构错误',
  downloadPreflightJson: '下载 JSON',
  reviewRepair: '打开修复工具',
  preflightDisclaimer:
    '此报告仅执行本地结构预检和标准声明检测，不等同于 veraPDF 或无障碍正式合规认证。',
  structure: '文件结构',
  strictParsing: '严格解析',
  headerOffset: 'PDF 文件头偏移',
  eofMarkers: 'EOF 标记',
  trailingBytes: '尾部数据',
  startXref: 'startxref 偏移',
  incrementalUpdates: '增量更新次数',
  standardDeclarations: '标准声明',
  noStandardDeclarations: '未发现 PDF/A 或 PDF/UA 声明',
  xmpMetadata: 'XMP 元数据',
  xmpValid: 'XMP 可解析',
  taggedPdf: '标签化 PDF',
  documentLanguage: '文档语言',
  outputIntents: '输出意图',
  findings: '检查结果',
  noPreflightFindings: '未发现结构问题。',
  findingMissingPdfHeader: '缺少 PDF 文件头。',
  findingHeaderNotAtStart: 'PDF 文件头不在文件起始位置。',
  findingMissingEofMarker: '缺少最终 EOF 标记。',
  findingLargeTrailingData: '最终 EOF 标记后存在大量尾部数据。',
  findingMissingStartXref: '缺少 startxref 指针。',
  findingStartXrefOutOfRange: 'startxref 指针超出文件范围。',
  findingStartXrefTargetInvalid: 'startxref 指针未指向交叉引用数据。',
  findingStrictParseFailed: 'PDF 严格解析失败。',
  findingEmptyPageTree: 'PDF 没有页面。',
  findingInvalidPageBox: '页面尺寸无效。',
  findingInvalidXmpMetadata: '存在 XMP 元数据，但无法解析。',
  findingStandardDeclaredOnly: '文档声明了标准，但本地预检不认证正式合规。',
  findingPdfaMissingOutputIntent: '声明了 PDF/A，但未发现输出意图。',
  findingPdfaAttachmentsRisk: '附件可能与声明的 PDF/A 版本不兼容。',
  findingPdfaJavaScriptRisk: 'JavaScript 不符合 PDF/A 的通常要求。',
  findingPdfuaNotTagged: '声明了 PDF/UA，但文档未标记为标签化 PDF。',
  findingPdfuaMissingLanguage: '声明了 PDF/UA，但缺少文档语言。',
  findingJavascriptPresent: '文档包含 JavaScript 动作。',
  findingAttachmentsPresent: '文档包含嵌入附件。',
  findingXfaPresent: '文档包含 XFA 表单，部分阅读器可能不支持。',
  findingSignaturesPresent: '文档包含数字签名字段。',
  metadataLoadFailed: '无法读取 PDF 元数据。',
  overview: '基本信息',
  pageCount: '页数',
  pdfVersion: 'PDF 版本',
  fileSize: '文件大小',
  encrypted: '已加密',
  yes: '是',
  no: '否',
  documentProperties: '文档属性',
  documentTitle: '标题',
  author: '作者',
  subject: '主题',
  keywords: '关键词',
  creator: '创建程序',
  producer: '生成程序',
  creationDate: '创建时间',
  modificationDate: '修改时间',
  trapped: '陷印状态',
  trappedTrue: '是',
  trappedFalse: '否',
  trappedUnknown: '未知',
  customMetadata: '自定义元数据',
  customMetadataKey: '键',
  customMetadataValue: '值',
  addCustomMetadata: '添加字段',
  deleteAllMetadata: '清除全部文档元数据',
  resources: '资源统计',
  fonts: '字体',
  images: '图片',
  embeddedFiles: '附件',
  bookmarkCount: '目录书签',
  javaScriptActions: '脚本动作',
  uniqueScripts: '不重复脚本',
  scriptCodeSize: '代码大小',
  loadingJavaScript: '正在审计 JavaScript…',
  javaScriptLoadFailed: '无法审计 PDF JavaScript。',
  noJavaScript: '未发现 JavaScript 动作。',
  loadingSignatures: '正在验证数字签名…',
  signatureLoadFailed: '无法验证 PDF 数字签名。',
  noSignatures: '未发现已签署的签名值。',
  validSignatures: '密码学有效',
  invalidSignatures: '无效或无法读取',
  unsupportedSignatures: '暂不支持',
  signatureValid: '有效',
  signatureInvalid: '无效',
  signatureUnsupported: '暂不支持',
  signatureError: '无法验证',
  signatureNotChecked: '未检查',
  signatureField: '签名字段',
  signatureSigner: '签署人',
  reason: '签署原因',
  location: '签署地点',
  signatureType: '类型',
  approvalSignature: '批准签名',
  certificationSignature: '认证签名',
  documentTimestamp: '文档时间戳',
  unknownSignature: '未知签名',
  signedAt: '声明签署时间',
  timestampAt: '时间戳时间',
  signatureCoverage: '覆盖范围',
  entireDocument: '整个文件',
  earlierRevision: '较早版本',
  unsignedTail: '未签名尾部数据',
  digestAlgorithm: '摘要算法',
  signatureAlgorithm: '签名算法',
  certificateSubject: '证书主体',
  certificateIssuer: '证书颁发者',
  certificateSerial: '证书序列号',
  certificateValidFrom: '证书生效时间',
  certificateValidUntil: '证书失效时间',
  certificateValidAtSigning: '签署时证书有效',
  certificateSelfSigned: '自签名证书',
  trustNotEvaluated: '本地未评估证书信任链和吊销状态。密码学有效仅表示签名覆盖的字节未被修改。',
  copy: '复制',
  copyAll: '全部复制',
  copied: '已复制',
  reviewRemoval: '检查并移除',
  sourceNamed: '命名脚本',
  sourceDocument: '文档',
  sourcePage: '页面',
  sourceAnnotation: '批注',
  sourceForm: '表单字段',
  trigger: '触发器',
  scriptName: '脚本',
  field: '字段',
  annotationNumber: '批注',
  interactiveContent: '交互内容',
  formFields: '表单字段',
  signatures: '签名字段',
  xfa: 'XFA 表单',
  annotations: '批注',
  annotationTypes: '类型',
  pageDetails: '页面明细',
  dimensions: '尺寸',
  rotation: '旋转',
  none: '无',
  close: '关闭',
  loadingBookmarks: '正在读取…',
  noBookmarks: '没有目录书签',
  moveUp: '上移',
  moveDown: '下移',
  indentBookmark: '降低层级',
  outdentBookmark: '提升层级',
  deleteBookmark: '删除书签',
  invalidSplit: '请输入有效页码或正整数。',
  invalidFileSize: '请输入有效的正数文件大小。',
  invalidChapterLevels: '目录层级必须是 1 到 21 的整数。',
  invalidMargins: '边距必须是非负数。',
  invalidAutoCrop: '白色阈值必须为 0–255，保留边距必须为 0–144 点。',
  invalidGrid: '横向和纵向份数必须是 1 到 10 的整数。',
  invalidNup: '请检查行列、边距和边框宽度。',
  invalidGutter: '装订间距必须是非负数。',
  invalidPages: '请输入文档内有效的页面范围。',
  invalidExtractPages: '请按需要的顺序输入至少一个有效页码。',
  invalidColors: '请选择有效的文字色和背景色。',
  invalidBlankDetection: '白色容差必须在 0 到 255 之间，白色比例必须大于 0% 且不超过 100%。',
  invalidRemoval: '请选择要删除的页面，并至少保留一页。',
  invalidOrder: '请输入至少包含一页的有效页面顺序。',
  invalidCopies: '每页份数必须是 1 到 100 的整数。',
  invalidRenderDpi: '渲染 DPI 必须是 72 到 600 的整数。',
  invalidScannerEffect: '请输入有效的扫描件效果参数。',
  invalidDeskew: '请检查页面范围、最大倾斜角度和渲染 DPI。',
  invalidScannerImageSplit: '请输入有效的扫描照片检测参数。',
  invalidImageQuality: '图片质量必须是 10% 到 100% 的整数。',
  invalidLineArt: '线条灵敏度必须为 0% 到 100%，边缘强度必须为 1 到 3。',
  invalidPdfToImages: '请检查页面范围、DPI 和图片质量。',
  invalidPdfToVideo: '请检查页面范围、时长、过渡和视频总时长。',
  invalidPdfToCbz: '请检查页面范围、格式、DPI 和图片质量。',
  invalidPdfToHtml: '请检查页面范围和渲染 DPI。',
  invalidPdfToEpub: '请检查页面范围和渲染 DPI。',
  invalidPdfToPptx: '请检查页面范围和渲染 DPI。',
  invalidPdfToDocx: '请检查页面范围和渲染 DPI。',
  invalidPdfToRtf: '请检查页面范围。',
  invalidRedactionPatterns: '请至少输入一个有效的文字匹配规则。',
  invalidRedactionAreas: '请填写位于页面范围内的有效涂黑区域。',
  invalidRedactionPadding: '额外留白必须在 0 到 72 点之间。',
  invalidPipeline: '请添加 1 到 12 个有效处理步骤。',
  invalidComments: '每条评论都需要有效页码、坐标和评论内容。',
  invalidFormFields: '请填写所有必填表单字段。',
  invalidCreatedFields: '请检查字段名称、页码、坐标、尺寸、选项和默认值。',
  modifyAtLeastOneField: '请至少修改一个表单字段属性。',
  invalidModifiedFieldNames: '字段名称必须唯一，且不能包含空分组。',
  invalidChoiceOptions: '每个选择字段都需要至少一个选项。',
  selectFieldsToDelete: '请至少选择一个要删除的表单字段。',
  selectSanitize: '请至少选择一项要移除的内容。',
  selectOverlayFiles: '请至少选择一个用于叠加的 PDF。',
  selectOverlayImage: '请选择 PNG 或 JPEG 图片。',
  selectImagesToPdf: '请至少选择一张 PNG、JPEG 或 WebP 图片。',
  selectMergeFiles: '请至少添加一个 PDF。',
  selectComparisonFile: '请选择一个用于对比的 PDF。',
  invalidComparison: '渲染 DPI 必须为 72 到 300，差异容差必须为 0% 到 100%。',
  invalidRepeatCounts: '请为每个叠加 PDF 输入一个正整数重复次数。',
  invalidOpacity: '透明度必须在 0% 到 100% 之间。',
  invalidImageOverlay: '请检查页码、尺寸、边距和自定义坐标。',
  invalidContentFilter: '请检查页面范围和筛选条件。',
  invalidDocumentFilter: '请检查候选文件和文档筛选条件。',
  invalidBulkReplace: '请选择有效页面，并至少填写一项非空的查找内容。',
  invalidPageNumbers: '请检查页面范围、编号、字体、颜色和页码格式。',
  invalidMetadata: '请检查日期和自定义字段名称。',
  selectAttachmentFiles: '请至少选择一个要添加的附件文件。',
  selectExistingAttachment: '请选择一个已有附件。',
  invalidAttachmentName: '请输入有效的新附件名称。',
  attachmentLoadFailed: '无法读取嵌入附件。',
  bookmarkLoadFailed: '无法读取 PDF 目录书签。',
  invalidBookmarks: '每个书签都需要标题和有效页码。',
  invalidBookmarkJson: '请选择页码属于当前文档的有效目录 JSON。',
  bookmarkLimitReached: '目录最多支持 1000 项和 20 层。',
  completed: 'PDF 工具处理完成',
  outputs: '个文件',
}

const ZH_TW: typeof EN = {
  ...ZH,
  title: 'PDF 工具',
  split: '拆分',
  merge: '合併 PDF',
  imagesToPdf: '圖片 / SVG 轉 PDF',
  cbzToPdf: 'CBZ / CBR 轉 PDF',
  emailToPdf: '郵件轉 PDF',
  epubToPdf: 'EPUB 轉 PDF',
  htmlToPdf: '網頁 / HTML 轉 PDF',
  markdownToPdf: 'Markdown / ZIP 轉 PDF',
  pdfToImages: 'PDF 轉圖片',
  pdfToVideo: 'PDF 轉影片',
  pdfToCbz: 'PDF 轉 CBZ',
  pdfToHtml: 'PDF 轉 HTML',
  pdfToEpub: 'PDF 轉 EPUB',
  pdfToPptx: 'PDF 轉 PowerPoint',
  pdfToDocx: 'PDF 轉 Word',
  pdfToOdt: 'PDF 轉 ODT',
  pdfToRtf: 'PDF 轉 RTF',
  pdfToPdfa: '封存 PDF（PDF/A）',
  pdfToMarkdown: 'PDF 轉 Markdown',
  pdfToXlsx: 'PDF 轉 Excel',
  pdfToXml: 'PDF 轉 XML',
  compare: 'PDF 對比',
  extractPages: '擷取頁面',
  splitSections: '頁面分區',
  crop: '裁邊',
  cropMode: '裁邊方式',
  cropAuto: '自動識別',
  cropManual: '手動邊距',
  whiteThreshold: '白色閾值',
  cropPadding: '保留邊距',
  autoCropHint: '在本機識別頁面白邊，並保留原始向量內容。',
  scale: '縮放',
  nup: '多頁合一',
  booklet: '小冊子',
  poster: '海報分頁',
  singlePage: '長頁合併',
  rotatePages: '批次旋轉',
  autoRotate: '自動旋轉',
  deskew: '掃描頁校正',
  autoRename: '自動重新命名',
  pageNumbers: '新增頁碼',
  scannerEffect: '掃描件效果',
  scannerImageSplit: '掃描照片分割',
  autoSplit: 'QR Code 自動拆分',
  javascriptAudit: 'JavaScript 稽核',
  signatureAudit: '驗證數位簽章',
  fontInfo: '字型清單',
  annotationInfo: '註解清單',
  securityInfo: '安全與權限',
  certificateSign: '憑證簽章',
  timestamp: '文件時間戳記',
  password: '密碼保護',
  timestampAuthority: '時間戳記服務機構',
  timestampPrivacyHint: '只會向所選機構傳送 SHA-256 摘要，PDF 不會離開目前裝置。',
  timestampSecurityHint: '輸出檔案包含數位簽章，後續編輯會使時間戳記失效。',
  runTimestamp: '新增時間戳記',
  ocr: 'OCR 文字辨識',
  protectPdf: '保護 PDF',
  unlockPdf: '解鎖 PDF',
  openPassword: '開啟密碼',
  ownerPassword: '擁有者密碼',
  currentPassword: '目前密碼',
  confirmPassword: '確認密碼',
  encryption: '加密演算法',
  aesRecommended: '建議用於現代 PDF 閱讀器，支援 Unicode 密碼。',
  rc4Compatibility: '僅用於相容舊閱讀器，密碼字元範圍有限。',
  pdfPermissions: '允許的操作',
  allowPrinting: '列印',
  allowHighQualityPrint: '高品質列印',
  allowModifying: '修改內容',
  allowCopying: '複製文字和圖片',
  allowAnnotating: '新增註解',
  allowFillingForms: '填寫表單和簽章',
  allowExtraction: '輔助功能擷取',
  allowAssembly: '頁面組合',
  permissionAdvisory: 'PDF 權限依賴閱讀器遵循。如需內容保密，請設定開啟密碼。',
  passwordRewritesPdf: '保護或解鎖會重寫 PDF，並使現有數位簽章失效。',
  unlockHint: '匯出未加密副本，目前文件保持不變。',
  passwordRequired: '請至少輸入開啟密碼或擁有者密碼。',
  passwordMismatch: '兩次輸入的密碼不一致。',
  ownerPasswordRequired: '限制操作時必須設定不同的擁有者密碼。',
  ocrMode: '辨識模式',
  ocrSkipText: '自動 · 跳過已有文字頁',
  ocrForce: '強制 · 取代全部頁面',
  ocrStrict: '嚴格 · 發現文字即停止',
  ocrLanguages: '辨識語言',
  ocrEnglish: '英文',
  ocrChinese: '簡體中文',
  ocrRenderDpi: '渲染 DPI',
  ocrClean: '辨識前清理掃描影像',
  ocrCleanHint: '辨識前增強對比度並轉為灰階，適合背景發灰的掃描文件。',
  ocrSidecar: '同時匯出旁掛文字',
  ocrSidecarHint: '將可搜尋 PDF 與按頁整理的純文字一起封裝為 ZIP。',
  ocrLocalHint: '辨識完全在本機執行，語言模型首次載入後會由瀏覽器快取。',
  ocrRewritesPdf: 'OCR 會重寫處理頁面，並使現有數位簽章失效。',
  ocrNoLanguage: '請至少選擇一種辨識語言。',
  invalidOcrDpi: 'OCR 渲染 DPI 必須為 100 到 300 的整數。',
  runOcr: '開始 OCR',
  ocrAnalyzing: '正在檢查已有文字層…',
  ocrLoading: '正在載入本機 OCR 引擎…',
  ocrRecognizing: '正在辨識第 {page}/{total} 頁…',
  ocrFinishing: '正在產生可搜尋 PDF…',
  certificateFile: 'P12 / PFX 憑證',
  chooseCertificate: '選擇憑證',
  noCertificate: '尚未選擇憑證',
  certificatePassword: '憑證密碼',
  signerNameInput: '簽署人名稱',
  contactInfo: '聯絡資訊',
  certificateSignerPlaceholder: '可選，用於 PDF 顯示',
  certificateReasonPlaceholder: '文件核准',
  certificateLocationPlaceholder: '城市或組織',
  certificateContactPlaceholder: '電子郵件或電話（可選）',
  certificateRequired: '請選擇 P12 或 PFX 憑證檔案。',
  certificateTooLarge: '憑證檔案不能超過 16 MB。',
  certificateLocalHint: '憑證和密碼僅用於本次本機簽章，不會被儲存或上傳。',
  certificateSecurityHint: '系統會建立已簽章副本；之後對該副本的任何編輯都會使數位簽章失效。',
  runCertificateSign: '簽署 PDF',
  pdfaFormat: '封存標準',
  pdfaMode: '封存方式',
  pdfaAuto: '自動保真',
  pdfaRaster: '圖像封存',
  pdfaRenderDpi: '封存算繪 DPI',
  pdfaImageQuality: '封存圖像品質',
  pdfaAutoHint:
    '所有已用字型均已嵌入時保留可選文字和向量圖形，否則自動回退為圖像封存。互動表單、連結、附件和指令碼均在本機移除。',
  pdfaRasterHint:
    '始終建立圖像化的 PDF/A-2b 副本，優先保持視覺一致性，但會扁平化可選文字和向量圖形。',
  invalidPdfa: '算繪 DPI 必須為 72 到 600 的整數，圖像品質必須為 10% 到 100%。',
  removePages: '批次刪除',
  removeImages: '移除圖片',
  bulkReplaceText: '批次替換文字',
  extractText: '匯出文字',
  pdfToJson: 'PDF 轉 JSON',
  jsonToPdf: 'JSON 轉 PDF',
  extractTables: '擷取表格',
  extractImages: '擷取圖片',
  removeAnnotations: '移除註解',
  removeBlanks: '移除空白頁',
  invertColors: '反轉顏色',
  replaceColors: '換色與高對比度',
  adjustColors: '色彩調整',
  rearrange: '頁面重排',
  redact: '安全塗黑',
  comments: '新增便箋評論',
  compress: '壓縮 PDF',
  flatten: 'PDF 扁平化',
  forms: '表單處理',
  repair: '修復 PDF',
  decompress: '解壓縮 PDF',
  removeSignatures: '移除數位簽章',
  sanitize: '文件清理',
  pipeline: '處理流水線',
  overlay: 'PDF 疊加',
  overlayImage: '新增圖片',
  filterPages: '篩選頁面',
  filterDocuments: '批次篩選 PDF',
  attachments: '附件管理',
  bookmarks: '目錄書籤',
  metadata: '編輯中繼資料',
  info: 'PDF 資訊',
  preflight: '本機預檢',
  splitMethod: '拆分方式',
  afterPages: '在指定頁後拆分',
  byFileSize: '按檔案大小',
  pagesPerDocument: '每份固定頁數',
  documentCount: '平均拆成固定份數',
  chapters: '按目錄章節拆分',
  splitAfter: '拆分位置',
  splitValue: '數量',
  targetFileSize: '目標大小',
  chapterLevels: '包含目錄層級',
  allowDuplicateChapters: '同頁書籤分別輸出',
  sectionArrangement: '分區輸出順序',
  rowMajor: '按行從左到右',
  columnMajor: '按列從上到下',
  mergeSections: '合併為一個 PDF',
  pageSize: '頁面尺寸',
  pagesPerSheet: '每張頁數',
  customLayout: '自訂版面',
  nupPageOrder: '頁面排列',
  readingDirection: '閱讀方向',
  leftToRight: '從左向右',
  outerMargins: '外邊距',
  innerMargin: '頁間邊距',
  borderWidth: '邊框寬度',
  orientation: '頁面方向',
  invalidNup: '請檢查行列、邊距和邊框寬度。',
  spine: '裝訂方向',
  spineLeft: '左側裝訂',
  spineRight: '右側裝訂',
  gutter: '裝訂間距',
  duplexPass: '輸出頁面',
  both: '正反面',
  front: '僅正面',
  back: '僅反面',
  shortEdge: '短邊翻轉',
  columns: '橫向份數',
  rows: '縱向份數',
  rightToLeft: '從右向左',
  joinDirection: '合併方向',
  pages: '頁面範圍',
  pagesPlaceholder: 'all 或 1, 3-5',
  colorMode: '換色方式',
  highContrast: '高對比度',
  customColors: '自訂顏色',
  colorPreset: '配色方案',
  whiteOnBlack: '黑底白字',
  blackOnWhite: '白底黑字',
  yellowOnBlack: '黑底黃字',
  greenOnBlack: '黑底綠字',
  textColor: '文字顏色',
  backgroundColor: '背景顏色',
  contrast: '對比度',
  brightness: '亮度',
  saturation: '飽和度',
  rgbChannels: 'RGB 色版',
  redChannel: '紅色',
  greenChannel: '綠色',
  blueChannel: '藍色',
  whitenessThreshold: '白色容差',
  whitePercent: '最小白色比例',
  includeBlankPages: '另存偵測到的空白頁',
  outputFormat: '輸出格式',
  outputMode: '輸出方式',
  separateImages: '分頁圖片',
  singleLongImage: '單張長圖',
  includeAnnotations: '包含評論與表單值',
  pdfToImagesHint: '頁面只在本機渲染；多張圖片會打包成一個 ZIP 檔案下載。',
  videoDuration: '每頁時長（秒）',
  videoResolution: '影片解析度',
  videoTransition: '淡入轉場（秒）',
  videoEncoding: '正在編碼第 {page}/{total} 頁…',
  pdfToVideoHint: '不需 FFmpeg 或伺服器，在本機產生 WebM 幻燈片；最多 100 頁或 5 分鐘。',
  invalidPdfToVideo: '請檢查頁面範圍、時長、轉場和影片總時長。',
  pdfToCbzHint: '在本機渲染頁面並依自然閱讀順序產生漫畫書封存檔，不會上傳檔案。',
  cbzFile: '漫畫壓縮檔',
  chooseCbz: '選擇 CBZ / CBR',
  replaceCbz: '更換壓縮檔',
  selectCbz: '請選擇一個 CBZ、CBR、ZIP 或 RAR 漫畫壓縮檔。',
  cbzToPdfHint: 'ZIP 和 RAR 漫畫檔會在本機解壓並自然排序；隱藏檔案和非圖片項目會被忽略。',
  emailFile: '郵件檔案',
  chooseEmail: '選擇 EML',
  replaceEmail: '更換郵件',
  selectEmail: '請選擇一個 EML 郵件檔案。',
  emailOutputFormat: '郵件輸出',
  emailPdf: 'PDF 文件',
  emailHtml: '離線 HTML',
  emailIncludeAttachments: '嵌入原始附件',
  emailMaxAttachmentSize: '單一附件上限',
  emailIncludeAllRecipients: '包含副本和密件副本',
  emailLocalHint: '郵件只在本機解析；渲染前會移除遠端圖片和主動內容。',
  emailFromLabel: '寄件者',
  emailToLabel: '收件者',
  emailCcLabel: '副本',
  emailBccLabel: '密件副本',
  emailDateLabel: '日期',
  emailAttachmentsLabel: '附件',
  emailEmptyBody: '無法解析郵件內容。',
  emailRemoteImageRemoved: '遠端圖片已移除',
  emailUntitled: '無主旨郵件',
  invalidEmailAttachmentSize: '單一附件上限必須是 1 到 50 MB 的整數。',
  epubFile: 'EPUB 電子書',
  chooseEpub: '選擇 EPUB',
  replaceEpub: '更換 EPUB',
  selectEpub: '請選擇一個 EPUB 電子書檔案。',
  epubEmbedFonts: '嵌入書籍字型',
  epubIncludeToc: '新增目錄頁',
  epubIncludePageNumbers: '新增頁碼',
  epubOptimize: '針對電子書閱讀器最佳化',
  epubLocalHint: '閱讀順序、章節和本機圖片只在目前裝置解析；遠端和主動內容會被移除。',
  epubTocLabel: '目錄',
  epubUntitled: '未命名章節',
  jsonFile: '結構化 PDF JSON',
  chooseJson: '選擇 JSON',
  replaceJson: '更換 JSON',
  selectJson: '請選擇一個 GenOffice 結構化 PDF JSON v1 檔案。',
  jsonToPdfHint:
    '只在本機還原 GenOffice 結構化 PDF JSON v1；版面模式恢復文字座標，語意模式重新排版文字。此格式目前不包含圖片、註解、連結位置和表單位置。',
  htmlFile: '網頁文件',
  htmlSource: '來源',
  htmlSourceFile: '本機檔案',
  htmlSourceWebsite: '網頁位址',
  htmlWebsiteUrl: '網頁位址',
  htmlWebsitePlaceholder: 'https://example.com/article',
  chooseHtml: '選擇 HTML / ZIP',
  replaceHtml: '更換檔案',
  selectHtml: '請選擇一個 HTML、HTM 或 ZIP 網頁文件。',
  selectWebsiteUrl: '請輸入有效的網頁位址。',
  htmlZoom: '頁面縮放',
  htmlIncludePageNumbers: '新增頁碼',
  htmlLocalHint: 'HTML、CSS 和本機圖片只在目前裝置處理；遠端和主動內容會被移除。',
  htmlRemoteHint: '只安全取得公開網頁、樣式和圖片；主動內容及內網資源會被移除。',
  htmlUntitled: '未命名網頁文件',
  invalidHtmlZoom: '頁面縮放必須是 50% 到 200% 的整數。',
  markdownFile: 'Markdown 文件',
  chooseMarkdown: '選擇 Markdown / ZIP',
  replaceMarkdown: '更換檔案',
  selectMarkdown: '請選擇一個 MD、Markdown 或 ZIP 文件。',
  markdownZoom: '頁面縮放',
  markdownIncludePageNumbers: '新增頁碼',
  markdownLocalHint:
    'GFM 表格、清單、程式碼區塊和 ZIP 內本機圖片只在目前裝置處理；原始主動內容會被移除。',
  markdownUntitled: '未命名 Markdown 文件',
  invalidMarkdownZoom: '頁面縮放必須是 50% 到 200% 的整數。',
  pdfToHtmlHint: '在本機產生包含頁面圖片和可搜尋文字層的離線 ZIP，不會上傳檔案。',
  epubMode: '電子書版式',
  epubReflowable: '閱讀版',
  epubFixed: '保真版',
  pdfToEpubHint: '閱讀版會重排可選取文字；保真版使用本機頁面影像保留原始外觀。',
  pptxMode: 'PowerPoint 模式',
  pptxEditableText: '可編輯文字',
  pptxFidelity: '頁面保真',
  pdfToPptxHint:
    '可編輯文字會產生帶座標的文字框，不會假裝轉換圖片和複雜圖形；頁面保真會用整頁影像保留完整外觀。',
  docxMode: 'Word 模式',
  odtMode: 'OpenDocument 模式',
  docxEditableText: '可編輯文字',
  docxFidelity: '頁面保真',
  pdfToDocxHint: '可編輯文字會在本機依閱讀順序重建段落；頁面保真會用每頁一張影像保留完整外觀。',
  pdfToOdtHint:
    '在本機產生標準 ODT；可編輯文字保留閱讀順序和常用文字樣式，頁面保真會嵌入每頁影像。',
  pdfToRtfHint:
    '在本機依閱讀順序匯出可編輯富文字，保留常用字型、字級、粗體、斜體、Unicode 文字和分頁。',
  pdfToMarkdownPageBreaks: '保留頁面邊界',
  pdfToMarkdownHint: '依閱讀順序重建標題、段落、清單和連結；掃描文件請先執行 OCR。',
  pdfToXlsxHint: '在本機識別對齊表格，並為每個表格產生一張可編輯的 Excel 工作表。',
  pdfToXmlHint: '在本機匯出標準 XML，包含中繼資料、書籤、表單、連結和所選頁面。',
  pageOrder: '頁面順序',
  inferUndetected: '根據文件推斷文字不足的頁面',
  deskewMaxAngle: '最大傾斜角度',
  deskewHint: '在本機偵測掃描頁的小角度傾斜；僅校正頁面會點陣化，未變更頁面保留原始 PDF 內容。',
  renameStrategy: '標題來源',
  largestHeading: '最大字級標題',
  firstText: '首段文字',
  pageNumberPosition: '頁碼位置',
  pageNumberMargin: '頁邊距',
  small: '小',
  medium: '中',
  large: '大',
  xLarge: '超大',
  pageNumberFont: '字型',
  fontSize: '字級',
  fontColor: '文字顏色',
  startingNumber: '起始編號',
  zeroPad: '補零位數',
  textPattern: '頁碼格式',
  pageNumbersHint: '可用變數：{n}、{total}、{filename}。',
  invalidPageNumbers: '請檢查頁面範圍、編號、字型、顏色和頁碼格式。',
  mode: '重排方式',
  custom: '自訂順序',
  oddEven: '奇數頁後偶數頁',
  duplex: '雙面列印順序',
  removeFirst: '移除首頁',
  removeLast: '移除末頁',
  removeFirstAndLast: '移除首尾頁',
  duplicate: '重複每一頁',
  copies: '每頁份數',
  flattenMode: '扁平化方式',
  flattenPages: '整份文件',
  flattenFormsOnly: '僅表單',
  renderDpi: '渲染 DPI',
  imageQuality: '圖片品質',
  lineArt: '轉換為線稿',
  lineArtThreshold: '線條靈敏度',
  lineArtEdgeLevel: '邊緣強度',
  redactPatterns: '文字或比對規則',
  redactPatternsPlaceholder: '每行輸入一個關鍵詞',
  redactTextMode: '文字比對',
  redactAreaMode: '頁面區域',
  redactionMode: '塗黑方式',
  addRedactionArea: '新增區域',
  redactionArea: '區域',
  deleteRedactionArea: '刪除區域',
  redactionPreview: '塗黑區域預覽',
  resizeRedactionArea: '調整塗黑區域大小',
  redactLeft: '左側',
  redactTop: '頂部',
  redactWidth: '寬度',
  redactHeight: '高度',
  redactionAreaHint:
    '座標為從頁面左上角開始的百分比。命中頁面會安全點陣化，其他頁面保留原始 PDF 內容。',
  useRegex: '使用正規表示式',
  wholeWord: '僅比對完整單字',
  redactColor: '塗黑顏色',
  redactPadding: '額外留白',
  pipelineStep: '流水線步驟',
  addPipelineStep: '新增步驟',
  deletePipelineStep: '刪除步驟',
  flattenForms: '扁平化表單',
  pipelineHint: '依列表順序在本機逐步處理，並只匯出一份 PDF。調整步驟順序可控制最終結果。',
  addComment: '新增評論',
  comment: '評論',
  commentText: '評論內容',
  commentTextPlaceholder: '輸入在 PDF 閱讀器中顯示的便箋內容',
  anchorText: '錨點文字（可選）',
  anchorTextPlaceholder: '放在本頁第一個符合文字旁',
  fallbackPosition: '備用座標',
  positionX: 'X',
  positionY: 'Y',
  deleteComment: '刪除評論',
  formAction: '處理方式',
  exportFormData: '匯出資料',
  flattenFields: '扁平化表單',
  fillFields: '填寫表單',
  createFields: '新增欄位',
  unlock: '解除唯讀',
  modifyFields: '修改屬性',
  deleteFields: '刪除欄位',
  selectAllFields: '全選',
  clearFieldSelection: '清空',
  selectedFields: '已選欄位',
  modifiedFields: '已修改欄位',
  createdFields: '新增欄位',
  addField: '新增欄位',
  removeField: '移除欄位',
  fieldName: '欄位名稱',
  fieldLabel: '欄位標籤',
  fieldType: '欄位類型',
  fieldPage: '頁碼',
  fieldWidth: '寬度',
  fieldHeight: '高度',
  defaultValue: '預設值',
  defaultValues: '預設值（每行一項）',
  defaultChecked: '預設選取',
  multilineField: '允許多行',
  optionSpacing: '選項間距',
  choiceOptions: '選項（每行一項）',
  multiselectField: '允許多選',
  formCreateHint: '座標使用 PDF 點，原點位於頁面左下角。',
  loadingFormFields: '正在讀取表單欄位…',
  formFieldsLoadFailed: '無法讀取 PDF 表單欄位。',
  noFormFields: '未發現 AcroForm 表單欄位。',
  readOnlyField: '唯讀',
  requiredField: '必填',
  selectEmpty: '不選擇',
  unsupportedFormField: '此欄位類型無法在這裡填寫。',
  formText: '文字',
  formCheckbox: '核取方塊',
  formRadio: '單選群組',
  formDropdown: '下拉清單',
  formOptionList: '多選清單',
  formButton: '按鈕',
  formSignature: '簽名',
  formUnknown: '未知欄位',
  textExportFormat: '輸出格式',
  plainText: '純文字',
  markdown: 'Markdown',
  exportBoth: '兩種格式',
  textExportHint: '在本機匯出所選頁面；掃描文件請先執行 OCR。',
  structuredExportMode: '結構詳細程度',
  semanticJson: '精簡語意',
  layoutJson: '完整版面',
  pdfToJsonHint: '在本機匯出中繼資料、書籤、表單、連結和所選頁面；完整版面還包含文字座標與字型。',
  tableExportFormat: '輸出格式',
  csv: 'CSV',
  excel: 'Excel',
  includeTwoColumnTables: '識別純文字雙欄表格',
  tableExportHint: '在本機識別對齊表格；複雜、旋轉或掃描表格建議匯出後複核。',
  downloadDivider: '下載分隔頁',
  duplexMode: '雙面掃描模式',
  autoSplitHint: '在本機掃描相容 QR Code，移除分隔頁，並將各份文件封裝為 ZIP。',
  removeJavaScript: '移除 JavaScript',
  removeEmbeddedFiles: '移除嵌入檔案',
  removeXmpMetadata: '移除 XMP 中繼資料',
  removeMetadata: '移除文件屬性',
  removeLinks: '移除外部連結動作',
  overlayFiles: '疊加 PDF',
  chooseFiles: '選擇 PDF',
  noFiles: '未選擇檔案',
  overlayMode: '頁面分配',
  sequential: '順序循環',
  interleaved: '檔案交錯',
  fixedRepeat: '固定次數',
  repeatCounts: '重複次數',
  layer: '疊加層級',
  foreground: '前景',
  background: '背景',
  opacity: '透明度',
  attachmentAction: '操作',
  addAttachments: '新增檔案',
  extractAttachments: '匯出 ZIP',
  renameAttachment: '重新命名',
  deleteAttachment: '刪除',
  attachmentFiles: '附件檔案',
  chooseAttachments: '選擇檔案',
  existingAttachments: '現有附件',
  loadingAttachments: '正在讀取…',
  noAttachments: '沒有嵌入附件',
  newAttachmentName: '新名稱',
  addBookmark: '新增書籤',
  addChildBookmark: '新增子層',
  bookmarkImportMode: '匯入方式',
  bookmarkReplace: '取代',
  bookmarkAppend: '附加',
  importBookmarkJson: '匯入 JSON',
  pasteBookmarkJson: '貼上 JSON',
  downloadBookmarkJson: '下載 JSON',
  copyBookmarkJson: '複製 JSON',
  bookmarkJsonHint: 'JSON 使用 title、pageNumber 與巢狀 children，最多 1000 項 / 20 層。',
  bookmarkJsonImported: '目錄 JSON 已匯入。',
  bookmarkJsonCopied: '目錄 JSON 已複製。',
  bookmarkTitle: '標題',
  bookmarkPage: '頁碼',
  loadingAnalysis: '正在分析…',
  fontInfoLoadFailed: '無法讀取 PDF 字型資訊。',
  fontSummary: '字型概覽',
  embeddedFonts: '已嵌入字型',
  subsetFonts: '子集字型',
  fontName: '字型名稱',
  fontSubtype: '類型',
  fontEncoding: '編碼',
  fontEmbedded: '嵌入',
  fontSubset: '子集',
  fontToUnicode: 'ToUnicode',
  fontPages: '使用頁碼',
  noFonts: '未發現 PDF 字型資源。',
  annotationInfoLoadFailed: '無法讀取 PDF 註解資訊。',
  annotationSummary: '註解概覽',
  annotationIndex: '序號',
  annotationSubtype: '類型',
  annotationContents: '內容',
  annotationModified: '修改時間',
  annotationPosition: '位置',
  annotationFlags: '旗標',
  noAnnotations: '未發現 PDF 註解。',
  securityInfoLoadFailed: '無法讀取 PDF 安全資訊。',
  securitySummary: '安全概覽',
  encryptionAlgorithm: '加密演算法',
  keyLength: '金鑰長度',
  securityVersion: '處理器版本 / 修訂版',
  restrictedActions: '受限操作',
  permissionAllowed: '允許',
  permissionRestricted: '受限',
  permissionUnavailable: '無法讀取加密權限字典。',
  analysisLoadFailed: '無法分析此 PDF。',
  preflightLoadFailed: '無法完成 PDF 本機預檢。',
  preflightStatus: '預檢狀態',
  preflightPass: '結構檢查通過',
  preflightWarning: '存在待檢查警告',
  preflightError: '發現結構錯誤',
  downloadPreflightJson: '下載 JSON',
  reviewRepair: '開啟修復工具',
  preflightDisclaimer:
    '此報告僅執行本機結構預檢和標準宣告偵測，不等同於 veraPDF 或無障礙正式合規認證。',
  structure: '檔案結構',
  strictParsing: '嚴格解析',
  headerOffset: 'PDF 檔頭偏移',
  eofMarkers: 'EOF 標記',
  trailingBytes: '尾部資料',
  startXref: 'startxref 偏移',
  incrementalUpdates: '增量更新次數',
  standardDeclarations: '標準宣告',
  noStandardDeclarations: '未發現 PDF/A 或 PDF/UA 宣告',
  xmpMetadata: 'XMP 中繼資料',
  xmpValid: 'XMP 可解析',
  taggedPdf: '標籤化 PDF',
  documentLanguage: '文件語言',
  outputIntents: '輸出意圖',
  findings: '檢查結果',
  noPreflightFindings: '未發現結構問題。',
  findingMissingPdfHeader: '缺少 PDF 檔頭。',
  findingHeaderNotAtStart: 'PDF 檔頭不在檔案起始位置。',
  findingMissingEofMarker: '缺少最終 EOF 標記。',
  findingLargeTrailingData: '最終 EOF 標記後存在大量尾部資料。',
  findingMissingStartXref: '缺少 startxref 指標。',
  findingStartXrefOutOfRange: 'startxref 指標超出檔案範圍。',
  findingStartXrefTargetInvalid: 'startxref 指標未指向交叉參照資料。',
  findingStrictParseFailed: 'PDF 嚴格解析失敗。',
  findingEmptyPageTree: 'PDF 沒有頁面。',
  findingInvalidPageBox: '頁面尺寸無效。',
  findingInvalidXmpMetadata: '存在 XMP 中繼資料，但無法解析。',
  findingStandardDeclaredOnly: '文件宣告了標準，但本機預檢不認證正式合規。',
  findingPdfaMissingOutputIntent: '宣告了 PDF/A，但未發現輸出意圖。',
  findingPdfaAttachmentsRisk: '附件可能與宣告的 PDF/A 版本不相容。',
  findingPdfaJavaScriptRisk: 'JavaScript 不符合 PDF/A 的通常要求。',
  findingPdfuaNotTagged: '宣告了 PDF/UA，但文件未標記為標籤化 PDF。',
  findingPdfuaMissingLanguage: '宣告了 PDF/UA，但缺少文件語言。',
  findingJavascriptPresent: '文件包含 JavaScript 動作。',
  findingAttachmentsPresent: '文件包含嵌入附件。',
  findingXfaPresent: '文件包含 XFA 表單，部分閱讀器可能不支援。',
  findingSignaturesPresent: '文件包含數位簽章欄位。',
  metadataLoadFailed: '無法讀取 PDF 中繼資料。',
  overview: '基本資訊',
  pageCount: '頁數',
  pdfVersion: 'PDF 版本',
  fileSize: '檔案大小',
  encrypted: '已加密',
  documentProperties: '文件屬性',
  documentTitle: '標題',
  author: '作者',
  subject: '主題',
  keywords: '關鍵字',
  creator: '建立程式',
  producer: '產生程式',
  creationDate: '建立時間',
  modificationDate: '修改時間',
  trapped: '補漏白狀態',
  trappedTrue: '是',
  trappedFalse: '否',
  trappedUnknown: '未知',
  customMetadata: '自訂中繼資料',
  customMetadataKey: '鍵',
  customMetadataValue: '值',
  addCustomMetadata: '新增欄位',
  deleteAllMetadata: '清除全部文件中繼資料',
  resources: '資源統計',
  fonts: '字體',
  images: '圖片',
  embeddedFiles: '附件',
  bookmarkCount: '目錄書籤',
  interactiveContent: '互動內容',
  formFields: '表單欄位',
  signatures: '簽名欄位',
  xfa: 'XFA 表單',
  annotations: '註解',
  annotationTypes: '類型',
  pageDetails: '頁面明細',
  dimensions: '尺寸',
  rotation: '旋轉',
  none: '無',
  close: '關閉',
  loadingBookmarks: '正在讀取…',
  noBookmarks: '沒有目錄書籤',
  moveUp: '上移',
  moveDown: '下移',
  indentBookmark: '降低層級',
  outdentBookmark: '提升層級',
  deleteBookmark: '刪除書籤',
  selectOverlayFiles: '請至少選擇一個用於疊加的 PDF。',
  mergeFiles: 'PDF 順序',
  currentPdf: '目前 PDF',
  addMergeFiles: '新增 PDF',
  deleteFile: '移除檔案',
  comparisonFile: '對比 PDF',
  chooseComparisonFile: '選擇 PDF',
  comparisonTolerance: '差異容差',
  selectMergeFiles: '請至少新增一個 PDF。',
  selectComparisonFile: '請選擇一個用於對比的 PDF。',
  invalidComparison: '渲染 DPI 必須為 72 到 300，差異容差必須為 0% 到 100%。',
  invalidRepeatCounts: '請為每個疊加 PDF 輸入一個正整數重複次數。',
  invalidOpacity: '透明度必須在 0% 到 100% 之間。',
  invalidImageQuality: '圖片品質必須是 10% 到 100% 的整數。',
  invalidLineArt: '線條靈敏度必須為 0% 到 100%，邊緣強度必須為 1 到 3。',
  invalidRedactionPatterns: '請至少輸入一個有效的文字比對規則。',
  invalidRedactionAreas: '請填寫位於頁面範圍內的有效塗黑區域。',
  invalidRedactionPadding: '額外留白必須在 0 到 72 點之間。',
  invalidPipeline: '請新增 1 到 12 個有效處理步驟。',
  invalidComments: '每則評論都需要有效頁碼、座標和評論內容。',
  invalidFormFields: '請填寫所有必填表單欄位。',
  invalidCreatedFields: '請檢查欄位名稱、頁碼、座標、尺寸、選項和預設值。',
  modifyAtLeastOneField: '請至少修改一個表單欄位屬性。',
  invalidModifiedFieldNames: '欄位名稱必須唯一，且不能包含空群組。',
  invalidChoiceOptions: '每個選擇欄位都需要至少一個選項。',
  selectFieldsToDelete: '請至少選擇一個要刪除的表單欄位。',
  invalidExtractPages: '請按需要的順序輸入至少一個有效頁碼。',
  scanQuality: '掃描品質',
  scanHigh: '高',
  scanMedium: '中',
  scanLow: '低',
  scanCustom: '自訂',
  scanAngleThreshold: '旋轉校正門檻',
  scanTolerance: '背景色容差',
  minimumPhotoArea: '最小照片面積',
  minimumForegroundArea: '最小前景面積',
  detectionBorder: '偵測邊緣',
  scanRotation: '頁面傾斜',
  scanNone: '無',
  scanSlight: '輕微',
  scanModerate: '中等',
  scanSevere: '明顯',
  scanColorspace: '色彩空間',
  grayscale: '灰階',
  color: '彩色',
  paperBorder: '紙張邊緣',
  baseRotation: '基礎角度',
  rotationVariance: '隨機偏移',
  invalidScannerImageSplit: '請輸入有效的掃描照片偵測參數。',
  blur: '模糊',
  noise: '雜訊',
  yellowedPaper: '泛黃紙張',
  textureSeed: '紋理種子',
  invalidScannerEffect: '請輸入有效的掃描件效果參數。',
  invalidDeskew: '請檢查頁面範圍、最大傾斜角度和渲染 DPI。',
  invalidPdfToHtml: '請檢查頁面範圍和渲染 DPI。',
  invalidPdfToEpub: '請檢查頁面範圍和渲染 DPI。',
  invalidPdfToPptx: '請檢查頁面範圍和渲染 DPI。',
  invalidPdfToDocx: '請檢查頁面範圍和渲染 DPI。',
  invalidPdfToRtf: '請檢查頁面範圍。',
  invalidPdfToCbz: '請檢查頁面範圍、格式、DPI 和圖片品質。',
  invalidColors: '請選擇有效的文字色和背景色。',
  invalidBlankDetection: '白色容差必須在 0 到 255 之間，白色比例必須大於 0% 且不超過 100%。',
  selectAttachmentFiles: '請至少選擇一個要新增的附件檔案。',
  selectExistingAttachment: '請選擇一個現有附件。',
  invalidAttachmentName: '請輸入有效的新附件名稱。',
  attachmentLoadFailed: '無法讀取嵌入附件。',
  overlayImageFile: '圖片',
  chooseImage: '選擇圖片',
  overlayImagePosition: '位置',
  topLeft: '左上',
  topCenter: '上中',
  topRight: '右上',
  middleLeft: '左中',
  center: '置中',
  middleRight: '右中',
  bottomLeft: '左下',
  bottomCenter: '下中',
  bottomRight: '右下',
  customPosition: '自訂',
  overlayImageWidth: '頁面寬度占比',
  overlayImageMargin: '邊距',
  customPositionHint: '座標以可見頁面左上角為起點。',
  overlayImageHint: '在本機將 PNG 或 JPEG 新增到所選頁面。',
  imagePdfFiles: '圖片 / SVG 順序',
  addImages: '新增圖片 / SVG',
  svgVectorHint: '彩色 SVG 會保留向量內容，縮放依然清晰；灰階模式會在本機點陣化。',
  cameraStart: '開啟相機',
  cameraStop: '關閉相機',
  cameraSwitch: '切換相機',
  cameraCapture: '拍攝此頁',
  cameraReadyHint: '相機預覽',
  cameraPrivacyHint: '在本機掃描頁面，相機畫面不會離開此裝置。',
  cameraDenied: '相機權限遭拒，請在瀏覽器或系統設定中允許後重試。',
  cameraUnavailable: '找不到可用相機，或相機正由其他應用程式使用。',
  cameraUnsupported: '相機掃描需要安全瀏覽器環境或受支援的桌面系統。',
  cameraCaptureFailed: '相機畫面尚未就緒，請重試。',
  mobileScannerStart: '手機掃碼',
  mobileScannerStop: '結束工作階段',
  mobileScannerTitle: '手機上傳',
  mobileScannerHint: '用可存取目前服務的手機掃描 QR Code。',
  mobileScannerWaiting: '等待手機傳送圖片…',
  mobileScannerReceived: '已接收圖片：',
  mobileScannerCopyLink: '複製連結',
  mobileScannerCopied: '已複製',
  mobileScannerExpiresIn: '剩餘時間',
  mobileScannerExpired: '手機上傳工作階段已過期。',
  mobileScannerUnavailable: '手機掃碼功能暫不可用。',
  imageFit: '頁面適配',
  fillPage: '鋪滿 A4 頁面',
  fitDocumentToImage: '使用圖片尺寸',
  maintainAspectRatio: '等比放入 A4',
  imageColorMode: '色彩模式',
  blackWhite: '黑白',
  autoRotateImages: '橫圖使用 A4 橫向頁面',
  imagePdfDestination: '輸出位置',
  newPdf: '新建 PDF',
  appendCurrent: '附加到目前 PDF',
  selectOverlayImage: '請選擇 PNG 或 JPEG 圖片。',
  selectImagesToPdf: '請至少選擇一張 PNG、JPEG 或 WebP 圖片。',
  invalidImageOverlay: '請檢查頁碼、尺寸、邊距和自訂座標。',
  filterCriterion: '篩選條件',
  filterText: '文字',
  filterImage: '圖片',
  filterPageSize: '頁面尺寸',
  filterOrientation: '頁面方向',
  filterRotation: '頁面旋轉',
  filterComparator: '比較方式',
  filterLess: '小於',
  filterEqual: '等於',
  filterGreater: '大於',
  filterQuery: '尋找文字',
  filterQueryPlaceholder: '輸入要尋找的原文',
  caseSensitive: '區分大小寫',
  filterAction: '輸出結果',
  keepMatches: '保留符合頁面',
  removeMatches: '移除符合頁面',
  filterPagesHint: '在本機檢查所選頁面並匯出新 PDF。',
  filterDocumentsHint: '在本機檢查目前 PDF 和加入的檔案，僅原樣匯出符合條件的文件。',
  filterDocumentsFiles: '候選 PDF',
  addFilterDocuments: '加入 PDF',
  currentDocument: '目前文件',
  removeFilterDocument: '移除 PDF',
  filterPageCount: '頁數',
  filterFileSize: '檔案大小',
  filterContainsText: '包含文字',
  filterContainsImage: '包含圖片',
  filterFirstPageSize: '首頁尺寸',
  filterFirstPageRotation: '首頁旋轉',
  filterExpectedValue: '目標值',
  filterSizeUnit: '單位',
  replacementRules: '替換規則',
  findText: '尋找',
  replaceWith: '替換為',
  addReplacement: '新增規則',
  deleteReplacement: '刪除規則',
  emptyReplacementDeletes: '替換內容留空即可刪除符合的文字。',
  bulkReplaceHint: '規則會由上到下依序執行，後續規則可看到前序結果，並在本機匯出新的 PDF 副本。',
  invalidContentFilter: '請檢查頁面範圍和篩選條件。',
  invalidDocumentFilter: '請檢查候選檔案和文件篩選條件。',
  invalidBulkReplace: '請選擇有效頁面，並至少填寫一項非空的尋找內容。',
  invalidMetadata: '請檢查日期和自訂欄位名稱。',
  bookmarkLoadFailed: '無法讀取 PDF 目錄書籤。',
  invalidBookmarks: '每個書籤都需要標題和有效頁碼。',
  invalidBookmarkJson: '請選擇頁碼屬於目前文件的有效目錄 JSON。',
  bookmarkLimitReached: '目錄最多支援 1000 項與 20 層。',
  javaScriptActions: '腳本動作',
  uniqueScripts: '不重複腳本',
  scriptCodeSize: '程式碼大小',
  loadingJavaScript: '正在稽核 JavaScript…',
  javaScriptLoadFailed: '無法稽核 PDF JavaScript。',
  noJavaScript: '未發現 JavaScript 動作。',
  loadingSignatures: '正在驗證數位簽章…',
  signatureLoadFailed: '無法驗證 PDF 數位簽章。',
  noSignatures: '未發現已簽署的簽章值。',
  validSignatures: '密碼學有效',
  invalidSignatures: '無效或無法讀取',
  unsupportedSignatures: '暫不支援',
  signatureValid: '有效',
  signatureInvalid: '無效',
  signatureUnsupported: '暫不支援',
  signatureError: '無法驗證',
  signatureNotChecked: '未檢查',
  signatureField: '簽章欄位',
  signatureSigner: '簽署人',
  reason: '簽署原因',
  location: '簽署地點',
  signatureType: '類型',
  approvalSignature: '核准簽章',
  certificationSignature: '認證簽章',
  documentTimestamp: '文件時間戳記',
  unknownSignature: '未知簽章',
  signedAt: '聲明簽署時間',
  timestampAt: '時間戳記時間',
  signatureCoverage: '涵蓋範圍',
  entireDocument: '整個檔案',
  earlierRevision: '較早版本',
  unsignedTail: '未簽章尾端資料',
  digestAlgorithm: '摘要演算法',
  signatureAlgorithm: '簽章演算法',
  certificateSubject: '憑證主體',
  certificateIssuer: '憑證頒發者',
  certificateSerial: '憑證序號',
  certificateValidFrom: '憑證生效時間',
  certificateValidUntil: '憑證失效時間',
  certificateValidAtSigning: '簽署時憑證有效',
  certificateSelfSigned: '自簽憑證',
  trustNotEvaluated: '本機未評估憑證信任鏈和撤銷狀態。密碼學有效僅表示簽章涵蓋的位元組未被修改。',
  copy: '複製',
  copyAll: '全部複製',
  copied: '已複製',
  reviewRemoval: '檢查並移除',
  sourceNamed: '命名腳本',
  sourceDocument: '文件',
  sourcePage: '頁面',
  sourceAnnotation: '註解',
  sourceForm: '表單欄位',
  trigger: '觸發器',
  scriptName: '腳本',
  field: '欄位',
  annotationNumber: '註解',
  portrait: '直向',
  landscape: '橫向',
  completed: 'PDF 工具處理完成',
  outputs: '個檔案',
}

const PREFLIGHT_FINDING_TEXT_KEYS: Record<PdfPreflightFinding['code'], keyof typeof EN> = {
  missingPdfHeader: 'findingMissingPdfHeader',
  headerNotAtStart: 'findingHeaderNotAtStart',
  missingEofMarker: 'findingMissingEofMarker',
  largeTrailingData: 'findingLargeTrailingData',
  missingStartXref: 'findingMissingStartXref',
  startXrefOutOfRange: 'findingStartXrefOutOfRange',
  startXrefTargetInvalid: 'findingStartXrefTargetInvalid',
  strictParseFailed: 'findingStrictParseFailed',
  emptyPageTree: 'findingEmptyPageTree',
  invalidPageBox: 'findingInvalidPageBox',
  invalidXmpMetadata: 'findingInvalidXmpMetadata',
  standardDeclaredOnly: 'findingStandardDeclaredOnly',
  pdfaMissingOutputIntent: 'findingPdfaMissingOutputIntent',
  pdfaAttachmentsRisk: 'findingPdfaAttachmentsRisk',
  pdfaJavaScriptRisk: 'findingPdfaJavaScriptRisk',
  pdfuaNotTagged: 'findingPdfuaNotTagged',
  pdfuaMissingLanguage: 'findingPdfuaMissingLanguage',
  javascriptPresent: 'findingJavascriptPresent',
  attachmentsPresent: 'findingAttachmentsPresent',
  xfaPresent: 'findingXfaPresent',
  signaturesPresent: 'findingSignaturesPresent',
}

function textFor(lang: Lang): typeof EN {
  if (lang === 'zh') return ZH
  if (lang === 'zh-TW') return ZH_TW
  return EN
}

export function pdfToolsTabLabel(lang: Lang): string {
  return textFor(lang).tab
}

export function pdfToolLabel(kind: PdfToolKind, lang: Lang): string {
  return textFor(lang)[kind]
}

export function pdfToolSuccessText(lang: Lang, count: number): string {
  const text = textFor(lang)
  return `${text.completed} · ${count} ${text.outputs}`
}

function parseSplitPoints(value: string, pageCount: number): number[] | null {
  const normalized = value.trim().toLowerCase()
  if (normalized === 'all') return Array.from({ length: Math.max(0, pageCount - 1) }, (_, i) => i)
  if (!normalized) return null
  const points = new Set<number>()
  for (const token of normalized.split(/[,，;；\s]+/).filter(Boolean)) {
    const match = /^(\d+)(?:-(\d+))?$/.exec(token)
    if (!match) return null
    const start = Number(match[1])
    const end = Number(match[2] ?? match[1])
    if (start < 1 || end < start || end >= pageCount) return null
    for (let page = start; page <= end; page++) points.add(page - 1)
  }
  return [...points].sort((left, right) => left - right)
}

function parsePageIndexes(
  value: string,
  pageCount: number,
  allowDuplicates = false,
): number[] | null {
  const normalized = value.trim().toLowerCase()
  if (normalized === 'all') return Array.from({ length: pageCount }, (_, pageIndex) => pageIndex)
  if (!normalized) return null
  const indexes: number[] = []
  for (const token of normalized.split(/[,，;；\s]+/).filter(Boolean)) {
    const match = /^(\d+)(?:-(\d+))?$/.exec(token)
    if (!match) return null
    const start = Number(match[1])
    const end = Number(match[2] ?? match[1])
    if (start < 1 || end < start || end > pageCount) return null
    for (let page = start; page <= end; page++) indexes.push(page - 1)
  }
  return allowDuplicates ? indexes : [...new Set(indexes)]
}

function parseRepeatCounts(value: string, fileCount: number): number[] | null {
  const counts = value
    .trim()
    .split(/[,，;；\s]+/)
    .filter(Boolean)
    .map(Number)
  if (
    counts.length !== fileCount ||
    counts.some((count) => !Number.isInteger(count) || count < 1)
  ) {
    return null
  }
  return counts
}

function attachmentSizeLabel(size: number): string {
  if (size < 1024) return `${size} B`
  if (size < 1024 * 1024) return `${Math.round(size / 1024)} KB`
  return `${(size / (1024 * 1024)).toFixed(1)} MB`
}

function analysisDateLabel(value: string, lang: Lang): string {
  const date = new Date(value)
  return Number.isNaN(date.getTime()) ? value : date.toLocaleString(lang)
}

const PDF_PERMISSION_KEYS = [
  'allowPrinting',
  'allowHighQualityPrint',
  'allowModifying',
  'allowCopying',
  'allowAnnotating',
  'allowFillingForms',
  'allowExtraction',
  'allowAssembly',
] as const

function emptyPdfMetadata(): PdfMetadataValues {
  return {
    title: '',
    author: '',
    subject: '',
    keywords: '',
    creator: '',
    producer: '',
    creationDate: '',
    modificationDate: '',
    trapped: '',
    custom: [],
  }
}

function metadataInputDate(value?: string): string {
  if (!value) return ''
  const date = new Date(value)
  if (Number.isNaN(date.getTime())) return ''
  const localDate = new Date(date.getTime() - date.getTimezoneOffset() * 60_000)
  return localDate.toISOString().slice(0, 16)
}

function hasPdfProperties(properties: PdfAnalysis['properties']): boolean {
  return (
    Boolean(
      properties.title ||
      properties.author ||
      properties.subject ||
      properties.keywords ||
      properties.creator ||
      properties.producer ||
      properties.creationDate ||
      properties.modificationDate ||
      properties.trapped,
    ) || properties.custom.length > 0
  )
}

function updateCustomMetadataField(
  fields: PdfMetadataCustomField[],
  index: number,
  patch: Partial<PdfMetadataCustomField>,
): PdfMetadataCustomField[] {
  return fields.map((field, fieldIndex) => (fieldIndex === index ? { ...field, ...patch } : field))
}

function javaScriptActionText(action: PdfJavaScriptAction): string {
  const details = [
    action.name ? `name=${action.name}` : '',
    action.pageNumber ? `page=${action.pageNumber}` : '',
    action.annotationNumber ? `annotation=${action.annotationNumber}` : '',
    action.fieldName ? `field=${action.fieldName}` : '',
  ].filter(Boolean)
  return `// ${action.source} / ${action.trigger}${details.length ? ` / ${details.join(' / ')}` : ''}\n${action.code}`
}

function initialFormFieldValue(field: PdfFormFieldInfo): PdfFormFieldValue['value'] {
  if (field.type === 'checkbox') return field.value === true
  if (field.type === 'text' || field.type === 'radio') {
    return typeof field.value === 'string' ? field.value : ''
  }
  if (field.type === 'dropdown' || field.type === 'optionList') {
    return Array.isArray(field.value) ? field.value : []
  }
  return ''
}

interface PdfFormFieldEdit {
  newName: string
  label: string
  readOnly: boolean
  required: boolean
  optionsText: string
  multiselect: boolean
}

interface PdfFormFieldCreationEdit {
  id: number
  name: string
  label: string
  type: PdfCreatableFormFieldType
  pageNumber: number
  x: number
  y: number
  width: number
  height: number
  required: boolean
  readOnly: boolean
  multiline: boolean
  multiselect: boolean
  optionsText: string
  defaultText: string
  defaultChecked: boolean
  optionSpacing: number
}

let nextPdfFormCreationId = 1

function emptyPdfFormFieldCreation(pageNumber: number): PdfFormFieldCreationEdit {
  return {
    id: nextPdfFormCreationId++,
    name: '',
    label: '',
    type: 'text',
    pageNumber,
    x: 72,
    y: 72,
    width: 180,
    height: 24,
    required: false,
    readOnly: false,
    multiline: false,
    multiselect: false,
    optionsText: '',
    defaultText: '',
    defaultChecked: false,
    optionSpacing: 30,
  }
}

function parsedFormFieldOptions(value: string): string[] {
  return [
    ...new Set(
      value
        .split(/\r?\n/)
        .map((option) => option.trim())
        .filter(Boolean),
    ),
  ]
}

function resolvedFormFieldName(originalName: string, requestedName: string): string | null {
  const name = requestedName.trim()
  const parts = name.split('.')
  if (!name || parts.some((part) => !part.trim())) return null
  if (parts.length > 1) {
    const originalParent = originalName.split('.').slice(0, -1).join('.')
    return parts.slice(0, -1).join('.') === originalParent ? name : null
  }
  const originalParent = originalName.split('.').slice(0, -1).join('.')
  return originalParent ? `${originalParent}.${name}` : name
}

function editableFormField(field: PdfFormFieldInfo): boolean {
  return (
    !field.readOnly && ['text', 'checkbox', 'radio', 'dropdown', 'optionList'].includes(field.type)
  )
}

async function copyPdfToolText(value: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(value)
    return true
  } catch {
    const textArea = document.createElement('textarea')
    textArea.value = value
    textArea.style.position = 'fixed'
    textArea.style.opacity = '0'
    document.body.appendChild(textArea)
    textArea.select()
    const copied = document.execCommand('copy')
    textArea.remove()
    return copied
  }
}

function downloadPdfToolText(fileName: string, value: string, type = 'text/plain'): void {
  const url = URL.createObjectURL(new Blob([value], { type }))
  const link = document.createElement('a')
  link.href = url
  link.download = fileName
  document.body.appendChild(link)
  link.click()
  link.remove()
  window.setTimeout(() => URL.revokeObjectURL(url), 0)
}

function bookmarkAt(bookmarks: PdfBookmark[], path: number[]): PdfBookmark | undefined {
  let level = bookmarks
  let bookmark: PdfBookmark | undefined
  for (const index of path) {
    bookmark = level[index]
    if (!bookmark) return undefined
    level = bookmark.children
  }
  return bookmark
}

function bookmarkSubtreeDepth(bookmark: PdfBookmark): number {
  return 1 + Math.max(0, ...bookmark.children.map(bookmarkSubtreeDepth))
}

function transformBookmarkLevel(
  bookmarks: PdfBookmark[],
  parentPath: number[],
  transform: (level: PdfBookmark[]) => PdfBookmark[],
): PdfBookmark[] {
  if (parentPath.length === 0) return transform(bookmarks)
  const [index, ...remainingPath] = parentPath
  return bookmarks.map((bookmark, bookmarkIndex) =>
    bookmarkIndex === index
      ? {
          ...bookmark,
          children: transformBookmarkLevel(bookmark.children, remainingPath, transform),
        }
      : bookmark,
  )
}

function updateBookmarkAt(
  bookmarks: PdfBookmark[],
  path: number[],
  update: (bookmark: PdfBookmark) => PdfBookmark,
): PdfBookmark[] {
  const index = path.at(-1)
  if (index === undefined) return bookmarks
  return transformBookmarkLevel(bookmarks, path.slice(0, -1), (level) =>
    level.map((bookmark, bookmarkIndex) => (bookmarkIndex === index ? update(bookmark) : bookmark)),
  )
}

function moveBookmarkAt(bookmarks: PdfBookmark[], path: number[], offset: -1 | 1): PdfBookmark[] {
  const index = path.at(-1)
  if (index === undefined) return bookmarks
  return transformBookmarkLevel(bookmarks, path.slice(0, -1), (level) => {
    const target = index + offset
    if (target < 0 || target >= level.length) return level
    const next = [...level]
    const moving = next[index]!
    next[index] = next[target]!
    next[target] = moving
    return next
  })
}

function deleteBookmarkAt(bookmarks: PdfBookmark[], path: number[]): PdfBookmark[] {
  const index = path.at(-1)
  if (index === undefined) return bookmarks
  return transformBookmarkLevel(bookmarks, path.slice(0, -1), (level) =>
    level.filter((_, bookmarkIndex) => bookmarkIndex !== index),
  )
}

function indentBookmarkAt(bookmarks: PdfBookmark[], path: number[]): PdfBookmark[] {
  const index = path.at(-1)
  if (index === undefined || index === 0) return bookmarks
  return transformBookmarkLevel(bookmarks, path.slice(0, -1), (level) => {
    const moving = level[index]
    const parent = level[index - 1]
    if (!moving || !parent) return level
    const next = level.filter((_, bookmarkIndex) => bookmarkIndex !== index)
    next[index - 1] = { ...parent, children: [...parent.children, moving] }
    return next
  })
}

function outdentBookmarkAt(bookmarks: PdfBookmark[], path: number[]): PdfBookmark[] {
  if (path.length < 2) return bookmarks
  const moving = bookmarkAt(bookmarks, path)
  const childIndex = path.at(-1)
  const parentPath = path.slice(0, -1)
  const parentIndex = parentPath.at(-1)
  if (!moving || childIndex === undefined || parentIndex === undefined) return bookmarks
  const withoutChild = transformBookmarkLevel(bookmarks, parentPath, (level) =>
    level.filter((_, bookmarkIndex) => bookmarkIndex !== childIndex),
  )
  return transformBookmarkLevel(withoutChild, parentPath.slice(0, -1), (level) => {
    const next = [...level]
    next.splice(parentIndex + 1, 0, moving)
    return next
  })
}

function validBookmarks(
  bookmarks: PdfBookmark[],
  pageCount: number,
  depth = 1,
  counter = { count: 0 },
): boolean {
  if (bookmarks.length === 0) return true
  if (depth > MAX_BOOKMARK_DEPTH) return false
  return bookmarks.every((bookmark) => {
    counter.count += 1
    return (
      counter.count <= MAX_BOOKMARK_COUNT &&
      bookmark.title.trim().length > 0 &&
      bookmark.title.trim().length <= MAX_BOOKMARK_TITLE_LENGTH &&
      Number.isInteger(bookmark.pageNumber) &&
      bookmark.pageNumber >= 1 &&
      bookmark.pageNumber <= pageCount &&
      validBookmarks(bookmark.children, pageCount, depth + 1, counter)
    )
  })
}

const MARGIN_KEYS = ['top', 'right', 'bottom', 'left'] as const
const SANITIZE_KEYS = [
  'removeJavaScript',
  'removeEmbeddedFiles',
  'removeXmpMetadata',
  'removeMetadata',
  'removeLinks',
] as const
const PIPELINE_STEP_KINDS = [
  'sanitize',
  'removeAnnotations',
  'removeSignatures',
  'flattenForms',
  'repair',
  'decompress',
] as const

const COLOR_PRESETS = [
  { id: 'whiteOnBlack', textColor: '#ffffff', backgroundColor: '#000000' },
  { id: 'blackOnWhite', textColor: '#000000', backgroundColor: '#ffffff' },
  { id: 'yellowOnBlack', textColor: '#ffff00', backgroundColor: '#000000' },
  { id: 'greenOnBlack', textColor: '#00ff00', backgroundColor: '#000000' },
] as const

const COLOR_ADJUSTMENT_FIELDS = [
  ['contrast', 'contrast'],
  ['brightness', 'brightness'],
  ['saturation', 'saturation'],
  ['red', 'redChannel'],
  ['green', 'greenChannel'],
  ['blue', 'blueChannel'],
] as const

const COLOR_ADJUSTMENT_GROUPS = [
  COLOR_ADJUSTMENT_FIELDS.slice(0, 3),
  COLOR_ADJUSTMENT_FIELDS.slice(3),
] as const

const METADATA_TEXT_FIELDS = [
  ['title', 'documentTitle'],
  ['author', 'author'],
  ['subject', 'subject'],
  ['keywords', 'keywords'],
  ['creator', 'creator'],
  ['producer', 'producer'],
] as const

const RESERVED_METADATA_KEYS = new Set([
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

type PdfColorPreset = (typeof COLOR_PRESETS)[number]['id']
type PdfMergeItem = { id: 'current'; kind: 'current' } | { id: string; kind: 'file'; file: File }
type PdfImageItem = { id: string; file: File }

function moveListItem<T>(items: T[], index: number, offset: -1 | 1): T[] {
  const destination = index + offset
  if (index < 0 || index >= items.length || destination < 0 || destination >= items.length) {
    return items
  }
  const next = [...items]
  const [item] = next.splice(index, 1)
  if (item !== undefined) next.splice(destination, 0, item)
  return next
}

function emptyPdfComment(): PdfCommentInput {
  return {
    pageIndex: 0,
    x: 24,
    y: 24,
    width: 20,
    height: 20,
    text: '',
    author: 'GenOffice AI',
    subject: 'GenOffice AI Comment',
    anchorText: '',
  }
}

function emptyPdfRedactionArea(pageIndex: number): PdfRedactionArea {
  return { pageIndex, x: 0.1, y: 0.1, width: 0.4, height: 0.1 }
}

function pdfPipelineStep(kind: (typeof PIPELINE_STEP_KINDS)[number]): PdfPipelineStep {
  return kind === 'sanitize'
    ? {
        kind,
        removeJavaScript: true,
        removeEmbeddedFiles: true,
        removeXmpMetadata: false,
        removeMetadata: false,
        removeLinks: true,
      }
    : { kind }
}

const emptyTextReplacement = (): PdfBulkTextReplacementRequest['rules'][number] => ({
  find: '',
  replace: '',
})

export function PdfToolsDialog({
  initialKind,
  filePath,
  lang,
  currentPage,
  pageCount,
  t,
  busy,
  encrypted,
  openingPassword,
  ocrProgress,
  videoProgress,
  onCancel,
  onApply,
  onBulkReplaceText,
}: {
  initialKind: PdfToolKind
  filePath: string
  lang: Lang
  currentPage: number
  pageCount: number
  t: TFunc
  busy: boolean
  encrypted: boolean
  openingPassword: string
  ocrProgress: PdfOcrProgress | null
  videoProgress: PdfVideoProgress | null
  onCancel: () => void
  onApply: (operation: PdfToolOperation) => void
  onBulkReplaceText: (request: PdfBulkTextReplacementRequest) => void
}): ReactElement {
  const text = textFor(lang)
  const emailDocumentLabels: EmailDocumentLabels = {
    from: text.emailFromLabel,
    to: text.emailToLabel,
    cc: text.emailCcLabel,
    bcc: text.emailBccLabel,
    date: text.emailDateLabel,
    attachments: text.emailAttachmentsLabel,
    emptyBody: text.emailEmptyBody,
    remoteImageRemoved: text.emailRemoteImageRemoved,
    untitled: text.emailUntitled,
  }
  const javaScriptSourceLabels: Record<PdfJavaScriptAction['source'], string> = {
    named: text.sourceNamed,
    document: text.sourceDocument,
    page: text.sourcePage,
    annotation: text.sourceAnnotation,
    form: text.sourceForm,
  }
  const signatureStatusLabels: Record<PdfSignatureAuditItem['cryptographicStatus'], string> = {
    valid: text.signatureValid,
    invalid: text.signatureInvalid,
    unsupported: text.signatureUnsupported,
    error: text.signatureError,
    notChecked: text.signatureNotChecked,
  }
  const signatureTypeLabels: Record<PdfSignatureAuditItem['type'], string> = {
    approval: text.approvalSignature,
    certification: text.certificationSignature,
    documentTimestamp: text.documentTimestamp,
    unknown: text.unknownSignature,
  }
  const [kind, setKind] = useState<PdfToolKind>(initialKind)
  const [splitMode, setSplitMode] = useState<PdfSplitMode>('afterPages')
  const [splitAfter, setSplitAfter] = useState('')
  const [mergeItems, setMergeItems] = useState<PdfMergeItem[]>([{ id: 'current', kind: 'current' }])
  const [preparingMerge, setPreparingMerge] = useState(false)
  const [imagePdfItems, setImagePdfItems] = useState<PdfImageItem[]>([])
  const [imagePdfFit, setImagePdfFit] = useState<PdfImagesToPdfFit>('maintainAspectRatio')
  const [imagePdfColor, setImagePdfColor] = useState<PdfImageColorMode>('color')
  const [imagePdfAutoRotate, setImagePdfAutoRotate] = useState(true)
  const [imagePdfAppend, setImagePdfAppend] = useState(false)
  const [preparingImagesToPdf, setPreparingImagesToPdf] = useState(false)
  const [cbzFile, setCbzFile] = useState<File | null>(null)
  const [cbzFit, setCbzFit] = useState<PdfImagesToPdfFit>('fitDocumentToImage')
  const [cbzColor, setCbzColor] = useState<PdfImageColorMode>('color')
  const [cbzAutoRotate, setCbzAutoRotate] = useState(true)
  const [preparingCbzToPdf, setPreparingCbzToPdf] = useState(false)
  const [emailFile, setEmailFile] = useState<File | null>(null)
  const [emailOutputFormat, setEmailOutputFormat] = useState<'pdf' | 'html'>('pdf')
  const [emailIncludeAttachments, setEmailIncludeAttachments] = useState(false)
  const [emailMaxAttachmentSize, setEmailMaxAttachmentSize] = useState(10)
  const [emailIncludeAllRecipients, setEmailIncludeAllRecipients] = useState(true)
  const [preparingEmailToPdf, setPreparingEmailToPdf] = useState(false)
  const [epubFile, setEpubFile] = useState<File | null>(null)
  const [epubEmbedFonts, setEpubEmbedFonts] = useState(false)
  const [epubIncludeToc, setEpubIncludeToc] = useState(false)
  const [epubIncludePageNumbers, setEpubIncludePageNumbers] = useState(false)
  const [epubOptimize, setEpubOptimize] = useState(false)
  const [preparingEpubToPdf, setPreparingEpubToPdf] = useState(false)
  const [jsonFile, setJsonFile] = useState<File | null>(null)
  const [htmlSource, setHtmlSource] = useState<'file' | 'website'>('file')
  const [htmlFile, setHtmlFile] = useState<File | null>(null)
  const [htmlWebsiteUrl, setHtmlWebsiteUrl] = useState('')
  const [htmlZoom, setHtmlZoom] = useState(100)
  const [htmlIncludePageNumbers, setHtmlIncludePageNumbers] = useState(false)
  const [preparingHtmlToPdf, setPreparingHtmlToPdf] = useState(false)
  const [markdownFile, setMarkdownFile] = useState<File | null>(null)
  const [markdownZoom, setMarkdownZoom] = useState(100)
  const [markdownIncludePageNumbers, setMarkdownIncludePageNumbers] = useState(false)
  const [preparingMarkdownToPdf, setPreparingMarkdownToPdf] = useState(false)
  const [pdfImageFormat, setPdfImageFormat] = useState<PdfPageImageFormat>('png')
  const [pdfImageOutputMode, setPdfImageOutputMode] = useState<PdfPageImageOutputMode>('multiple')
  const [pdfImageColor, setPdfImageColor] = useState<PdfPageImageColorMode>('color')
  const [pdfImageRenderDpi, setPdfImageRenderDpi] = useState(150)
  const [pdfImageQuality, setPdfImageQuality] = useState(92)
  const [pdfImageAnnotations, setPdfImageAnnotations] = useState(true)
  const [pdfVideoSeconds, setPdfVideoSeconds] = useState(3)
  const [pdfVideoResolution, setPdfVideoResolution] = useState<PdfVideoResolution>('720p')
  const [pdfVideoTransition, setPdfVideoTransition] = useState(0.4)
  const [pdfVideoAnnotations, setPdfVideoAnnotations] = useState(true)
  const [pdfCbzFormat, setPdfCbzFormat] = useState<PdfComicImageFormat>('png')
  const [pdfCbzColor, setPdfCbzColor] = useState<PdfPageImageColorMode>('color')
  const [pdfCbzRenderDpi, setPdfCbzRenderDpi] = useState(150)
  const [pdfCbzQuality, setPdfCbzQuality] = useState(90)
  const [pdfCbzAnnotations, setPdfCbzAnnotations] = useState(true)
  const [pdfHtmlRenderDpi, setPdfHtmlRenderDpi] = useState(150)
  const [pdfHtmlAnnotations, setPdfHtmlAnnotations] = useState(true)
  const [pdfEpubMode, setPdfEpubMode] = useState<PdfEpubMode>('reflowable')
  const [pdfEpubRenderDpi, setPdfEpubRenderDpi] = useState(150)
  const [pdfEpubAnnotations, setPdfEpubAnnotations] = useState(true)
  const [pdfPptxMode, setPdfPptxMode] = useState<PdfPptxMode>('editableText')
  const [pdfPptxRenderDpi, setPdfPptxRenderDpi] = useState(150)
  const [pdfPptxAnnotations, setPdfPptxAnnotations] = useState(true)
  const [pdfDocxMode, setPdfDocxMode] = useState<PdfDocxMode>('editableText')
  const [pdfDocxRenderDpi, setPdfDocxRenderDpi] = useState(150)
  const [pdfDocxAnnotations, setPdfDocxAnnotations] = useState(true)
  const [pdfJsonLightweight, setPdfJsonLightweight] = useState(true)
  const [comparisonFile, setComparisonFile] = useState<File | null>(null)
  const [comparisonRenderDpi, setComparisonRenderDpi] = useState(150)
  const [comparisonThreshold, setComparisonThreshold] = useState(0.1)
  const [preparingComparison, setPreparingComparison] = useState(false)
  const [extractPageOrder, setExtractPageOrder] = useState('')
  const [splitValue, setSplitValue] = useState(2)
  const [splitSizeValue, setSplitSizeValue] = useState(10)
  const [splitSizeUnit, setSplitSizeUnit] = useState<'KB' | 'MB' | 'GB'>('MB')
  const [chapterLevels, setChapterLevels] = useState(1)
  const [allowDuplicateChapters, setAllowDuplicateChapters] = useState(false)
  const [sectionRows, setSectionRows] = useState(1)
  const [sectionColumns, setSectionColumns] = useState(2)
  const [sectionArrangement, setSectionArrangement] = useState<PdfSectionArrangement>('rows')
  const [mergeSections, setMergeSections] = useState(true)
  const [cropMode, setCropMode] = useState<'auto' | 'manual'>('auto')
  const [cropWhiteThreshold, setCropWhiteThreshold] = useState(250)
  const [cropPadding, setCropPadding] = useState(6)
  const [margins, setMargins] = useState({ top: 0, right: 0, bottom: 0, left: 0 })
  const [pageSize, setPageSize] = useState<PdfPageSize>('A4')
  const [scaleFactor, setScaleFactor] = useState(1)
  const [nupRows, setNupRows] = useState(1)
  const [nupColumns, setNupColumns] = useState(2)
  const [nupArrangement, setNupArrangement] = useState<'rows' | 'columns'>('rows')
  const [nupReadingDirection, setNupReadingDirection] = useState<'ltr' | 'rtl'>('ltr')
  const [nupInnerMargin, setNupInnerMargin] = useState(8)
  const [nupMargins, setNupMargins] = useState({ top: 0, right: 0, bottom: 0, left: 0 })
  const [nupBorder, setNupBorder] = useState(false)
  const [nupBorderWidth, setNupBorderWidth] = useState(1)
  const [orientation, setOrientation] = useState<PdfOrientation>('portrait')
  const [border, setBorder] = useState(false)
  const [bookletSpine, setBookletSpine] = useState<'left' | 'right'>('left')
  const [bookletGutter, setBookletGutter] = useState(12)
  const [duplexPass, setDuplexPass] = useState<'both' | 'front' | 'back'>('both')
  const [flipOnShortEdge, setFlipOnShortEdge] = useState(false)
  const [posterPageSize, setPosterPageSize] = useState<Exclude<PdfPageSize, 'KEEP'>>('A4')
  const [posterRows, setPosterRows] = useState(2)
  const [posterColumns, setPosterColumns] = useState(2)
  const [posterRtl, setPosterRtl] = useState(false)
  const [singlePageDirection, setSinglePageDirection] = useState<'vertical' | 'horizontal'>(
    'vertical',
  )
  const [targetPages, setTargetPages] = useState('all')
  const [removePages, setRemovePages] = useState('')
  const [rotation, setRotation] = useState<PdfPageRotation>(90)
  const [inferUndetectedRotation, setInferUndetectedRotation] = useState(true)
  const [deskewMaxAngle, setDeskewMaxAngle] = useState(8)
  const [deskewRenderDpi, setDeskewRenderDpi] = useState(150)
  const [deskewAnnotations, setDeskewAnnotations] = useState(true)
  const [autoRenameStrategy, setAutoRenameStrategy] =
    useState<PdfAutoRenameStrategy>('largestHeading')
  const [pageNumberPosition, setPageNumberPosition] =
    useState<PdfPageNumberPosition>('bottomCenter')
  const [pageNumberMargin, setPageNumberMargin] = useState<PdfPageNumberMargin>('medium')
  const [pageNumberFont, setPageNumberFont] = useState<PdfPageNumberFont>('times')
  const [pageNumberFontSize, setPageNumberFontSize] = useState(12)
  const [pageNumberColor, setPageNumberColor] = useState('#000000')
  const [pageNumberStart, setPageNumberStart] = useState(1)
  const [pageNumberZeroPad, setPageNumberZeroPad] = useState(0)
  const [pageNumberPattern, setPageNumberPattern] = useState('{n}')
  const [scannerQuality, setScannerQuality] = useState<PdfScannerQuality>('high')
  const [scannerRotation, setScannerRotation] = useState<PdfScannerRotation>('slight')
  const [scannerColorspace, setScannerColorspace] = useState<PdfScannerColorspace>('grayscale')
  const [scannerBorder, setScannerBorder] = useState(20)
  const [scannerRotate, setScannerRotate] = useState(0)
  const [scannerRotateVariance, setScannerRotateVariance] = useState(2)
  const [scannerBrightness, setScannerBrightness] = useState(1.03)
  const [scannerContrast, setScannerContrast] = useState(1.06)
  const [scannerBlur, setScannerBlur] = useState(0.1)
  const [scannerNoise, setScannerNoise] = useState(1)
  const [scannerYellowish, setScannerYellowish] = useState(false)
  const [scannerRenderDpi, setScannerRenderDpi] = useState(150)
  const [scannerSeed, setScannerSeed] = useState(2026)
  const [scanSplitAngleThreshold, setScanSplitAngleThreshold] = useState(10)
  const [scanSplitTolerance, setScanSplitTolerance] = useState(30)
  const [scanSplitMinArea, setScanSplitMinArea] = useState(10000)
  const [scanSplitMinContourArea, setScanSplitMinContourArea] = useState(500)
  const [scanSplitBorderSize, setScanSplitBorderSize] = useState(1)
  const [scanSplitRenderDpi, setScanSplitRenderDpi] = useState(150)
  const [autoSplitDuplexMode, setAutoSplitDuplexMode] = useState(false)
  const [colorMode, setColorMode] = useState<'preset' | 'custom'>('preset')
  const [colorPreset, setColorPreset] = useState<PdfColorPreset>('whiteOnBlack')
  const [textColor, setTextColor] = useState('#ffffff')
  const [backgroundColor, setBackgroundColor] = useState('#000000')
  const [colorAdjustments, setColorAdjustments] = useState<PdfColorAdjustments>({
    contrast: 100,
    brightness: 100,
    saturation: 100,
    red: 100,
    green: 100,
    blue: 100,
  })
  const [blankThreshold, setBlankThreshold] = useState(10)
  const [blankWhitePercent, setBlankWhitePercent] = useState(99.9)
  const [includeBlankPages, setIncludeBlankPages] = useState(false)
  const [extractImageFormat, setExtractImageFormat] = useState<PdfExtractImageFormat>('png')
  const [textExportFormat, setTextExportFormat] = useState<PdfTextExportFormat>('both')
  const [pdfMarkdownPageBreaks, setPdfMarkdownPageBreaks] = useState(true)
  const [tableExportFormat, setTableExportFormat] = useState<PdfTableExportFormat>('both')
  const [includeTwoColumnTables, setIncludeTwoColumnTables] = useState(false)
  const [rearrangeMode, setRearrangeMode] = useState<PdfRearrangeMode>('reverse')
  const [customOrder, setCustomOrder] = useState('')
  const [duplicateCount, setDuplicateCount] = useState(2)
  const [redactionMode, setRedactionMode] = useState<'text' | 'areas'>('text')
  const [redactionPatterns, setRedactionPatterns] = useState('')
  const [redactionUseRegex, setRedactionUseRegex] = useState(false)
  const [redactionWholeWord, setRedactionWholeWord] = useState(false)
  const [redactionColor, setRedactionColor] = useState('#000000')
  const [redactionPadding, setRedactionPadding] = useState(2)
  const [redactionRenderDpi, setRedactionRenderDpi] = useState(144)
  const [redactionAreas, setRedactionAreas] = useState<PdfRedactionArea[]>([
    emptyPdfRedactionArea(Math.max(0, currentPage - 1)),
  ])
  const [selectedRedactionAreaIndex, setSelectedRedactionAreaIndex] = useState(0)
  const [comments, setComments] = useState<PdfCommentInput[]>([emptyPdfComment()])
  const [compressRenderDpi, setCompressRenderDpi] = useState(120)
  const [compressImageQuality, setCompressImageQuality] = useState(72)
  const [compressLineArt, setCompressLineArt] = useState(false)
  const [compressLineArtThreshold, setCompressLineArtThreshold] = useState(55)
  const [compressLineArtEdgeLevel, setCompressLineArtEdgeLevel] = useState<1 | 2 | 3>(1)
  const [pdfaArchiveMode, setPdfaArchiveMode] = useState<PdfArchiveMode>('auto')
  const [pdfaRenderDpi, setPdfaRenderDpi] = useState(150)
  const [pdfaImageQuality, setPdfaImageQuality] = useState(92)
  const [flattenMode, setFlattenMode] = useState<PdfFlattenMode>('pages')
  const [flattenRenderDpi, setFlattenRenderDpi] = useState(100)
  const [formAction, setFormAction] = useState<PdfFormAction>('fill')
  const [formFields, setFormFields] = useState<PdfFormFieldInfo[]>([])
  const [formValues, setFormValues] = useState<Record<string, PdfFormFieldValue['value']>>({})
  const [formFieldEdits, setFormFieldEdits] = useState<Record<string, PdfFormFieldEdit>>({})
  const [formCreations, setFormCreations] = useState<PdfFormFieldCreationEdit[]>([
    emptyPdfFormFieldCreation(currentPage),
  ])
  const [selectedFormFields, setSelectedFormFields] = useState<string[]>([])
  const [loadingFormFields, setLoadingFormFields] = useState(false)
  const [sanitizeOptions, setSanitizeOptions] = useState({
    removeJavaScript: true,
    removeEmbeddedFiles: true,
    removeXmpMetadata: false,
    removeMetadata: false,
    removeLinks: false,
  })
  const [pipelineSteps, setPipelineSteps] = useState<PdfPipelineStep[]>([
    pdfPipelineStep('sanitize'),
    pdfPipelineStep('removeAnnotations'),
  ])
  const [overlayFiles, setOverlayFiles] = useState<File[]>([])
  const [overlayMode, setOverlayMode] = useState<PdfOverlayMode>('sequential')
  const [overlayPosition, setOverlayPosition] = useState<PdfOverlayPosition>('foreground')
  const [overlayOpacity, setOverlayOpacity] = useState(1)
  const [overlayRepeatCounts, setOverlayRepeatCounts] = useState('')
  const [preparingOverlay, setPreparingOverlay] = useState(false)
  const [overlayImageFile, setOverlayImageFile] = useState<File | null>(null)
  const [overlayImagePosition, setOverlayImagePosition] =
    useState<PdfImageOverlayPosition>('bottomRight')
  const [overlayImageWidth, setOverlayImageWidth] = useState(20)
  const [overlayImageMargin, setOverlayImageMargin] = useState(18)
  const [overlayImageOpacity, setOverlayImageOpacity] = useState(1)
  const [overlayImageLayer, setOverlayImageLayer] = useState<PdfOverlayPosition>('foreground')
  const [overlayImageX, setOverlayImageX] = useState(0)
  const [overlayImageY, setOverlayImageY] = useState(0)
  const [preparingOverlayImage, setPreparingOverlayImage] = useState(false)
  const [filterCriterion, setFilterCriterion] = useState<PdfContentFilterCriterion>('text')
  const [filterAction, setFilterAction] = useState<PdfContentFilterAction>('keep')
  const [filterQuery, setFilterQuery] = useState('')
  const [filterCaseSensitive, setFilterCaseSensitive] = useState(false)
  const [filterWholeWord, setFilterWholeWord] = useState(false)
  const [filterPageSize, setFilterPageSize] = useState<Exclude<PdfPageSize, 'KEEP'>>('A4')
  const [filterOrientation, setFilterOrientation] = useState<PdfOrientation>('portrait')
  const [filterRotation, setFilterRotation] = useState<PdfPageFilterRotation>(0)
  const [filterComparator, setFilterComparator] = useState<PdfPageFilterComparator>('equal')
  const [documentFilterFiles, setDocumentFilterFiles] = useState<File[]>([])
  const [documentFilterCriterion, setDocumentFilterCriterion] =
    useState<PdfDocumentFilterCriterion>('pageCount')
  const [documentFilterComparator, setDocumentFilterComparator] =
    useState<PdfPageFilterComparator>('equal')
  const [documentFilterQuery, setDocumentFilterQuery] = useState('')
  const [documentFilterCaseSensitive, setDocumentFilterCaseSensitive] = useState(false)
  const [documentFilterWholeWord, setDocumentFilterWholeWord] = useState(false)
  const [documentFilterPageCount, setDocumentFilterPageCount] = useState(1)
  const [documentFilterFileSize, setDocumentFilterFileSize] = useState(10)
  const [documentFilterSizeUnit, setDocumentFilterSizeUnit] = useState<'KB' | 'MB' | 'GB'>('MB')
  const [documentFilterPageSize, setDocumentFilterPageSize] =
    useState<Exclude<PdfPageSize, 'KEEP'>>('A4')
  const [documentFilterRotation, setDocumentFilterRotation] = useState<PdfPageFilterRotation>(0)
  const [preparingDocumentFilter, setPreparingDocumentFilter] = useState(false)
  const [replacementRules, setReplacementRules] = useState([emptyTextReplacement()])
  const [replacementCaseSensitive, setReplacementCaseSensitive] = useState(true)
  const [replacementWholeWord, setReplacementWholeWord] = useState(false)
  const [attachmentAction, setAttachmentAction] = useState<PdfAttachmentAction>('add')
  const [attachmentFiles, setAttachmentFiles] = useState<File[]>([])
  const [attachmentItems, setAttachmentItems] = useState<PdfAttachmentInfo[]>([])
  const [selectedAttachment, setSelectedAttachment] = useState('')
  const [newAttachmentName, setNewAttachmentName] = useState('')
  const [loadingAttachments, setLoadingAttachments] = useState(false)
  const [preparingAttachments, setPreparingAttachments] = useState(false)
  const [bookmarks, setBookmarks] = useState<PdfBookmark[]>([])
  const [loadingBookmarks, setLoadingBookmarks] = useState(false)
  const [bookmarkImportMode, setBookmarkImportMode] = useState<'replace' | 'append'>('replace')
  const [bookmarkNotice, setBookmarkNotice] = useState('')
  const bookmarkJsonInputRef = useRef<HTMLInputElement>(null)
  const [metadataValues, setMetadataValues] = useState<PdfMetadataValues>(emptyPdfMetadata)
  const [clearAllMetadata, setClearAllMetadata] = useState(false)
  const [loadingMetadata, setLoadingMetadata] = useState(false)
  const [analysis, setAnalysis] = useState<PdfAnalysis | null>(null)
  const [loadingAnalysis, setLoadingAnalysis] = useState(false)
  const [fontReport, setFontReport] = useState<PdfFontReport | null>(null)
  const [loadingFonts, setLoadingFonts] = useState(false)
  const [annotationReport, setAnnotationReport] = useState<PdfAnnotationReport | null>(null)
  const [loadingAnnotations, setLoadingAnnotations] = useState(false)
  const [securityReport, setSecurityReport] = useState<PdfSecurityReport | null>(null)
  const [loadingSecurity, setLoadingSecurity] = useState(false)
  const [preflightReport, setPreflightReport] = useState<PdfPreflightReport | null>(null)
  const [loadingPreflight, setLoadingPreflight] = useState(false)
  const [javaScriptAudit, setJavaScriptAudit] = useState<PdfJavaScriptAudit | null>(null)
  const [loadingJavaScript, setLoadingJavaScript] = useState(false)
  const [copiedScript, setCopiedScript] = useState<number | 'all' | null>(null)
  const [signatureAudit, setSignatureAudit] = useState<PdfSignatureAudit | null>(null)
  const [loadingSignatures, setLoadingSignatures] = useState(false)
  const [passwordAlgorithm, setPasswordAlgorithm] = useState<PdfPasswordAlgorithm>('AES-256')
  const [userPassword, setUserPassword] = useState('')
  const [userPasswordConfirm, setUserPasswordConfirm] = useState('')
  const [ownerPassword, setOwnerPassword] = useState('')
  const [ownerPasswordConfirm, setOwnerPasswordConfirm] = useState('')
  const [currentPassword, setCurrentPassword] = useState(openingPassword)
  const [certificateFile, setCertificateFile] = useState<File | null>(null)
  const [certificatePassword, setCertificatePassword] = useState('')
  const [certificateSignerName, setCertificateSignerName] = useState('')
  const [certificateReason, setCertificateReason] = useState('')
  const [certificateLocation, setCertificateLocation] = useState('')
  const [certificateContact, setCertificateContact] = useState('')
  const [timestampTsaUrl, setTimestampTsaUrl] = useState<string>(PDF_TIMESTAMP_TSA_PRESETS[0].url)
  const [ocrMode, setOcrMode] = useState<PdfOcrMode>('skipText')
  const [ocrLanguages, setOcrLanguages] = useState<PdfOcrLanguage[]>(['eng', 'chi_sim'])
  const [ocrRenderDpi, setOcrRenderDpi] = useState(200)
  const [ocrClean, setOcrClean] = useState(true)
  const [ocrSidecar, setOcrSidecar] = useState(false)
  const [passwordPermissions, setPasswordPermissions] = useState<PdfPasswordPermissions>({
    allowPrinting: true,
    allowHighQualityPrint: true,
    allowModifying: true,
    allowCopying: true,
    allowAnnotating: true,
    allowFillingForms: true,
    allowExtraction: true,
    allowAssembly: true,
  })
  const [error, setError] = useState('')

  const splitPoints = useMemo(
    () => parseSplitPoints(splitAfter, pageCount),
    [pageCount, splitAfter],
  )
  const extractedPageIndexes = useMemo(
    () => parsePageIndexes(extractPageOrder, pageCount, true),
    [extractPageOrder, pageCount],
  )
  const targetPageIndexes = useMemo(
    () => parsePageIndexes(targetPages, pageCount),
    [pageCount, targetPages],
  )
  const removedPageIndexes = useMemo(
    () => parsePageIndexes(removePages, pageCount),
    [pageCount, removePages],
  )
  const customPageOrder = useMemo(
    () => parsePageIndexes(customOrder, pageCount, true),
    [customOrder, pageCount],
  )
  const formModifications = useMemo<PdfFormFieldModification[]>(
    () =>
      formFields.flatMap((field) => {
        const edit = formFieldEdits[field.name]
        if (
          !edit ||
          (edit.newName.trim() === field.name &&
            edit.label.trim() === (field.label ?? '') &&
            edit.readOnly === field.readOnly &&
            edit.required === field.required &&
            edit.optionsText === (field.options ?? []).join('\n') &&
            edit.multiselect === (field.multiselect ?? false))
        ) {
          return []
        }
        const choiceField = field.type === 'dropdown' || field.type === 'optionList'
        const options = parsedFormFieldOptions(edit.optionsText)
        return [
          {
            name: field.name,
            newName: edit.newName.trim(),
            label: edit.label.trim(),
            readOnly: edit.readOnly,
            required: edit.required,
            ...(choiceField && edit.optionsText !== (field.options ?? []).join('\n')
              ? { options }
              : {}),
            ...(choiceField && edit.multiselect !== (field.multiselect ?? false)
              ? { multiselect: edit.multiselect }
              : {}),
          },
        ]
      }),
    [formFieldEdits, formFields],
  )
  const preparedFormCreations = useMemo<PdfFormFieldCreation[]>(
    () =>
      formCreations.map((creation) => {
        const options = parsedFormFieldOptions(creation.optionsText)
        const defaultValues = parsedFormFieldOptions(creation.defaultText)
        return {
          name: creation.name.trim(),
          label: creation.label.trim(),
          type: creation.type,
          pageIndex: creation.pageNumber - 1,
          x: creation.x,
          y: creation.y,
          width: creation.width,
          height: creation.height,
          required: creation.required,
          readOnly: creation.readOnly,
          ...(creation.type === 'text'
            ? { multiline: creation.multiline, defaultValue: creation.defaultText }
            : creation.type === 'checkbox'
              ? { defaultValue: creation.defaultChecked }
              : {
                  options,
                  defaultValue: creation.multiselect ? defaultValues : (defaultValues[0] ?? ''),
                  ...(creation.type === 'dropdown' || creation.type === 'optionList'
                    ? { multiselect: creation.multiselect }
                    : {}),
                  ...(creation.type === 'radio' ? { optionSpacing: creation.optionSpacing } : {}),
                }),
        }
      }),
    [formCreations],
  )

  useEffect(() => {
    if (kind !== 'attachments' || !filePath) return
    let active = true
    setLoadingAttachments(true)
    void window.pdfApi
      .readFile(filePath)
      .then((data) => listPdfAttachmentsBytes(data))
      .then((items) => {
        if (!active) return
        setAttachmentItems(items)
        setSelectedAttachment((current) =>
          items.some((item) => item.name === current) ? current : (items[0]?.name ?? ''),
        )
      })
      .catch(() => {
        if (!active) return
        setAttachmentItems([])
        setSelectedAttachment('')
        setError(text.attachmentLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingAttachments(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.attachmentLoadFailed])

  useEffect(() => {
    if (kind !== 'bookmarks' || !filePath) return
    let active = true
    setLoadingBookmarks(true)
    void window.pdfApi
      .readFile(filePath)
      .then((data) => listPdfBookmarksBytes(data))
      .then((items) => {
        if (active) setBookmarks(items)
      })
      .catch(() => {
        if (!active) return
        setBookmarks([])
        setError(text.bookmarkLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingBookmarks(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.bookmarkLoadFailed])

  const importBookmarkText = (value: string) => {
    try {
      const imported = parseBookmarkJson(value, pageCount)
      const next = bookmarkImportMode === 'append' ? [...bookmarks, ...imported] : imported
      if (bookmarkTreeCount(next) > MAX_BOOKMARK_COUNT) throw new Error('invalid-bookmarks')
      setBookmarks(next)
      setError('')
      setBookmarkNotice(text.bookmarkJsonImported)
    } catch {
      setBookmarkNotice('')
      setError(text.invalidBookmarkJson)
    }
  }

  const importBookmarkFile = async (file: File | undefined) => {
    if (!file) return
    if (file.size > MAX_BOOKMARK_JSON_BYTES) {
      setBookmarkNotice('')
      setError(text.invalidBookmarkJson)
      return
    }
    try {
      importBookmarkText(await file.text())
    } catch {
      setBookmarkNotice('')
      setError(text.invalidBookmarkJson)
    }
  }

  const pasteBookmarkJson = async () => {
    try {
      if (!navigator.clipboard?.readText) throw new Error('clipboard-unavailable')
      importBookmarkText(await navigator.clipboard.readText())
    } catch {
      setBookmarkNotice('')
      setError(text.invalidBookmarkJson)
    }
  }

  const copyBookmarkJson = async () => {
    if (await copyPdfToolText(bookmarkJsonText(bookmarks))) {
      setError('')
      setBookmarkNotice(text.bookmarkJsonCopied)
      return
    }
    setBookmarkNotice('')
    setError(text.invalidBookmarkJson)
  }

  useEffect(() => {
    if (kind !== 'metadata' || !filePath) return
    let active = true
    setLoadingMetadata(true)
    setClearAllMetadata(false)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => analyzePdfBytes(data))
      .then((result) => {
        if (!active) return
        setMetadataValues({
          title: result.properties.title ?? '',
          author: result.properties.author ?? '',
          subject: result.properties.subject ?? '',
          keywords: result.properties.keywords ?? '',
          creator: result.properties.creator ?? '',
          producer: result.properties.producer ?? '',
          creationDate: metadataInputDate(result.properties.creationDate),
          modificationDate: metadataInputDate(result.properties.modificationDate),
          trapped: result.properties.trapped ?? '',
          custom: result.properties.custom,
        })
      })
      .catch(() => {
        if (!active) return
        setMetadataValues(emptyPdfMetadata())
        setError(text.metadataLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingMetadata(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.metadataLoadFailed])

  useEffect(() => {
    if (kind !== 'info' || !filePath) return
    let active = true
    setLoadingAnalysis(true)
    setAnalysis(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => analyzePdfBytes(data))
      .then((result) => {
        if (active) setAnalysis(result)
      })
      .catch(() => {
        if (!active) return
        setAnalysis(null)
        setError(text.analysisLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingAnalysis(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.analysisLoadFailed])

  useEffect(() => {
    if (kind !== 'preflight' || !filePath) return
    let active = true
    setLoadingPreflight(true)
    setPreflightReport(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => preflightPdfBytes(data))
      .then((result) => {
        if (active) setPreflightReport(result)
      })
      .catch(() => {
        if (!active) return
        setPreflightReport(null)
        setError(text.preflightLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingPreflight(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.preflightLoadFailed])

  useEffect(() => {
    if (kind !== 'fontInfo' || !filePath) return
    let active = true
    setLoadingFonts(true)
    setFontReport(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => analyzePdfFontsBytes(data))
      .then((result) => {
        if (active) setFontReport(result)
      })
      .catch(() => {
        if (!active) return
        setFontReport(null)
        setError(text.fontInfoLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingFonts(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.fontInfoLoadFailed])

  useEffect(() => {
    if (kind !== 'annotationInfo' || !filePath) return
    let active = true
    setLoadingAnnotations(true)
    setAnnotationReport(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => analyzePdfAnnotationsBytes(data))
      .then((result) => {
        if (active) setAnnotationReport(result)
      })
      .catch(() => {
        if (!active) return
        setAnnotationReport(null)
        setError(text.annotationInfoLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingAnnotations(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.annotationInfoLoadFailed])

  useEffect(() => {
    if (kind !== 'securityInfo' || !filePath) return
    let active = true
    setLoadingSecurity(true)
    setSecurityReport(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => analyzePdfSecurityBytes(data))
      .then((result) => {
        if (active) setSecurityReport(result)
      })
      .catch(() => {
        if (!active) return
        setSecurityReport(null)
        setError(text.securityInfoLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingSecurity(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.securityInfoLoadFailed])

  useEffect(() => {
    if (kind !== 'forms' || !filePath) return
    let active = true
    setLoadingFormFields(true)
    setFormFields([])
    setFormValues({})
    setFormFieldEdits({})
    setSelectedFormFields([])
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => listPdfFormFieldsBytes(data))
      .then((fields) => {
        if (!active) return
        setFormFields(fields)
        setFormValues(
          Object.fromEntries(fields.map((field) => [field.name, initialFormFieldValue(field)])),
        )
        setFormFieldEdits(
          Object.fromEntries(
            fields.map((field) => [
              field.name,
              {
                newName: field.name,
                label: field.label ?? '',
                readOnly: field.readOnly,
                required: field.required,
                optionsText: (field.options ?? []).join('\n'),
                multiselect: field.multiselect ?? false,
              },
            ]),
          ),
        )
      })
      .catch(() => {
        if (!active) return
        setFormFields([])
        setFormValues({})
        setFormFieldEdits({})
        setSelectedFormFields([])
        setError(text.formFieldsLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingFormFields(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.formFieldsLoadFailed])

  useEffect(() => {
    if (kind !== 'javascriptAudit' || !filePath) return
    let active = true
    setLoadingJavaScript(true)
    setJavaScriptAudit(null)
    setCopiedScript(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => auditPdfJavaScriptBytes(data))
      .then((result) => {
        if (active) setJavaScriptAudit(result)
      })
      .catch(() => {
        if (!active) return
        setJavaScriptAudit(null)
        setError(text.javaScriptLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingJavaScript(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.javaScriptLoadFailed])

  useEffect(() => {
    if (kind !== 'signatureAudit' || !filePath) return
    let active = true
    setLoadingSignatures(true)
    setSignatureAudit(null)
    setError('')
    void window.pdfApi
      .readFile(filePath)
      .then((data) => auditPdfSignaturesBytes(data))
      .then((result) => {
        if (active) setSignatureAudit(result)
      })
      .catch(() => {
        if (!active) return
        setSignatureAudit(null)
        setError(text.signatureLoadFailed)
      })
      .finally(() => {
        if (active) setLoadingSignatures(false)
      })
    return () => {
      active = false
    }
  }, [filePath, kind, text.signatureLoadFailed])

  const apply = async () => {
    setError('')
    if (kind === 'jsonToPdf') {
      if (!jsonFile) return setError(text.selectJson)
      try {
        onApply({
          kind,
          jsonBytes: new Uint8Array(await jsonFile.arrayBuffer()),
          baseName: jsonFile.name,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      }
      return
    }
    if (kind === 'merge') {
      const additionalFiles = mergeItems.filter(
        (item): item is Extract<PdfMergeItem, { kind: 'file' }> => item.kind === 'file',
      )
      if (additionalFiles.length === 0) return setError(text.selectMergeFiles)
      const currentDocumentIndex = mergeItems.findIndex((item) => item.kind === 'current')
      setPreparingMerge(true)
      try {
        const documents = await Promise.all(
          additionalFiles.map(async (item) => new Uint8Array(await item.file.arrayBuffer())),
        )
        onApply({ kind, documents, currentDocumentIndex })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingMerge(false)
      }
      return
    }
    if (kind === 'imagesToPdf') {
      if (imagePdfItems.length === 0) return setError(text.selectImagesToPdf)
      setPreparingImagesToPdf(true)
      try {
        const images = await prepareImagesForPdf(
          imagePdfItems.map((item) => item.file),
          imagePdfColor,
        )
        onApply({
          kind,
          images,
          fitOption: imagePdfFit,
          autoRotate: imagePdfAutoRotate,
          appendToCurrent: imagePdfAppend,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingImagesToPdf(false)
      }
      return
    }
    if (kind === 'cbzToPdf') {
      if (!cbzFile) return setError(text.selectCbz)
      setPreparingCbzToPdf(true)
      try {
        onApply({
          kind,
          images: await prepareComicArchiveForPdf(cbzFile, cbzColor),
          fitOption: cbzFit,
          autoRotate: cbzAutoRotate,
          baseName: cbzFile.name,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingCbzToPdf(false)
      }
      return
    }
    if (kind === 'emailToPdf') {
      if (!emailFile) return setError(text.selectEmail)
      if (
        !Number.isInteger(emailMaxAttachmentSize) ||
        emailMaxAttachmentSize < 1 ||
        emailMaxAttachmentSize > 50
      ) {
        return setError(text.invalidEmailAttachmentSize)
      }
      setPreparingEmailToPdf(true)
      try {
        const options = {
          includeAttachments: emailIncludeAttachments,
          maxAttachmentSizeMb: emailMaxAttachmentSize,
          includeAllRecipients: emailIncludeAllRecipients,
          labels: emailDocumentLabels,
          locale: lang === 'zh-TW' ? 'zh-TW' : lang === 'zh' ? 'zh-CN' : 'en-US',
        }
        if (emailOutputFormat === 'html') {
          const prepared = await prepareEmailMarkup(
            new Uint8Array(await emailFile.arrayBuffer()),
            options,
          )
          onApply({
            kind,
            outputFormat: 'html',
            htmlBytes: new TextEncoder().encode(prepared.html),
            baseName: emailFile.name,
          })
        } else {
          const prepared = await prepareEmailDocumentForPdf(emailFile, options)
          onApply({
            kind,
            outputFormat: 'pdf',
            pages: prepared.pages,
            attachments: prepared.attachments,
            baseName: emailFile.name,
          })
        }
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingEmailToPdf(false)
      }
      return
    }
    if (kind === 'epubToPdf') {
      if (!epubFile) return setError(text.selectEpub)
      setPreparingEpubToPdf(true)
      try {
        const prepared = await prepareEpubDocumentForPdf(epubFile, {
          embedAllFonts: epubEmbedFonts,
          includeTableOfContents: epubIncludeToc,
          includePageNumbers: epubIncludePageNumbers,
          optimizeForEbook: epubOptimize,
          tableOfContentsLabel: text.epubTocLabel,
          untitledLabel: text.epubUntitled,
        })
        onApply({
          kind,
          pages: prepared.pages,
          baseName: epubFile.name,
          title: prepared.title,
          author: prepared.author,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingEpubToPdf(false)
      }
      return
    }
    if (kind === 'htmlToPdf') {
      if (htmlSource === 'file' && !htmlFile) return setError(text.selectHtml)
      if (htmlSource === 'website' && !htmlWebsiteUrl.trim()) {
        return setError(text.selectWebsiteUrl)
      }
      if (!Number.isInteger(htmlZoom) || htmlZoom < 50 || htmlZoom > 200) {
        return setError(text.invalidHtmlZoom)
      }
      setPreparingHtmlToPdf(true)
      try {
        const options = {
          includePageNumbers: htmlIncludePageNumbers,
          zoom: htmlZoom / 100,
          untitledLabel: text.htmlUntitled,
        }
        const prepared =
          htmlSource === 'file'
            ? await prepareLocalWebDocumentForPdf(htmlFile!, options)
            : await prepareRemoteWebDocumentForPdf(
                htmlWebsiteUrl,
                (request) => window.pdfApi.fetchWebResource(request),
                options,
              )
        onApply({
          kind,
          pages: prepared.pages,
          baseName: htmlSource === 'file' ? htmlFile!.name : `${prepared.title}.html`,
          title: prepared.title,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingHtmlToPdf(false)
      }
      return
    }
    if (kind === 'markdownToPdf') {
      if (!markdownFile) return setError(text.selectMarkdown)
      if (!Number.isInteger(markdownZoom) || markdownZoom < 50 || markdownZoom > 200) {
        return setError(text.invalidMarkdownZoom)
      }
      setPreparingMarkdownToPdf(true)
      try {
        const prepared = await prepareMarkdownDocumentForPdf(markdownFile, {
          includePageNumbers: markdownIncludePageNumbers,
          zoom: markdownZoom / 100,
          untitledLabel: text.markdownUntitled,
        })
        onApply({
          kind,
          pages: prepared.pages,
          baseName: markdownFile.name,
          title: prepared.title,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingMarkdownToPdf(false)
      }
      return
    }
    if (kind === 'compare') {
      if (!comparisonFile) return setError(text.selectComparisonFile)
      if (
        !Number.isInteger(comparisonRenderDpi) ||
        comparisonRenderDpi < 72 ||
        comparisonRenderDpi > 300 ||
        !Number.isFinite(comparisonThreshold) ||
        comparisonThreshold < 0 ||
        comparisonThreshold > 1
      ) {
        return setError(text.invalidComparison)
      }
      setPreparingComparison(true)
      try {
        onApply({
          kind,
          comparisonDocument: new Uint8Array(await comparisonFile.arrayBuffer()),
          renderDpi: comparisonRenderDpi,
          threshold: comparisonThreshold,
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingComparison(false)
      }
      return
    }
    if (kind === 'split') {
      if (splitMode === 'chapters') {
        if (!Number.isInteger(chapterLevels) || chapterLevels < 1 || chapterLevels > 21) {
          return setError(text.invalidChapterLevels)
        }
        onApply({
          kind,
          mode: splitMode,
          bookmarkLevel: chapterLevels - 1,
          allowDuplicates: allowDuplicateChapters,
        })
        return
      }
      if (splitMode === 'afterPages') {
        if (!splitPoints || splitPoints.length === 0) return setError(text.invalidSplit)
        onApply({ kind, mode: splitMode, splitAfterPages: splitPoints })
        return
      }
      if (splitMode === 'fileSize') {
        const multiplier =
          splitSizeUnit === 'KB' ? 1024 : splitSizeUnit === 'MB' ? 1024 * 1024 : 1024 * 1024 * 1024
        const maxBytes = Math.round(splitSizeValue * multiplier)
        if (
          !Number.isFinite(splitSizeValue) ||
          splitSizeValue <= 0 ||
          !Number.isSafeInteger(maxBytes)
        ) {
          return setError(text.invalidFileSize)
        }
        onApply({ kind, mode: splitMode, maxBytes })
        return
      }
      if (!Number.isInteger(splitValue) || splitValue < 1) return setError(text.invalidSplit)
      onApply({ kind, mode: splitMode, value: splitValue })
      return
    }
    if (kind === 'extractPages') {
      if (!extractedPageIndexes || extractedPageIndexes.length === 0) {
        return setError(text.invalidExtractPages)
      }
      onApply({ kind, pageIndexes: extractedPageIndexes })
      return
    }
    if (kind === 'splitSections') {
      if (
        !Number.isInteger(sectionRows) ||
        !Number.isInteger(sectionColumns) ||
        sectionRows < 1 ||
        sectionColumns < 1 ||
        sectionRows > 10 ||
        sectionColumns > 10
      ) {
        return setError(text.invalidGrid)
      }
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        rows: sectionRows,
        columns: sectionColumns,
        arrangement: sectionArrangement,
        merge: mergeSections,
      })
      return
    }
    if (kind === 'crop') {
      if (cropMode === 'auto') {
        if (
          !Number.isInteger(cropWhiteThreshold) ||
          cropWhiteThreshold < 0 ||
          cropWhiteThreshold > 255 ||
          !Number.isFinite(cropPadding) ||
          cropPadding < 0 ||
          cropPadding > 144
        ) {
          return setError(text.invalidAutoCrop)
        }
        onApply({
          kind,
          mode: cropMode,
          whiteThreshold: cropWhiteThreshold,
          padding: cropPadding,
        })
        return
      }
      if (Object.values(margins).some((value) => !Number.isFinite(value) || value < 0)) {
        return setError(text.invalidMargins)
      }
      onApply({ kind, mode: cropMode, margins })
      return
    }
    if (kind === 'scale') {
      onApply({ kind, pageSize, orientation, scaleFactor })
      return
    }
    if (kind === 'nup') {
      if (
        !Number.isInteger(nupRows) ||
        !Number.isInteger(nupColumns) ||
        nupRows < 1 ||
        nupColumns < 1 ||
        nupRows > 30 ||
        nupColumns > 30 ||
        Object.values(nupMargins).some(
          (value) => !Number.isFinite(value) || value < 0 || value > 1_000,
        ) ||
        !Number.isFinite(nupInnerMargin) ||
        nupInnerMargin < 0 ||
        nupInnerMargin > 1_000 ||
        (nupBorder &&
          (!Number.isFinite(nupBorderWidth) || nupBorderWidth <= 0 || nupBorderWidth > 72))
      ) {
        return setError(text.invalidNup)
      }
      onApply({
        kind,
        rows: nupRows,
        columns: nupColumns,
        orientation,
        arrangement: nupArrangement,
        readingDirection: nupReadingDirection,
        innerMargin: nupInnerMargin,
        ...nupMargins,
        borderWidth: nupBorder ? nupBorderWidth : 0,
      })
      return
    }
    if (kind === 'booklet') {
      if (!Number.isFinite(bookletGutter) || bookletGutter < 0) {
        return setError(text.invalidGutter)
      }
      onApply({
        kind,
        spine: bookletSpine,
        gutter: bookletGutter,
        border,
        duplexPass,
        flipOnShortEdge,
      })
      return
    }
    if (kind === 'poster') {
      if (
        !Number.isInteger(posterRows) ||
        !Number.isInteger(posterColumns) ||
        posterRows < 1 ||
        posterRows > 10 ||
        posterColumns < 1 ||
        posterColumns > 10
      ) {
        return setError(text.invalidGrid)
      }
      onApply({
        kind,
        pageSize: posterPageSize,
        rows: posterRows,
        columns: posterColumns,
        readingDirection: posterRtl ? 'rtl' : 'ltr',
      })
      return
    }
    if (kind === 'singlePage') {
      onApply({ kind, direction: singlePageDirection })
      return
    }
    if (kind === 'rotatePages') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({ kind, pageIndexes: targetPageIndexes, angle: rotation })
      return
    }
    if (kind === 'autoRotate') {
      onApply({ kind, inferUndetected: inferUndetectedRotation })
      return
    }
    if (kind === 'deskew') {
      if (
        !targetPageIndexes ||
        !Number.isFinite(deskewMaxAngle) ||
        deskewMaxAngle < 0.5 ||
        deskewMaxAngle > 15 ||
        !Number.isInteger(deskewRenderDpi) ||
        deskewRenderDpi < 72 ||
        deskewRenderDpi > 300
      ) {
        return setError(text.invalidDeskew)
      }
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        maxAngle: deskewMaxAngle,
        renderDpi: deskewRenderDpi,
        includeAnnotations: deskewAnnotations,
      })
      return
    }
    if (kind === 'scannerEffect') {
      const valuesInRange =
        Number.isInteger(scannerBorder) &&
        scannerBorder >= 0 &&
        scannerBorder <= 200 &&
        Number.isFinite(scannerRotate) &&
        scannerRotate >= -15 &&
        scannerRotate <= 15 &&
        Number.isFinite(scannerRotateVariance) &&
        scannerRotateVariance >= 0 &&
        scannerRotateVariance <= 10 &&
        Number.isFinite(scannerBrightness) &&
        scannerBrightness >= 0.5 &&
        scannerBrightness <= 2 &&
        Number.isFinite(scannerContrast) &&
        scannerContrast >= 0.5 &&
        scannerContrast <= 2 &&
        Number.isFinite(scannerBlur) &&
        scannerBlur >= 0 &&
        scannerBlur <= 10 &&
        Number.isFinite(scannerNoise) &&
        scannerNoise >= 0 &&
        scannerNoise <= 50 &&
        Number.isInteger(scannerRenderDpi) &&
        scannerRenderDpi >= 72 &&
        scannerRenderDpi <= 500 &&
        Number.isInteger(scannerSeed)
      if (!valuesInRange) return setError(text.invalidScannerEffect)
      onApply({
        kind,
        quality: scannerQuality,
        rotation: scannerRotation,
        colorspace: scannerColorspace,
        border: scannerBorder,
        rotate: scannerRotate,
        rotateVariance: scannerRotateVariance,
        brightness: scannerBrightness,
        contrast: scannerContrast,
        blur: scannerBlur,
        noise: scannerNoise,
        yellowish: scannerYellowish,
        renderDpi: scannerRenderDpi,
        seed: scannerSeed,
      })
      return
    }
    if (kind === 'scannerImageSplit') {
      const settings = [
        [scanSplitAngleThreshold, 0, 45],
        [scanSplitTolerance, 0, 255],
        [scanSplitMinArea, 1, 100_000_000],
        [scanSplitMinContourArea, 1, 100_000_000],
        [scanSplitBorderSize, 0, 200],
        [scanSplitRenderDpi, 72, 300],
      ] as const
      if (
        settings.some(
          ([value, minimum, maximum]) =>
            !Number.isInteger(value) || value < minimum || value > maximum,
        )
      ) {
        return setError(text.invalidScannerImageSplit)
      }
      onApply({
        kind,
        angleThreshold: scanSplitAngleThreshold,
        tolerance: scanSplitTolerance,
        minArea: scanSplitMinArea,
        minContourArea: scanSplitMinContourArea,
        borderSize: scanSplitBorderSize,
        renderDpi: scanSplitRenderDpi,
      })
      return
    }
    if (kind === 'autoSplit') {
      onApply({ kind, action: 'split', duplexMode: autoSplitDuplexMode })
      return
    }
    if (kind === 'autoRename') {
      onApply({ kind, strategy: autoRenameStrategy })
      return
    }
    if (kind === 'pageNumbers') {
      if (
        !targetPageIndexes ||
        targetPageIndexes.length === 0 ||
        !Number.isFinite(pageNumberFontSize) ||
        pageNumberFontSize < 1 ||
        pageNumberFontSize > 200 ||
        !Number.isSafeInteger(pageNumberStart) ||
        pageNumberStart < 1 ||
        !Number.isInteger(pageNumberZeroPad) ||
        pageNumberZeroPad < 0 ||
        pageNumberZeroPad > 12 ||
        !/^#[\da-f]{6}$/i.test(pageNumberColor) ||
        !pageNumberPattern.trim()
      ) {
        return setError(text.invalidPageNumbers)
      }
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        position: pageNumberPosition,
        margin: pageNumberMargin,
        fontSize: pageNumberFontSize,
        font: pageNumberFont,
        fontColor: pageNumberColor,
        startingNumber: pageNumberStart,
        zeroPad: pageNumberZeroPad,
        textPattern: pageNumberPattern,
      })
      return
    }
    if (kind === 'replaceColors') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      const preset = COLOR_PRESETS.find((item) => item.id === colorPreset) ?? COLOR_PRESETS[0]
      const selectedTextColor = colorMode === 'preset' ? preset.textColor : textColor
      const selectedBackgroundColor =
        colorMode === 'preset' ? preset.backgroundColor : backgroundColor
      if (
        !/^#[\da-f]{6}$/i.test(selectedTextColor) ||
        !/^#[\da-f]{6}$/i.test(selectedBackgroundColor)
      ) {
        return setError(text.invalidColors)
      }
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        textColor: selectedTextColor,
        backgroundColor: selectedBackgroundColor,
      })
      return
    }
    if (kind === 'adjustColors') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({ kind, pageIndexes: targetPageIndexes, ...colorAdjustments })
      return
    }
    if (kind === 'removeImages' || kind === 'invertColors') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({ kind, pageIndexes: targetPageIndexes })
      return
    }
    if (kind === 'extractText') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({ kind, format: textExportFormat, pageIndexes: targetPageIndexes })
      return
    }
    if (kind === 'pdfToMarkdown') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        includePageBreaks: pdfMarkdownPageBreaks,
      })
      return
    }
    if (kind === 'pdfToJson' || kind === 'pdfToXml') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({ kind, pageIndexes: targetPageIndexes, lightweight: pdfJsonLightweight })
      return
    }
    if (kind === 'extractTables') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({
        kind,
        format: tableExportFormat,
        pageIndexes: targetPageIndexes,
        includeTwoColumnTextTables: includeTwoColumnTables,
      })
      return
    }
    if (kind === 'pdfToXlsx') {
      if (!targetPageIndexes || targetPageIndexes.length === 0) return setError(text.invalidPages)
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        includeTwoColumnTextTables: includeTwoColumnTables,
      })
      return
    }
    if (kind === 'removePages') {
      if (
        !removedPageIndexes ||
        removedPageIndexes.length === 0 ||
        removedPageIndexes.length >= pageCount
      ) {
        return setError(text.invalidRemoval)
      }
      onApply({ kind, pageIndexes: removedPageIndexes })
      return
    }
    if (kind === 'forms') {
      if (formAction === 'fill') {
        const requiredMissing = formFields.some((field) => {
          if (!field.required || !editableFormField(field)) return false
          const value = formValues[field.name]
          return value === false || value === '' || (Array.isArray(value) && value.length === 0)
        })
        if (requiredMissing) return setError(text.invalidFormFields)
        onApply({
          kind,
          action: formAction,
          fields: formFields
            .filter(editableFormField)
            .map((field) => ({ name: field.name, value: formValues[field.name] ?? '' })),
        })
      } else if (formAction === 'export') {
        if (formFields.length === 0) return setError(text.noFormFields)
        onApply({ kind, action: formAction })
      } else if (formAction === 'create') {
        const names = formCreations.map((creation) => creation.name.trim())
        const allNames = [...formFields.map((field) => field.name), ...names]
        const validNames = names.every((name) => {
          const parts = name.split('.')
          return name.length > 0 && name.length <= 500 && parts.every((part) => part.trim())
        })
        const uniqueNames = allNames.every(
          (name, index) =>
            !allNames.some(
              (otherName, otherIndex) =>
                otherIndex !== index &&
                (otherName === name ||
                  otherName.startsWith(`${name}.`) ||
                  name.startsWith(`${otherName}.`)),
            ),
        )
        const validCreations = formCreations.every((creation) => {
          const options = parsedFormFieldOptions(creation.optionsText)
          const defaults = parsedFormFieldOptions(creation.defaultText)
          const choiceField = ['radio', 'dropdown', 'optionList'].includes(creation.type)
          return (
            Number.isInteger(creation.pageNumber) &&
            creation.pageNumber >= 1 &&
            creation.pageNumber <= pageCount &&
            [creation.x, creation.y, creation.width, creation.height].every(Number.isFinite) &&
            creation.x >= 0 &&
            creation.y >= 0 &&
            creation.width > 0 &&
            creation.height > 0 &&
            (!choiceField ||
              (options.length > 0 &&
                options.length <= 100 &&
                defaults.every((value) => options.includes(value)) &&
                ((creation.type !== 'radio' && creation.multiselect) || defaults.length <= 1))) &&
            (creation.type !== 'radio' ||
              (Number.isFinite(creation.optionSpacing) &&
                creation.optionSpacing >= creation.height &&
                creation.y - creation.optionSpacing * (options.length - 1) >= 0))
          )
        })
        if (!validNames || !uniqueNames || !validCreations) {
          return setError(text.invalidCreatedFields)
        }
        onApply({ kind, action: formAction, creations: preparedFormCreations })
      } else if (formAction === 'modify') {
        if (formModifications.length === 0) return setError(text.modifyAtLeastOneField)
        const finalNames = formFields.map((field) => {
          const edit = formFieldEdits[field.name]
          return resolvedFormFieldName(field.name, edit?.newName ?? field.name)
        })
        if (
          finalNames.some((name) => name === null) ||
          new Set(finalNames).size !== finalNames.length
        ) {
          return setError(text.invalidModifiedFieldNames)
        }
        if (
          formFields.some(
            (field) =>
              (field.type === 'dropdown' || field.type === 'optionList') &&
              parsedFormFieldOptions(formFieldEdits[field.name]?.optionsText ?? '').length === 0,
          )
        ) {
          return setError(text.invalidChoiceOptions)
        }
        onApply({ kind, action: formAction, modifications: formModifications })
      } else if (formAction === 'delete') {
        if (selectedFormFields.length === 0) return setError(text.selectFieldsToDelete)
        onApply({ kind, action: formAction, fieldNames: selectedFormFields })
      } else {
        onApply({ kind, action: formAction })
      }
      return
    }
    if (kind === 'redact') {
      const patterns = redactionPatterns
        .split(/\r?\n/)
        .map((pattern) => pattern.trim())
        .filter(Boolean)
      if (!/^#[\da-f]{6}$/i.test(redactionColor)) {
        return setError(text.invalidRedactionPatterns)
      }
      if (redactionMode === 'text' && patterns.length === 0) {
        return setError(text.invalidRedactionPatterns)
      }
      if (redactionMode === 'text' && redactionUseRegex) {
        try {
          for (const pattern of patterns) new RegExp(pattern, 'giu')
        } catch {
          return setError(text.invalidRedactionPatterns)
        }
      }
      if (
        redactionMode === 'text' &&
        (!Number.isFinite(redactionPadding) || redactionPadding < 0 || redactionPadding > 72)
      ) {
        return setError(text.invalidRedactionPadding)
      }
      if (
        redactionMode === 'areas' &&
        (redactionAreas.length === 0 ||
          redactionAreas.some(
            (area) =>
              !Number.isInteger(area.pageIndex) ||
              area.pageIndex < 0 ||
              area.pageIndex >= pageCount ||
              !Number.isFinite(area.x) ||
              !Number.isFinite(area.y) ||
              !Number.isFinite(area.width) ||
              !Number.isFinite(area.height) ||
              area.x < 0 ||
              area.y < 0 ||
              area.width <= 0 ||
              area.height <= 0 ||
              area.x + area.width > 1 ||
              area.y + area.height > 1,
          ))
      ) {
        return setError(text.invalidRedactionAreas)
      }
      if (
        !Number.isInteger(redactionRenderDpi) ||
        redactionRenderDpi < 72 ||
        redactionRenderDpi > 600
      ) {
        return setError(text.invalidRenderDpi)
      }
      onApply({
        kind,
        mode: redactionMode,
        patterns,
        useRegex: redactionUseRegex,
        wholeWord: redactionWholeWord,
        color: redactionColor,
        padding: redactionPadding,
        renderDpi: redactionRenderDpi,
        areas: redactionMode === 'areas' ? redactionAreas : undefined,
      })
      return
    }
    if (kind === 'comments') {
      if (
        comments.length === 0 ||
        comments.some(
          (comment) =>
            !Number.isInteger(comment.pageIndex) ||
            comment.pageIndex < 0 ||
            comment.pageIndex >= pageCount ||
            !Number.isFinite(comment.x) ||
            !Number.isFinite(comment.y) ||
            comment.text.trim().length === 0 ||
            comment.text.length > 100_000,
        )
      ) {
        return setError(text.invalidComments)
      }
      onApply({
        kind,
        comments: comments.map((comment) => ({
          ...comment,
          text: comment.text.trim(),
          author: comment.author?.trim() || undefined,
          subject: comment.subject?.trim() || undefined,
          anchorText: comment.anchorText?.trim() || undefined,
        })),
      })
      return
    }
    if (kind === 'compress') {
      if (
        !Number.isInteger(compressRenderDpi) ||
        compressRenderDpi < 72 ||
        compressRenderDpi > 600
      ) {
        return setError(text.invalidRenderDpi)
      }
      if (
        !Number.isInteger(compressImageQuality) ||
        compressImageQuality < 10 ||
        compressImageQuality > 100
      ) {
        return setError(text.invalidImageQuality)
      }
      if (
        compressLineArt &&
        (!Number.isFinite(compressLineArtThreshold) ||
          compressLineArtThreshold < 0 ||
          compressLineArtThreshold > 100 ||
          !Number.isInteger(compressLineArtEdgeLevel) ||
          compressLineArtEdgeLevel < 1 ||
          compressLineArtEdgeLevel > 3)
      ) {
        return setError(text.invalidLineArt)
      }
      onApply({
        kind,
        renderDpi: compressRenderDpi,
        imageQuality: compressImageQuality,
        lineArt: compressLineArt,
        lineArtThreshold: compressLineArtThreshold,
        lineArtEdgeLevel: compressLineArtEdgeLevel,
      })
      return
    }
    if (kind === 'pdfToPdfa') {
      if (
        !Number.isInteger(pdfaRenderDpi) ||
        pdfaRenderDpi < 72 ||
        pdfaRenderDpi > 600 ||
        !Number.isInteger(pdfaImageQuality) ||
        pdfaImageQuality < 10 ||
        pdfaImageQuality > 100
      ) {
        return setError(text.invalidPdfa)
      }
      onApply({
        kind,
        format: 'PDF/A-2b',
        archiveMode: pdfaArchiveMode,
        renderDpi: pdfaRenderDpi,
        imageQuality: pdfaImageQuality,
      })
      return
    }
    if (kind === 'flatten') {
      if (flattenMode === 'forms') {
        onApply({ kind, mode: 'forms' })
        return
      }
      if (!Number.isInteger(flattenRenderDpi) || flattenRenderDpi < 72 || flattenRenderDpi > 600) {
        return setError(text.invalidRenderDpi)
      }
      onApply({ kind, mode: 'pages', renderDpi: flattenRenderDpi })
      return
    }
    if (kind === 'repair' || kind === 'decompress' || kind === 'removeSignatures') {
      onApply({ kind })
      return
    }
    if (kind === 'certificateSign') {
      if (!certificateFile) return setError(text.certificateRequired)
      if (certificateFile.size > 16 * 1024 * 1024) return setError(text.certificateTooLarge)
      onApply({
        kind,
        certificate: new Uint8Array(await certificateFile.arrayBuffer()),
        password: certificatePassword,
        signerName: certificateSignerName,
        reason: certificateReason,
        location: certificateLocation,
        contactInfo: certificateContact,
      })
      return
    }
    if (kind === 'timestamp') {
      onApply({ kind, tsaUrl: timestampTsaUrl })
      return
    }
    if (kind === 'password') {
      if (encrypted) {
        onApply({ kind, action: 'unlock', password: currentPassword })
        return
      }
      if (!userPassword && !ownerPassword) return setError(text.passwordRequired)
      if (userPassword !== userPasswordConfirm || ownerPassword !== ownerPasswordConfirm) {
        return setError(text.passwordMismatch)
      }
      const hasRestrictions = Object.values(passwordPermissions).some((allowed) => !allowed)
      if (
        hasRestrictions &&
        (!ownerPassword || (userPassword.length > 0 && ownerPassword === userPassword))
      ) {
        return setError(text.ownerPasswordRequired)
      }
      onApply({
        kind,
        action: 'protect',
        userPassword,
        ownerPassword,
        algorithm: passwordAlgorithm,
        permissions: passwordPermissions,
      })
      return
    }
    if (kind === 'ocr') {
      if (ocrLanguages.length === 0) return setError(text.ocrNoLanguage)
      if (!Number.isInteger(ocrRenderDpi) || ocrRenderDpi < 100 || ocrRenderDpi > 300) {
        return setError(text.invalidOcrDpi)
      }
      onApply({
        kind,
        mode: ocrMode,
        languages: ocrLanguages,
        renderDpi: ocrRenderDpi,
        clean: ocrClean,
        sidecar: ocrSidecar,
      })
      return
    }
    if (kind === 'sanitize') {
      if (!Object.values(sanitizeOptions).some(Boolean)) return setError(text.selectSanitize)
      onApply({ kind, ...sanitizeOptions })
      return
    }
    if (kind === 'pipeline') {
      if (
        pipelineSteps.length === 0 ||
        pipelineSteps.length > 12 ||
        pipelineSteps.some(
          (step) => step.kind === 'sanitize' && !SANITIZE_KEYS.some((key) => step[key]),
        )
      ) {
        return setError(text.invalidPipeline)
      }
      onApply({ kind, steps: pipelineSteps })
      return
    }
    if (kind === 'overlay') {
      if (overlayFiles.length === 0) return setError(text.selectOverlayFiles)
      if (!Number.isFinite(overlayOpacity) || overlayOpacity < 0 || overlayOpacity > 1) {
        return setError(text.invalidOpacity)
      }
      const repeatCounts =
        overlayMode === 'fixedRepeat'
          ? parseRepeatCounts(overlayRepeatCounts, overlayFiles.length)
          : undefined
      if (overlayMode === 'fixedRepeat' && !repeatCounts) {
        return setError(text.invalidRepeatCounts)
      }
      setPreparingOverlay(true)
      try {
        const overlayDocuments = await Promise.all(
          overlayFiles.map(async (file) => new Uint8Array(await file.arrayBuffer())),
        )
        onApply({
          kind,
          overlayDocuments,
          mode: overlayMode,
          position: overlayPosition,
          opacity: overlayOpacity,
          ...(repeatCounts ? { repeatCounts } : {}),
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingOverlay(false)
      }
      return
    }
    if (kind === 'overlayImage') {
      if (!overlayImageFile) return setError(text.selectOverlayImage)
      if (
        !targetPageIndexes ||
        targetPageIndexes.length === 0 ||
        !Number.isFinite(overlayImageWidth) ||
        overlayImageWidth <= 0 ||
        overlayImageWidth > 100 ||
        !Number.isFinite(overlayImageMargin) ||
        overlayImageMargin < 0 ||
        !Number.isFinite(overlayImageOpacity) ||
        overlayImageOpacity < 0 ||
        overlayImageOpacity > 1 ||
        (overlayImagePosition === 'custom' &&
          (!Number.isFinite(overlayImageX) ||
            overlayImageX < 0 ||
            !Number.isFinite(overlayImageY) ||
            overlayImageY < 0))
      ) {
        return setError(text.invalidImageOverlay)
      }
      setPreparingOverlayImage(true)
      try {
        onApply({
          kind,
          image: new Uint8Array(await overlayImageFile.arrayBuffer()),
          pageIndexes: targetPageIndexes,
          position: overlayImagePosition,
          widthPercent: overlayImageWidth,
          margin: overlayImageMargin,
          opacity: overlayImageOpacity,
          layer: overlayImageLayer,
          ...(overlayImagePosition === 'custom' ? { x: overlayImageX, y: overlayImageY } : {}),
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingOverlayImage(false)
      }
      return
    }
    if (kind === 'filterPages') {
      if (
        !targetPageIndexes ||
        targetPageIndexes.length === 0 ||
        (filterCriterion === 'text' && !filterQuery.trim())
      ) {
        return setError(text.invalidContentFilter)
      }
      onApply({
        kind,
        criterion: filterCriterion,
        action: filterAction,
        pageIndexes: targetPageIndexes,
        ...(filterCriterion === 'text' ? { text: filterQuery.trim() } : {}),
        ...(filterCriterion === 'pageSize'
          ? { pageSize: filterPageSize, comparator: filterComparator }
          : {}),
        ...(filterCriterion === 'orientation' ? { orientation: filterOrientation } : {}),
        ...(filterCriterion === 'rotation' ? { rotation: filterRotation } : {}),
        caseSensitive: filterCaseSensitive,
        wholeWord: filterWholeWord,
      })
      return
    }
    if (kind === 'filterDocuments') {
      const multiplier =
        documentFilterSizeUnit === 'KB'
          ? 1024
          : documentFilterSizeUnit === 'MB'
            ? 1024 ** 2
            : 1024 ** 3
      if (
        (documentFilterCriterion === 'text' && !documentFilterQuery.trim()) ||
        (documentFilterCriterion === 'pageCount' &&
          (!Number.isInteger(documentFilterPageCount) || documentFilterPageCount < 1)) ||
        (documentFilterCriterion === 'fileSize' &&
          (!Number.isFinite(documentFilterFileSize) || documentFilterFileSize <= 0))
      ) {
        return setError(text.invalidDocumentFilter)
      }
      setPreparingDocumentFilter(true)
      try {
        const documents = await Promise.all(
          documentFilterFiles.map(async (file) => ({
            fileName: file.name,
            bytes: new Uint8Array(await file.arrayBuffer()),
          })),
        )
        onApply({
          kind,
          currentFileName: filePath.split(/[/\\]/).at(-1) || 'Document.pdf',
          documents,
          criterion: documentFilterCriterion,
          comparator: documentFilterComparator,
          text: documentFilterQuery.trim() || undefined,
          caseSensitive: documentFilterCaseSensitive,
          wholeWord: documentFilterWholeWord,
          ...(documentFilterCriterion === 'pageCount'
            ? { pageCount: documentFilterPageCount }
            : {}),
          ...(documentFilterCriterion === 'fileSize'
            ? { fileSizeBytes: Math.round(documentFilterFileSize * multiplier) }
            : {}),
          ...(documentFilterCriterion === 'pageSize' ? { pageSize: documentFilterPageSize } : {}),
          ...(documentFilterCriterion === 'rotation' ? { rotation: documentFilterRotation } : {}),
        })
      } catch (readError) {
        setError(readError instanceof Error ? readError.message : String(readError))
      } finally {
        setPreparingDocumentFilter(false)
      }
      return
    }
    if (kind === 'bulkReplaceText') {
      const rules = replacementRules.map((rule) => ({
        find: rule.find,
        replace: rule.replace,
      }))
      if (
        !targetPageIndexes ||
        targetPageIndexes.length === 0 ||
        rules.length === 0 ||
        rules.some((rule) => !rule.find)
      ) {
        return setError(text.invalidBulkReplace)
      }
      onBulkReplaceText({
        rules,
        pageIndexes: targetPageIndexes,
        caseSensitive: replacementCaseSensitive,
        wholeWord: replacementWholeWord,
      })
      return
    }
    if (kind === 'attachments') {
      if (attachmentAction === 'add') {
        if (attachmentFiles.length === 0) return setError(text.selectAttachmentFiles)
        setPreparingAttachments(true)
        try {
          const attachments = await Promise.all(
            attachmentFiles.map(async (file) => ({
              name: file.name,
              bytes: new Uint8Array(await file.arrayBuffer()),
              mimeType: file.type || undefined,
            })),
          )
          onApply({ kind, action: 'add', attachments })
        } catch (readError) {
          setError(readError instanceof Error ? readError.message : String(readError))
        } finally {
          setPreparingAttachments(false)
        }
        return
      }
      if (!selectedAttachment) return setError(text.selectExistingAttachment)
      if (attachmentAction === 'extract') {
        onApply({ kind, action: 'extract' })
        return
      }
      if (attachmentAction === 'rename') {
        if (!newAttachmentName.trim()) return setError(text.invalidAttachmentName)
        onApply({
          kind,
          action: 'rename',
          attachmentName: selectedAttachment,
          newName: newAttachmentName,
        })
        return
      }
      onApply({ kind, action: 'delete', attachmentName: selectedAttachment })
      return
    }
    if (kind === 'bookmarks') {
      if (!validBookmarks(bookmarks, pageCount)) return setError(text.invalidBookmarks)
      onApply({ kind, bookmarks })
      return
    }
    if (kind === 'metadata') {
      const dates = [metadataValues.creationDate, metadataValues.modificationDate]
      const customKeys = metadataValues.custom.map((field) => field.key.trim().toLowerCase())
      if (
        !clearAllMetadata &&
        (dates.some((value) => value && Number.isNaN(new Date(value).getTime())) ||
          metadataValues.custom.some(
            (field) =>
              !field.key.trim() ||
              field.key.trim().length > 128 ||
              /\p{Cc}/u.test(field.key) ||
              RESERVED_METADATA_KEYS.has(field.key.trim().toLowerCase()),
          ) ||
          new Set(customKeys).size !== customKeys.length)
      ) {
        return setError(text.invalidMetadata)
      }
      onApply({
        kind,
        deleteAll: clearAllMetadata,
        metadata: {
          ...metadataValues,
          creationDate: metadataValues.creationDate
            ? new Date(metadataValues.creationDate).toISOString()
            : '',
          modificationDate: metadataValues.modificationDate
            ? new Date(metadataValues.modificationDate).toISOString()
            : '',
          custom: metadataValues.custom.map((field) => ({
            key: field.key.trim(),
            value: field.value,
          })),
        },
      })
      return
    }
    if (kind === 'removeAnnotations') {
      onApply({ kind })
      return
    }
    if (kind === 'extractImages') {
      onApply({ kind, format: extractImageFormat })
      return
    }
    if (kind === 'pdfToImages') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfImageRenderDpi) ||
        pdfImageRenderDpi < 72 ||
        pdfImageRenderDpi > 300 ||
        !Number.isInteger(pdfImageQuality) ||
        pdfImageQuality < 10 ||
        pdfImageQuality > 100
      ) {
        return setError(text.invalidPdfToImages)
      }
      onApply({
        kind,
        pageCount,
        pageIndexes: targetPageIndexes,
        format: pdfImageFormat,
        outputMode: pdfImageOutputMode,
        renderDpi: pdfImageRenderDpi,
        imageQuality: pdfImageQuality,
        colorMode: pdfImageColor,
        includeAnnotations: pdfImageAnnotations,
      })
      return
    }
    if (kind === 'pdfToVideo') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfVideoSeconds) ||
        pdfVideoSeconds < 1 ||
        pdfVideoSeconds > 10 ||
        !Number.isFinite(pdfVideoTransition) ||
        pdfVideoTransition < 0 ||
        pdfVideoTransition > 1 ||
        targetPageIndexes.length > 100 ||
        targetPageIndexes.length * pdfVideoSeconds > 300
      ) {
        return setError(text.invalidPdfToVideo)
      }
      onApply({
        kind,
        pageIndexes: targetPageIndexes,
        secondsPerPage: pdfVideoSeconds,
        resolution: pdfVideoResolution,
        transitionSeconds: pdfVideoTransition,
        includeAnnotations: pdfVideoAnnotations,
      })
      return
    }
    if (kind === 'pdfToCbz') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfCbzRenderDpi) ||
        pdfCbzRenderDpi < 72 ||
        pdfCbzRenderDpi > 300 ||
        !Number.isInteger(pdfCbzQuality) ||
        pdfCbzQuality < 10 ||
        pdfCbzQuality > 100
      ) {
        return setError(text.invalidPdfToCbz)
      }
      onApply({
        kind,
        pageCount,
        pageIndexes: targetPageIndexes,
        format: pdfCbzFormat,
        renderDpi: pdfCbzRenderDpi,
        imageQuality: pdfCbzQuality,
        colorMode: pdfCbzColor,
        includeAnnotations: pdfCbzAnnotations,
      })
      return
    }
    if (kind === 'pdfToHtml') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfHtmlRenderDpi) ||
        pdfHtmlRenderDpi < 72 ||
        pdfHtmlRenderDpi > 300
      ) {
        return setError(text.invalidPdfToHtml)
      }
      onApply({
        kind,
        pageCount,
        pageIndexes: targetPageIndexes,
        renderDpi: pdfHtmlRenderDpi,
        includeAnnotations: pdfHtmlAnnotations,
      })
      return
    }
    if (kind === 'pdfToEpub') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfEpubRenderDpi) ||
        pdfEpubRenderDpi < 72 ||
        pdfEpubRenderDpi > 300
      ) {
        return setError(text.invalidPdfToEpub)
      }
      onApply({
        kind,
        pageCount,
        pageIndexes: targetPageIndexes,
        mode: pdfEpubMode,
        renderDpi: pdfEpubRenderDpi,
        includeAnnotations: pdfEpubAnnotations,
      })
      return
    }
    if (kind === 'pdfToPptx') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfPptxRenderDpi) ||
        pdfPptxRenderDpi < 72 ||
        pdfPptxRenderDpi > 300
      ) {
        return setError(text.invalidPdfToPptx)
      }
      onApply({
        kind,
        pageCount,
        pageIndexes: targetPageIndexes,
        mode: pdfPptxMode,
        renderDpi: pdfPptxRenderDpi,
        includeAnnotations: pdfPptxAnnotations,
      })
      return
    }
    if (kind === 'pdfToDocx' || kind === 'pdfToOdt') {
      if (
        !targetPageIndexes ||
        !Number.isInteger(pdfDocxRenderDpi) ||
        pdfDocxRenderDpi < 72 ||
        pdfDocxRenderDpi > 300
      ) {
        return setError(text.invalidPdfToDocx)
      }
      onApply({
        kind,
        pageCount,
        pageIndexes: targetPageIndexes,
        mode: pdfDocxMode,
        renderDpi: pdfDocxRenderDpi,
        includeAnnotations: pdfDocxAnnotations,
      })
      return
    }
    if (kind === 'pdfToRtf') {
      if (!targetPageIndexes) return setError(text.invalidPdfToRtf)
      onApply({ kind, pageCount, pageIndexes: targetPageIndexes })
      return
    }
    if (kind === 'removeBlanks') {
      if (
        !Number.isFinite(blankThreshold) ||
        blankThreshold < 0 ||
        blankThreshold > 255 ||
        !Number.isFinite(blankWhitePercent) ||
        blankWhitePercent <= 0 ||
        blankWhitePercent > 100
      ) {
        return setError(text.invalidBlankDetection)
      }
      onApply({
        kind,
        threshold: blankThreshold,
        whitePercent: blankWhitePercent,
        includeBlankPages,
      })
      return
    }
    if (
      kind === 'info' ||
      kind === 'preflight' ||
      kind === 'javascriptAudit' ||
      kind === 'signatureAudit'
    )
      return
    if (kind !== 'rearrange') return
    if (rearrangeMode === 'custom' && (!customPageOrder || customPageOrder.length === 0)) {
      return setError(text.invalidOrder)
    }
    if (
      rearrangeMode === 'duplicate' &&
      (!Number.isInteger(duplicateCount) || duplicateCount < 1 || duplicateCount > 100)
    ) {
      return setError(text.invalidCopies)
    }
    if ((rearrangeMode === 'removeFirst' || rearrangeMode === 'removeLast') && pageCount < 2) {
      return setError(text.invalidOrder)
    }
    if (rearrangeMode === 'removeFirstAndLast' && pageCount < 3) {
      return setError(text.invalidOrder)
    }
    onApply({
      kind,
      mode: rearrangeMode,
      ...(rearrangeMode === 'custom' ? { customOrder: customPageOrder! } : {}),
      ...(rearrangeMode === 'duplicate' ? { duplicateCount } : {}),
    })
  }

  const renderBookmarkRows = (items: PdfBookmark[], parentPath: number[] = []): ReactElement[] =>
    items.map((bookmark, index) => {
      const path = [...parentPath, index]
      const depth = path.length - 1
      const mutate = (transform: (current: PdfBookmark[]) => PdfBookmark[]) => {
        setBookmarks(transform)
        setBookmarkNotice('')
        setError('')
      }
      const update = (next: Partial<PdfBookmark>) => {
        mutate((current) => updateBookmarkAt(current, path, (item) => ({ ...item, ...next })))
      }
      return (
        <div className="pdf-bookmark-branch" key={path.join('-')}>
          <div className="pdf-bookmark-row" style={{ paddingLeft: Math.min(depth, 8) * 12 }}>
            <input
              className={`pdf-modal-input pdf-bookmark-title${error ? ' invalid' : ''}`}
              aria-label={text.bookmarkTitle}
              value={bookmark.title}
              maxLength={MAX_BOOKMARK_TITLE_LENGTH}
              disabled={busy || loadingBookmarks}
              onChange={(event) => update({ title: event.target.value })}
            />
            <input
              className={`pdf-modal-input pdf-bookmark-page${error ? ' invalid' : ''}`}
              aria-label={text.bookmarkPage}
              type="number"
              min={1}
              max={pageCount}
              step={1}
              value={bookmark.pageNumber}
              disabled={busy || loadingBookmarks}
              onChange={(event) => update({ pageNumber: Number(event.target.value) })}
            />
            <div className="pdf-bookmark-controls">
              <button
                className="pdf-bookmark-icon"
                type="button"
                title={text.moveUp}
                aria-label={text.moveUp}
                disabled={busy || loadingBookmarks || index === 0}
                onClick={() => mutate((current) => moveBookmarkAt(current, path, -1))}
              >
                ↑
              </button>
              <button
                className="pdf-bookmark-icon"
                type="button"
                title={text.moveDown}
                aria-label={text.moveDown}
                disabled={busy || loadingBookmarks || index === items.length - 1}
                onClick={() => mutate((current) => moveBookmarkAt(current, path, 1))}
              >
                ↓
              </button>
              <button
                className="pdf-bookmark-icon"
                type="button"
                title={text.outdentBookmark}
                aria-label={text.outdentBookmark}
                disabled={busy || loadingBookmarks || parentPath.length === 0}
                onClick={() => mutate((current) => outdentBookmarkAt(current, path))}
              >
                ←
              </button>
              <button
                className="pdf-bookmark-icon"
                type="button"
                title={text.indentBookmark}
                aria-label={text.indentBookmark}
                disabled={
                  busy ||
                  loadingBookmarks ||
                  index === 0 ||
                  depth + bookmarkSubtreeDepth(bookmark) >= MAX_BOOKMARK_DEPTH
                }
                onClick={() => mutate((current) => indentBookmarkAt(current, path))}
              >
                →
              </button>
              <button
                className="pdf-bookmark-icon"
                type="button"
                title={text.addChildBookmark}
                aria-label={text.addChildBookmark}
                disabled={
                  busy ||
                  loadingBookmarks ||
                  depth + 1 >= MAX_BOOKMARK_DEPTH ||
                  bookmarkTreeCount(bookmarks) >= MAX_BOOKMARK_COUNT
                }
                onClick={() => {
                  mutate((current) =>
                    updateBookmarkAt(current, path, (item) => ({
                      ...item,
                      children: [
                        ...item.children,
                        { title: '', pageNumber: item.pageNumber, children: [] },
                      ],
                    })),
                  )
                }}
              >
                +
              </button>
              <button
                className="pdf-bookmark-icon"
                type="button"
                title={text.deleteBookmark}
                aria-label={text.deleteBookmark}
                disabled={busy || loadingBookmarks}
                onClick={() => mutate((current) => deleteBookmarkAt(current, path))}
              >
                ×
              </button>
            </div>
          </div>
          {renderBookmarkRows(bookmark.children, path)}
        </div>
      )
    })

  return createPortal(
    <div className="pdf-modal-mask" onClick={busy ? undefined : onCancel}>
      <div
        className={`pdf-modal pdf-modal-wide pdf-tools-modal${kind === 'info' || kind === 'fontInfo' || kind === 'annotationInfo' || kind === 'securityInfo' || kind === 'preflight' || kind === 'javascriptAudit' || kind === 'signatureAudit' ? ' pdf-tools-modal-info' : ''}`}
        role="dialog"
        aria-modal="true"
        aria-label={text.title}
        onClick={(e) => e.stopPropagation()}
      >
        <div className="pdf-modal-title">{text.title}</div>
        <div className="pdf-sign-tabs pdf-tools-tabs">
          {(
            [
              'split',
              'merge',
              'imagesToPdf',
              'cbzToPdf',
              'emailToPdf',
              'epubToPdf',
              'htmlToPdf',
              'markdownToPdf',
              'pdfToImages',
              'pdfToVideo',
              'pdfToCbz',
              'pdfToHtml',
              'pdfToEpub',
              'pdfToPptx',
              'pdfToDocx',
              'pdfToOdt',
              'pdfToRtf',
              'pdfToPdfa',
              'pdfToMarkdown',
              'pdfToXlsx',
              'pdfToXml',
              'compare',
              'extractPages',
              'splitSections',
              'crop',
              'scale',
              'nup',
              'booklet',
              'poster',
              'singlePage',
              'rotatePages',
              'autoRotate',
              'deskew',
              'autoRename',
              'pageNumbers',
              'scannerEffect',
              'scannerImageSplit',
              'autoSplit',
              'removePages',
              'removeImages',
              'bulkReplaceText',
              'extractText',
              'pdfToJson',
              'jsonToPdf',
              'extractTables',
              'extractImages',
              'removeAnnotations',
              'removeBlanks',
              'invertColors',
              'replaceColors',
              'adjustColors',
              'rearrange',
              'redact',
              'comments',
              'compress',
              'flatten',
              'forms',
              'repair',
              'decompress',
              'ocr',
              'password',
              'certificateSign',
              'timestamp',
              'removeSignatures',
              'signatureAudit',
              'javascriptAudit',
              'sanitize',
              'pipeline',
              'overlay',
              'overlayImage',
              'filterPages',
              'filterDocuments',
              'attachments',
              'bookmarks',
              'metadata',
              'fontInfo',
              'annotationInfo',
              'securityInfo',
              'preflight',
              'info',
            ] as const
          ).map((tool) => (
            <button
              key={tool}
              className={`pdf-sign-tab${kind === tool ? ' active' : ''}`}
              disabled={busy || (encrypted && tool !== 'password' && tool !== 'securityInfo')}
              onClick={() => {
                setKind(tool)
                setError('')
              }}
            >
              {text[tool]}
            </button>
          ))}
        </div>

        {kind === 'merge' && (
          <div className="pdf-merge-editor">
            <div className="pdf-bookmark-toolbar">
              <span>{text.mergeFiles}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept="application/pdf,.pdf"
                  multiple
                  disabled={busy || preparingMerge}
                  onChange={(event) => {
                    const files = Array.from(event.currentTarget.files ?? [])
                    setMergeItems((current) => [
                      ...current,
                      ...files.map((file, index) => ({
                        id: `${file.name}-${file.size}-${file.lastModified}-${index}-${Date.now()}`,
                        kind: 'file' as const,
                        file,
                      })),
                    ])
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                + {text.addMergeFiles}
              </label>
            </div>
            <div className="pdf-merge-list">
              {mergeItems.map((item, index) => (
                <div className="pdf-merge-row" key={item.id}>
                  <span className="pdf-merge-order">{index + 1}</span>
                  <span
                    className="pdf-merge-name"
                    title={item.kind === 'current' ? filePath : item.file.name}
                  >
                    {item.kind === 'current'
                      ? `${filePath.split(/[\\/]/).pop() ?? filePath} · ${text.currentPdf}`
                      : item.file.name}
                  </span>
                  <div className="pdf-bookmark-controls">
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.moveUp}
                      aria-label={text.moveUp}
                      disabled={busy || preparingMerge || index === 0}
                      onClick={() => setMergeItems((current) => moveListItem(current, index, -1))}
                    >
                      ↑
                    </button>
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.moveDown}
                      aria-label={text.moveDown}
                      disabled={busy || preparingMerge || index === mergeItems.length - 1}
                      onClick={() => setMergeItems((current) => moveListItem(current, index, 1))}
                    >
                      ↓
                    </button>
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.deleteFile}
                      aria-label={text.deleteFile}
                      disabled={busy || preparingMerge || item.kind === 'current'}
                      onClick={() =>
                        setMergeItems((current) => current.filter((entry) => entry.id !== item.id))
                      }
                    >
                      ×
                    </button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {kind === 'imagesToPdf' && (
          <div className="pdf-merge-editor">
            <div className="pdf-bookmark-toolbar">
              <span>{text.imagePdfFiles}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept="image/png,image/jpeg,image/webp,image/svg+xml,.png,.jpg,.jpeg,.webp,.svg"
                  multiple
                  disabled={busy || preparingImagesToPdf}
                  onChange={(event) => {
                    const files = Array.from(event.currentTarget.files ?? [])
                    setImagePdfItems((current) => [
                      ...current,
                      ...files.map((file, index) => ({
                        id: `${file.name}-${file.size}-${file.lastModified}-${index}-${Date.now()}`,
                        file,
                      })),
                    ])
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                + {text.addImages}
              </label>
            </div>
            <CameraCapture
              disabled={busy || preparingImagesToPdf}
              labels={{
                start: text.cameraStart,
                stop: text.cameraStop,
                switchCamera: text.cameraSwitch,
                capture: text.cameraCapture,
                readyHint: text.cameraReadyHint,
                privacyHint: text.cameraPrivacyHint,
                captureFailed: text.cameraCaptureFailed,
                failures: {
                  denied: text.cameraDenied,
                  unavailable: text.cameraUnavailable,
                  unsupported: text.cameraUnsupported,
                },
              }}
              onCapture={(file) => {
                setImagePdfItems((current) => [
                  ...current,
                  {
                    id: `${file.name}-${file.size}-${file.lastModified}-${Date.now()}`,
                    file,
                  },
                ])
                setError('')
              }}
            />
            <MobileScannerCapture
              disabled={busy || preparingImagesToPdf}
              labels={{
                start: text.mobileScannerStart,
                stop: text.mobileScannerStop,
                title: text.mobileScannerTitle,
                hint: text.mobileScannerHint,
                waiting: text.mobileScannerWaiting,
                received: text.mobileScannerReceived,
                copyLink: text.mobileScannerCopyLink,
                copied: text.mobileScannerCopied,
                expiresIn: text.mobileScannerExpiresIn,
                expired: text.mobileScannerExpired,
                unavailable: text.mobileScannerUnavailable,
              }}
              onFiles={(files) => {
                const stamp = Date.now()
                setImagePdfItems((current) => [
                  ...current,
                  ...files.map((file, index) => ({
                    id: `${file.name}-${file.size}-${file.lastModified}-${index}-${stamp}`,
                    file,
                  })),
                ])
                setError('')
              }}
            />
            <div className="pdf-merge-list">
              {imagePdfItems.length === 0 && (
                <div className="pdf-bookmark-empty">{text.noFiles}</div>
              )}
              {imagePdfItems.map((item, index) => (
                <div className="pdf-merge-row" key={item.id}>
                  <span className="pdf-merge-order">{index + 1}</span>
                  <span className="pdf-merge-name" title={item.file.name}>
                    {item.file.name}
                  </span>
                  <div className="pdf-bookmark-controls">
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.moveUp}
                      aria-label={text.moveUp}
                      disabled={busy || preparingImagesToPdf || index === 0}
                      onClick={() =>
                        setImagePdfItems((current) => moveListItem(current, index, -1))
                      }
                    >
                      ↑
                    </button>
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.moveDown}
                      aria-label={text.moveDown}
                      disabled={busy || preparingImagesToPdf || index === imagePdfItems.length - 1}
                      onClick={() => setImagePdfItems((current) => moveListItem(current, index, 1))}
                    >
                      ↓
                    </button>
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.deleteFile}
                      aria-label={text.deleteFile}
                      disabled={busy || preparingImagesToPdf}
                      onClick={() =>
                        setImagePdfItems((current) =>
                          current.filter((entry) => entry.id !== item.id),
                        )
                      }
                    >
                      ×
                    </button>
                  </div>
                </div>
              ))}
            </div>
            <label className="pdf-field">
              <span>{text.imageFit}</span>
              <select
                className="pdf-modal-input"
                value={imagePdfFit}
                disabled={busy || preparingImagesToPdf}
                onChange={(event) => setImagePdfFit(event.target.value as PdfImagesToPdfFit)}
              >
                <option value="maintainAspectRatio">{text.maintainAspectRatio}</option>
                <option value="fillPage">{text.fillPage}</option>
                <option value="fitDocumentToImage">{text.fitDocumentToImage}</option>
              </select>
            </label>
            <label className="pdf-field">
              <span>{text.imageColorMode}</span>
              <span className="pdf-tools-segments">
                {(['color', 'greyscale', 'blackwhite'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${imagePdfColor === mode ? ' active' : ''}`}
                    disabled={busy || preparingImagesToPdf}
                    onClick={() => setImagePdfColor(mode)}
                  >
                    {mode === 'color'
                      ? text.color
                      : mode === 'greyscale'
                        ? text.grayscale
                        : text.blackWhite}
                  </button>
                ))}
              </span>
            </label>
            <div className="pdf-password-note">{text.svgVectorHint}</div>
            <label className="pdf-field pdf-field-check">
              <span>{text.autoRotateImages}</span>
              <input
                type="checkbox"
                checked={imagePdfAutoRotate}
                disabled={busy || preparingImagesToPdf || imagePdfFit === 'fitDocumentToImage'}
                onChange={(event) => setImagePdfAutoRotate(event.target.checked)}
              />
            </label>
            <label className="pdf-field">
              <span>{text.imagePdfDestination}</span>
              <span className="pdf-tools-segments">
                {([false, true] as const).map((append) => (
                  <button
                    key={String(append)}
                    className={`pdf-sign-tab${imagePdfAppend === append ? ' active' : ''}`}
                    disabled={busy || preparingImagesToPdf}
                    onClick={() => setImagePdfAppend(append)}
                  >
                    {append ? text.appendCurrent : text.newPdf}
                  </button>
                ))}
              </span>
            </label>
          </div>
        )}

        {kind === 'cbzToPdf' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <div className="pdf-bookmark-toolbar">
              <span>{text.cbzFile}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept=".cbz,.cbr,.zip,.rar,application/zip,application/vnd.comicbook+zip,application/vnd.rar,application/x-rar-compressed"
                  disabled={busy || preparingCbzToPdf}
                  onChange={(event) => {
                    setCbzFile(event.currentTarget.files?.[0] ?? null)
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                {cbzFile ? text.replaceCbz : text.chooseCbz}
              </label>
            </div>
            <div className="pdf-merge-list">
              {cbzFile ? (
                <div className="pdf-merge-row">
                  <span className="pdf-merge-order">1</span>
                  <span className="pdf-merge-name" title={cbzFile.name}>
                    {cbzFile.name}
                  </span>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.deleteFile}
                    aria-label={text.deleteFile}
                    disabled={busy || preparingCbzToPdf}
                    onClick={() => setCbzFile(null)}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <div className="pdf-bookmark-empty">{text.noFiles}</div>
              )}
            </div>
            <label className="pdf-field">
              <span>{text.imageFit}</span>
              <select
                className="pdf-modal-input"
                value={cbzFit}
                disabled={busy || preparingCbzToPdf}
                onChange={(event) => setCbzFit(event.target.value as PdfImagesToPdfFit)}
              >
                <option value="fitDocumentToImage">{text.fitDocumentToImage}</option>
                <option value="maintainAspectRatio">{text.maintainAspectRatio}</option>
                <option value="fillPage">{text.fillPage}</option>
              </select>
            </label>
            <label className="pdf-field">
              <span>{text.imageColorMode}</span>
              <span className="pdf-tools-segments">
                {(['color', 'greyscale', 'blackwhite'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${cbzColor === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy || preparingCbzToPdf}
                    onClick={() => setCbzColor(mode)}
                  >
                    {mode === 'color'
                      ? text.color
                      : mode === 'greyscale'
                        ? text.grayscale
                        : text.blackWhite}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.autoRotateImages}</span>
              <input
                type="checkbox"
                checked={cbzAutoRotate}
                disabled={busy || preparingCbzToPdf || cbzFit === 'fitDocumentToImage'}
                onChange={(event) => setCbzAutoRotate(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.cbzToPdfHint}</div>
          </div>
        )}

        {kind === 'emailToPdf' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <div className="pdf-bookmark-toolbar">
              <span>{text.emailFile}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept=".eml,message/rfc822"
                  disabled={busy || preparingEmailToPdf}
                  onChange={(event) => {
                    setEmailFile(event.currentTarget.files?.[0] ?? null)
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                {emailFile ? text.replaceEmail : text.chooseEmail}
              </label>
            </div>
            <div className="pdf-merge-list">
              {emailFile ? (
                <div className="pdf-merge-row">
                  <span className="pdf-merge-order">1</span>
                  <span className="pdf-merge-name" title={emailFile.name}>
                    {emailFile.name}
                  </span>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.deleteFile}
                    aria-label={text.deleteFile}
                    disabled={busy || preparingEmailToPdf}
                    onClick={() => setEmailFile(null)}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <div className="pdf-bookmark-empty">{text.noFiles}</div>
              )}
            </div>
            <label className="pdf-field">
              <span>{text.emailOutputFormat}</span>
              <span className="pdf-tools-segments">
                {(['pdf', 'html'] as const).map((format) => (
                  <button
                    key={format}
                    className={`pdf-sign-tab${emailOutputFormat === format ? ' active' : ''}`}
                    type="button"
                    disabled={busy || preparingEmailToPdf}
                    onClick={() => setEmailOutputFormat(format)}
                  >
                    {format === 'pdf' ? text.emailPdf : text.emailHtml}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.emailIncludeAllRecipients}</span>
              <input
                type="checkbox"
                checked={emailIncludeAllRecipients}
                disabled={busy || preparingEmailToPdf}
                onChange={(event) => setEmailIncludeAllRecipients(event.target.checked)}
              />
            </label>
            {emailOutputFormat === 'pdf' && (
              <>
                <label className="pdf-field pdf-field-check">
                  <span>{text.emailIncludeAttachments}</span>
                  <input
                    type="checkbox"
                    checked={emailIncludeAttachments}
                    disabled={busy || preparingEmailToPdf}
                    onChange={(event) => setEmailIncludeAttachments(event.target.checked)}
                  />
                </label>
                {emailIncludeAttachments && (
                  <label className="pdf-field">
                    <span>{text.emailMaxAttachmentSize}</span>
                    <span className="pdf-tools-size-control">
                      <input
                        className={`pdf-modal-input${error ? ' invalid' : ''}`}
                        type="number"
                        min={1}
                        max={50}
                        step={1}
                        value={emailMaxAttachmentSize}
                        disabled={busy || preparingEmailToPdf}
                        onChange={(event) => setEmailMaxAttachmentSize(Number(event.target.value))}
                      />
                      <span>MB</span>
                    </span>
                  </label>
                )}
              </>
            )}
            <div className="pdf-bookmark-empty">{text.emailLocalHint}</div>
          </div>
        )}

        {kind === 'epubToPdf' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <div className="pdf-bookmark-toolbar">
              <span>{text.epubFile}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept=".epub,application/epub+zip"
                  disabled={busy || preparingEpubToPdf}
                  onChange={(event) => {
                    setEpubFile(event.currentTarget.files?.[0] ?? null)
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                {epubFile ? text.replaceEpub : text.chooseEpub}
              </label>
            </div>
            <div className="pdf-merge-list">
              {epubFile ? (
                <div className="pdf-merge-row">
                  <span className="pdf-merge-order">1</span>
                  <span className="pdf-merge-name" title={epubFile.name}>
                    {epubFile.name}
                  </span>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.deleteFile}
                    aria-label={text.deleteFile}
                    disabled={busy || preparingEpubToPdf}
                    onClick={() => setEpubFile(null)}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <div className="pdf-bookmark-empty">{text.noFiles}</div>
              )}
            </div>
            <div className="pdf-tools-toggle-grid">
              {(
                [
                  [text.epubEmbedFonts, epubEmbedFonts, setEpubEmbedFonts],
                  [text.epubIncludeToc, epubIncludeToc, setEpubIncludeToc],
                  [text.epubIncludePageNumbers, epubIncludePageNumbers, setEpubIncludePageNumbers],
                  [text.epubOptimize, epubOptimize, setEpubOptimize],
                ] as const
              ).map(([label, checked, setter]) => (
                <label className="pdf-field pdf-field-check" key={label}>
                  <span>{label}</span>
                  <input
                    type="checkbox"
                    checked={checked}
                    disabled={busy || preparingEpubToPdf}
                    onChange={(event) => setter(event.target.checked)}
                  />
                </label>
              ))}
            </div>
            <div className="pdf-bookmark-empty">{text.epubLocalHint}</div>
          </div>
        )}

        {kind === 'jsonToPdf' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <div className="pdf-bookmark-toolbar">
              <span>{text.jsonFile}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept=".json,application/json"
                  disabled={busy}
                  onChange={(event) => {
                    setJsonFile(event.currentTarget.files?.[0] ?? null)
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                {jsonFile ? text.replaceJson : text.chooseJson}
              </label>
            </div>
            <div className="pdf-merge-list">
              {jsonFile ? (
                <div className="pdf-merge-row">
                  <span className="pdf-merge-order">1</span>
                  <span className="pdf-merge-name" title={jsonFile.name}>
                    {jsonFile.name}
                  </span>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.deleteFile}
                    aria-label={text.deleteFile}
                    disabled={busy}
                    onClick={() => setJsonFile(null)}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <div className="pdf-bookmark-empty">{text.noFiles}</div>
              )}
            </div>
            <div className="pdf-bookmark-empty">{text.jsonToPdfHint}</div>
          </div>
        )}

        {kind === 'htmlToPdf' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{text.htmlSource}</span>
              <span className="pdf-tools-segments">
                {(['file', 'website'] as const).map((source) => (
                  <button
                    key={source}
                    type="button"
                    className={`pdf-sign-tab${htmlSource === source ? ' active' : ''}`}
                    disabled={busy || preparingHtmlToPdf}
                    onClick={() => {
                      setHtmlSource(source)
                      setError('')
                    }}
                  >
                    {source === 'file' ? text.htmlSourceFile : text.htmlSourceWebsite}
                  </button>
                ))}
              </span>
            </label>
            {htmlSource === 'file' ? (
              <>
                <div className="pdf-bookmark-toolbar">
                  <span>{text.htmlFile}</span>
                  <label className="pdf-modal-btn pdf-tools-file-button">
                    <input
                      type="file"
                      accept=".html,.htm,.zip,text/html,application/zip"
                      disabled={busy || preparingHtmlToPdf}
                      onChange={(event) => {
                        setHtmlFile(event.currentTarget.files?.[0] ?? null)
                        event.currentTarget.value = ''
                        setError('')
                      }}
                    />
                    {htmlFile ? text.replaceHtml : text.chooseHtml}
                  </label>
                </div>
                <div className="pdf-merge-list">
                  {htmlFile ? (
                    <div className="pdf-merge-row">
                      <span className="pdf-merge-order">1</span>
                      <span className="pdf-merge-name" title={htmlFile.name}>
                        {htmlFile.name}
                      </span>
                      <button
                        className="pdf-bookmark-icon"
                        type="button"
                        title={text.deleteFile}
                        aria-label={text.deleteFile}
                        disabled={busy || preparingHtmlToPdf}
                        onClick={() => setHtmlFile(null)}
                      >
                        ×
                      </button>
                    </div>
                  ) : (
                    <div className="pdf-bookmark-empty">{text.noFiles}</div>
                  )}
                </div>
              </>
            ) : (
              <label className="pdf-field">
                <span>{text.htmlWebsiteUrl}</span>
                <input
                  className="pdf-modal-input"
                  type="url"
                  inputMode="url"
                  autoCapitalize="none"
                  autoCorrect="off"
                  spellCheck={false}
                  value={htmlWebsiteUrl}
                  placeholder={text.htmlWebsitePlaceholder}
                  disabled={busy || preparingHtmlToPdf}
                  onChange={(event) => {
                    setHtmlWebsiteUrl(event.target.value)
                    setError('')
                  }}
                />
              </label>
            )}
            <label className="pdf-field">
              <span>{text.htmlZoom}</span>
              <span className="pdf-tools-size-control">
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={50}
                  max={200}
                  step={10}
                  value={htmlZoom}
                  disabled={busy || preparingHtmlToPdf}
                  onChange={(event) => setHtmlZoom(Number(event.target.value))}
                />
                <span>%</span>
              </span>
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.htmlIncludePageNumbers}</span>
              <input
                type="checkbox"
                checked={htmlIncludePageNumbers}
                disabled={busy || preparingHtmlToPdf}
                onChange={(event) => setHtmlIncludePageNumbers(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">
              {htmlSource === 'file' ? text.htmlLocalHint : text.htmlRemoteHint}
            </div>
          </div>
        )}

        {kind === 'markdownToPdf' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <div className="pdf-bookmark-toolbar">
              <span>{text.markdownFile}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept=".md,.markdown,.zip,text/markdown,text/x-markdown,application/zip"
                  disabled={busy || preparingMarkdownToPdf}
                  onChange={(event) => {
                    setMarkdownFile(event.currentTarget.files?.[0] ?? null)
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                {markdownFile ? text.replaceMarkdown : text.chooseMarkdown}
              </label>
            </div>
            <div className="pdf-merge-list">
              {markdownFile ? (
                <div className="pdf-merge-row">
                  <span className="pdf-merge-order">1</span>
                  <span className="pdf-merge-name" title={markdownFile.name}>
                    {markdownFile.name}
                  </span>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.deleteFile}
                    aria-label={text.deleteFile}
                    disabled={busy || preparingMarkdownToPdf}
                    onClick={() => setMarkdownFile(null)}
                  >
                    ×
                  </button>
                </div>
              ) : (
                <div className="pdf-bookmark-empty">{text.noFiles}</div>
              )}
            </div>
            <label className="pdf-field">
              <span>{text.markdownZoom}</span>
              <span className="pdf-tools-size-control">
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={50}
                  max={200}
                  step={10}
                  value={markdownZoom}
                  disabled={busy || preparingMarkdownToPdf}
                  onChange={(event) => setMarkdownZoom(Number(event.target.value))}
                />
                <span>%</span>
              </span>
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.markdownIncludePageNumbers}</span>
              <input
                type="checkbox"
                checked={markdownIncludePageNumbers}
                disabled={busy || preparingMarkdownToPdf}
                onChange={(event) => setMarkdownIncludePageNumbers(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.markdownLocalHint}</div>
          </div>
        )}

        {kind === 'pdfToImages' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.outputFormat}</span>
              <span className="pdf-tools-segments">
                {(['png', 'jpg', 'gif', 'webp'] as const).map((format) => (
                  <button
                    key={format}
                    className={`pdf-sign-tab${pdfImageFormat === format ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfImageFormat(format)
                      setError('')
                    }}
                  >
                    {format === 'jpg' ? 'JPEG' : format.toUpperCase()}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.outputMode}</span>
              <span className="pdf-tools-segments">
                {(['multiple', 'single'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${pdfImageOutputMode === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfImageOutputMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'multiple' ? text.separateImages : text.singleLongImage}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.colorMode}</span>
              <span className="pdf-tools-segments">
                {(['color', 'greyscale', 'blackwhite'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${pdfImageColor === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfImageColor(mode)
                      setError('')
                    }}
                  >
                    {mode === 'color'
                      ? text.color
                      : mode === 'greyscale'
                        ? text.grayscale
                        : text.blackWhite}
                  </button>
                ))}
              </span>
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.renderDpi}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={72}
                  max={300}
                  step={12}
                  value={pdfImageRenderDpi}
                  disabled={busy}
                  onChange={(event) => {
                    setPdfImageRenderDpi(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.imageQuality}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={10}
                  max={100}
                  step={1}
                  value={pdfImageQuality}
                  disabled={busy || pdfImageFormat === 'png' || pdfImageFormat === 'gif'}
                  onChange={(event) => {
                    setPdfImageQuality(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeAnnotations}</span>
              <input
                type="checkbox"
                checked={pdfImageAnnotations}
                disabled={busy}
                onChange={(event) => setPdfImageAnnotations(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToImagesHint}</div>
          </div>
        )}

        {kind === 'pdfToVideo' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            {busy && videoProgress && (
              <div className="pdf-ocr-progress" aria-live="polite">
                <div className="pdf-ocr-progress-label">
                  {text.videoEncoding
                    .replace('{page}', String(videoProgress.pageNumber))
                    .replace('{total}', String(videoProgress.pageCount))}
                </div>
                <progress
                  max={1}
                  value={Math.min(
                    1,
                    (videoProgress.pageNumber - 1 + videoProgress.progress) /
                      videoProgress.pageCount,
                  )}
                />
              </div>
            )}
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.videoResolution}</span>
              <span className="pdf-tools-segments">
                {(['480p', '720p', '1080p'] as const).map((resolution) => (
                  <button
                    key={resolution}
                    className={`pdf-sign-tab${pdfVideoResolution === resolution ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfVideoResolution(resolution)
                      setError('')
                    }}
                  >
                    {resolution.toUpperCase()}
                  </button>
                ))}
              </span>
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.videoDuration}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={10}
                  step={1}
                  value={pdfVideoSeconds}
                  disabled={busy}
                  onChange={(event) => {
                    setPdfVideoSeconds(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.videoTransition}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0}
                  max={1}
                  step={0.1}
                  value={pdfVideoTransition}
                  disabled={busy}
                  onChange={(event) => {
                    setPdfVideoTransition(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeAnnotations}</span>
              <input
                type="checkbox"
                checked={pdfVideoAnnotations}
                disabled={busy}
                onChange={(event) => setPdfVideoAnnotations(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToVideoHint}</div>
          </div>
        )}

        {kind === 'pdfToCbz' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.outputFormat}</span>
              <span className="pdf-tools-segments">
                {(['png', 'jpg', 'webp'] as const).map((format) => (
                  <button
                    key={format}
                    className={`pdf-sign-tab${pdfCbzFormat === format ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfCbzFormat(format)
                      setError('')
                    }}
                  >
                    {format === 'jpg' ? 'JPEG' : format.toUpperCase()}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.colorMode}</span>
              <span className="pdf-tools-segments">
                {(['color', 'greyscale', 'blackwhite'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${pdfCbzColor === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfCbzColor(mode)
                      setError('')
                    }}
                  >
                    {mode === 'color'
                      ? text.color
                      : mode === 'greyscale'
                        ? text.grayscale
                        : text.blackWhite}
                  </button>
                ))}
              </span>
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.renderDpi}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={72}
                  max={300}
                  step={12}
                  value={pdfCbzRenderDpi}
                  disabled={busy}
                  onChange={(event) => {
                    setPdfCbzRenderDpi(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.imageQuality}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={10}
                  max={100}
                  step={1}
                  value={pdfCbzQuality}
                  disabled={busy || pdfCbzFormat === 'png'}
                  onChange={(event) => {
                    setPdfCbzQuality(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeAnnotations}</span>
              <input
                type="checkbox"
                checked={pdfCbzAnnotations}
                disabled={busy}
                onChange={(event) => setPdfCbzAnnotations(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToCbzHint}</div>
          </div>
        )}

        {kind === 'pdfToHtml' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.renderDpi}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                type="number"
                min={72}
                max={300}
                step={12}
                value={pdfHtmlRenderDpi}
                disabled={busy}
                onChange={(event) => {
                  setPdfHtmlRenderDpi(Number(event.target.value))
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeAnnotations}</span>
              <input
                type="checkbox"
                checked={pdfHtmlAnnotations}
                disabled={busy}
                onChange={(event) => setPdfHtmlAnnotations(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToHtmlHint}</div>
          </div>
        )}

        {kind === 'pdfToEpub' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{text.epubMode}</span>
              <span className="pdf-tools-segments">
                {(['reflowable', 'fixed'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${pdfEpubMode === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfEpubMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'reflowable' ? text.epubReflowable : text.epubFixed}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            {pdfEpubMode === 'fixed' && (
              <>
                <label className="pdf-field">
                  <span>{text.renderDpi}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={72}
                    max={300}
                    step={12}
                    value={pdfEpubRenderDpi}
                    disabled={busy}
                    onChange={(event) => {
                      setPdfEpubRenderDpi(Number(event.target.value))
                      setError('')
                    }}
                  />
                </label>
                <label className="pdf-field pdf-field-check">
                  <span>{text.includeAnnotations}</span>
                  <input
                    type="checkbox"
                    checked={pdfEpubAnnotations}
                    disabled={busy}
                    onChange={(event) => setPdfEpubAnnotations(event.target.checked)}
                  />
                </label>
              </>
            )}
            <div className="pdf-bookmark-empty">{text.pdfToEpubHint}</div>
          </div>
        )}

        {kind === 'pdfToPptx' && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{text.pptxMode}</span>
              <span className="pdf-tools-segments">
                {(['editableText', 'fidelity'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${pdfPptxMode === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfPptxMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'editableText' ? text.pptxEditableText : text.pptxFidelity}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            {pdfPptxMode === 'fidelity' && (
              <>
                <label className="pdf-field">
                  <span>{text.renderDpi}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={72}
                    max={300}
                    step={12}
                    value={pdfPptxRenderDpi}
                    disabled={busy}
                    onChange={(event) => {
                      setPdfPptxRenderDpi(Number(event.target.value))
                      setError('')
                    }}
                  />
                </label>
                <label className="pdf-field pdf-field-check">
                  <span>{text.includeAnnotations}</span>
                  <input
                    type="checkbox"
                    checked={pdfPptxAnnotations}
                    disabled={busy}
                    onChange={(event) => setPdfPptxAnnotations(event.target.checked)}
                  />
                </label>
              </>
            )}
            <div className="pdf-bookmark-empty">{text.pdfToPptxHint}</div>
          </div>
        )}

        {(kind === 'pdfToDocx' || kind === 'pdfToOdt') && (
          <div className="pdf-form-editor pdf-to-images-tool">
            <label className="pdf-field">
              <span>{kind === 'pdfToOdt' ? text.odtMode : text.docxMode}</span>
              <span className="pdf-tools-segments">
                {(['editableText', 'fidelity'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${pdfDocxMode === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfDocxMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'editableText' ? text.docxEditableText : text.docxFidelity}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            {pdfDocxMode === 'fidelity' && (
              <>
                <label className="pdf-field">
                  <span>{text.renderDpi}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={72}
                    max={300}
                    step={12}
                    value={pdfDocxRenderDpi}
                    disabled={busy}
                    onChange={(event) => {
                      setPdfDocxRenderDpi(Number(event.target.value))
                      setError('')
                    }}
                  />
                </label>
                <label className="pdf-field pdf-field-check">
                  <span>{text.includeAnnotations}</span>
                  <input
                    type="checkbox"
                    checked={pdfDocxAnnotations}
                    disabled={busy}
                    onChange={(event) => setPdfDocxAnnotations(event.target.checked)}
                  />
                </label>
              </>
            )}
            <div className="pdf-bookmark-empty">
              {kind === 'pdfToOdt' ? text.pdfToOdtHint : text.pdfToDocxHint}
            </div>
          </div>
        )}

        {kind === 'pdfToRtf' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToRtfHint}</div>
          </div>
        )}

        {kind === 'pdfToXlsx' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeTwoColumnTables}</span>
              <input
                type="checkbox"
                checked={includeTwoColumnTables}
                disabled={busy}
                onChange={(event) => {
                  setIncludeTwoColumnTables(event.target.checked)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToXlsxHint}</div>
          </div>
        )}

        {kind === 'split' && (
          <>
            <label className="pdf-field">
              <span>{text.splitMethod}</span>
              <select
                className="pdf-modal-input"
                value={splitMode}
                autoFocus
                disabled={busy || pageCount < 2}
                onChange={(event) => {
                  setSplitMode(event.target.value as PdfSplitMode)
                  setError('')
                }}
              >
                <option value="afterPages">{text.afterPages}</option>
                <option value="fileSize">{text.byFileSize}</option>
                <option value="pagesPerDocument">{text.pagesPerDocument}</option>
                <option value="documentCount">{text.documentCount}</option>
                <option value="chapters">{text.chapters}</option>
              </select>
            </label>
            {splitMode === 'afterPages' ? (
              <label className="pdf-field">
                <span>{text.splitAfter}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  value={splitAfter}
                  placeholder={text.splitPlaceholder}
                  disabled={busy || pageCount < 2}
                  onChange={(event) => {
                    setSplitAfter(event.target.value)
                    setError('')
                  }}
                />
              </label>
            ) : splitMode === 'chapters' ? (
              <>
                <label className="pdf-field">
                  <span>{text.chapterLevels}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={1}
                    max={21}
                    step={1}
                    value={chapterLevels}
                    disabled={busy || pageCount < 2}
                    onChange={(event) => {
                      setChapterLevels(Number(event.target.value))
                      setError('')
                    }}
                  />
                </label>
                <label className="pdf-field">
                  <span>{text.allowDuplicateChapters}</span>
                  <input
                    type="checkbox"
                    checked={allowDuplicateChapters}
                    disabled={busy || pageCount < 2}
                    onChange={(event) => setAllowDuplicateChapters(event.target.checked)}
                  />
                </label>
              </>
            ) : splitMode === 'fileSize' ? (
              <label className="pdf-field">
                <span>{text.targetFileSize}</span>
                <span className="pdf-tools-size-control">
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={0.01}
                    step={0.1}
                    value={splitSizeValue}
                    disabled={busy || pageCount < 2}
                    onChange={(event) => {
                      setSplitSizeValue(Number(event.target.value))
                      setError('')
                    }}
                  />
                  <select
                    className="pdf-modal-input"
                    value={splitSizeUnit}
                    disabled={busy || pageCount < 2}
                    onChange={(event) => {
                      setSplitSizeUnit(event.target.value as 'KB' | 'MB' | 'GB')
                      setError('')
                    }}
                  >
                    <option value="KB">KB</option>
                    <option value="MB">MB</option>
                    <option value="GB">GB</option>
                  </select>
                </span>
              </label>
            ) : (
              <label className="pdf-field">
                <span>{text.splitValue}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  step={1}
                  value={splitValue}
                  disabled={busy || pageCount < 2}
                  onChange={(event) => {
                    setSplitValue(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            )}
          </>
        )}

        {kind === 'extractPages' && (
          <label className="pdf-field">
            <span>{text.pages}</span>
            <input
              className={`pdf-modal-input${error ? ' invalid' : ''}`}
              value={extractPageOrder}
              placeholder={text.pageOrderPlaceholder}
              autoFocus
              disabled={busy}
              onChange={(event) => {
                setExtractPageOrder(event.target.value)
                setError('')
              }}
            />
          </label>
        )}

        {kind === 'compare' && (
          <>
            <div className="pdf-field">
              <span>{text.comparisonFile}</span>
              <div className="pdf-tools-file-control">
                <label className="pdf-modal-btn pdf-tools-file-button">
                  <input
                    type="file"
                    accept="application/pdf,.pdf"
                    disabled={busy || preparingComparison}
                    onChange={(event) => {
                      setComparisonFile(event.target.files?.[0] ?? null)
                      event.currentTarget.value = ''
                      setError('')
                    }}
                  />
                  {text.chooseComparisonFile}
                </label>
                <span className="pdf-tools-file-summary" title={comparisonFile?.name}>
                  {comparisonFile?.name ?? text.noFiles}
                </span>
              </div>
            </div>
            <label className="pdf-field">
              <span>{text.renderDpi}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                type="number"
                min={72}
                max={300}
                step={12}
                value={comparisonRenderDpi}
                disabled={busy || preparingComparison}
                onChange={(event) => {
                  setComparisonRenderDpi(Number(event.target.value))
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.comparisonTolerance}</span>
              <input
                type="range"
                min={0}
                max={1}
                step={0.01}
                value={comparisonThreshold}
                disabled={busy || preparingComparison}
                onChange={(event) => {
                  setComparisonThreshold(Number(event.target.value))
                  setError('')
                }}
              />
              <em>{Math.round(comparisonThreshold * 100)}%</em>
            </label>
          </>
        )}

        {kind === 'splitSections' && (
          <>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.rows}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={10}
                  step={1}
                  value={sectionRows}
                  disabled={busy}
                  onChange={(event) => {
                    setSectionRows(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.columns}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={10}
                  step={1}
                  value={sectionColumns}
                  disabled={busy}
                  onChange={(event) => {
                    setSectionColumns(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.sectionArrangement}</span>
              <select
                className="pdf-modal-input"
                value={sectionArrangement}
                disabled={busy}
                onChange={(event) =>
                  setSectionArrangement(event.target.value as PdfSectionArrangement)
                }
              >
                <option value="rows">{text.rowMajor}</option>
                <option value="columns">{text.columnMajor}</option>
              </select>
            </label>
            <label className="pdf-field">
              <span>{text.mergeSections}</span>
              <input
                type="checkbox"
                checked={mergeSections}
                disabled={busy}
                onChange={(event) => setMergeSections(event.target.checked)}
              />
            </label>
          </>
        )}

        {kind === 'crop' && (
          <>
            <label className="pdf-field">
              <span>{text.cropMode}</span>
              <span className="pdf-tools-segments">
                {(['auto', 'manual'] as const).map((mode) => (
                  <button
                    key={mode}
                    type="button"
                    className={`pdf-sign-tab${cropMode === mode ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => {
                      setCropMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'auto' ? text.cropAuto : text.cropManual}
                  </button>
                ))}
              </span>
            </label>
            {cropMode === 'auto' ? (
              <>
                <div className="pdf-tools-number-grid">
                  <label className="pdf-field">
                    <span>{text.whiteThreshold}</span>
                    <input
                      className="pdf-modal-input"
                      type="number"
                      min={0}
                      max={255}
                      step={1}
                      value={cropWhiteThreshold}
                      disabled={busy}
                      onChange={(event) => setCropWhiteThreshold(Number(event.target.value))}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.cropPadding}</span>
                    <input
                      className="pdf-modal-input"
                      type="number"
                      min={0}
                      max={144}
                      step={1}
                      value={cropPadding}
                      disabled={busy}
                      onChange={(event) => setCropPadding(Number(event.target.value))}
                    />
                    <em>{text.points}</em>
                  </label>
                </div>
                <div className="pdf-bookmark-empty">{text.autoCropHint}</div>
              </>
            ) : (
              <div className="pdf-tools-number-grid">
                {MARGIN_KEYS.map((key) => (
                  <label className="pdf-field" key={key}>
                    <span>{text[key]}</span>
                    <input
                      className="pdf-modal-input"
                      type="number"
                      min={0}
                      step={1}
                      value={margins[key]}
                      disabled={busy}
                      onChange={(event) =>
                        setMargins({ ...margins, [key]: Number(event.target.value) })
                      }
                    />
                    <em>{text.points}</em>
                  </label>
                ))}
              </div>
            )}
          </>
        )}

        {kind === 'scale' && (
          <>
            <label className="pdf-field">
              <span>{text.pageSize}</span>
              <select
                className="pdf-modal-input"
                value={pageSize}
                disabled={busy}
                onChange={(event) => setPageSize(event.target.value as PdfPageSize)}
              >
                <option value="KEEP">Keep</option>
                <option value="A3">A3</option>
                <option value="A4">A4</option>
                <option value="A5">A5</option>
                <option value="LETTER">Letter</option>
                <option value="LEGAL">Legal</option>
                <option value="TABLOID">Tabloid</option>
              </select>
            </label>
            <label className="pdf-field">
              <span>{text.contentScale}</span>
              <input
                type="range"
                min={25}
                max={125}
                step={5}
                value={Math.round(scaleFactor * 100)}
                disabled={busy}
                onChange={(event) => setScaleFactor(Number(event.target.value) / 100)}
              />
              <em>{Math.round(scaleFactor * 100)}%</em>
            </label>
          </>
        )}

        {kind === 'nup' && (
          <>
            <label className="pdf-field">
              <span>{text.pagesPerSheet}</span>
              <span className="pdf-tools-segments">
                {(
                  [
                    [2, 1, 2],
                    [4, 2, 2],
                    [9, 3, 3],
                    [16, 4, 4],
                  ] as const
                ).map(([count, rows, columns]) => (
                  <button
                    key={count}
                    className={`pdf-sign-tab${nupRows === rows && nupColumns === columns ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => {
                      setNupRows(rows)
                      setNupColumns(columns)
                      setError('')
                    }}
                  >
                    {count}
                  </button>
                ))}
              </span>
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.rows}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={30}
                  step={1}
                  value={nupRows}
                  disabled={busy}
                  onChange={(event) => {
                    setNupRows(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.columns}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={30}
                  step={1}
                  value={nupColumns}
                  disabled={busy}
                  onChange={(event) => {
                    setNupColumns(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.nupPageOrder}</span>
                <select
                  className="pdf-modal-input"
                  value={nupArrangement}
                  disabled={busy}
                  onChange={(event) => setNupArrangement(event.target.value as 'rows' | 'columns')}
                >
                  <option value="rows">{text.rowMajor}</option>
                  <option value="columns">{text.columnMajor}</option>
                </select>
              </label>
              <label className="pdf-field">
                <span>{text.readingDirection}</span>
                <select
                  className="pdf-modal-input"
                  value={nupReadingDirection}
                  disabled={busy}
                  onChange={(event) => setNupReadingDirection(event.target.value as 'ltr' | 'rtl')}
                >
                  <option value="ltr">{text.leftToRight}</option>
                  <option value="rtl">{text.rightToLeft}</option>
                </select>
              </label>
            </div>
            <label className="pdf-field">
              <span>{text.innerMargin}</span>
              <span className="pdf-tools-size-control">
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0}
                  max={1000}
                  step={1}
                  value={nupInnerMargin}
                  disabled={busy}
                  onChange={(event) => {
                    setNupInnerMargin(Number(event.target.value))
                    setError('')
                  }}
                />
                <span>{text.points}</span>
              </span>
            </label>
            <div className="pdf-field">
              <span>{text.outerMargins}</span>
              <div className="pdf-margin-grid">
                {(['top', 'right', 'bottom', 'left'] as const).map((side) => (
                  <label key={side}>
                    <span>{text[side]}</span>
                    <input
                      className={`pdf-modal-input${error ? ' invalid' : ''}`}
                      type="number"
                      min={0}
                      max={1000}
                      step={1}
                      value={nupMargins[side]}
                      disabled={busy}
                      onChange={(event) => {
                        setNupMargins((current) => ({
                          ...current,
                          [side]: Number(event.target.value),
                        }))
                        setError('')
                      }}
                    />
                  </label>
                ))}
              </div>
            </div>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field pdf-field-check">
                <span>{text.border}</span>
                <input
                  type="checkbox"
                  checked={nupBorder}
                  disabled={busy}
                  onChange={(event) => setNupBorder(event.target.checked)}
                />
              </label>
              <label className={`pdf-field${nupBorder ? '' : ' is-disabled'}`}>
                <span>{text.borderWidth}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0.1}
                  max={72}
                  step={0.5}
                  value={nupBorderWidth}
                  disabled={busy || !nupBorder}
                  onChange={(event) => {
                    setNupBorderWidth(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
          </>
        )}

        {kind === 'booklet' && (
          <>
            <label className="pdf-field">
              <span>{text.spine}</span>
              <span className="pdf-tools-segments">
                {(['left', 'right'] as const).map((value) => (
                  <button
                    key={value}
                    className={`pdf-sign-tab${bookletSpine === value ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => setBookletSpine(value)}
                  >
                    {value === 'left' ? text.spineLeft : text.spineRight}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.gutter}</span>
              <input
                className="pdf-modal-input"
                type="number"
                min={0}
                step={1}
                value={bookletGutter}
                disabled={busy}
                onChange={(event) => {
                  setBookletGutter(Number(event.target.value))
                  setError('')
                }}
              />
              <em>{text.points}</em>
            </label>
            <label className="pdf-field">
              <span>{text.duplexPass}</span>
              <select
                className="pdf-modal-input"
                value={duplexPass}
                disabled={busy}
                onChange={(event) => setDuplexPass(event.target.value as 'both' | 'front' | 'back')}
              >
                <option value="both">{text.both}</option>
                <option value="front">{text.front}</option>
                <option value="back">{text.back}</option>
              </select>
            </label>
            <div className="pdf-tools-toggle-grid">
              <label className="pdf-field">
                <span>{text.border}</span>
                <input
                  type="checkbox"
                  checked={border}
                  disabled={busy}
                  onChange={(event) => setBorder(event.target.checked)}
                />
              </label>
              <label className="pdf-field">
                <span>{text.shortEdge}</span>
                <input
                  type="checkbox"
                  checked={flipOnShortEdge}
                  disabled={busy || duplexPass === 'front'}
                  onChange={(event) => setFlipOnShortEdge(event.target.checked)}
                />
              </label>
            </div>
          </>
        )}

        {kind === 'poster' && (
          <>
            <label className="pdf-field">
              <span>{text.pageSize}</span>
              <select
                className="pdf-modal-input"
                value={posterPageSize}
                disabled={busy}
                onChange={(event) =>
                  setPosterPageSize(event.target.value as Exclude<PdfPageSize, 'KEEP'>)
                }
              >
                <option value="A3">A3</option>
                <option value="A4">A4</option>
                <option value="A5">A5</option>
                <option value="LETTER">Letter</option>
                <option value="LEGAL">Legal</option>
                <option value="TABLOID">Tabloid</option>
              </select>
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.columns}</span>
                <input
                  className="pdf-modal-input"
                  type="number"
                  min={1}
                  max={10}
                  step={1}
                  value={posterColumns}
                  disabled={busy}
                  onChange={(event) => {
                    setPosterColumns(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.rows}</span>
                <input
                  className="pdf-modal-input"
                  type="number"
                  min={1}
                  max={10}
                  step={1}
                  value={posterRows}
                  disabled={busy}
                  onChange={(event) => {
                    setPosterRows(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <label className="pdf-field">
              <span>{text.rightToLeft}</span>
              <input
                type="checkbox"
                checked={posterRtl}
                disabled={busy}
                onChange={(event) => setPosterRtl(event.target.checked)}
              />
            </label>
          </>
        )}

        {kind === 'singlePage' && (
          <label className="pdf-field">
            <span>{text.joinDirection}</span>
            <span className="pdf-tools-segments">
              {(['vertical', 'horizontal'] as const).map((value) => (
                <button
                  key={value}
                  className={`pdf-sign-tab${singlePageDirection === value ? ' active' : ''}`}
                  disabled={busy}
                  onClick={() => setSinglePageDirection(value)}
                >
                  {text[value]}
                </button>
              ))}
            </span>
          </label>
        )}

        {kind === 'replaceColors' && (
          <>
            <div className="pdf-field">
              <span>{text.colorMode}</span>
              <span className="pdf-tools-segments">
                {(['preset', 'custom'] as const).map((value) => (
                  <button
                    key={value}
                    className={`pdf-sign-tab${colorMode === value ? ' active' : ''}`}
                    disabled={busy}
                    aria-pressed={colorMode === value}
                    onClick={() => setColorMode(value)}
                  >
                    {value === 'preset' ? text.highContrast : text.customColors}
                  </button>
                ))}
              </span>
            </div>
            {colorMode === 'preset' ? (
              <div className="pdf-tools-color-presets" aria-label={text.colorPreset}>
                {COLOR_PRESETS.map((preset) => (
                  <button
                    key={preset.id}
                    type="button"
                    className={`pdf-tools-color-preset${colorPreset === preset.id ? ' active' : ''}`}
                    disabled={busy}
                    aria-pressed={colorPreset === preset.id}
                    onClick={() => setColorPreset(preset.id)}
                  >
                    <span className="pdf-tools-color-pair" aria-hidden>
                      <span style={{ background: preset.textColor }} />
                      <span style={{ background: preset.backgroundColor }} />
                    </span>
                    <span>{text[preset.id]}</span>
                  </button>
                ))}
              </div>
            ) : (
              <div className="pdf-tools-custom-colors">
                <label className="pdf-field">
                  <span>{text.textColor}</span>
                  <span className="pdf-tools-color-input">
                    <input
                      type="color"
                      value={textColor}
                      disabled={busy}
                      aria-label={text.textColor}
                      onChange={(event) => setTextColor(event.target.value)}
                    />
                    <span>{textColor.toUpperCase()}</span>
                  </span>
                </label>
                <label className="pdf-field">
                  <span>{text.backgroundColor}</span>
                  <span className="pdf-tools-color-input">
                    <input
                      type="color"
                      value={backgroundColor}
                      disabled={busy}
                      aria-label={text.backgroundColor}
                      onChange={(event) => setBackgroundColor(event.target.value)}
                    />
                    <span>{backgroundColor.toUpperCase()}</span>
                  </span>
                </label>
              </div>
            )}
          </>
        )}

        {kind === 'adjustColors' && (
          <div className="pdf-tools-adjustments">
            {COLOR_ADJUSTMENT_GROUPS.map((fields, groupIndex) => (
              <div className="pdf-tools-adjustment-group" key={groupIndex}>
                {groupIndex === 1 && (
                  <span className="pdf-tools-adjustments-label">{text.rgbChannels}</span>
                )}
                {fields.map(([field, label]) => (
                  <label className="pdf-field" key={field}>
                    <span>{text[label]}</span>
                    <input
                      type="range"
                      min={0}
                      max={200}
                      step={1}
                      value={colorAdjustments[field]}
                      disabled={busy}
                      aria-label={text[label]}
                      onChange={(event) =>
                        setColorAdjustments((current) => ({
                          ...current,
                          [field]: Number(event.target.value),
                        }))
                      }
                    />
                    <em>{colorAdjustments[field]}%</em>
                  </label>
                ))}
              </div>
            ))}
          </div>
        )}

        {kind === 'removeBlanks' && (
          <>
            <label className="pdf-field">
              <span>{text.whitenessThreshold}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                type="number"
                min={0}
                max={255}
                step={1}
                value={blankThreshold}
                disabled={busy}
                onChange={(event) => {
                  setBlankThreshold(Number(event.target.value))
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.whitePercent}</span>
              <input
                type="range"
                min={0.1}
                max={100}
                step={0.1}
                value={blankWhitePercent}
                disabled={busy}
                aria-label={text.whitePercent}
                onChange={(event) => {
                  setBlankWhitePercent(Number(event.target.value))
                  setError('')
                }}
              />
              <em>{blankWhitePercent}%</em>
            </label>
            <label className="pdf-field">
              <span>{text.includeBlankPages}</span>
              <input
                type="checkbox"
                checked={includeBlankPages}
                disabled={busy}
                onChange={(event) => setIncludeBlankPages(event.target.checked)}
              />
            </label>
          </>
        )}

        {kind === 'extractImages' && (
          <label className="pdf-field">
            <span>{text.outputFormat}</span>
            <select
              className="pdf-modal-input"
              value={extractImageFormat}
              disabled={busy}
              onChange={(event) =>
                setExtractImageFormat(event.target.value as PdfExtractImageFormat)
              }
            >
              <option value="png">PNG</option>
              <option value="jpg">JPG</option>
              <option value="gif">GIF</option>
            </select>
          </label>
        )}

        {(kind === 'rotatePages' ||
          kind === 'removeImages' ||
          kind === 'invertColors' ||
          kind === 'replaceColors' ||
          kind === 'adjustColors') && (
          <>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            {kind === 'rotatePages' && (
              <label className="pdf-field">
                <span>{text.angle}</span>
                <span className="pdf-tools-segments">
                  {([90, 180, 270] as const).map((value) => (
                    <button
                      key={value}
                      className={`pdf-sign-tab${rotation === value ? ' active' : ''}`}
                      disabled={busy}
                      onClick={() => setRotation(value)}
                    >
                      {value}°
                    </button>
                  ))}
                </span>
              </label>
            )}
          </>
        )}

        {kind === 'autoRotate' && (
          <label className="pdf-field">
            <span>{text.inferUndetected}</span>
            <input
              type="checkbox"
              checked={inferUndetectedRotation}
              disabled={busy}
              onChange={(event) => setInferUndetectedRotation(event.target.checked)}
            />
          </label>
        )}

        {kind === 'deskew' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.deskewMaxAngle}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0.5}
                  max={15}
                  step={0.5}
                  value={deskewMaxAngle}
                  disabled={busy}
                  onChange={(event) => {
                    setDeskewMaxAngle(Number(event.target.value))
                    setError('')
                  }}
                />
                <em>°</em>
              </label>
              <label className="pdf-field">
                <span>{text.renderDpi}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={72}
                  max={300}
                  step={12}
                  value={deskewRenderDpi}
                  disabled={busy}
                  onChange={(event) => {
                    setDeskewRenderDpi(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeAnnotations}</span>
              <input
                type="checkbox"
                checked={deskewAnnotations}
                disabled={busy}
                onChange={(event) => setDeskewAnnotations(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.deskewHint}</div>
          </div>
        )}

        {kind === 'scannerEffect' && (
          <>
            <label className="pdf-field">
              <span>{text.scanQuality}</span>
              <span className="pdf-tools-segments">
                {(
                  [
                    ['high', 'scanHigh'],
                    ['medium', 'scanMedium'],
                    ['low', 'scanLow'],
                    ['custom', 'scanCustom'],
                  ] as const
                ).map(([value, label]) => (
                  <button
                    key={value}
                    type="button"
                    className={`pdf-sign-tab${scannerQuality === value ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => setScannerQuality(value)}
                  >
                    {text[label]}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.scanRotation}</span>
              <select
                className="pdf-modal-input"
                value={scannerRotation}
                disabled={busy}
                onChange={(event) => setScannerRotation(event.target.value as PdfScannerRotation)}
              >
                {(
                  [
                    ['none', 'scanNone'],
                    ['slight', 'scanSlight'],
                    ['moderate', 'scanModerate'],
                    ['severe', 'scanSevere'],
                  ] as const
                ).map(([value, label]) => (
                  <option key={value} value={value}>
                    {text[label]}
                  </option>
                ))}
              </select>
            </label>
            <label className="pdf-field">
              <span>{text.scanColorspace}</span>
              <span className="pdf-tools-segments">
                {(['grayscale', 'color'] as const).map((value) => (
                  <button
                    key={value}
                    type="button"
                    className={`pdf-sign-tab${scannerColorspace === value ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => setScannerColorspace(value)}
                  >
                    {text[value]}
                  </button>
                ))}
              </span>
            </label>
            {scannerQuality === 'custom' && (
              <div className="pdf-tools-number-grid">
                {(
                  [
                    ['paperBorder', scannerBorder, setScannerBorder, 0, 200, 1],
                    ['baseRotation', scannerRotate, setScannerRotate, -15, 15, 0.1],
                    [
                      'rotationVariance',
                      scannerRotateVariance,
                      setScannerRotateVariance,
                      0,
                      10,
                      0.1,
                    ],
                    ['brightness', scannerBrightness, setScannerBrightness, 0.5, 2, 0.01],
                    ['contrast', scannerContrast, setScannerContrast, 0.5, 2, 0.01],
                    ['blur', scannerBlur, setScannerBlur, 0, 10, 0.1],
                    ['noise', scannerNoise, setScannerNoise, 0, 50, 0.1],
                    ['renderDpi', scannerRenderDpi, setScannerRenderDpi, 72, 500, 1],
                    ['textureSeed', scannerSeed, setScannerSeed, -2147483648, 2147483647, 1],
                  ] as const
                ).map(([label, value, update, minimum, maximum, step]) => (
                  <label className="pdf-field" key={label}>
                    <span>{text[label]}</span>
                    <input
                      className={`pdf-modal-input${error ? ' invalid' : ''}`}
                      type="number"
                      min={minimum}
                      max={maximum}
                      step={step}
                      value={value}
                      disabled={busy}
                      onChange={(event) => {
                        update(Number(event.target.value))
                        setError('')
                      }}
                    />
                  </label>
                ))}
                <label className="pdf-field">
                  <span>{text.yellowedPaper}</span>
                  <input
                    type="checkbox"
                    checked={scannerYellowish}
                    disabled={busy}
                    onChange={(event) => setScannerYellowish(event.target.checked)}
                  />
                </label>
              </div>
            )}
          </>
        )}

        {kind === 'scannerImageSplit' && (
          <>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.scanAngleThreshold}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0}
                  max={45}
                  step={1}
                  value={scanSplitAngleThreshold}
                  disabled={busy}
                  onChange={(event) => {
                    setScanSplitAngleThreshold(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.scanTolerance}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0}
                  max={255}
                  step={1}
                  value={scanSplitTolerance}
                  disabled={busy}
                  onChange={(event) => {
                    setScanSplitTolerance(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.minimumPhotoArea}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={100000000}
                  step={100}
                  value={scanSplitMinArea}
                  disabled={busy}
                  onChange={(event) => {
                    setScanSplitMinArea(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.minimumForegroundArea}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={100000000}
                  step={10}
                  value={scanSplitMinContourArea}
                  disabled={busy}
                  onChange={(event) => {
                    setScanSplitMinContourArea(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.detectionBorder}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0}
                  max={200}
                  step={1}
                  value={scanSplitBorderSize}
                  disabled={busy}
                  onChange={(event) => {
                    setScanSplitBorderSize(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.renderDpi}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={72}
                  max={300}
                  step={12}
                  value={scanSplitRenderDpi}
                  disabled={busy}
                  onChange={(event) => {
                    setScanSplitRenderDpi(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
          </>
        )}

        {kind === 'removePages' && (
          <label className="pdf-field">
            <span>{text.pages}</span>
            <input
              className={`pdf-modal-input${error ? ' invalid' : ''}`}
              value={removePages}
              placeholder={text.pagesPlaceholder}
              autoFocus
              disabled={busy || pageCount < 2}
              onChange={(event) => {
                setRemovePages(event.target.value)
                setError('')
              }}
            />
          </label>
        )}

        {kind === 'rearrange' && (
          <>
            <label className="pdf-field">
              <span>{text.redactionMode}</span>
              <select
                className="pdf-modal-input"
                value={rearrangeMode}
                disabled={busy}
                onChange={(event) => {
                  setRearrangeMode(event.target.value as PdfRearrangeMode)
                  setError('')
                }}
              >
                {(
                  [
                    'custom',
                    'reverse',
                    'oddEven',
                    'duplex',
                    'removeFirst',
                    'removeLast',
                    'removeFirstAndLast',
                    'duplicate',
                  ] as const
                ).map((mode) => (
                  <option key={mode} value={mode}>
                    {text[mode]}
                  </option>
                ))}
              </select>
            </label>
            {rearrangeMode === 'custom' && (
              <label className="pdf-field">
                <span>{text.pageOrder}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  value={customOrder}
                  placeholder={text.pageOrderPlaceholder}
                  autoFocus
                  disabled={busy}
                  onChange={(event) => {
                    setCustomOrder(event.target.value)
                    setError('')
                  }}
                />
              </label>
            )}
            {rearrangeMode === 'duplicate' && (
              <label className="pdf-field">
                <span>{text.copies}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={100}
                  value={duplicateCount}
                  disabled={busy}
                  onChange={(event) => {
                    setDuplicateCount(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            )}
          </>
        )}

        {kind === 'redact' && (
          <>
            <label className="pdf-field">
              <span>{text.redactionMode}</span>
              <span className="pdf-tools-segments">
                {(['text', 'areas'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${redactionMode === mode ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setRedactionMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'text' ? text.redactTextMode : text.redactAreaMode}
                  </button>
                ))}
              </span>
            </label>
            {redactionMode === 'text' ? (
              <>
                <label className="pdf-field pdf-tools-field-column">
                  <span>{text.redactPatterns}</span>
                  <textarea
                    className={`pdf-modal-input pdf-tools-textarea${error ? ' invalid' : ''}`}
                    value={redactionPatterns}
                    placeholder={text.redactPatternsPlaceholder}
                    rows={4}
                    autoFocus
                    disabled={busy}
                    onChange={(event) => {
                      setRedactionPatterns(event.target.value)
                      setError('')
                    }}
                  />
                </label>
                <div className="pdf-tools-toggle-grid">
                  <label className="pdf-field">
                    <span>{text.useRegex}</span>
                    <input
                      type="checkbox"
                      checked={redactionUseRegex}
                      disabled={busy}
                      onChange={(event) => {
                        setRedactionUseRegex(event.target.checked)
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.wholeWord}</span>
                    <input
                      type="checkbox"
                      checked={redactionWholeWord}
                      disabled={busy}
                      onChange={(event) => setRedactionWholeWord(event.target.checked)}
                    />
                  </label>
                </div>
              </>
            ) : (
              <div className="pdf-comment-editor pdf-redaction-area-editor">
                <RedactionAreaPreview
                  filePath={filePath}
                  areas={redactionAreas}
                  selectedIndex={selectedRedactionAreaIndex}
                  color={redactionColor}
                  label={text.redactionPreview}
                  resizeLabel={text.resizeRedactionArea}
                  disabled={busy}
                  onChange={(nextArea) => {
                    setRedactionAreas((current) =>
                      current.map((item, index) =>
                        index === selectedRedactionAreaIndex ? nextArea : item,
                      ),
                    )
                    setError('')
                  }}
                />
                <div className="pdf-comment-toolbar">
                  <span>{text.redactAreaMode}</span>
                  <button
                    className="pdf-modal-btn"
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setSelectedRedactionAreaIndex(redactionAreas.length)
                      setRedactionAreas((current) => [
                        ...current,
                        emptyPdfRedactionArea(Math.max(0, currentPage - 1)),
                      ])
                      setError('')
                    }}
                  >
                    + {text.addRedactionArea}
                  </button>
                </div>
                <div className="pdf-comment-list">
                  {redactionAreas.map((area, areaIndex) => (
                    <section
                      className={`pdf-comment-row pdf-redaction-area-row${
                        areaIndex === selectedRedactionAreaIndex ? ' selected' : ''
                      }`}
                      key={areaIndex}
                      onClick={() => setSelectedRedactionAreaIndex(areaIndex)}
                      onFocusCapture={() => setSelectedRedactionAreaIndex(areaIndex)}
                    >
                      <div className="pdf-comment-row-header">
                        <span>
                          {text.redactionArea} {areaIndex + 1}
                        </span>
                        <button
                          className="pdf-bookmark-icon"
                          type="button"
                          title={text.deleteRedactionArea}
                          aria-label={text.deleteRedactionArea}
                          disabled={busy || redactionAreas.length === 1}
                          onClick={(event) => {
                            event.stopPropagation()
                            const nextAreas = redactionAreas.filter(
                              (_, index) => index !== areaIndex,
                            )
                            setRedactionAreas(nextAreas)
                            setSelectedRedactionAreaIndex((current) =>
                              Math.min(
                                current > areaIndex ? current - 1 : current,
                                nextAreas.length - 1,
                              ),
                            )
                            setError('')
                          }}
                        >
                          ×
                        </button>
                      </div>
                      <div className="pdf-redaction-area-fields">
                        {(
                          [
                            ['pageIndex', text.bookmarkPage, 1, pageCount, 1],
                            ['x', text.redactLeft, 0, 100, 1],
                            ['y', text.redactTop, 0, 100, 1],
                            ['width', text.redactWidth, 0.1, 100, 1],
                            ['height', text.redactHeight, 0.1, 100, 1],
                          ] as const
                        ).map(([key, label, min, max, step]) => (
                          <label className="pdf-field" key={key}>
                            <span>{label}</span>
                            <input
                              className={`pdf-modal-input${error ? ' invalid' : ''}`}
                              type="number"
                              min={min}
                              max={max}
                              step={step}
                              value={
                                key === 'pageIndex'
                                  ? area.pageIndex + 1
                                  : Math.round(area[key] * 1000) / 10
                              }
                              disabled={busy}
                              onChange={(event) => {
                                const raw = Number(event.target.value)
                                const value = key === 'pageIndex' ? raw - 1 : raw / 100
                                setRedactionAreas((current) =>
                                  current.map((item, index) =>
                                    index === areaIndex ? { ...item, [key]: value } : item,
                                  ),
                                )
                                setError('')
                              }}
                            />
                            {key !== 'pageIndex' && <em>%</em>}
                          </label>
                        ))}
                      </div>
                    </section>
                  ))}
                </div>
                <p className="pdf-tools-hint">{text.redactionAreaHint}</p>
              </div>
            )}
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.redactColor}</span>
                <input
                  className="pdf-tools-color-input"
                  type="color"
                  value={redactionColor}
                  disabled={busy}
                  onChange={(event) => setRedactionColor(event.target.value)}
                />
              </label>
              {redactionMode === 'text' && (
                <label className="pdf-field">
                  <span>{text.redactPadding}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={0}
                    max={72}
                    step={1}
                    value={redactionPadding}
                    disabled={busy}
                    onChange={(event) => {
                      setRedactionPadding(Number(event.target.value))
                      setError('')
                    }}
                  />
                  <em>{text.points}</em>
                </label>
              )}
              <label className="pdf-field">
                <span>{text.renderDpi}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={72}
                  max={600}
                  step={12}
                  value={redactionRenderDpi}
                  disabled={busy}
                  onChange={(event) => {
                    setRedactionRenderDpi(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            </div>
          </>
        )}

        {kind === 'comments' && (
          <div className="pdf-comment-editor">
            <div className="pdf-comment-toolbar">
              <span>{text.comments}</span>
              <button
                className="pdf-modal-btn"
                type="button"
                disabled={busy}
                onClick={() => {
                  setComments((current) => [...current, emptyPdfComment()])
                  setError('')
                }}
              >
                + {text.addComment}
              </button>
            </div>
            <div className="pdf-comment-list">
              {comments.map((comment, commentIndex) => (
                <section className="pdf-comment-row" key={commentIndex}>
                  <div className="pdf-comment-row-header">
                    <span>
                      {text.comment} {commentIndex + 1}
                    </span>
                    <button
                      className="pdf-bookmark-icon"
                      type="button"
                      title={text.deleteComment}
                      aria-label={text.deleteComment}
                      disabled={busy || comments.length === 1}
                      onClick={() => {
                        setComments((current) =>
                          current.filter((_, index) => index !== commentIndex),
                        )
                        setError('')
                      }}
                    >
                      ×
                    </button>
                  </div>
                  <div className="pdf-comment-fields">
                    <label className="pdf-field">
                      <span>{text.bookmarkPage}</span>
                      <input
                        className={`pdf-modal-input${error ? ' invalid' : ''}`}
                        type="number"
                        min={1}
                        max={pageCount}
                        value={comment.pageIndex + 1}
                        disabled={busy}
                        onChange={(event) => {
                          const pageIndex = Number(event.target.value) - 1
                          setComments((current) =>
                            current.map((item, index) =>
                              index === commentIndex ? { ...item, pageIndex } : item,
                            ),
                          )
                          setError('')
                        }}
                      />
                    </label>
                    <label className="pdf-field pdf-comment-anchor">
                      <span>{text.anchorText}</span>
                      <input
                        className="pdf-modal-input"
                        value={comment.anchorText ?? ''}
                        placeholder={text.anchorTextPlaceholder}
                        disabled={busy}
                        onChange={(event) => {
                          const anchorText = event.target.value
                          setComments((current) =>
                            current.map((item, index) =>
                              index === commentIndex ? { ...item, anchorText } : item,
                            ),
                          )
                        }}
                      />
                    </label>
                    <label className="pdf-field pdf-tools-field-column pdf-comment-body">
                      <span>{text.commentText}</span>
                      <textarea
                        className={`pdf-modal-input pdf-tools-textarea${error ? ' invalid' : ''}`}
                        value={comment.text}
                        placeholder={text.commentTextPlaceholder}
                        rows={3}
                        autoFocus={commentIndex === 0}
                        disabled={busy}
                        onChange={(event) => {
                          const commentText = event.target.value
                          setComments((current) =>
                            current.map((item, index) =>
                              index === commentIndex ? { ...item, text: commentText } : item,
                            ),
                          )
                          setError('')
                        }}
                      />
                    </label>
                    <label className="pdf-field">
                      <span>{text.author}</span>
                      <input
                        className="pdf-modal-input"
                        value={comment.author ?? ''}
                        disabled={busy}
                        onChange={(event) => {
                          const author = event.target.value
                          setComments((current) =>
                            current.map((item, index) =>
                              index === commentIndex ? { ...item, author } : item,
                            ),
                          )
                        }}
                      />
                    </label>
                    <label className="pdf-field">
                      <span>{text.subject}</span>
                      <input
                        className="pdf-modal-input"
                        value={comment.subject ?? ''}
                        disabled={busy}
                        onChange={(event) => {
                          const subject = event.target.value
                          setComments((current) =>
                            current.map((item, index) =>
                              index === commentIndex ? { ...item, subject } : item,
                            ),
                          )
                        }}
                      />
                    </label>
                    <div className="pdf-comment-position">
                      <span>{text.fallbackPosition}</span>
                      {(['x', 'y'] as const).map((axis) => (
                        <label className="pdf-field" key={axis}>
                          <span>{axis === 'x' ? text.positionX : text.positionY}</span>
                          <input
                            className={`pdf-modal-input${error ? ' invalid' : ''}`}
                            type="number"
                            value={comment[axis]}
                            disabled={busy}
                            onChange={(event) => {
                              const value = Number(event.target.value)
                              setComments((current) =>
                                current.map((item, index) =>
                                  index === commentIndex ? { ...item, [axis]: value } : item,
                                ),
                              )
                              setError('')
                            }}
                          />
                        </label>
                      ))}
                    </div>
                  </div>
                </section>
              ))}
            </div>
          </div>
        )}

        {kind === 'compress' && (
          <div className="pdf-tools-stack">
            <label className="pdf-field">
              <span>{text.renderDpi}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                type="number"
                min={72}
                max={600}
                step={12}
                value={compressRenderDpi}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setCompressRenderDpi(Number(event.target.value))
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.lineArt}</span>
              <input
                type="checkbox"
                checked={compressLineArt}
                disabled={busy}
                onChange={(event) => {
                  setCompressLineArt(event.target.checked)
                  setError('')
                }}
              />
            </label>
            {compressLineArt ? (
              <>
                <label className="pdf-field">
                  <span>{text.lineArtThreshold}</span>
                  <input
                    type="range"
                    min={0}
                    max={100}
                    step={1}
                    value={compressLineArtThreshold}
                    disabled={busy}
                    onChange={(event) => {
                      setCompressLineArtThreshold(Number(event.target.value))
                      setError('')
                    }}
                  />
                  <em>{compressLineArtThreshold}%</em>
                </label>
                <label className="pdf-field">
                  <span>{text.lineArtEdgeLevel}</span>
                  <span className="pdf-tools-segments">
                    {([1, 2, 3] as const).map((value) => (
                      <button
                        key={value}
                        type="button"
                        className={`pdf-sign-tab${compressLineArtEdgeLevel === value ? ' active' : ''}`}
                        aria-pressed={compressLineArtEdgeLevel === value}
                        disabled={busy}
                        onClick={() => {
                          setCompressLineArtEdgeLevel(value)
                          setError('')
                        }}
                      >
                        {value}
                      </button>
                    ))}
                  </span>
                </label>
              </>
            ) : (
              <label className="pdf-field">
                <span>{text.imageQuality}</span>
                <input
                  type="range"
                  min={10}
                  max={100}
                  step={1}
                  value={compressImageQuality}
                  disabled={busy}
                  onChange={(event) => {
                    setCompressImageQuality(Number(event.target.value))
                    setError('')
                  }}
                />
                <em>{compressImageQuality}%</em>
              </label>
            )}
          </div>
        )}

        {kind === 'pdfToPdfa' && (
          <div className="pdf-tools-stack">
            <label className="pdf-field">
              <span>{text.pdfaFormat}</span>
              <select className="pdf-modal-input" value="PDF/A-2b" disabled>
                <option value="PDF/A-2b">PDF/A-2b</option>
              </select>
            </label>
            <label className="pdf-field">
              <span>{text.pdfaMode}</span>
              <span className="pdf-tools-segments">
                {(['auto', 'raster'] as const).map((archiveMode) => (
                  <button
                    key={archiveMode}
                    type="button"
                    className={`pdf-sign-tab${pdfaArchiveMode === archiveMode ? ' active' : ''}`}
                    aria-pressed={pdfaArchiveMode === archiveMode}
                    disabled={busy}
                    onClick={() => {
                      setPdfaArchiveMode(archiveMode)
                      setError('')
                    }}
                  >
                    {archiveMode === 'auto' ? text.pdfaAuto : text.pdfaRaster}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pdfaRenderDpi}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                type="number"
                min={72}
                max={600}
                step={25}
                value={pdfaRenderDpi}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setPdfaRenderDpi(Number(event.target.value))
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field">
              <span>{text.pdfaImageQuality}</span>
              <input
                type="range"
                min={10}
                max={100}
                step={1}
                value={pdfaImageQuality}
                disabled={busy}
                onChange={(event) => {
                  setPdfaImageQuality(Number(event.target.value))
                  setError('')
                }}
              />
              <em>{pdfaImageQuality}%</em>
            </label>
            <div className="pdf-password-note warning">
              {pdfaArchiveMode === 'auto' ? text.pdfaAutoHint : text.pdfaRasterHint}
            </div>
          </div>
        )}

        {kind === 'flatten' && (
          <>
            <label className="pdf-field">
              <span>{text.flattenMode}</span>
              <span className="pdf-tools-segments">
                {(['pages', 'forms'] as const).map((mode) => (
                  <button
                    key={mode}
                    className={`pdf-sign-tab${flattenMode === mode ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => {
                      setFlattenMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'pages' ? text.flattenPages : text.flattenFormsOnly}
                  </button>
                ))}
              </span>
            </label>
            {flattenMode === 'pages' && (
              <label className="pdf-field">
                <span>{text.renderDpi}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={72}
                  max={600}
                  step={25}
                  value={flattenRenderDpi}
                  disabled={busy}
                  onChange={(event) => {
                    setFlattenRenderDpi(Number(event.target.value))
                    setError('')
                  }}
                />
              </label>
            )}
          </>
        )}

        {kind === 'forms' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.formAction}</span>
              <span className="pdf-tools-segments pdf-form-actions">
                {(
                  ['fill', 'export', 'create', 'modify', 'flatten', 'unlock', 'delete'] as const
                ).map((action) => (
                  <button
                    key={action}
                    className={`pdf-sign-tab${formAction === action ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => {
                      setFormAction(action)
                      setError('')
                    }}
                  >
                    {action === 'fill'
                      ? text.fillFields
                      : action === 'export'
                        ? text.exportFormData
                        : action === 'create'
                          ? text.createFields
                          : action === 'modify'
                            ? text.modifyFields
                            : action === 'flatten'
                              ? text.flattenFields
                              : action === 'unlock'
                                ? text.unlock
                                : text.deleteFields}
                  </button>
                ))}
              </span>
            </label>
            {formAction === 'fill' && (
              <div className="pdf-form-field-list">
                {loadingFormFields ? (
                  <div className="pdf-bookmark-empty">{text.loadingFormFields}</div>
                ) : formFields.length === 0 ? (
                  <div className="pdf-bookmark-empty">{text.noFormFields}</div>
                ) : (
                  formFields.map((field) => {
                    const fieldTypeLabel =
                      field.type === 'text'
                        ? text.formText
                        : field.type === 'checkbox'
                          ? text.formCheckbox
                          : field.type === 'radio'
                            ? text.formRadio
                            : field.type === 'dropdown'
                              ? text.formDropdown
                              : field.type === 'optionList'
                                ? text.formOptionList
                                : field.type === 'button'
                                  ? text.formButton
                                  : field.type === 'signature'
                                    ? text.formSignature
                                    : text.formUnknown
                    const disabled = busy || !editableFormField(field)
                    const setFieldValue = (value: PdfFormFieldValue['value']) => {
                      setFormValues((current) => ({ ...current, [field.name]: value }))
                      setError('')
                    }
                    return (
                      <div className="pdf-form-field-row" key={field.name}>
                        <div className="pdf-form-field-heading">
                          <strong>{field.name}</strong>
                          <span>
                            {fieldTypeLabel}
                            {field.required ? ` · ${text.requiredField}` : ''}
                            {field.readOnly ? ` · ${text.readOnlyField}` : ''}
                          </span>
                        </div>
                        {field.type === 'text' ? (
                          field.multiline ? (
                            <textarea
                              className="pdf-modal-input pdf-form-textarea"
                              value={String(formValues[field.name] ?? '')}
                              disabled={disabled}
                              onChange={(event) => setFieldValue(event.target.value)}
                            />
                          ) : (
                            <input
                              className="pdf-modal-input"
                              value={String(formValues[field.name] ?? '')}
                              disabled={disabled}
                              onChange={(event) => setFieldValue(event.target.value)}
                            />
                          )
                        ) : field.type === 'checkbox' ? (
                          <input
                            type="checkbox"
                            checked={formValues[field.name] === true}
                            disabled={disabled}
                            onChange={(event) => setFieldValue(event.target.checked)}
                          />
                        ) : field.type === 'radio' ? (
                          <span className="pdf-tools-segments pdf-form-radio-options">
                            <button
                              className={`pdf-sign-tab${formValues[field.name] === '' ? ' active' : ''}`}
                              type="button"
                              disabled={disabled}
                              onClick={() => setFieldValue('')}
                            >
                              {text.selectEmpty}
                            </button>
                            {(field.options ?? []).map((option) => (
                              <button
                                className={`pdf-sign-tab${formValues[field.name] === option ? ' active' : ''}`}
                                type="button"
                                key={option}
                                disabled={disabled}
                                onClick={() => setFieldValue(option)}
                              >
                                {option}
                              </button>
                            ))}
                          </span>
                        ) : field.type === 'dropdown' || field.type === 'optionList' ? (
                          <select
                            className="pdf-modal-input"
                            multiple={field.multiselect === true}
                            value={
                              field.multiselect
                                ? Array.isArray(formValues[field.name])
                                  ? (formValues[field.name] as string[])
                                  : []
                                : Array.isArray(formValues[field.name])
                                  ? ((formValues[field.name] as string[])[0] ?? '')
                                  : ''
                            }
                            disabled={disabled}
                            onChange={(event) =>
                              setFieldValue(
                                Array.from(
                                  event.currentTarget.selectedOptions,
                                  (option) => option.value,
                                ),
                              )
                            }
                          >
                            {!field.multiselect && <option value="">{text.selectEmpty}</option>}
                            {(field.options ?? []).map((option) => (
                              <option value={option} key={option}>
                                {option}
                              </option>
                            ))}
                          </select>
                        ) : (
                          <div className="pdf-form-unsupported">{text.unsupportedFormField}</div>
                        )}
                      </div>
                    )
                  })
                )}
              </div>
            )}
            {formAction === 'create' && (
              <>
                <div className="pdf-tools-inline-actions">
                  <button
                    className="pdf-modal-btn"
                    type="button"
                    disabled={busy || formCreations.length >= 500}
                    onClick={() => {
                      setFormCreations((current) => [
                        ...current,
                        emptyPdfFormFieldCreation(currentPage),
                      ])
                      setError('')
                    }}
                  >
                    + {text.addField}
                  </button>
                  <span className="pdf-tools-selection-count">
                    {text.createdFields}: {formCreations.length}
                  </span>
                </div>
                <div className="pdf-form-create-hint">{text.formCreateHint}</div>
                <div className="pdf-form-field-list pdf-form-create-list">
                  {formCreations.map((creation, creationIndex) => {
                    const choiceField = ['radio', 'dropdown', 'optionList'].includes(creation.type)
                    const updateCreation = (update: Partial<PdfFormFieldCreationEdit>) => {
                      setFormCreations((current) =>
                        current.map((item) =>
                          item.id === creation.id ? { ...item, ...update } : item,
                        ),
                      )
                      setError('')
                    }
                    return (
                      <section className="pdf-form-create-row" key={creation.id}>
                        <div className="pdf-form-create-heading">
                          <strong>
                            {text.createdFields} {creationIndex + 1}
                          </strong>
                          <button
                            className="pdf-bookmark-icon"
                            type="button"
                            title={text.removeField}
                            aria-label={`${text.removeField} ${creationIndex + 1}`}
                            disabled={busy || formCreations.length === 1}
                            onClick={() => {
                              setFormCreations((current) =>
                                current.filter((item) => item.id !== creation.id),
                              )
                              setError('')
                            }}
                          >
                            ×
                          </button>
                        </div>
                        <div className="pdf-form-create-grid">
                          <label className="pdf-form-modify-name">
                            <span>{text.fieldName}</span>
                            <input
                              className="pdf-modal-input"
                              value={creation.name}
                              disabled={busy}
                              onChange={(event) => updateCreation({ name: event.target.value })}
                            />
                          </label>
                          <label className="pdf-form-modify-name">
                            <span>{text.fieldLabel}</span>
                            <input
                              className="pdf-modal-input"
                              value={creation.label}
                              disabled={busy}
                              onChange={(event) => updateCreation({ label: event.target.value })}
                            />
                          </label>
                          <label className="pdf-form-modify-name">
                            <span>{text.fieldType}</span>
                            <select
                              className="pdf-modal-input"
                              value={creation.type}
                              disabled={busy}
                              onChange={(event) =>
                                updateCreation({
                                  type: event.target.value as PdfCreatableFormFieldType,
                                })
                              }
                            >
                              <option value="text">{text.formText}</option>
                              <option value="checkbox">{text.formCheckbox}</option>
                              <option value="radio">{text.formRadio}</option>
                              <option value="dropdown">{text.formDropdown}</option>
                              <option value="optionList">{text.formOptionList}</option>
                            </select>
                          </label>
                          <label className="pdf-form-modify-name">
                            <span>{text.fieldPage}</span>
                            <input
                              className="pdf-modal-input"
                              type="number"
                              min={1}
                              max={pageCount}
                              step={1}
                              value={creation.pageNumber}
                              disabled={busy}
                              onChange={(event) =>
                                updateCreation({ pageNumber: Number(event.target.value) })
                              }
                            />
                          </label>
                          {(
                            [
                              ['x', text.positionX],
                              ['y', text.positionY],
                              ['width', text.fieldWidth],
                              ['height', text.fieldHeight],
                            ] as const
                          ).map(([property, label]) => (
                            <label className="pdf-form-modify-name" key={property}>
                              <span>{label}</span>
                              <input
                                className="pdf-modal-input"
                                type="number"
                                min={0}
                                step={1}
                                value={creation[property]}
                                disabled={busy}
                                onChange={(event) =>
                                  updateCreation({ [property]: Number(event.target.value) })
                                }
                              />
                            </label>
                          ))}
                          <div className="pdf-form-modify-flags pdf-form-create-flags">
                            <label>
                              <input
                                type="checkbox"
                                checked={creation.readOnly}
                                disabled={busy}
                                onChange={(event) =>
                                  updateCreation({ readOnly: event.target.checked })
                                }
                              />
                              <span>{text.readOnlyField}</span>
                            </label>
                            <label>
                              <input
                                type="checkbox"
                                checked={creation.required}
                                disabled={busy}
                                onChange={(event) =>
                                  updateCreation({ required: event.target.checked })
                                }
                              />
                              <span>{text.requiredField}</span>
                            </label>
                            {creation.type === 'text' && (
                              <label>
                                <input
                                  type="checkbox"
                                  checked={creation.multiline}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateCreation({ multiline: event.target.checked })
                                  }
                                />
                                <span>{text.multilineField}</span>
                              </label>
                            )}
                            {(creation.type === 'dropdown' || creation.type === 'optionList') && (
                              <label>
                                <input
                                  type="checkbox"
                                  checked={creation.multiselect}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateCreation({ multiselect: event.target.checked })
                                  }
                                />
                                <span>{text.multiselectField}</span>
                              </label>
                            )}
                            {creation.type === 'checkbox' && (
                              <label>
                                <input
                                  type="checkbox"
                                  checked={creation.defaultChecked}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateCreation({ defaultChecked: event.target.checked })
                                  }
                                />
                                <span>{text.defaultChecked}</span>
                              </label>
                            )}
                          </div>
                          {choiceField && (
                            <label className="pdf-form-modify-options">
                              <span>{text.choiceOptions}</span>
                              <textarea
                                className="pdf-modal-input pdf-tools-textarea"
                                value={creation.optionsText}
                                disabled={busy}
                                onChange={(event) =>
                                  updateCreation({ optionsText: event.target.value })
                                }
                              />
                            </label>
                          )}
                          {creation.type !== 'checkbox' && (
                            <label className="pdf-form-modify-options">
                              <span>{choiceField ? text.defaultValues : text.defaultValue}</span>
                              {choiceField ? (
                                <textarea
                                  className="pdf-modal-input pdf-tools-textarea pdf-form-create-defaults"
                                  value={creation.defaultText}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateCreation({ defaultText: event.target.value })
                                  }
                                />
                              ) : (
                                <input
                                  className="pdf-modal-input"
                                  value={creation.defaultText}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateCreation({ defaultText: event.target.value })
                                  }
                                />
                              )}
                            </label>
                          )}
                          {creation.type === 'radio' && (
                            <label className="pdf-form-modify-name">
                              <span>{text.optionSpacing}</span>
                              <input
                                className="pdf-modal-input"
                                type="number"
                                min={creation.height}
                                step={1}
                                value={creation.optionSpacing}
                                disabled={busy}
                                onChange={(event) =>
                                  updateCreation({ optionSpacing: Number(event.target.value) })
                                }
                              />
                            </label>
                          )}
                        </div>
                      </section>
                    )
                  })}
                </div>
              </>
            )}
            {formAction === 'modify' && (
              <>
                <div className="pdf-tools-selection-count">
                  {text.modifiedFields}: {formModifications.length}/{formFields.length}
                </div>
                <div className="pdf-form-field-list">
                  {loadingFormFields ? (
                    <div className="pdf-bookmark-empty">{text.loadingFormFields}</div>
                  ) : formFields.length === 0 ? (
                    <div className="pdf-bookmark-empty">{text.noFormFields}</div>
                  ) : (
                    formFields.map((field) => {
                      const edit = formFieldEdits[field.name] ?? {
                        newName: field.name,
                        label: field.label ?? '',
                        readOnly: field.readOnly,
                        required: field.required,
                        optionsText: (field.options ?? []).join('\n'),
                        multiselect: field.multiselect ?? false,
                      }
                      const updateEdit = (update: Partial<PdfFormFieldEdit>) => {
                        setFormFieldEdits((current) => ({
                          ...current,
                          [field.name]: { ...edit, ...update },
                        }))
                        setError('')
                      }
                      return (
                        <div className="pdf-form-modify-row" key={field.name}>
                          <span className="pdf-form-field-heading">
                            <strong>{field.name}</strong>
                            <span>{field.type}</span>
                          </span>
                          <div className="pdf-form-modify-controls">
                            <label className="pdf-form-modify-name">
                              <span>{text.fieldName}</span>
                              <input
                                className="pdf-modal-input"
                                value={edit.newName}
                                disabled={busy}
                                onChange={(event) => updateEdit({ newName: event.target.value })}
                              />
                            </label>
                            <label className="pdf-form-modify-name">
                              <span>{text.fieldLabel}</span>
                              <input
                                className="pdf-modal-input"
                                value={edit.label}
                                disabled={busy}
                                onChange={(event) => updateEdit({ label: event.target.value })}
                              />
                            </label>
                            <div className="pdf-form-modify-flags">
                              <label>
                                <input
                                  type="checkbox"
                                  checked={edit.readOnly}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateEdit({ readOnly: event.target.checked })
                                  }
                                />
                                <span>{text.readOnlyField}</span>
                              </label>
                              <label>
                                <input
                                  type="checkbox"
                                  checked={edit.required}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateEdit({ required: event.target.checked })
                                  }
                                />
                                <span>{text.requiredField}</span>
                              </label>
                              {(field.type === 'dropdown' || field.type === 'optionList') && (
                                <label>
                                  <input
                                    type="checkbox"
                                    checked={edit.multiselect}
                                    disabled={busy}
                                    onChange={(event) =>
                                      updateEdit({ multiselect: event.target.checked })
                                    }
                                  />
                                  <span>{text.multiselectField}</span>
                                </label>
                              )}
                            </div>
                            {(field.type === 'dropdown' || field.type === 'optionList') && (
                              <label className="pdf-form-modify-options">
                                <span>{text.choiceOptions}</span>
                                <textarea
                                  className="pdf-modal-input pdf-tools-textarea"
                                  value={edit.optionsText}
                                  disabled={busy}
                                  onChange={(event) =>
                                    updateEdit({ optionsText: event.target.value })
                                  }
                                />
                              </label>
                            )}
                          </div>
                        </div>
                      )
                    })
                  )}
                </div>
              </>
            )}
            {formAction === 'delete' && (
              <>
                <div className="pdf-tools-inline-actions">
                  <button
                    className="pdf-modal-btn"
                    type="button"
                    disabled={busy || formFields.length === 0}
                    onClick={() => {
                      setSelectedFormFields(formFields.map((field) => field.name))
                      setError('')
                    }}
                  >
                    {text.selectAllFields}
                  </button>
                  <button
                    className="pdf-modal-btn"
                    type="button"
                    disabled={busy || selectedFormFields.length === 0}
                    onClick={() => setSelectedFormFields([])}
                  >
                    {text.clearFieldSelection}
                  </button>
                  <span className="pdf-tools-selection-count">
                    {text.selectedFields}: {selectedFormFields.length}/{formFields.length}
                  </span>
                </div>
                <div className="pdf-form-field-list">
                  {loadingFormFields ? (
                    <div className="pdf-bookmark-empty">{text.loadingFormFields}</div>
                  ) : formFields.length === 0 ? (
                    <div className="pdf-bookmark-empty">{text.noFormFields}</div>
                  ) : (
                    formFields.map((field) => (
                      <label className="pdf-form-delete-row" key={field.name}>
                        <input
                          type="checkbox"
                          checked={selectedFormFields.includes(field.name)}
                          disabled={busy}
                          onChange={(event) => {
                            setSelectedFormFields((current) =>
                              event.target.checked
                                ? [...current, field.name]
                                : current.filter((name) => name !== field.name),
                            )
                            setError('')
                          }}
                        />
                        <span className="pdf-form-field-heading">
                          <strong>{field.name}</strong>
                          <span>
                            {field.type}
                            {field.required ? ` · ${text.requiredField}` : ''}
                            {field.readOnly ? ` · ${text.readOnlyField}` : ''}
                          </span>
                        </span>
                      </label>
                    ))
                  )}
                </div>
              </>
            )}
          </div>
        )}

        {kind === 'extractText' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.textExportFormat}</span>
              <span className="pdf-tools-segments">
                {(['txt', 'markdown', 'both'] as const).map((format) => (
                  <button
                    key={format}
                    className={`pdf-sign-tab${textExportFormat === format ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setTextExportFormat(format)
                      setError('')
                    }}
                  >
                    {format === 'txt'
                      ? text.plainText
                      : format === 'markdown'
                        ? text.markdown
                        : text.exportBoth}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.textExportHint}</div>
          </div>
        )}

        {kind === 'pdfToMarkdown' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.pdfToMarkdownPageBreaks}</span>
              <input
                type="checkbox"
                checked={pdfMarkdownPageBreaks}
                disabled={busy}
                onChange={(event) => setPdfMarkdownPageBreaks(event.target.checked)}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pdfToMarkdownHint}</div>
          </div>
        )}

        {(kind === 'pdfToJson' || kind === 'pdfToXml') && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.structuredExportMode}</span>
              <span className="pdf-tools-segments">
                {([true, false] as const).map((lightweight) => (
                  <button
                    key={String(lightweight)}
                    className={`pdf-sign-tab${pdfJsonLightweight === lightweight ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setPdfJsonLightweight(lightweight)
                      setError('')
                    }}
                  >
                    {lightweight ? text.semanticJson : text.layoutJson}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                autoFocus
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">
              {kind === 'pdfToXml' ? text.pdfToXmlHint : text.pdfToJsonHint}
            </div>
          </div>
        )}

        {kind === 'autoRename' && (
          <label className="pdf-field">
            <span>{text.renameStrategy}</span>
            <span className="pdf-tools-segments">
              {(['largestHeading', 'firstText'] as const).map((strategy) => (
                <button
                  key={strategy}
                  className={`pdf-sign-tab${autoRenameStrategy === strategy ? ' active' : ''}`}
                  type="button"
                  disabled={busy}
                  onClick={() => {
                    setAutoRenameStrategy(strategy)
                    setError('')
                  }}
                >
                  {text[strategy]}
                </button>
              ))}
            </span>
          </label>
        )}

        {kind === 'pageNumbers' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-field">
              <span>{text.pageNumberPosition}</span>
              <div className="pdf-tools-position-grid">
                {(
                  [
                    'topLeft',
                    'topCenter',
                    'topRight',
                    'middleLeft',
                    'center',
                    'middleRight',
                    'bottomLeft',
                    'bottomCenter',
                    'bottomRight',
                  ] as const
                ).map((position) => (
                  <button
                    key={position}
                    type="button"
                    className={`pdf-tools-position-button${pageNumberPosition === position ? ' active' : ''}`}
                    disabled={busy}
                    title={text[position]}
                    aria-label={text[position]}
                    onClick={() => {
                      setPageNumberPosition(position)
                      setError('')
                    }}
                  >
                    <span className={`pdf-tools-position-dot ${position}`} />
                  </button>
                ))}
              </div>
            </div>
            <label className="pdf-field">
              <span>{text.pageNumberMargin}</span>
              <span className="pdf-tools-segments">
                {(['small', 'medium', 'large', 'xLarge'] as const).map((margin) => (
                  <button
                    key={margin}
                    type="button"
                    className={`pdf-sign-tab${pageNumberMargin === margin ? ' active' : ''}`}
                    disabled={busy}
                    onClick={() => setPageNumberMargin(margin)}
                  >
                    {text[margin]}
                  </button>
                ))}
              </span>
            </label>
            <div className="pdf-tools-number-grid">
              <label className="pdf-field">
                <span>{text.pageNumberFont}</span>
                <select
                  className="pdf-modal-input"
                  value={pageNumberFont}
                  disabled={busy}
                  onChange={(event) => setPageNumberFont(event.target.value as PdfPageNumberFont)}
                >
                  {(['helvetica', 'times', 'courier'] as const).map((font) => (
                    <option key={font} value={font}>
                      {text[font]}
                    </option>
                  ))}
                </select>
              </label>
              <label className="pdf-field">
                <span>{text.fontSize}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  max={200}
                  value={pageNumberFontSize}
                  disabled={busy}
                  onChange={(event) => setPageNumberFontSize(Number(event.target.value))}
                />
              </label>
              <label className="pdf-field">
                <span>{text.startingNumber}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={1}
                  step={1}
                  value={pageNumberStart}
                  disabled={busy}
                  onChange={(event) => setPageNumberStart(Number(event.target.value))}
                />
              </label>
              <label className="pdf-field">
                <span>{text.zeroPad}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  type="number"
                  min={0}
                  max={12}
                  step={1}
                  value={pageNumberZeroPad}
                  disabled={busy}
                  onChange={(event) => setPageNumberZeroPad(Number(event.target.value))}
                />
              </label>
            </div>
            <label className="pdf-field">
              <span>{text.fontColor}</span>
              <span className="pdf-tools-color-input">
                <input
                  type="color"
                  value={pageNumberColor}
                  disabled={busy}
                  onChange={(event) => setPageNumberColor(event.target.value)}
                />
                <span>{pageNumberColor.toUpperCase()}</span>
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.textPattern}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={pageNumberPattern}
                placeholder={text.textPatternPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setPageNumberPattern(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.pageNumbersHint}</div>
          </div>
        )}

        {kind === 'autoSplit' && (
          <div className="pdf-form-editor">
            <button
              className="pdf-modal-btn pdf-tools-inline-action"
              type="button"
              disabled={busy}
              onClick={() => onApply({ kind: 'autoSplit', action: 'divider' })}
            >
              {text.downloadDivider}
            </button>
            <label className="pdf-field pdf-field-check">
              <span>{text.duplexMode}</span>
              <input
                type="checkbox"
                checked={autoSplitDuplexMode}
                disabled={busy}
                onChange={(event) => {
                  setAutoSplitDuplexMode(event.target.checked)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.autoSplitHint}</div>
          </div>
        )}

        {kind === 'extractTables' && (
          <div className="pdf-form-editor">
            <label className="pdf-field">
              <span>{text.tableExportFormat}</span>
              <span className="pdf-tools-segments">
                {(['csv', 'xlsx', 'both'] as const).map((format) => (
                  <button
                    key={format}
                    className={`pdf-sign-tab${tableExportFormat === format ? ' active' : ''}`}
                    type="button"
                    disabled={busy}
                    onClick={() => {
                      setTableExportFormat(format)
                      setError('')
                    }}
                  >
                    {format === 'csv' ? text.csv : format === 'xlsx' ? text.excel : text.exportBoth}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field pdf-field-check">
              <span>{text.includeTwoColumnTables}</span>
              <input
                type="checkbox"
                checked={includeTwoColumnTables}
                disabled={busy}
                onChange={(event) => {
                  setIncludeTwoColumnTables(event.target.checked)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-empty">{text.tableExportHint}</div>
          </div>
        )}

        {kind === 'sanitize' && (
          <div className="pdf-tools-toggle-grid">
            {SANITIZE_KEYS.map((key) => (
              <label className="pdf-field" key={key}>
                <span>{text[key]}</span>
                <input
                  type="checkbox"
                  checked={sanitizeOptions[key]}
                  disabled={busy}
                  onChange={(event) => {
                    setSanitizeOptions({ ...sanitizeOptions, [key]: event.target.checked })
                    setError('')
                  }}
                />
              </label>
            ))}
          </div>
        )}

        {kind === 'pipeline' && (
          <div className="pdf-comment-editor pdf-pipeline-editor">
            <div className="pdf-comment-toolbar">
              <span>{text.pipelineStep}</span>
              <button
                className="pdf-modal-btn"
                type="button"
                disabled={busy || pipelineSteps.length >= 12}
                onClick={() => {
                  setPipelineSteps((current) => [...current, pdfPipelineStep('sanitize')])
                  setError('')
                }}
              >
                + {text.addPipelineStep}
              </button>
            </div>
            <div className="pdf-comment-list">
              {pipelineSteps.map((step, stepIndex) => (
                <section className="pdf-comment-row" key={stepIndex}>
                  <div className="pdf-comment-row-header">
                    <span>
                      {stepIndex + 1}. {text[step.kind]}
                    </span>
                    <div className="pdf-bookmark-actions">
                      <button
                        className="pdf-bookmark-icon"
                        type="button"
                        title={text.moveUp}
                        aria-label={text.moveUp}
                        disabled={busy || stepIndex === 0}
                        onClick={() =>
                          setPipelineSteps((current) => moveListItem(current, stepIndex, -1))
                        }
                      >
                        ↑
                      </button>
                      <button
                        className="pdf-bookmark-icon"
                        type="button"
                        title={text.moveDown}
                        aria-label={text.moveDown}
                        disabled={busy || stepIndex === pipelineSteps.length - 1}
                        onClick={() =>
                          setPipelineSteps((current) => moveListItem(current, stepIndex, 1))
                        }
                      >
                        ↓
                      </button>
                      <button
                        className="pdf-bookmark-icon"
                        type="button"
                        title={text.deletePipelineStep}
                        aria-label={text.deletePipelineStep}
                        disabled={busy || pipelineSteps.length === 1}
                        onClick={() => {
                          setPipelineSteps((current) =>
                            current.filter((_, index) => index !== stepIndex),
                          )
                          setError('')
                        }}
                      >
                        ×
                      </button>
                    </div>
                  </div>
                  <label className="pdf-field">
                    <span>{text.pipelineStep}</span>
                    <select
                      className="pdf-modal-input"
                      value={step.kind}
                      disabled={busy}
                      onChange={(event) => {
                        const next = pdfPipelineStep(
                          event.target.value as (typeof PIPELINE_STEP_KINDS)[number],
                        )
                        setPipelineSteps((current) =>
                          current.map((item, index) => (index === stepIndex ? next : item)),
                        )
                        setError('')
                      }}
                    >
                      {PIPELINE_STEP_KINDS.map((stepKind) => (
                        <option value={stepKind} key={stepKind}>
                          {text[stepKind]}
                        </option>
                      ))}
                    </select>
                  </label>
                  {step.kind === 'sanitize' && (
                    <div className="pdf-tools-toggle-grid pdf-pipeline-sanitize-options">
                      {SANITIZE_KEYS.map((key) => (
                        <label className="pdf-field" key={key}>
                          <span>{text[key]}</span>
                          <input
                            type="checkbox"
                            checked={step[key]}
                            disabled={busy}
                            onChange={(event) => {
                              const checked = event.target.checked
                              setPipelineSteps((current) =>
                                current.map((item, index) =>
                                  index === stepIndex && item.kind === 'sanitize'
                                    ? { ...item, [key]: checked }
                                    : item,
                                ),
                              )
                              setError('')
                            }}
                          />
                        </label>
                      ))}
                    </div>
                  )}
                </section>
              ))}
            </div>
            <p className="pdf-tools-hint">{text.pipelineHint}</p>
          </div>
        )}

        {kind === 'password' && (
          <div className="pdf-password-tool">
            {encrypted ? (
              <>
                <label className="pdf-field">
                  <span>{text.currentPassword}</span>
                  <input
                    className="pdf-modal-input"
                    type="password"
                    autoComplete="current-password"
                    value={currentPassword}
                    disabled={busy}
                    onChange={(event) => {
                      setCurrentPassword(event.target.value)
                      setError('')
                    }}
                  />
                </label>
                <div className="pdf-password-note">{text.unlockHint}</div>
                <div className="pdf-password-note warning">{text.passwordRewritesPdf}</div>
              </>
            ) : (
              <>
                <div className="pdf-password-fields">
                  <label className="pdf-field">
                    <span>{text.openPassword}</span>
                    <input
                      className="pdf-modal-input"
                      type="password"
                      autoComplete="new-password"
                      value={userPassword}
                      disabled={busy}
                      onChange={(event) => {
                        setUserPassword(event.target.value)
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.confirmPassword}</span>
                    <input
                      className="pdf-modal-input"
                      type="password"
                      autoComplete="new-password"
                      value={userPasswordConfirm}
                      disabled={busy}
                      onChange={(event) => {
                        setUserPasswordConfirm(event.target.value)
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.ownerPassword}</span>
                    <input
                      className="pdf-modal-input"
                      type="password"
                      autoComplete="new-password"
                      value={ownerPassword}
                      disabled={busy}
                      onChange={(event) => {
                        setOwnerPassword(event.target.value)
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.confirmPassword}</span>
                    <input
                      className="pdf-modal-input"
                      type="password"
                      autoComplete="new-password"
                      value={ownerPasswordConfirm}
                      disabled={busy}
                      onChange={(event) => {
                        setOwnerPasswordConfirm(event.target.value)
                        setError('')
                      }}
                    />
                  </label>
                </div>
                <div className="pdf-password-section">
                  <span>{text.encryption}</span>
                  <div className="pdf-tools-segments">
                    {(['AES-256', 'RC4'] as const).map((algorithm) => (
                      <button
                        className={`pdf-sign-tab${passwordAlgorithm === algorithm ? ' active' : ''}`}
                        type="button"
                        key={algorithm}
                        disabled={busy}
                        onClick={() => {
                          setPasswordAlgorithm(algorithm)
                          setError('')
                        }}
                      >
                        {algorithm === 'AES-256' ? text.aes256 : text.rc4}
                      </button>
                    ))}
                  </div>
                  <div className="pdf-password-note">
                    {passwordAlgorithm === 'AES-256' ? text.aesRecommended : text.rc4Compatibility}
                  </div>
                </div>
                <div className="pdf-password-section">
                  <span>{text.pdfPermissions}</span>
                  <div className="pdf-tools-toggle-grid pdf-password-permissions">
                    {(
                      [
                        'allowPrinting',
                        'allowHighQualityPrint',
                        'allowModifying',
                        'allowCopying',
                        'allowAnnotating',
                        'allowFillingForms',
                        'allowExtraction',
                        'allowAssembly',
                      ] as const
                    ).map((permission) => (
                      <label className="pdf-field" key={permission}>
                        <span>{text[permission]}</span>
                        <input
                          type="checkbox"
                          checked={passwordPermissions[permission]}
                          disabled={busy}
                          onChange={(event) => {
                            setPasswordPermissions({
                              ...passwordPermissions,
                              [permission]: event.target.checked,
                            })
                            setError('')
                          }}
                        />
                      </label>
                    ))}
                  </div>
                  <div className="pdf-password-note">{text.permissionAdvisory}</div>
                  <div className="pdf-password-note warning">{text.passwordRewritesPdf}</div>
                </div>
              </>
            )}
          </div>
        )}

        {kind === 'ocr' && (
          <div className="pdf-ocr-tool">
            {busy && ocrProgress && (
              <div className="pdf-ocr-progress" aria-live="polite">
                <div className="pdf-ocr-progress-label">
                  {ocrProgress.stage === 'analyzing'
                    ? text.ocrAnalyzing
                    : ocrProgress.stage === 'loading'
                      ? text.ocrLoading
                      : ocrProgress.stage === 'finishing'
                        ? text.ocrFinishing
                        : text.ocrRecognizing
                            .replace('{page}', String(ocrProgress.pageNumber ?? 1))
                            .replace('{total}', String(ocrProgress.pageCount))}
                </div>
                <progress
                  max={1}
                  value={
                    ocrProgress.stage === 'finishing'
                      ? 1
                      : ocrProgress.pageNumber
                        ? Math.min(
                            1,
                            (ocrProgress.pageNumber - 1 + ocrProgress.progress) /
                              ocrProgress.pageCount,
                          )
                        : ocrProgress.progress
                  }
                />
              </div>
            )}
            <div className="pdf-password-section">
              <span>{text.ocrMode}</span>
              <div className="pdf-tools-segments pdf-ocr-mode">
                {(['skipText', 'force', 'strict'] as const).map((mode) => (
                  <button
                    className={`pdf-sign-tab${ocrMode === mode ? ' active' : ''}`}
                    type="button"
                    key={mode}
                    disabled={busy}
                    onClick={() => {
                      setOcrMode(mode)
                      setError('')
                    }}
                  >
                    {mode === 'skipText'
                      ? text.ocrSkipText
                      : mode === 'force'
                        ? text.ocrForce
                        : text.ocrStrict}
                  </button>
                ))}
              </div>
            </div>
            <div className="pdf-password-section">
              <span>{text.ocrLanguages}</span>
              <div className="pdf-tools-toggle-grid">
                {(
                  [
                    ['eng', text.ocrEnglish],
                    ['chi_sim', text.ocrChinese],
                  ] as const
                ).map(([language, label]) => (
                  <label className="pdf-field" key={language}>
                    <span>{label}</span>
                    <input
                      type="checkbox"
                      checked={ocrLanguages.includes(language)}
                      disabled={busy}
                      onChange={(event) => {
                        setOcrLanguages((current) =>
                          event.target.checked
                            ? [...current, language]
                            : current.filter((item) => item !== language),
                        )
                        setError('')
                      }}
                    />
                  </label>
                ))}
              </div>
            </div>
            <label className="pdf-field">
              <span>{text.ocrRenderDpi}</span>
              <input
                className="pdf-modal-input"
                type="number"
                min={100}
                max={300}
                step={25}
                value={ocrRenderDpi}
                disabled={busy}
                onChange={(event) => {
                  setOcrRenderDpi(Number(event.target.value))
                  setError('')
                }}
              />
            </label>
            <label className="pdf-field pdf-ocr-clean">
              <span>{text.ocrClean}</span>
              <input
                type="checkbox"
                checked={ocrClean}
                disabled={busy}
                onChange={(event) => setOcrClean(event.target.checked)}
              />
            </label>
            {ocrClean && <div className="pdf-password-note">{text.ocrCleanHint}</div>}
            <label className="pdf-field pdf-ocr-clean">
              <span>{text.ocrSidecar}</span>
              <input
                type="checkbox"
                checked={ocrSidecar}
                disabled={busy}
                onChange={(event) => setOcrSidecar(event.target.checked)}
              />
            </label>
            {ocrSidecar && <div className="pdf-password-note">{text.ocrSidecarHint}</div>}
            <div className="pdf-password-note">{text.ocrLocalHint}</div>
            <div className="pdf-password-note warning">{text.ocrRewritesPdf}</div>
          </div>
        )}

        {kind === 'certificateSign' && (
          <div className="pdf-certificate-sign-tool">
            <div className="pdf-password-note">{text.certificateLocalHint}</div>
            <label className="pdf-field">
              <span>{text.certificateFile}</span>
              <span className="pdf-tools-file-control">
                <label className="pdf-modal-btn pdf-tools-file-button">
                  {text.chooseCertificate}
                  <input
                    type="file"
                    accept=".p12,.pfx,application/x-pkcs12"
                    disabled={busy}
                    onChange={(event) => {
                      const file = event.currentTarget.files?.[0] ?? null
                      setCertificateFile(file)
                      setError(file && file.size > 16 * 1024 * 1024 ? text.certificateTooLarge : '')
                    }}
                  />
                </label>
                <span className="pdf-tools-file-summary" title={certificateFile?.name}>
                  {certificateFile?.name ?? text.noCertificate}
                </span>
              </span>
            </label>
            <div className="pdf-certificate-fields">
              <label className="pdf-field">
                <span>{text.certificatePassword}</span>
                <input
                  className="pdf-modal-input"
                  type="password"
                  autoComplete="off"
                  value={certificatePassword}
                  disabled={busy}
                  onChange={(event) => {
                    setCertificatePassword(event.target.value)
                    setError('')
                  }}
                />
              </label>
              <label className="pdf-field">
                <span>{text.signerNameInput}</span>
                <input
                  className="pdf-modal-input"
                  value={certificateSignerName}
                  placeholder={text.certificateSignerPlaceholder}
                  disabled={busy}
                  maxLength={256}
                  onChange={(event) => setCertificateSignerName(event.target.value)}
                />
              </label>
              <label className="pdf-field">
                <span>{text.reason}</span>
                <input
                  className="pdf-modal-input"
                  value={certificateReason}
                  placeholder={text.certificateReasonPlaceholder}
                  disabled={busy}
                  maxLength={512}
                  onChange={(event) => setCertificateReason(event.target.value)}
                />
              </label>
              <label className="pdf-field">
                <span>{text.location}</span>
                <input
                  className="pdf-modal-input"
                  value={certificateLocation}
                  placeholder={text.certificateLocationPlaceholder}
                  disabled={busy}
                  maxLength={256}
                  onChange={(event) => setCertificateLocation(event.target.value)}
                />
              </label>
              <label className="pdf-field pdf-certificate-contact">
                <span>{text.contactInfo}</span>
                <input
                  className="pdf-modal-input"
                  value={certificateContact}
                  placeholder={text.certificateContactPlaceholder}
                  disabled={busy}
                  maxLength={256}
                  onChange={(event) => setCertificateContact(event.target.value)}
                />
              </label>
            </div>
            <div className="pdf-password-note warning">{text.certificateSecurityHint}</div>
          </div>
        )}

        {kind === 'timestamp' && (
          <div className="pdf-certificate-sign-tool">
            <label className="pdf-field">
              <span>{text.timestampAuthority}</span>
              <select
                className="pdf-modal-input"
                value={timestampTsaUrl}
                disabled={busy}
                onChange={(event) => {
                  setTimestampTsaUrl(event.target.value)
                  setError('')
                }}
              >
                {PDF_TIMESTAMP_TSA_PRESETS.map((preset) => (
                  <option value={preset.url} key={preset.id}>
                    {preset.label}
                  </option>
                ))}
              </select>
            </label>
            <div className="pdf-password-note">{text.timestampPrivacyHint}</div>
            <div className="pdf-password-note warning">{text.timestampSecurityHint}</div>
          </div>
        )}

        {kind === 'overlay' && (
          <>
            <div className="pdf-field">
              <span>{text.overlayFiles}</span>
              <div className="pdf-tools-file-control">
                <label className="pdf-modal-btn pdf-tools-file-button">
                  <input
                    type="file"
                    accept="application/pdf,.pdf"
                    multiple
                    disabled={busy || preparingOverlay}
                    onChange={(event) => {
                      const files = Array.from(event.currentTarget.files ?? [])
                      setOverlayFiles(files)
                      setOverlayRepeatCounts(files.map(() => '1').join(', '))
                      setError('')
                    }}
                  />
                  {text.chooseFiles}
                </label>
                <span
                  className="pdf-tools-file-summary"
                  title={overlayFiles.map((file) => file.name).join(', ')}
                >
                  {overlayFiles.length > 0
                    ? overlayFiles.map((file) => file.name).join(', ')
                    : text.noFiles}
                </span>
              </div>
            </div>
            <label className="pdf-field">
              <span>{text.overlayMode}</span>
              <select
                className="pdf-modal-input"
                value={overlayMode}
                disabled={busy || preparingOverlay}
                onChange={(event) => {
                  setOverlayMode(event.target.value as PdfOverlayMode)
                  setError('')
                }}
              >
                <option value="sequential">{text.sequential}</option>
                <option value="interleaved">{text.interleaved}</option>
                <option value="fixedRepeat">{text.fixedRepeat}</option>
              </select>
            </label>
            {overlayMode === 'fixedRepeat' && (
              <label className="pdf-field">
                <span>{text.repeatCounts}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  value={overlayRepeatCounts}
                  placeholder={text.repeatCountsPlaceholder}
                  disabled={busy || preparingOverlay}
                  onChange={(event) => {
                    setOverlayRepeatCounts(event.target.value)
                    setError('')
                  }}
                />
              </label>
            )}
            <label className="pdf-field">
              <span>{text.layer}</span>
              <span className="pdf-tools-segments">
                {(['foreground', 'background'] as const).map((position) => (
                  <button
                    key={position}
                    className={`pdf-sign-tab${overlayPosition === position ? ' active' : ''}`}
                    disabled={busy || preparingOverlay}
                    onClick={() => setOverlayPosition(position)}
                  >
                    {text[position]}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.opacity}</span>
              <input
                type="range"
                min={0}
                max={100}
                step={5}
                value={Math.round(overlayOpacity * 100)}
                disabled={busy || preparingOverlay}
                onChange={(event) => {
                  setOverlayOpacity(Number(event.target.value) / 100)
                  setError('')
                }}
              />
              <em>{Math.round(overlayOpacity * 100)}%</em>
            </label>
          </>
        )}

        {kind === 'overlayImage' && (
          <>
            <div className="pdf-field">
              <span>{text.overlayImageFile}</span>
              <div className="pdf-tools-file-control">
                <label className="pdf-modal-btn pdf-tools-file-button">
                  <input
                    type="file"
                    accept="image/png,image/jpeg,.png,.jpg,.jpeg"
                    disabled={busy || preparingOverlayImage}
                    onChange={(event) => {
                      setOverlayImageFile(event.currentTarget.files?.[0] ?? null)
                      setError('')
                    }}
                  />
                  {text.chooseImage}
                </label>
                <span className="pdf-tools-file-summary" title={overlayImageFile?.name ?? ''}>
                  {overlayImageFile?.name ?? text.noFiles}
                </span>
              </div>
            </div>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy || preparingOverlayImage}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-field">
              <span>{text.overlayImagePosition}</span>
              <div className="pdf-tools-position-grid">
                {(
                  [
                    'topLeft',
                    'topCenter',
                    'topRight',
                    'middleLeft',
                    'center',
                    'middleRight',
                    'bottomLeft',
                    'bottomCenter',
                    'bottomRight',
                  ] as const
                ).map((position) => (
                  <button
                    key={position}
                    className={`pdf-tools-position-button${overlayImagePosition === position ? ' active' : ''}`}
                    disabled={busy || preparingOverlayImage}
                    title={text[position]}
                    aria-label={text[position]}
                    onClick={() => {
                      setOverlayImagePosition(position)
                      setError('')
                    }}
                  >
                    <span className={`pdf-tools-position-dot ${position}`} />
                  </button>
                ))}
              </div>
              <button
                className={`pdf-modal-btn pdf-tools-custom-position${overlayImagePosition === 'custom' ? ' active' : ''}`}
                disabled={busy || preparingOverlayImage}
                onClick={() => {
                  setOverlayImagePosition('custom')
                  setError('')
                }}
              >
                {text.customPosition}
              </button>
            </div>
            {overlayImagePosition === 'custom' && (
              <>
                <div className="pdf-tools-number-grid">
                  <label className="pdf-field">
                    <span>{text.positionX}</span>
                    <input
                      className="pdf-modal-input"
                      type="number"
                      min={0}
                      step={1}
                      value={overlayImageX}
                      disabled={busy || preparingOverlayImage}
                      onChange={(event) => setOverlayImageX(Number(event.target.value))}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.positionY}</span>
                    <input
                      className="pdf-modal-input"
                      type="number"
                      min={0}
                      step={1}
                      value={overlayImageY}
                      disabled={busy || preparingOverlayImage}
                      onChange={(event) => setOverlayImageY(Number(event.target.value))}
                    />
                  </label>
                </div>
                <div className="pdf-bookmark-empty">{text.customPositionHint}</div>
              </>
            )}
            <label className="pdf-field">
              <span>{text.overlayImageWidth}</span>
              <input
                type="range"
                min={1}
                max={100}
                step={1}
                value={overlayImageWidth}
                disabled={busy || preparingOverlayImage}
                onChange={(event) => setOverlayImageWidth(Number(event.target.value))}
              />
              <em>{overlayImageWidth}%</em>
            </label>
            <label className="pdf-field">
              <span>{text.overlayImageMargin}</span>
              <input
                className="pdf-modal-input"
                type="number"
                min={0}
                step={1}
                value={overlayImageMargin}
                disabled={busy || preparingOverlayImage}
                onChange={(event) => setOverlayImageMargin(Number(event.target.value))}
              />
            </label>
            <label className="pdf-field">
              <span>{text.layer}</span>
              <span className="pdf-tools-segments">
                {(['foreground', 'background'] as const).map((position) => (
                  <button
                    key={position}
                    className={`pdf-sign-tab${overlayImageLayer === position ? ' active' : ''}`}
                    disabled={busy || preparingOverlayImage}
                    onClick={() => setOverlayImageLayer(position)}
                  >
                    {text[position]}
                  </button>
                ))}
              </span>
            </label>
            <label className="pdf-field">
              <span>{text.opacity}</span>
              <input
                type="range"
                min={0}
                max={100}
                step={5}
                value={Math.round(overlayImageOpacity * 100)}
                disabled={busy || preparingOverlayImage}
                onChange={(event) => setOverlayImageOpacity(Number(event.target.value) / 100)}
              />
              <em>{Math.round(overlayImageOpacity * 100)}%</em>
            </label>
            <div className="pdf-bookmark-empty">{text.overlayImageHint}</div>
          </>
        )}

        {kind === 'filterPages' && (
          <>
            <div className="pdf-field">
              <span>{text.filterCriterion}</span>
              <span className="pdf-tools-segments pdf-filter-criterion-segments">
                {(['text', 'image', 'pageSize', 'orientation', 'rotation'] as const).map(
                  (criterion) => (
                    <button
                      key={criterion}
                      className={`pdf-sign-tab${filterCriterion === criterion ? ' active' : ''}`}
                      disabled={busy}
                      onClick={() => {
                        setFilterCriterion(criterion)
                        setError('')
                      }}
                    >
                      {
                        {
                          text: text.filterText,
                          image: text.filterImage,
                          pageSize: text.filterPageSize,
                          orientation: text.filterOrientation,
                          rotation: text.filterRotation,
                        }[criterion]
                      }
                    </button>
                  ),
                )}
              </span>
            </div>
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            {filterCriterion === 'text' && (
              <>
                <label className="pdf-field">
                  <span>{text.filterQuery}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    value={filterQuery}
                    placeholder={text.filterQueryPlaceholder}
                    disabled={busy}
                    onChange={(event) => {
                      setFilterQuery(event.target.value)
                      setError('')
                    }}
                  />
                </label>
                <div className="pdf-tools-toggle-grid">
                  <label className="pdf-field pdf-field-check">
                    <span>{text.caseSensitive}</span>
                    <input
                      type="checkbox"
                      checked={filterCaseSensitive}
                      disabled={busy}
                      onChange={(event) => setFilterCaseSensitive(event.target.checked)}
                    />
                  </label>
                  <label className="pdf-field pdf-field-check">
                    <span>{text.wholeWord}</span>
                    <input
                      type="checkbox"
                      checked={filterWholeWord}
                      disabled={busy}
                      onChange={(event) => setFilterWholeWord(event.target.checked)}
                    />
                  </label>
                </div>
              </>
            )}
            {filterCriterion === 'pageSize' && (
              <div className="pdf-tools-number-grid">
                <label className="pdf-field">
                  <span>{text.filterPageSize}</span>
                  <select
                    className="pdf-modal-input"
                    value={filterPageSize}
                    disabled={busy}
                    onChange={(event) => {
                      setFilterPageSize(event.target.value as Exclude<PdfPageSize, 'KEEP'>)
                      setError('')
                    }}
                  >
                    {(['A3', 'A4', 'A5', 'LETTER', 'LEGAL', 'TABLOID'] as const).map((value) => (
                      <option key={value} value={value}>
                        {value}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="pdf-field">
                  <span>{text.filterComparator}</span>
                  <select
                    className="pdf-modal-input"
                    value={filterComparator}
                    disabled={busy}
                    onChange={(event) => {
                      setFilterComparator(event.target.value as PdfPageFilterComparator)
                      setError('')
                    }}
                  >
                    <option value="less">{text.filterLess}</option>
                    <option value="equal">{text.filterEqual}</option>
                    <option value="greater">{text.filterGreater}</option>
                  </select>
                </label>
              </div>
            )}
            {filterCriterion === 'orientation' && (
              <div className="pdf-field">
                <span>{text.filterOrientation}</span>
                <span className="pdf-tools-segments">
                  {(['portrait', 'landscape'] as const).map((value) => (
                    <button
                      key={value}
                      className={`pdf-sign-tab${filterOrientation === value ? ' active' : ''}`}
                      disabled={busy}
                      onClick={() => {
                        setFilterOrientation(value)
                        setError('')
                      }}
                    >
                      {text[value]}
                    </button>
                  ))}
                </span>
              </div>
            )}
            {filterCriterion === 'rotation' && (
              <div className="pdf-field">
                <span>{text.filterRotation}</span>
                <span className="pdf-tools-segments">
                  {([0, 90, 180, 270] as const).map((value) => (
                    <button
                      key={value}
                      className={`pdf-sign-tab${filterRotation === value ? ' active' : ''}`}
                      disabled={busy}
                      onClick={() => {
                        setFilterRotation(value)
                        setError('')
                      }}
                    >
                      {value}°
                    </button>
                  ))}
                </span>
              </div>
            )}
            <label className="pdf-field">
              <span>{text.filterAction}</span>
              <select
                className="pdf-modal-input"
                value={filterAction}
                disabled={busy}
                onChange={(event) => {
                  setFilterAction(event.target.value as PdfContentFilterAction)
                  setError('')
                }}
              >
                <option value="keep">{text.keepMatches}</option>
                <option value="remove">{text.removeMatches}</option>
              </select>
            </label>
            <div className="pdf-bookmark-empty">{text.filterPagesHint}</div>
          </>
        )}

        {kind === 'filterDocuments' && (
          <div className="pdf-document-filter-editor">
            <div className="pdf-bookmark-toolbar">
              <span>{text.filterDocumentsFiles}</span>
              <label className="pdf-modal-btn pdf-tools-file-button">
                <input
                  type="file"
                  accept="application/pdf,.pdf"
                  multiple
                  disabled={busy || preparingDocumentFilter}
                  onChange={(event) => {
                    const files = Array.from(event.currentTarget.files ?? [])
                    setDocumentFilterFiles((current) => [...current, ...files].slice(0, 100))
                    event.currentTarget.value = ''
                    setError('')
                  }}
                />
                + {text.addFilterDocuments}
              </label>
            </div>
            <div className="pdf-merge-list pdf-document-filter-list">
              <div className="pdf-merge-row">
                <span className="pdf-merge-order">1</span>
                <span className="pdf-merge-name">
                  {filePath.split(/[/\\]/).at(-1) || text.currentDocument}
                </span>
                <span className="pdf-document-filter-current">{text.currentDocument}</span>
              </div>
              {documentFilterFiles.map((file, index) => (
                <div
                  className="pdf-merge-row"
                  key={`${file.name}-${file.size}-${file.lastModified}-${index}`}
                >
                  <span className="pdf-merge-order">{index + 2}</span>
                  <span className="pdf-merge-name" title={file.name}>
                    {file.name}
                  </span>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.removeFilterDocument}
                    aria-label={`${text.removeFilterDocument} ${file.name}`}
                    disabled={busy || preparingDocumentFilter}
                    onClick={() => {
                      setDocumentFilterFiles((current) =>
                        current.filter((_, fileIndex) => fileIndex !== index),
                      )
                      setError('')
                    }}
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
            <div className="pdf-field">
              <span>{text.filterCriterion}</span>
              <span className="pdf-tools-segments pdf-document-filter-criteria">
                {(['text', 'image', 'pageCount', 'fileSize', 'pageSize', 'rotation'] as const).map(
                  (criterion) => (
                    <button
                      key={criterion}
                      className={`pdf-sign-tab${documentFilterCriterion === criterion ? ' active' : ''}`}
                      type="button"
                      disabled={busy || preparingDocumentFilter}
                      onClick={() => {
                        setDocumentFilterCriterion(criterion)
                        setError('')
                      }}
                    >
                      {
                        {
                          text: text.filterContainsText,
                          image: text.filterContainsImage,
                          pageCount: text.filterPageCount,
                          fileSize: text.filterFileSize,
                          pageSize: text.filterFirstPageSize,
                          rotation: text.filterFirstPageRotation,
                        }[criterion]
                      }
                    </button>
                  ),
                )}
              </span>
            </div>
            {documentFilterCriterion === 'text' && (
              <>
                <label className="pdf-field">
                  <span>{text.filterQuery}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    value={documentFilterQuery}
                    placeholder={text.filterQueryPlaceholder}
                    disabled={busy || preparingDocumentFilter}
                    onChange={(event) => {
                      setDocumentFilterQuery(event.target.value)
                      setError('')
                    }}
                  />
                </label>
                <div className="pdf-tools-toggle-grid">
                  <label className="pdf-field pdf-field-check">
                    <span>{text.caseSensitive}</span>
                    <input
                      type="checkbox"
                      checked={documentFilterCaseSensitive}
                      disabled={busy || preparingDocumentFilter}
                      onChange={(event) => setDocumentFilterCaseSensitive(event.target.checked)}
                    />
                  </label>
                  <label className="pdf-field pdf-field-check">
                    <span>{text.wholeWord}</span>
                    <input
                      type="checkbox"
                      checked={documentFilterWholeWord}
                      disabled={busy || preparingDocumentFilter}
                      onChange={(event) => setDocumentFilterWholeWord(event.target.checked)}
                    />
                  </label>
                </div>
              </>
            )}
            <div className="pdf-tools-number-grid">
              {documentFilterCriterion === 'pageCount' && (
                <label className="pdf-field">
                  <span>{text.filterExpectedValue}</span>
                  <input
                    className={`pdf-modal-input${error ? ' invalid' : ''}`}
                    type="number"
                    min={1}
                    step={1}
                    value={documentFilterPageCount}
                    disabled={busy || preparingDocumentFilter}
                    onChange={(event) => {
                      setDocumentFilterPageCount(Number(event.target.value))
                      setError('')
                    }}
                  />
                </label>
              )}
              {documentFilterCriterion === 'fileSize' && (
                <>
                  <label className="pdf-field">
                    <span>{text.filterExpectedValue}</span>
                    <input
                      className={`pdf-modal-input${error ? ' invalid' : ''}`}
                      type="number"
                      min={0.01}
                      step={0.01}
                      value={documentFilterFileSize}
                      disabled={busy || preparingDocumentFilter}
                      onChange={(event) => {
                        setDocumentFilterFileSize(Number(event.target.value))
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field">
                    <span>{text.filterSizeUnit}</span>
                    <select
                      className="pdf-modal-input"
                      value={documentFilterSizeUnit}
                      disabled={busy || preparingDocumentFilter}
                      onChange={(event) =>
                        setDocumentFilterSizeUnit(event.target.value as 'KB' | 'MB' | 'GB')
                      }
                    >
                      <option value="KB">KB</option>
                      <option value="MB">MB</option>
                      <option value="GB">GB</option>
                    </select>
                  </label>
                </>
              )}
              {documentFilterCriterion === 'pageSize' && (
                <label className="pdf-field">
                  <span>{text.filterFirstPageSize}</span>
                  <select
                    className="pdf-modal-input"
                    value={documentFilterPageSize}
                    disabled={busy || preparingDocumentFilter}
                    onChange={(event) =>
                      setDocumentFilterPageSize(event.target.value as Exclude<PdfPageSize, 'KEEP'>)
                    }
                  >
                    {(['A3', 'A4', 'A5', 'LETTER', 'LEGAL', 'TABLOID'] as const).map((value) => (
                      <option key={value} value={value}>
                        {value}
                      </option>
                    ))}
                  </select>
                </label>
              )}
              {documentFilterCriterion === 'rotation' && (
                <label className="pdf-field">
                  <span>{text.filterFirstPageRotation}</span>
                  <select
                    className="pdf-modal-input"
                    value={documentFilterRotation}
                    disabled={busy || preparingDocumentFilter}
                    onChange={(event) =>
                      setDocumentFilterRotation(Number(event.target.value) as PdfPageFilterRotation)
                    }
                  >
                    {([0, 90, 180, 270] as const).map((value) => (
                      <option key={value} value={value}>
                        {value}°
                      </option>
                    ))}
                  </select>
                </label>
              )}
              {documentFilterCriterion !== 'text' && documentFilterCriterion !== 'image' && (
                <label className="pdf-field">
                  <span>{text.filterComparator}</span>
                  <select
                    className="pdf-modal-input"
                    value={documentFilterComparator}
                    disabled={busy || preparingDocumentFilter}
                    onChange={(event) => {
                      setDocumentFilterComparator(event.target.value as PdfPageFilterComparator)
                      setError('')
                    }}
                  >
                    <option value="less">{text.filterLess}</option>
                    <option value="equal">{text.filterEqual}</option>
                    <option value="greater">{text.filterGreater}</option>
                  </select>
                </label>
              )}
            </div>
            <div className="pdf-bookmark-empty">{text.filterDocumentsHint}</div>
          </div>
        )}

        {kind === 'bulkReplaceText' && (
          <div className="pdf-replacement-editor">
            <label className="pdf-field">
              <span>{text.pages}</span>
              <input
                className={`pdf-modal-input${error ? ' invalid' : ''}`}
                value={targetPages}
                placeholder={text.pagesPlaceholder}
                disabled={busy}
                onChange={(event) => {
                  setTargetPages(event.target.value)
                  setError('')
                }}
              />
            </label>
            <div className="pdf-bookmark-toolbar">
              <span>{text.replacementRules}</span>
              <button
                className="pdf-modal-btn"
                type="button"
                disabled={busy || replacementRules.length >= 50}
                onClick={() => {
                  setReplacementRules((current) => [...current, emptyTextReplacement()])
                  setError('')
                }}
              >
                + {text.addReplacement}
              </button>
            </div>
            <div className="pdf-replacement-list">
              {replacementRules.map((rule, index) => (
                <div className="pdf-replacement-row" key={index}>
                  <span className="pdf-merge-order">{index + 1}</span>
                  <div className="pdf-replacement-fields">
                    <label className="pdf-tools-field-column pdf-field">
                      <span>{text.findText}</span>
                      <textarea
                        className={`pdf-modal-input pdf-replacement-textarea${error && !rule.find ? ' invalid' : ''}`}
                        value={rule.find}
                        rows={2}
                        disabled={busy}
                        onChange={(event) => {
                          const find = event.target.value
                          setReplacementRules((current) =>
                            current.map((item, ruleIndex) =>
                              ruleIndex === index ? { ...item, find } : item,
                            ),
                          )
                          setError('')
                        }}
                      />
                    </label>
                    <label className="pdf-tools-field-column pdf-field">
                      <span>{text.replaceWith}</span>
                      <textarea
                        className="pdf-modal-input pdf-replacement-textarea"
                        value={rule.replace}
                        rows={2}
                        disabled={busy}
                        onChange={(event) => {
                          const replace = event.target.value
                          setReplacementRules((current) =>
                            current.map((item, ruleIndex) =>
                              ruleIndex === index ? { ...item, replace } : item,
                            ),
                          )
                          setError('')
                        }}
                      />
                    </label>
                  </div>
                  <button
                    className="pdf-bookmark-icon"
                    type="button"
                    title={text.deleteReplacement}
                    aria-label={text.deleteReplacement}
                    disabled={busy || replacementRules.length === 1}
                    onClick={() => {
                      setReplacementRules((current) =>
                        current.filter((_, ruleIndex) => ruleIndex !== index),
                      )
                      setError('')
                    }}
                  >
                    ×
                  </button>
                </div>
              ))}
            </div>
            <div className="pdf-tools-toggle-grid">
              <label className="pdf-field pdf-field-check">
                <span>{text.caseSensitive}</span>
                <input
                  type="checkbox"
                  checked={replacementCaseSensitive}
                  disabled={busy}
                  onChange={(event) => setReplacementCaseSensitive(event.target.checked)}
                />
              </label>
              <label className="pdf-field pdf-field-check">
                <span>{text.wholeWord}</span>
                <input
                  type="checkbox"
                  checked={replacementWholeWord}
                  disabled={busy}
                  onChange={(event) => setReplacementWholeWord(event.target.checked)}
                />
              </label>
            </div>
            <div className="pdf-password-note">{text.emptyReplacementDeletes}</div>
            <div className="pdf-bookmark-empty">{text.bulkReplaceHint}</div>
          </div>
        )}

        {kind === 'attachments' && (
          <>
            <label className="pdf-field">
              <span>{text.attachmentAction}</span>
              <select
                className="pdf-modal-input"
                value={attachmentAction}
                disabled={busy || preparingAttachments}
                onChange={(event) => {
                  setAttachmentAction(event.target.value as PdfAttachmentAction)
                  setError('')
                }}
              >
                <option value="add">{text.addAttachments}</option>
                <option value="extract">{text.extractAttachments}</option>
                <option value="rename">{text.renameAttachment}</option>
                <option value="delete">{text.deleteAttachment}</option>
              </select>
            </label>
            {attachmentAction === 'add' ? (
              <div className="pdf-field">
                <span>{text.attachmentFiles}</span>
                <div className="pdf-tools-file-control">
                  <label className="pdf-modal-btn pdf-tools-file-button">
                    <input
                      type="file"
                      multiple
                      disabled={busy || preparingAttachments}
                      onChange={(event) => {
                        setAttachmentFiles(Array.from(event.currentTarget.files ?? []))
                        setError('')
                      }}
                    />
                    {text.chooseAttachments}
                  </label>
                  <span
                    className="pdf-tools-file-summary"
                    title={attachmentFiles.map((file) => file.name).join(', ')}
                  >
                    {attachmentFiles.length > 0
                      ? attachmentFiles.map((file) => file.name).join(', ')
                      : text.noFiles}
                  </span>
                </div>
              </div>
            ) : (
              <label className="pdf-field">
                <span>{text.existingAttachments}</span>
                <select
                  className="pdf-modal-input"
                  value={selectedAttachment}
                  disabled={busy || preparingAttachments || loadingAttachments}
                  onChange={(event) => {
                    setSelectedAttachment(event.target.value)
                    setError('')
                  }}
                >
                  {loadingAttachments ? (
                    <option value="">{text.loadingAttachments}</option>
                  ) : attachmentItems.length === 0 ? (
                    <option value="">{text.noAttachments}</option>
                  ) : (
                    attachmentItems.map((item) => (
                      <option key={item.name} value={item.name}>
                        {item.name} · {attachmentSizeLabel(item.size)}
                      </option>
                    ))
                  )}
                </select>
              </label>
            )}
            {attachmentAction === 'rename' && (
              <label className="pdf-field">
                <span>{text.newAttachmentName}</span>
                <input
                  className={`pdf-modal-input${error ? ' invalid' : ''}`}
                  value={newAttachmentName}
                  disabled={busy || preparingAttachments || loadingAttachments}
                  onChange={(event) => {
                    setNewAttachmentName(event.target.value)
                    setError('')
                  }}
                />
              </label>
            )}
          </>
        )}

        {kind === 'bookmarks' && (
          <div className="pdf-bookmark-editor">
            <div className="pdf-bookmark-toolbar">
              <span>
                {text.bookmarks} · {bookmarkTreeCount(bookmarks)}/{MAX_BOOKMARK_COUNT}
              </span>
              <button
                className="pdf-modal-btn"
                type="button"
                disabled={
                  busy || loadingBookmarks || bookmarkTreeCount(bookmarks) >= MAX_BOOKMARK_COUNT
                }
                onClick={() => {
                  setBookmarks((current) => [
                    ...current,
                    { title: '', pageNumber: 1, children: [] },
                  ])
                  setBookmarkNotice('')
                  setError('')
                }}
              >
                + {text.addBookmark}
              </button>
            </div>
            <div className="pdf-bookmark-exchange">
              <label className="pdf-field pdf-bookmark-import-mode">
                <span>{text.bookmarkImportMode}</span>
                <span className="pdf-tools-segments">
                  <button
                    type="button"
                    className={`pdf-sign-tab${bookmarkImportMode === 'replace' ? ' active' : ''}`}
                    disabled={busy || loadingBookmarks}
                    onClick={() => setBookmarkImportMode('replace')}
                  >
                    {text.bookmarkReplace}
                  </button>
                  <button
                    type="button"
                    className={`pdf-sign-tab${bookmarkImportMode === 'append' ? ' active' : ''}`}
                    disabled={busy || loadingBookmarks}
                    onClick={() => setBookmarkImportMode('append')}
                  >
                    {text.bookmarkAppend}
                  </button>
                </span>
              </label>
              <input
                ref={bookmarkJsonInputRef}
                className="pdf-hidden-input"
                type="file"
                accept="application/json,.json"
                disabled={busy || loadingBookmarks}
                onChange={(event) => {
                  void importBookmarkFile(event.target.files?.[0])
                  event.target.value = ''
                }}
              />
              <div className="pdf-bookmark-exchange-actions">
                <button
                  type="button"
                  className="pdf-modal-btn"
                  disabled={busy || loadingBookmarks}
                  onClick={() => bookmarkJsonInputRef.current?.click()}
                >
                  {text.importBookmarkJson}
                </button>
                <button
                  type="button"
                  className="pdf-modal-btn"
                  disabled={busy || loadingBookmarks}
                  onClick={() => void pasteBookmarkJson()}
                >
                  {text.pasteBookmarkJson}
                </button>
                <button
                  type="button"
                  className="pdf-modal-btn"
                  disabled={busy || loadingBookmarks || bookmarks.length === 0}
                  onClick={() =>
                    downloadPdfToolText(
                      'bookmarks.json',
                      bookmarkJsonText(bookmarks),
                      'application/json',
                    )
                  }
                >
                  {text.downloadBookmarkJson}
                </button>
                <button
                  type="button"
                  className="pdf-modal-btn"
                  disabled={busy || loadingBookmarks || bookmarks.length === 0}
                  onClick={() => void copyBookmarkJson()}
                >
                  {text.copyBookmarkJson}
                </button>
              </div>
              <div className="pdf-sign-hint">{text.bookmarkJsonHint}</div>
              {bookmarkNotice && <div className="pdf-bookmark-notice">{bookmarkNotice}</div>}
            </div>
            <div className="pdf-bookmark-list">
              {loadingBookmarks ? (
                <div className="pdf-bookmark-empty">{text.loadingBookmarks}</div>
              ) : bookmarks.length === 0 ? (
                <div className="pdf-bookmark-empty">{text.noBookmarks}</div>
              ) : (
                renderBookmarkRows(bookmarks)
              )}
            </div>
          </div>
        )}

        {kind === 'metadata' && (
          <div className="pdf-metadata-editor">
            {loadingMetadata ? (
              <div className="pdf-bookmark-empty">{text.loadingAnalysis}</div>
            ) : (
              <>
                <label className="pdf-field pdf-field-check pdf-metadata-clear">
                  <input
                    type="checkbox"
                    checked={clearAllMetadata}
                    disabled={busy}
                    onChange={(event) => {
                      setClearAllMetadata(event.target.checked)
                      setError('')
                    }}
                  />
                  <span>{text.deleteAllMetadata}</span>
                </label>

                <div className="pdf-metadata-fields">
                  {METADATA_TEXT_FIELDS.map(([field, label]) => (
                    <label className="pdf-field pdf-tools-field-column" key={field}>
                      <span>{text[label]}</span>
                      <input
                        className="pdf-modal-input"
                        value={metadataValues[field]}
                        disabled={busy || clearAllMetadata}
                        onChange={(event) => {
                          setMetadataValues((current) => ({
                            ...current,
                            [field]: event.target.value,
                          }))
                          setError('')
                        }}
                      />
                    </label>
                  ))}
                  <label className="pdf-field pdf-tools-field-column">
                    <span>{text.creationDate}</span>
                    <input
                      className="pdf-modal-input"
                      type="datetime-local"
                      value={metadataValues.creationDate}
                      disabled={busy || clearAllMetadata}
                      onChange={(event) => {
                        setMetadataValues((current) => ({
                          ...current,
                          creationDate: event.target.value,
                        }))
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field pdf-tools-field-column">
                    <span>{text.modificationDate}</span>
                    <input
                      className="pdf-modal-input"
                      type="datetime-local"
                      value={metadataValues.modificationDate}
                      disabled={busy || clearAllMetadata}
                      onChange={(event) => {
                        setMetadataValues((current) => ({
                          ...current,
                          modificationDate: event.target.value,
                        }))
                        setError('')
                      }}
                    />
                  </label>
                  <label className="pdf-field pdf-tools-field-column">
                    <span>{text.trapped}</span>
                    <select
                      className="pdf-modal-input"
                      value={metadataValues.trapped}
                      disabled={busy || clearAllMetadata}
                      onChange={(event) => {
                        setMetadataValues((current) => ({
                          ...current,
                          trapped: event.target.value as PdfMetadataTrapped,
                        }))
                        setError('')
                      }}
                    >
                      <option value="">{text.none}</option>
                      <option value="True">{text.trappedTrue}</option>
                      <option value="False">{text.trappedFalse}</option>
                      <option value="Unknown">{text.trappedUnknown}</option>
                    </select>
                  </label>
                </div>

                <div className="pdf-metadata-custom">
                  <div className="pdf-bookmark-toolbar">
                    <span>{text.customMetadata}</span>
                    <button
                      className="pdf-modal-btn"
                      type="button"
                      disabled={busy || clearAllMetadata}
                      onClick={() => {
                        setMetadataValues((current) => ({
                          ...current,
                          custom: [...current.custom, { key: '', value: '' }],
                        }))
                        setError('')
                      }}
                    >
                      + {text.addCustomMetadata}
                    </button>
                  </div>
                  {metadataValues.custom.length > 0 && (
                    <div className="pdf-metadata-custom-list">
                      {metadataValues.custom.map((field, index) => (
                        <div className="pdf-metadata-custom-row" key={index}>
                          <input
                            className="pdf-modal-input"
                            aria-label={`${text.customMetadataKey} ${index + 1}`}
                            placeholder={text.customMetadataKey}
                            value={field.key}
                            disabled={busy || clearAllMetadata}
                            onChange={(event) => {
                              setMetadataValues((current) => ({
                                ...current,
                                custom: updateCustomMetadataField(current.custom, index, {
                                  key: event.target.value,
                                }),
                              }))
                              setError('')
                            }}
                          />
                          <input
                            className="pdf-modal-input"
                            aria-label={`${text.customMetadataValue} ${index + 1}`}
                            placeholder={text.customMetadataValue}
                            value={field.value}
                            disabled={busy || clearAllMetadata}
                            onChange={(event) => {
                              setMetadataValues((current) => ({
                                ...current,
                                custom: updateCustomMetadataField(current.custom, index, {
                                  value: event.target.value,
                                }),
                              }))
                              setError('')
                            }}
                          />
                          <button
                            className="pdf-bookmark-icon"
                            type="button"
                            title={text.deleteAttachment}
                            aria-label={`${text.deleteAttachment} ${index + 1}`}
                            disabled={busy || clearAllMetadata}
                            onClick={() => {
                              setMetadataValues((current) => ({
                                ...current,
                                custom: current.custom.filter(
                                  (_, fieldIndex) => fieldIndex !== index,
                                ),
                              }))
                              setError('')
                            }}
                          >
                            ×
                          </button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              </>
            )}
          </div>
        )}

        {kind === 'info' && (
          <div className="pdf-analysis">
            {loadingAnalysis ? (
              <div className="pdf-bookmark-empty">{text.loadingAnalysis}</div>
            ) : analysis ? (
              <>
                <section className="pdf-analysis-section">
                  <h3>{text.overview}</h3>
                  <dl className="pdf-analysis-grid">
                    <dt>{text.pageCount}</dt>
                    <dd>{analysis.pageCount}</dd>
                    <dt>{text.pdfVersion}</dt>
                    <dd>{analysis.pdfVersion ?? text.none}</dd>
                    <dt>{text.fileSize}</dt>
                    <dd>{attachmentSizeLabel(analysis.fileSize)}</dd>
                    <dt>{text.encrypted}</dt>
                    <dd>{analysis.isEncrypted ? text.yes : text.no}</dd>
                  </dl>
                </section>

                <section className="pdf-analysis-section">
                  <h3>{text.documentProperties}</h3>
                  <dl className="pdf-analysis-grid">
                    {analysis.properties.title && (
                      <>
                        <dt>{text.documentTitle}</dt>
                        <dd>{analysis.properties.title}</dd>
                      </>
                    )}
                    {analysis.properties.author && (
                      <>
                        <dt>{text.author}</dt>
                        <dd>{analysis.properties.author}</dd>
                      </>
                    )}
                    {analysis.properties.subject && (
                      <>
                        <dt>{text.subject}</dt>
                        <dd>{analysis.properties.subject}</dd>
                      </>
                    )}
                    {analysis.properties.keywords && (
                      <>
                        <dt>{text.keywords}</dt>
                        <dd>{analysis.properties.keywords}</dd>
                      </>
                    )}
                    {analysis.properties.creator && (
                      <>
                        <dt>{text.creator}</dt>
                        <dd>{analysis.properties.creator}</dd>
                      </>
                    )}
                    {analysis.properties.producer && (
                      <>
                        <dt>{text.producer}</dt>
                        <dd>{analysis.properties.producer}</dd>
                      </>
                    )}
                    {analysis.properties.creationDate && (
                      <>
                        <dt>{text.creationDate}</dt>
                        <dd>{analysisDateLabel(analysis.properties.creationDate, lang)}</dd>
                      </>
                    )}
                    {analysis.properties.modificationDate && (
                      <>
                        <dt>{text.modificationDate}</dt>
                        <dd>{analysisDateLabel(analysis.properties.modificationDate, lang)}</dd>
                      </>
                    )}
                    {analysis.properties.trapped && (
                      <>
                        <dt>{text.trapped}</dt>
                        <dd>{text[`trapped${analysis.properties.trapped}`]}</dd>
                      </>
                    )}
                    {analysis.properties.custom.map((field) => (
                      <div className="pdf-analysis-custom-property" key={field.key}>
                        <dt>{field.key}</dt>
                        <dd>{field.value || text.none}</dd>
                      </div>
                    ))}
                    {!hasPdfProperties(analysis.properties) && (
                      <>
                        <dt>{text.documentProperties}</dt>
                        <dd>{text.none}</dd>
                      </>
                    )}
                  </dl>
                </section>

                <section className="pdf-analysis-section pdf-analysis-columns">
                  <div>
                    <h3>{text.resources}</h3>
                    <dl className="pdf-analysis-grid">
                      <dt>{text.fonts}</dt>
                      <dd>{analysis.fonts.join(', ') || text.none}</dd>
                      <dt>{text.images}</dt>
                      <dd>{analysis.imageCount}</dd>
                      <dt>{text.embeddedFiles}</dt>
                      <dd>{analysis.attachmentCount}</dd>
                      <dt>{text.bookmarkCount}</dt>
                      <dd>{analysis.bookmarkCount}</dd>
                    </dl>
                  </div>
                  <div>
                    <h3>{text.interactiveContent}</h3>
                    <dl className="pdf-analysis-grid">
                      <dt>{text.formFields}</dt>
                      <dd>{analysis.form.fieldCount}</dd>
                      <dt>{text.signatures}</dt>
                      <dd>{analysis.form.signatureCount}</dd>
                      <dt>{text.xfa}</dt>
                      <dd>{analysis.form.hasXfa ? text.yes : text.no}</dd>
                      <dt>{text.annotations}</dt>
                      <dd>{analysis.annotations.totalCount}</dd>
                      <dt>{text.annotationTypes}</dt>
                      <dd>
                        {Object.entries(analysis.annotations.typeBreakdown)
                          .map(([type, count]) => `${type} ${count}`)
                          .join(', ') || text.none}
                      </dd>
                    </dl>
                  </div>
                </section>

                <section className="pdf-analysis-section">
                  <h3>{text.pageDetails}</h3>
                  <div className="pdf-analysis-table-wrap">
                    <table className="pdf-analysis-table">
                      <thead>
                        <tr>
                          <th>{text.bookmarkPage}</th>
                          <th>{text.dimensions}</th>
                          <th>{text.rotation}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {analysis.pages.map((page) => (
                          <tr key={page.pageNumber}>
                            <td>{page.pageNumber}</td>
                            <td>
                              {page.width.toFixed(1)} × {page.height.toFixed(1)} pt
                            </td>
                            <td>{page.rotation}°</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              </>
            ) : null}
          </div>
        )}

        {kind === 'fontInfo' && (
          <div className="pdf-analysis">
            {loadingFonts ? (
              <div className="pdf-bookmark-empty">{text.loadingAnalysis}</div>
            ) : fontReport ? (
              <>
                <section className="pdf-analysis-section">
                  <h3>{text.fontSummary}</h3>
                  <dl className="pdf-analysis-grid">
                    <dt>{text.fonts}</dt>
                    <dd>{fontReport.fontCount}</dd>
                    <dt>{text.embeddedFonts}</dt>
                    <dd>{fontReport.embeddedCount}</dd>
                    <dt>{text.subsetFonts}</dt>
                    <dd>{fontReport.subsetCount}</dd>
                  </dl>
                </section>
                {fontReport.fonts.length === 0 ? (
                  <div className="pdf-bookmark-empty">{text.noFonts}</div>
                ) : (
                  <section className="pdf-analysis-section">
                    <div className="pdf-analysis-table-wrap pdf-font-table-wrap">
                      <table className="pdf-analysis-table pdf-font-table">
                        <thead>
                          <tr>
                            <th>{text.fontName}</th>
                            <th>{text.fontSubtype}</th>
                            <th>{text.fontEncoding}</th>
                            <th>{text.fontEmbedded}</th>
                            <th>{text.fontSubset}</th>
                            <th>{text.fontToUnicode}</th>
                            <th>{text.fontPages}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {fontReport.fonts.map((font) => (
                            <tr
                              key={`${font.name}:${font.subtype}:${font.encoding ?? ''}:${font.pages.join(',')}`}
                            >
                              <td>{font.name}</td>
                              <td>{font.subtype}</td>
                              <td>{font.encoding ?? text.none}</td>
                              <td>{font.embedded ? text.yes : text.no}</td>
                              <td>{font.subset ? text.yes : text.no}</td>
                              <td>{font.hasToUnicode ? text.yes : text.no}</td>
                              <td>{font.pages.join(', ')}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </section>
                )}
              </>
            ) : null}
          </div>
        )}

        {kind === 'annotationInfo' && (
          <div className="pdf-analysis">
            {loadingAnnotations ? (
              <div className="pdf-bookmark-empty">{text.loadingAnalysis}</div>
            ) : annotationReport ? (
              <>
                <section className="pdf-analysis-section">
                  <h3>{text.annotationSummary}</h3>
                  <dl className="pdf-analysis-grid">
                    <dt>{text.annotations}</dt>
                    <dd>{annotationReport.totalCount}</dd>
                    <dt>{text.annotationTypes}</dt>
                    <dd>
                      {Object.entries(annotationReport.typeBreakdown)
                        .sort(([left], [right]) => left.localeCompare(right))
                        .map(([type, count]) => `${type} ${count}`)
                        .join(', ') || text.none}
                    </dd>
                  </dl>
                </section>
                {annotationReport.annotations.length === 0 ? (
                  <div className="pdf-bookmark-empty">{text.noAnnotations}</div>
                ) : (
                  <section className="pdf-analysis-section">
                    <div className="pdf-analysis-table-wrap pdf-annotation-table-wrap">
                      <table className="pdf-analysis-table pdf-annotation-table">
                        <thead>
                          <tr>
                            <th>{text.bookmarkPage}</th>
                            <th>{text.annotationIndex}</th>
                            <th>{text.annotationSubtype}</th>
                            <th>{text.author}</th>
                            <th>{text.subject}</th>
                            <th>{text.annotationContents}</th>
                            <th>{text.annotationModified}</th>
                            <th>{text.annotationPosition}</th>
                            <th>{text.annotationFlags}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {annotationReport.annotations.map((annotation) => (
                            <tr
                              key={`${annotation.pageNumber}:${annotation.annotationNumber}:${annotation.name ?? ''}`}
                            >
                              <td>{annotation.pageNumber}</td>
                              <td>{annotation.annotationNumber}</td>
                              <td>{annotation.subtype}</td>
                              <td>{annotation.author ?? text.none}</td>
                              <td>{annotation.subject ?? text.none}</td>
                              <td className="pdf-annotation-content">
                                {annotation.contents ?? text.none}
                              </td>
                              <td>
                                {annotation.modifiedAt
                                  ? analysisDateLabel(annotation.modifiedAt, lang)
                                  : text.none}
                              </td>
                              <td>
                                {annotation.rectangle
                                  ? `${annotation.rectangle.x.toFixed(1)}, ${annotation.rectangle.y.toFixed(1)} · ${annotation.rectangle.width.toFixed(1)} × ${annotation.rectangle.height.toFixed(1)} pt`
                                  : text.none}
                              </td>
                              <td>{annotation.flags ?? text.none}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </section>
                )}
              </>
            ) : null}
          </div>
        )}

        {kind === 'securityInfo' && (
          <div className="pdf-analysis">
            {loadingSecurity ? (
              <div className="pdf-bookmark-empty">{text.loadingAnalysis}</div>
            ) : securityReport ? (
              <>
                <section className="pdf-analysis-section">
                  <h3>{text.securitySummary}</h3>
                  <dl className="pdf-analysis-grid">
                    <dt>{text.encrypted}</dt>
                    <dd>{securityReport.encrypted ? text.yes : text.no}</dd>
                    <dt>{text.encryptionAlgorithm}</dt>
                    <dd>
                      {securityReport.encrypted
                        ? (securityReport.algorithm ?? text.none)
                        : text.none}
                    </dd>
                    <dt>{text.keyLength}</dt>
                    <dd>
                      {securityReport.keyLength ? `${securityReport.keyLength} bit` : text.none}
                    </dd>
                    <dt>{text.securityVersion}</dt>
                    <dd>
                      {securityReport.version !== undefined || securityReport.revision !== undefined
                        ? `V${securityReport.version ?? '?'} / R${securityReport.revision ?? '?'}`
                        : text.none}
                    </dd>
                    <dt>{text.restrictedActions}</dt>
                    <dd>
                      {securityReport.permissions ? securityReport.restrictedCount : text.none}
                    </dd>
                  </dl>
                </section>
                <section className="pdf-analysis-section">
                  <h3>{text.pdfPermissions}</h3>
                  {securityReport.permissions ? (
                    <div className="pdf-security-permissions">
                      {PDF_PERMISSION_KEYS.map((permission) => {
                        const allowed = securityReport.permissions?.[permission] ?? false
                        return (
                          <div className="pdf-security-permission" key={permission}>
                            <span>{text[permission]}</span>
                            <span
                              className={`pdf-security-permission-status ${allowed ? 'allowed' : 'restricted'}`}
                            >
                              {allowed ? text.permissionAllowed : text.permissionRestricted}
                            </span>
                          </div>
                        )
                      })}
                    </div>
                  ) : (
                    <div className="pdf-bookmark-empty">{text.permissionUnavailable}</div>
                  )}
                  <div className="pdf-password-note">{text.permissionAdvisory}</div>
                </section>
              </>
            ) : null}
          </div>
        )}

        {kind === 'preflight' && (
          <div className="pdf-preflight">
            {loadingPreflight ? (
              <div className="pdf-bookmark-empty">{text.loadingAnalysis}</div>
            ) : preflightReport ? (
              <>
                <div className="pdf-preflight-summary">
                  <div>
                    <span className={`pdf-preflight-status ${preflightReport.status}`}>
                      {preflightReport.status === 'pass'
                        ? text.preflightPass
                        : preflightReport.status === 'warning'
                          ? text.preflightWarning
                          : text.preflightError}
                    </span>
                    <p>{text.preflightDisclaimer}</p>
                  </div>
                  <div className="pdf-preflight-actions">
                    {preflightReport.status === 'error' && (
                      <button
                        className="pdf-modal-btn"
                        type="button"
                        onClick={() => {
                          setKind('repair')
                          setError('')
                        }}
                      >
                        {text.reviewRepair}
                      </button>
                    )}
                    <button
                      className="pdf-modal-btn"
                      type="button"
                      onClick={() =>
                        downloadPdfToolText(
                          'pdf-preflight-report.json',
                          `${JSON.stringify(preflightReport, null, 2)}\n`,
                          'application/json',
                        )
                      }
                    >
                      {text.downloadPreflightJson}
                    </button>
                  </div>
                </div>

                <section className="pdf-analysis-section pdf-analysis-columns">
                  <div>
                    <h3>{text.structure}</h3>
                    <dl className="pdf-analysis-grid">
                      <dt>{text.pdfVersion}</dt>
                      <dd>{preflightReport.pdfVersion ?? text.none}</dd>
                      <dt>{text.fileSize}</dt>
                      <dd>{attachmentSizeLabel(preflightReport.fileSize)}</dd>
                      <dt>{text.pageCount}</dt>
                      <dd>{preflightReport.pageCount}</dd>
                      <dt>{text.strictParsing}</dt>
                      <dd>{preflightReport.strictParsing ? text.yes : text.no}</dd>
                      <dt>{text.headerOffset}</dt>
                      <dd>{preflightReport.structure.headerOffset}</dd>
                      <dt>{text.eofMarkers}</dt>
                      <dd>{preflightReport.structure.eofMarkerCount}</dd>
                      <dt>{text.trailingBytes}</dt>
                      <dd>{attachmentSizeLabel(preflightReport.structure.trailingBytes)}</dd>
                      <dt>{text.startXref}</dt>
                      <dd>{preflightReport.structure.startXrefOffset ?? text.none}</dd>
                      <dt>{text.incrementalUpdates}</dt>
                      <dd>{preflightReport.structure.incrementalUpdates}</dd>
                    </dl>
                  </div>
                  <div>
                    <h3>{text.standardDeclarations}</h3>
                    <dl className="pdf-analysis-grid">
                      <dt>PDF/A · PDF/UA</dt>
                      <dd>
                        {preflightReport.standards.map((standard) => standard.label).join(', ') ||
                          text.noStandardDeclarations}
                      </dd>
                      <dt>{text.xmpMetadata}</dt>
                      <dd>{preflightReport.features.hasXmpMetadata ? text.yes : text.no}</dd>
                      <dt>{text.xmpValid}</dt>
                      <dd>{preflightReport.features.xmpValid ? text.yes : text.no}</dd>
                      <dt>{text.taggedPdf}</dt>
                      <dd>{preflightReport.features.tagged ? text.yes : text.no}</dd>
                      <dt>{text.documentLanguage}</dt>
                      <dd>{preflightReport.features.language ?? text.none}</dd>
                      <dt>{text.outputIntents}</dt>
                      <dd>{preflightReport.features.outputIntentCount}</dd>
                    </dl>
                  </div>
                </section>

                <section className="pdf-analysis-section">
                  <h3>{text.interactiveContent}</h3>
                  <dl className="pdf-analysis-grid pdf-preflight-features">
                    <dt>{text.javaScriptActions}</dt>
                    <dd>{preflightReport.features.javaScriptActionCount}</dd>
                    <dt>{text.embeddedFiles}</dt>
                    <dd>{preflightReport.features.attachmentCount}</dd>
                    <dt>{text.formFields}</dt>
                    <dd>{preflightReport.features.formFieldCount}</dd>
                    <dt>{text.xfa}</dt>
                    <dd>{preflightReport.features.hasXfa ? text.yes : text.no}</dd>
                    <dt>{text.signatures}</dt>
                    <dd>{preflightReport.features.signatureCount}</dd>
                    <dt>{text.encrypted}</dt>
                    <dd>{preflightReport.features.encrypted ? text.yes : text.no}</dd>
                  </dl>
                </section>

                <section className="pdf-analysis-section">
                  <h3>{text.findings}</h3>
                  {preflightReport.findings.length === 0 ? (
                    <div className="pdf-bookmark-empty">{text.noPreflightFindings}</div>
                  ) : (
                    <div className="pdf-preflight-findings">
                      {preflightReport.findings.map((finding, index) => (
                        <div
                          className={`pdf-preflight-finding ${finding.severity}`}
                          key={`${finding.code}-${finding.pageNumber ?? 0}-${index}`}
                        >
                          <span>{text[PREFLIGHT_FINDING_TEXT_KEYS[finding.code]]}</span>
                          {(finding.pageNumber || finding.detail) && (
                            <small>
                              {finding.pageNumber
                                ? `${text.bookmarkPage} ${finding.pageNumber}${finding.detail ? ' · ' : ''}`
                                : ''}
                              {finding.detail?.slice(0, 300)}
                            </small>
                          )}
                        </div>
                      ))}
                    </div>
                  )}
                </section>
              </>
            ) : null}
          </div>
        )}

        {kind === 'javascriptAudit' && (
          <div className="pdf-javascript-audit">
            {loadingJavaScript ? (
              <div className="pdf-bookmark-empty">{text.loadingJavaScript}</div>
            ) : javaScriptAudit ? (
              <>
                <div className="pdf-javascript-summary">
                  <dl className="pdf-analysis-grid">
                    <dt>{text.javaScriptActions}</dt>
                    <dd>{javaScriptAudit.actions.length}</dd>
                    <dt>{text.uniqueScripts}</dt>
                    <dd>{javaScriptAudit.uniqueScriptCount}</dd>
                    <dt>{text.scriptCodeSize}</dt>
                    <dd>{attachmentSizeLabel(javaScriptAudit.totalCodeBytes)}</dd>
                  </dl>
                  <div className="pdf-javascript-summary-actions">
                    <button
                      className="pdf-modal-btn"
                      type="button"
                      disabled={javaScriptAudit.actions.length === 0}
                      onClick={() => {
                        const value = javaScriptAudit.actions.map(javaScriptActionText).join('\n\n')
                        void copyPdfToolText(value).then((copied) => {
                          if (!copied) return
                          setCopiedScript('all')
                          window.setTimeout(
                            () =>
                              setCopiedScript((current) => (current === 'all' ? null : current)),
                            1200,
                          )
                        })
                      }}
                    >
                      {copiedScript === 'all' ? text.copied : text.copyAll}
                    </button>
                    <button
                      className="pdf-modal-btn"
                      type="button"
                      disabled={javaScriptAudit.actions.length === 0}
                      onClick={() => {
                        setSanitizeOptions({
                          removeJavaScript: true,
                          removeEmbeddedFiles: false,
                          removeXmpMetadata: false,
                          removeMetadata: false,
                          removeLinks: false,
                        })
                        setKind('sanitize')
                        setError('')
                      }}
                    >
                      {text.reviewRemoval}
                    </button>
                  </div>
                </div>
                {javaScriptAudit.actions.length === 0 ? (
                  <div className="pdf-bookmark-empty">{text.noJavaScript}</div>
                ) : (
                  <div className="pdf-javascript-list">
                    {javaScriptAudit.actions.map((action, index) => (
                      <section
                        className="pdf-javascript-action"
                        key={`${action.source}-${action.trigger}-${index}`}
                      >
                        <div className="pdf-javascript-action-heading">
                          <div>
                            <strong>{javaScriptSourceLabels[action.source]}</strong>
                            <span>
                              {text.trigger}: {action.trigger}
                              {action.name ? ` · ${text.scriptName}: ${action.name}` : ''}
                              {action.pageNumber
                                ? ` · ${text.bookmarkPage}: ${action.pageNumber}`
                                : ''}
                              {action.annotationNumber
                                ? ` · ${text.annotationNumber}: ${action.annotationNumber}`
                                : ''}
                              {action.fieldName ? ` · ${text.field}: ${action.fieldName}` : ''}
                            </span>
                          </div>
                          <button
                            className="pdf-modal-btn pdf-javascript-copy"
                            type="button"
                            onClick={() => {
                              void copyPdfToolText(action.code).then((copied) => {
                                if (!copied) return
                                setCopiedScript(index)
                                window.setTimeout(
                                  () =>
                                    setCopiedScript((current) =>
                                      current === index ? null : current,
                                    ),
                                  1200,
                                )
                              })
                            }}
                          >
                            {copiedScript === index ? text.copied : text.copy}
                          </button>
                        </div>
                        <pre>{action.code}</pre>
                      </section>
                    ))}
                  </div>
                )}
              </>
            ) : null}
          </div>
        )}

        {kind === 'signatureAudit' && (
          <div className="pdf-signature-audit">
            {loadingSignatures ? (
              <div className="pdf-bookmark-empty">{text.loadingSignatures}</div>
            ) : signatureAudit ? (
              <>
                <div className="pdf-signature-summary">
                  <dl className="pdf-analysis-grid">
                    <dt>{text.signatures}</dt>
                    <dd>{signatureAudit.signatures.length}</dd>
                    <dt>{text.validSignatures}</dt>
                    <dd>{signatureAudit.validCount}</dd>
                    <dt>{text.invalidSignatures}</dt>
                    <dd>{signatureAudit.invalidCount}</dd>
                    <dt>{text.unsupportedSignatures}</dt>
                    <dd>{signatureAudit.unsupportedCount}</dd>
                  </dl>
                  <p>{text.trustNotEvaluated}</p>
                </div>
                {signatureAudit.signatures.length === 0 ? (
                  <div className="pdf-bookmark-empty">{text.noSignatures}</div>
                ) : (
                  <div className="pdf-signature-list">
                    {signatureAudit.signatures.map((signature, index) => (
                      <section
                        className="pdf-signature-entry"
                        key={`${signature.fieldName}-${index}`}
                      >
                        <div className="pdf-signature-heading">
                          <div>
                            <strong>{signature.signerName ?? signature.fieldName}</strong>
                            <span>{signature.fieldName}</span>
                          </div>
                          <span className={`pdf-signature-status ${signature.cryptographicStatus}`}>
                            {signatureStatusLabels[signature.cryptographicStatus]}
                          </span>
                        </div>
                        <dl className="pdf-analysis-grid">
                          <dt>{text.signatureType}</dt>
                          <dd>{signatureTypeLabels[signature.type]}</dd>
                          <dt>{text.signatureCoverage}</dt>
                          <dd>
                            {signature.coversEntireDocument
                              ? text.entireDocument
                              : text.earlierRevision}
                            {signature.unsignedTailBytes > 0
                              ? ` · ${text.unsignedTail} ${attachmentSizeLabel(signature.unsignedTailBytes)}`
                              : ''}
                          </dd>
                          {signature.claimedSigningTime && (
                            <>
                              <dt>{text.signedAt}</dt>
                              <dd>{analysisDateLabel(signature.claimedSigningTime, lang)}</dd>
                            </>
                          )}
                          {signature.timestampTime && (
                            <>
                              <dt>{text.timestampAt}</dt>
                              <dd>{analysisDateLabel(signature.timestampTime, lang)}</dd>
                            </>
                          )}
                          {signature.reason && (
                            <>
                              <dt>{text.reason}</dt>
                              <dd>{signature.reason}</dd>
                            </>
                          )}
                          {signature.location && (
                            <>
                              <dt>{text.location}</dt>
                              <dd>{signature.location}</dd>
                            </>
                          )}
                          {signature.digestAlgorithm && (
                            <>
                              <dt>{text.digestAlgorithm}</dt>
                              <dd>{signature.digestAlgorithm}</dd>
                            </>
                          )}
                          {signature.signatureAlgorithm && (
                            <>
                              <dt>{text.signatureAlgorithm}</dt>
                              <dd>{signature.signatureAlgorithm}</dd>
                            </>
                          )}
                          {signature.certificate && (
                            <>
                              <dt>{text.certificateSubject}</dt>
                              <dd>{signature.certificate.subject}</dd>
                              <dt>{text.certificateIssuer}</dt>
                              <dd>{signature.certificate.issuer}</dd>
                              <dt>{text.certificateSerial}</dt>
                              <dd>{signature.certificate.serialNumber}</dd>
                              <dt>{text.certificateValidFrom}</dt>
                              <dd>{analysisDateLabel(signature.certificate.validFrom, lang)}</dd>
                              <dt>{text.certificateValidUntil}</dt>
                              <dd>{analysisDateLabel(signature.certificate.validUntil, lang)}</dd>
                              {signature.certificate.validAtSigningTime !== undefined && (
                                <>
                                  <dt>{text.certificateValidAtSigning}</dt>
                                  <dd>
                                    {signature.certificate.validAtSigningTime ? text.yes : text.no}
                                  </dd>
                                </>
                              )}
                              <dt>{text.certificateSelfSigned}</dt>
                              <dd>{signature.certificate.selfSigned ? text.yes : text.no}</dd>
                            </>
                          )}
                        </dl>
                        {signature.error && (
                          <div className="pdf-signature-message error">{signature.error}</div>
                        )}
                        {signature.warnings.map((warning) => (
                          <div className="pdf-signature-message" key={warning}>
                            {warning}
                          </div>
                        ))}
                      </section>
                    ))}
                  </div>
                )}
              </>
            ) : null}
          </div>
        )}

        {(kind === 'nup' || (kind === 'scale' && pageSize !== 'KEEP')) && (
          <label className="pdf-field">
            <span>{text.orientation}</span>
            <span className="pdf-tools-segments">
              {(['portrait', 'landscape'] as const).map((value) => (
                <button
                  key={value}
                  className={`pdf-sign-tab${orientation === value ? ' active' : ''}`}
                  disabled={busy}
                  onClick={() => setOrientation(value)}
                >
                  {text[value]}
                </button>
              ))}
            </span>
          </label>
        )}

        {error && <div className="pdf-tools-error">{error}</div>}
        <div className="pdf-modal-actions">
          {kind === 'info' ||
          kind === 'fontInfo' ||
          kind === 'annotationInfo' ||
          kind === 'securityInfo' ||
          kind === 'preflight' ||
          kind === 'javascriptAudit' ||
          kind === 'signatureAudit' ? (
            <button className="pdf-modal-btn primary" onClick={onCancel}>
              {text.close}
            </button>
          ) : (
            <>
              <button
                className="pdf-modal-btn"
                disabled={
                  busy ||
                  preparingMerge ||
                  preparingImagesToPdf ||
                  preparingCbzToPdf ||
                  preparingEmailToPdf ||
                  preparingEpubToPdf ||
                  preparingHtmlToPdf ||
                  preparingMarkdownToPdf ||
                  preparingComparison ||
                  preparingOverlay ||
                  preparingOverlayImage ||
                  preparingDocumentFilter ||
                  preparingAttachments
                }
                onClick={onCancel}
              >
                {t('cancel')}
              </button>
              <button
                className="pdf-modal-btn primary"
                disabled={
                  busy ||
                  preparingMerge ||
                  preparingImagesToPdf ||
                  preparingCbzToPdf ||
                  preparingEmailToPdf ||
                  preparingEpubToPdf ||
                  preparingHtmlToPdf ||
                  preparingMarkdownToPdf ||
                  preparingComparison ||
                  preparingOverlay ||
                  preparingOverlayImage ||
                  preparingDocumentFilter ||
                  preparingAttachments ||
                  (kind === 'bookmarks' && loadingBookmarks) ||
                  (kind === 'metadata' && loadingMetadata) ||
                  (kind === 'attachments' && attachmentAction !== 'add' && loadingAttachments)
                }
                onClick={() => void apply()}
              >
                {busy ||
                preparingMerge ||
                preparingImagesToPdf ||
                preparingCbzToPdf ||
                preparingEmailToPdf ||
                preparingEpubToPdf ||
                preparingHtmlToPdf ||
                preparingMarkdownToPdf ||
                preparingComparison ||
                preparingOverlay ||
                preparingOverlayImage ||
                preparingDocumentFilter ||
                preparingAttachments
                  ? t('saving')
                  : kind === 'password'
                    ? encrypted
                      ? text.unlockPdf
                      : text.protectPdf
                    : kind === 'ocr'
                      ? text.runOcr
                      : kind === 'certificateSign'
                        ? text.runCertificateSign
                        : kind === 'timestamp'
                          ? text.runTimestamp
                          : t('ok')}
              </button>
            </>
          )}
        </div>
      </div>
    </div>,
    document.body,
  )
}
