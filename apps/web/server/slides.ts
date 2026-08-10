import { randomUUID } from 'node:crypto'

import {
  addChart,
  addElement,
  addMedia,
  addModel3d,
  addPicture,
  addSection,
  addSlideComment,
  addSmartArt,
  addTable,
  applyHeaderFooter,
  applyThemeToArchive,
  builtinLayoutInfos,
  BUILTIN_LAYOUT_PREFIX,
  commitSaved,
  copyElementData,
  copySlide,
  createBlankPptx,
  deleteElement,
  deleteSlide,
  deleteSlideComment,
  duplicateSlide,
  editChartElement,
  editGroupChildFill,
  editGroupChildStroke,
  editGroupChildTransform,
  editTableStyle,
  editPictureSrcRect,
  editTableCellText,
  editTableStructure,
  elementSpid,
  EMU_PER_PT,
  ensureBuiltinLayout,
  ensureTableStylePart,
  findGroupChild,
  getChartElementData,
  getElementLink,
  getSections,
  getSlideAnimations,
  getSlideComments,
  getSlideLinks,
  getSlideNotes,
  getSlideTransition,
  getRunLinks,
  groupElements,
  insertBlankSlide,
  insertSlideWithLayout,
  listSlideLayouts,
  listMasterParts,
  materializeSlide,
  mergeTableCells,
  moveSection,
  moveSlide,
  openPptx,
  parseTheme,
  parseMasterPart,
  patchSlideXml,
  patchGroupChildText,
  promoteSlideBackground,
  pasteElements,
  pasteSlide,
  readHeaderFooter,
  remapDeckColors,
  reorderElement,
  replaceAllInDeck,
  replacePictureBytes,
  resizeTable,
  reparseDeck,
  resetSlideLayout,
  removeSection,
  renameSection,
  savePptx,
  setElementConnection,
  setElementFont,
  setElementImageFill,
  setElementLink,
  setElementParagraphFormat,
  setElementTextAnchor,
  setGroupChildFont,
  setGroupChildParagraphFormat,
  setPictureOpacity,
  setSections,
  setSlideLayout,
  setSlideSize,
  setSlideAdvanceTime,
  setSlideAnimations,
  setSlideBackground,
  setSlideHidden,
  setSlideNotes,
  setSlideTransition,
  setTableCellAnchor,
  setTableColWidth,
  setTableRowHeight,
  shouldOfferBuiltinLayouts,
  TABLE_STYLE_PRESETS,
  ungroupElement,
  updateConnectorsForMoved,
  type ElementClipboardItem,
  type OpenedPptx,
  type Paragraph,
  type Slide,
  type SlideBundle,
  type SlideAnimation,
  type TextElement,
} from '@genoffice/pptx-engine'
import { buildRenderSlide, EMU_PER_PX_96, type RenderSlide } from '@genoffice/pptx-render'
import { applyEditParagraphs, collectParagraphFormatPatches } from '../../slides/src/main/edit-text'
import { tiffToPng } from '../../slides/src/main/tiff-decode'
import type {
  AddBlankSlideOp,
  AddChartOp,
  AddCommentOp,
  AddElementOp,
  AddImageBytesOp,
  AddInkOp,
  AddMediaBytesOp,
  AddSectionOp,
  AddSlideOp,
  AddSlideWithLayoutOp,
  AddSmartArtOp,
  AddTableOp,
  AnimationItem,
  ApplyThemeOp,
  BatchEditTransformOp,
  CopyElementsOp,
  DeleteCommentOp,
  DeleteElementOp,
  DuplicateElementsOp,
  EditBackgroundOp,
  EditChartOp,
  EditConnectorEndpointsOp,
  EditFillOp,
  EditPictureOpacityOp,
  EditPictureSrcRectOp,
  EditStrokeOp,
  EditTableCellOp,
  EditTableStyleOp,
  EditTextOp,
  EditTransformOp,
  FindReplaceOp,
  FlipElementOp,
  GroupElementsOp,
  HeaderFooterOp,
  MoveSectionOp,
  MoveSlideOp,
  MasterDeleteElementOp,
  MasterEditFillOp,
  MasterEditStrokeOp,
  MasterEditTextOp,
  MasterEditTransformOp,
  MasterEnterResult,
  OpenResult,
  PasteElementsOp,
  PasteSlideOp,
  RepasteSlideOp,
  ReorderElementOp,
  ReplacePictureBytesOp,
  RemoveSectionOp,
  RenameSectionOp,
  SectionInfo,
  SetAdvanceTimesOp,
  SetAnimationsOp,
  SetElementFontOp,
  SetElementParagraphFormatOp,
  SetLinkOp,
  SetNotesOp,
  SetSlideLayoutOp,
  SetSlideSizeOp,
  SetSlideHiddenOp,
  SetTableCellAnchorOp,
  SetTableColWidthOp,
  SetTableRowHeightOp,
  SetTransitionOp,
  TableMergeIpcOp,
  TableStructureIpcOp,
  UngroupElementOp,
} from '../../slides/src/shared/ipc'

interface HistorySnapshot {
  slides: Slide[]
  entries: Map<string, Uint8Array>
  size: { cx: number; cy: number }
  metaDirty: boolean
  name: string
  webPath: string
}

interface HistoryBatch {
  depth: number
  undoStart: number
  before: HistorySnapshot
}

interface SlidesSession {
  id: string
  name: string
  webPath: string
  opened: OpenedPptx
  fitWidthPx: number
  undo: HistorySnapshot[]
  redo: HistorySnapshot[]
  metaDirty: boolean
  transformPreview: boolean
  clipboard: { items: ElementClipboardItem[]; pasteCount: number } | null
  lastAccessAt: number
  masterEdit: { partPath: string; slide: Slide } | null
  historyBatch?: HistoryBatch
  aiSnapshots: Map<number, HistorySnapshot>
  lastSlidePaste: { afterIndex: number; undoLen: number } | null
}

interface OpenSlidesInput {
  name: string
  webPath?: string
  fitWidthPx: number
  pptxBase64: string
}

interface CallSlidesInput {
  sessionId: string
  action: string
  args?: unknown[]
}

const MAX_HISTORY = 50
const MAX_AI_SNAPSHOTS = 20
let nextAiSnapshotId = 1

function safeName(value: string): string {
  const name = value
    .split(/[\\/]/)
    .pop()
    ?.replace(/[\u0000-\u001f]/g, '')
    .trim()
  if (!name || name.length > 255) throw new Error('演示文稿名称无效')
  return name.toLowerCase().endsWith('.pptx') ? name : `${name}.pptx`
}

function decodePptx(value: string, maxBytes: number): Uint8Array {
  if (typeof value !== 'string') throw new Error('演示文稿内容无效')
  const bytes = Buffer.from(value, 'base64')
  if (bytes.length < 4 || bytes.length > maxBytes || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error('PPTX 文件无效或过大')
  }
  return new Uint8Array(bytes)
}

function mediaResolver(opened: OpenedPptx): (mediaRef: string) => string | undefined {
  const cache = new Map<string, string | undefined>()
  return (mediaRef) => {
    if (cache.has(mediaRef)) return cache.get(mediaRef)
    const bytes = opened.archive.readBytes(mediaRef)
    if (!bytes) return undefined
    const ext = mediaRef.split('.').pop()?.toLowerCase() ?? 'png'
    if (ext === 'tif' || ext === 'tiff') {
      const decoded = tiffToPng(bytes)
      const value = decoded
        ? `data:image/png;base64,${Buffer.from(decoded.png).toString('base64')}`
        : undefined
      cache.set(mediaRef, value)
      return value
    }
    const mime =
      ext === 'svg'
        ? 'image/svg+xml'
        : ext === 'gif'
          ? 'image/gif'
          : ext === 'webp'
            ? 'image/webp'
            : ext === 'jpg' || ext === 'jpeg'
              ? 'image/jpeg'
              : 'image/png'
    const value = `data:${mime};base64,${Buffer.from(bytes).toString('base64')}`
    cache.set(mediaRef, value)
    return value
  }
}

function renderSlides(session: SlidesSession): RenderSlide[] {
  const media = mediaResolver(session.opened)
  return session.opened.deck.slides.map((slide, index) =>
    buildRenderSlide(slide, session.opened.deck.size, {
      fitWidthPx: session.fitWidthPx,
      media,
      slideNo: index + 1,
    }),
  )
}

function renderSlide(session: SlidesSession, slideIndex: number): RenderSlide | null {
  return renderSlides(session)[slideIndex] ?? null
}

function defaultFont(opened: OpenedPptx): string | undefined {
  try {
    const slidePath = opened.archive.readPresentation().slidePaths[0]
    if (!slidePath) return undefined
    const themePath = opened.archive.resolveSlideChain(slidePath).themePath
    const xml = themePath ? opened.archive.readText(themePath) : undefined
    return xml ? parseTheme(xml).minorFont : undefined
  } catch {
    return undefined
  }
}

const FALLBACK_ACCENTS = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47']

function mixHex(hex: string, target: number, ratio: number): string {
  const match = /^#?([0-9a-fA-F]{6})$/.exec(hex)
  if (!match) return hex
  const value = Number.parseInt(match[1]!, 16)
  const channel = (source: number) => Math.round(source + (target - source) * ratio)
  const red = channel((value >> 16) & 255)
  const green = channel((value >> 8) & 255)
  const blue = channel(value & 255)
  return `#${((red << 16) | (green << 8) | blue).toString(16).padStart(6, '0').toUpperCase()}`
}

function deckAccents(opened: OpenedPptx): string[] {
  const slide = opened.deck.slides[0]
  if (!slide) return FALLBACK_ACCENTS
  try {
    const themePath = opened.archive.resolveSlideChain(slide.path).themePath
    const colors = themePath
      ? parseTheme(opened.archive.readText(themePath) ?? '').colors
      : undefined
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
      .map((key) => colors?.[key])
      .filter((color): color is string => !!color)
    return accents.length >= 3 ? accents : FALLBACK_ACCENTS
  } catch {
    return FALLBACK_ACCENTS
  }
}

function chartColorSchemes(opened: OpenedPptx): Array<{
  key: string
  label: string
  colors: string[]
}> {
  const accents = deckAccents(opened)
  const rotated = [...accents.slice(3), ...accents.slice(0, 3)]
  const mono = (color: string) => [
    mixHex(color, 0, 0.25),
    color,
    mixHex(color, 255, 0.25),
    mixHex(color, 255, 0.45),
    mixHex(color, 255, 0.65),
  ]
  return [
    { key: 'default', label: '主题默认', colors: [] },
    { key: 'colorful', label: '彩色 1', colors: accents },
    { key: 'colorful2', label: '彩色 2', colors: rotated },
    ...accents.map((color, index) => ({
      key: `mono-accent${index + 1}`,
      label: `单色 ${index + 1}`,
      colors: mono(color),
    })),
  ]
}

function openResult(session: SlidesSession): OpenResult {
  return {
    path: session.webPath,
    slides: renderSlides(session),
    size: { ...session.opened.deck.size },
    defaultFont: defaultFont(session.opened),
  }
}

function snapshot(session: SlidesSession): HistorySnapshot {
  return {
    slides: structuredClone(session.opened.deck.slides),
    entries: new Map(session.opened.archive.entries),
    size: { ...session.opened.deck.size },
    metaDirty: session.metaDirty,
    name: session.name,
    webPath: session.webPath,
  }
}

function cloneSnapshot(value: HistorySnapshot): HistorySnapshot {
  return {
    slides: structuredClone(value.slides),
    entries: new Map(value.entries),
    size: { ...value.size },
    metaDirty: value.metaDirty,
    name: value.name,
    webPath: value.webPath,
  }
}

function restore(session: SlidesSession, value: HistorySnapshot): void {
  session.opened.deck.slides = structuredClone(value.slides)
  session.opened.deck.size = { ...value.size }
  session.opened.archive.entries.clear()
  for (const [path, bytes] of value.entries) session.opened.archive.entries.set(path, bytes)
  session.metaDirty = value.metaDirty
  session.name = value.name
  session.webPath = value.webPath
  session.transformPreview = false
  session.masterEdit = null
}

function pushHistory(session: SlidesSession): void {
  session.undo.push(snapshot(session))
  while (session.undo.length > MAX_HISTORY) session.undo.shift()
  session.redo = []
}

function beginHistoryBatch(session: SlidesSession): void {
  if (session.historyBatch) {
    session.historyBatch.depth += 1
    return
  }
  session.historyBatch = {
    depth: 1,
    undoStart: session.undo.length,
    before: snapshot(session),
  }
}

function endHistoryBatch(session: SlidesSession): HistorySnapshot | null {
  const batch = session.historyBatch
  if (!batch) return null
  batch.depth -= 1
  if (batch.depth > 0) return null
  session.historyBatch = undefined
  if (session.undo.length <= batch.undoStart) return null
  session.undo.splice(batch.undoStart)
  session.undo.push(batch.before)
  while (session.undo.length > MAX_HISTORY) session.undo.shift()
  return batch.before
}

function registerAiSnapshot(session: SlidesSession, value: HistorySnapshot): number {
  const id = nextAiSnapshotId++
  session.aiSnapshots.set(id, cloneSnapshot(value))
  while (session.aiSnapshots.size > MAX_AI_SNAPSHOTS) {
    const oldest = session.aiSnapshots.keys().next().value
    if (oldest === undefined) break
    session.aiSnapshots.delete(oldest)
  }
  return id
}

function isDirty(session: SlidesSession): boolean {
  return (
    session.metaDirty ||
    session.opened.deck.slides.some(
      (slide) =>
        slide.structureDirty ||
        slide.elements.some((element) =>
          Object.keys(element).some(
            (key) => key.startsWith('dirty') && element[key as keyof typeof element],
          ),
        ),
    )
  )
}

function findTextElement(slide: Slide, sourceId: string): TextElement | null {
  const element = slide.elements.find((candidate) => candidate.id === sourceId)
  return element && (element.type === 'text' || element.type === 'shape')
    ? (element as TextElement)
    : null
}

function toEmu(session: SlidesSession, fitWidthPx: number, pixels: number): number {
  const baseWidthPx = session.opened.deck.size.cx / EMU_PER_PX_96
  return Math.round((pixels / (fitWidthPx / baseWidthPx)) * EMU_PER_PX_96)
}

function safeFitWidth(value: number): number {
  if (!Number.isFinite(value) || value < 320 || value > 4096) {
    throw new Error('演示文稿画布宽度无效')
  }
  return value
}

export class SlidesWebService {
  private readonly sessions = new Map<string, SlidesSession>()
  private slideClipboard: { bundle: SlideBundle; png?: string } | null = null

  constructor(
    private readonly maxPresentationBytes = 96 * 1024 * 1024,
    private readonly maxSessions = 64,
    private readonly sessionTtlMs = 30 * 60 * 1000,
  ) {}

  async open(input: OpenSlidesInput): Promise<{ sessionId: string; result: OpenResult }> {
    this.prepareSessionSlot()
    const name = safeName(input.name)
    const opened = await openPptx(decodePptx(input.pptxBase64, this.maxPresentationBytes))
    const now = Date.now()
    const session: SlidesSession = {
      id: randomUUID(),
      name,
      webPath: input.webPath || `webfile://${randomUUID()}/${encodeURIComponent(name)}`,
      opened,
      fitWidthPx: safeFitWidth(input.fitWidthPx),
      undo: [],
      redo: [],
      metaDirty: false,
      transformPreview: false,
      clipboard: null,
      lastAccessAt: now,
      masterEdit: null,
      aiSnapshots: new Map(),
      lastSlidePaste: null,
    }
    this.sessions.set(session.id, session)
    return { sessionId: session.id, result: openResult(session) }
  }

  async blank(fitWidthPx: number): Promise<{ sessionId: string; result: OpenResult }> {
    this.prepareSessionSlot()
    const opened = await openPptx(await createBlankPptx())
    const name = 'Presentation.pptx'
    const now = Date.now()
    const session: SlidesSession = {
      id: randomUUID(),
      name,
      webPath: `webfile://${randomUUID()}/${name}`,
      opened,
      fitWidthPx: safeFitWidth(fitWidthPx),
      undo: [],
      redo: [],
      metaDirty: false,
      transformPreview: false,
      clipboard: null,
      lastAccessAt: now,
      masterEdit: null,
      aiSnapshots: new Map(),
      lastSlidePaste: null,
    }
    this.sessions.set(session.id, session)
    return { sessionId: session.id, result: openResult(session) }
  }

  async save(
    sessionId: string,
    name?: string,
    webPath?: string,
  ): Promise<{
    ok: true
    name: string
    path: string
    slides: RenderSlide[]
    pptxBase64: string
  }> {
    const session = this.requireSession(sessionId)
    if (name) session.name = safeName(name)
    if (webPath) session.webPath = webPath
    const bytes = await savePptx(session.opened)
    commitSaved(session.opened)
    session.metaDirty = false
    return {
      ok: true,
      name: session.name,
      path: session.webPath,
      slides: renderSlides(session),
      pptxBase64: Buffer.from(bytes).toString('base64'),
    }
  }

  async call(input: CallSlidesInput): Promise<unknown> {
    const session = this.requireSession(input.sessionId)
    const args = input.args ?? []
    const op = args[0]
    switch (input.action) {
      case 'getRenderSlides':
        return renderSlides(session)
      case 'getSlideSize':
        return { ...session.opened.deck.size }
      case 'isDirty':
        return isDirty(session)
      case 'getLayouts': {
        const layouts = listSlideLayouts(session.opened.archive)
        if (shouldOfferBuiltinLayouts(layouts)) {
          layouts.push(
            ...builtinLayoutInfos(
              session.opened.deck.size,
              new Set(layouts.map((layout) => layout.name)),
            ),
          )
        }
        return { layouts, size: { ...session.opened.deck.size } }
      }
      case 'importGenerated':
        return this.importGenerated(
          session,
          op as {
            pptxBase64: string
            fitWidthPx: number
            mode?: 'replace' | 'append' | 'replace_at' | 'insert_at'
            atIndex?: number
            deckName?: string
          },
        )
      case 'findReplace':
        return this.findReplace(session, op as FindReplaceOp)
      case 'setSlideLayout':
        return this.setSlideLayout(session, op as SetSlideLayoutOp)
      case 'setSlideSize':
        return this.setSlideSize(session, op as SetSlideSizeOp)
      case 'editText':
        return this.editText(session, op as EditTextOp)
      case 'setElementFont':
        return this.setElementFont(session, op as SetElementFontOp)
      case 'setElementParagraphFormat':
        return this.setElementParagraphFormat(session, op as SetElementParagraphFormatOp)
      case 'editTransform':
        return this.editTransform(session, op as EditTransformOp)
      case 'editConnectorEndpoints':
        return this.editConnectorEndpoints(session, op as EditConnectorEndpointsOp)
      case 'batchEditTransform':
        return this.batchEditTransform(session, op as BatchEditTransformOp)
      case 'editPictureSrcRect':
        return this.editPictureSrcRect(session, op as EditPictureSrcRectOp)
      case 'editFill':
        return this.editFill(session, op as EditFillOp)
      case 'editStroke':
        return this.editStroke(session, op as EditStrokeOp)
      case 'editPictureOpacity':
        return this.editPictureOpacity(session, op as EditPictureOpacityOp)
      case 'setTextAnchor':
        return this.setTextAnchor(
          session,
          op as { slideIndex: number; sourceId: string; anchor: 'top' | 'middle' | 'bottom' },
        )
      case 'groupElements':
        return this.groupElements(session, op as GroupElementsOp)
      case 'ungroupElement':
        return this.ungroupElement(session, op as UngroupElementOp)
      case 'editBackground':
        return this.editBackground(session, op as EditBackgroundOp)
      case 'addElement':
        return this.addElement(session, op as AddElementOp)
      case 'deleteElement':
        return this.deleteElement(session, op as DeleteElementOp)
      case 'reorderElement':
        return this.reorderElement(session, op as ReorderElementOp)
      case 'flipElements':
        return this.flipElements(session, op as FlipElementOp)
      case 'copyElements':
        return this.copyElements(session, op as CopyElementsOp)
      case 'pasteElements':
        return this.pasteElements(session, op as PasteElementsOp)
      case 'duplicateElements':
        return this.duplicateElements(session, op as DuplicateElementsOp)
      case 'addTable':
        return this.addTable(session, op as AddTableOp)
      case 'editTableCell':
        return this.editTableCell(session, op as EditTableCellOp)
      case 'editTableStyle':
        return this.editTableStyle(session, op as EditTableStyleOp)
      case 'tableStructure':
        return this.tableStructure(session, op as TableStructureIpcOp)
      case 'tableMerge':
        return this.tableMerge(session, op as TableMergeIpcOp)
      case 'setTableColWidth':
        return this.setTableColWidth(session, op as SetTableColWidthOp)
      case 'setTableRowHeight':
        return this.setTableRowHeight(session, op as SetTableRowHeightOp)
      case 'setTableCellAnchor':
        return this.setTableCellAnchor(session, op as SetTableCellAnchorOp)
      case 'addInk':
        return this.addInk(session, op as AddInkOp)
      case 'addChart':
        return this.addChart(session, op as AddChartOp)
      case 'editChart':
        return this.editChart(session, op as EditChartOp)
      case 'getChartData': {
        const slide = session.opened.deck.slides[Number(args[0])]
        return slide ? getChartElementData(slide, String(args[1])) : null
      }
      case 'addSmartArt':
        return this.addSmartArt(session, op as AddSmartArtOp)
      case 'addImageBytes':
        return this.addImageBytes(session, op as AddImageBytesOp)
      case 'replacePictureBytes':
        return this.replacePictureBytes(session, op as ReplacePictureBytesOp)
      case 'addMediaBytes':
        return this.addMediaBytes(session, op as AddMediaBytesOp)
      case 'editImageFillBytes':
        return this.editImageFillBytes(
          session,
          op as { slideIndex: number; sourceId: string; base64: string; ext: string },
        )
      case 'addModel3dBytes':
        return this.addModel3dBytes(
          session,
          op as {
            slideIndex: number
            base64: string
            ext: string
            fitWidthPx: number
            name?: string
          },
        )
      case 'getMediaData':
        return this.getMediaData(session, Number(args[0]), String(args[1]))
      case 'addSlide':
        return this.addSlide(session, op as AddSlideOp)
      case 'addSlideWithLayout':
        return this.addSlideWithLayout(session, op as AddSlideWithLayoutOp)
      case 'addBlankSlide':
        return this.addBlankSlide(session, op as AddBlankSlideOp)
      case 'deleteSlide':
        return this.deleteSlide(session, Number(op))
      case 'moveSlide':
        return this.moveSlide(session, op as MoveSlideOp)
      case 'copySlide':
        return this.copySlide(session, Number(args[0]), args[1] as string | undefined)
      case 'pasteSlide':
        return this.pasteSlide(session, op as PasteSlideOp)
      case 'repasteSlide':
        return this.repasteSlide(session, op as RepasteSlideOp)
      case 'hasSlideClipboard':
        return this.slideClipboard !== null
      case 'setLink':
        return this.setLink(session, op as SetLinkOp)
      case 'getLink':
        return getElementLink(session.opened, Number(args[0]), String(args[1]))
      case 'getNotes': {
        const slide = session.opened.deck.slides[Number(op)]
        return slide ? getSlideNotes(session.opened.archive, slide.path) : ''
      }
      case 'setNotes':
        return this.setNotes(session, op as SetNotesOp)
      case 'getComments': {
        const slide = session.opened.deck.slides[Number(op)]
        return slide ? getSlideComments(session.opened.archive, slide.path) : []
      }
      case 'addComment':
        return this.addComment(session, op as AddCommentOp)
      case 'deleteComment':
        return this.deleteComment(session, op as DeleteCommentOp)
      case 'getTransition':
        return getSlideTransition(session.opened.deck.slides[Number(op)]!)
      case 'setTransition':
        return this.setTransition(session, op as SetTransitionOp)
      case 'getAnimations':
        return this.getAnimations(session, Number(op))
      case 'getShapeKeys': {
        const slide = session.opened.deck.slides[Number(op)]
        return slide
          ? slide.elements.map((element) => ({
              sourceId: element.id,
              spid: elementSpid(element),
              name: element.name ?? '',
            }))
          : []
      }
      case 'setAnimations':
        return this.setAnimations(session, op as SetAnimationsOp)
      case 'setAdvanceTimes':
        return this.setAdvanceTimes(session, op as SetAdvanceTimesOp)
      case 'setSlideHidden':
        return this.setSlideHidden(session, op as SetSlideHiddenOp)
      case 'getSections':
        return getSections(session.opened)
      case 'applyHeaderFooter':
        return this.applyHeaderFooter(session, op as HeaderFooterOp)
      case 'getHeaderFooter': {
        const slide = session.opened.deck.slides[Number(op)]
        return slide ? readHeaderFooter(slide) : { footer: null, slideNum: false, date: null }
      }
      case 'applyTheme':
        return this.applyTheme(session, op as ApplyThemeOp)
      case 'masterEnter':
        return this.masterEnter(session, Number(op))
      case 'masterOpen':
        return this.masterOpen(session, String(op))
      case 'masterClose':
        session.masterEdit = null
        return renderSlides(session)
      case 'masterEditText':
        return this.masterEditText(session, op as MasterEditTextOp)
      case 'masterEditTransform':
        return this.masterEditTransform(session, op as MasterEditTransformOp)
      case 'masterEditFill':
        return this.masterEditFill(session, op as MasterEditFillOp)
      case 'masterEditStroke':
        return this.masterEditStroke(session, op as MasterEditStrokeOp)
      case 'masterDeleteElement':
        return this.masterDeleteElement(session, op as MasterDeleteElementOp)
      case 'setSections':
        return this.setSections(session, op as SectionInfo[])
      case 'addSection':
        return this.addSection(session, op as AddSectionOp)
      case 'renameSection':
        return this.renameSection(session, op as RenameSectionOp)
      case 'removeSection':
        return this.removeSection(session, op as RemoveSectionOp)
      case 'moveSection':
        return this.moveSection(session, op as MoveSectionOp)
      case 'undo':
        return this.undo(session)
      case 'redo':
        return this.redo(session)
      case 'beginHistoryBatch':
        beginHistoryBatch(session)
        return true
      case 'endHistoryBatch': {
        const before = endHistoryBatch(session)
        return before ? registerAiSnapshot(session, before) : null
      }
      case 'aiSnapshotRestore':
        return this.restoreAiSnapshot(session, Number(op))
      case 'getRecentFiles':
        return []
      case 'getSlideLinks':
        return getSlideLinks(session.opened, Number(op)).map(({ elementId, target }) => ({
          sourceId: elementId,
          target,
        }))
      case 'getRunLinks':
        return getRunLinks(session.opened, Number(op)).map(({ elementId, ...rest }) => ({
          sourceId: elementId,
          ...rest,
        }))
      case 'getCommentsForSlide':
        return []
      case 'cloudGenStatus':
        return { enabled: false }
      case 'getChartColorSchemes':
        return chartColorSchemes(session.opened)
      default:
        return null
    }
  }

  close(sessionId: unknown): void {
    if (typeof sessionId === 'string') this.sessions.delete(sessionId)
  }

  private findReplace(session: SlidesSession, op: FindReplaceOp): unknown {
    pushHistory(session)
    const { count } = replaceAllInDeck(session.opened.deck, op.find, op.replace, {
      matchCase: op.matchCase,
      firstOnly: op.firstOnly,
      slideIndex: op.slideIndex,
      elementId: op.elementId,
    })
    if (!count) {
      session.undo.pop()
      return { count: 0, slides: null }
    }
    return { count, slides: renderSlides(session) }
  }

  private setSlideLayout(session: SlidesSession, op: SetSlideLayoutOp): RenderSlide | null {
    pushHistory(session)
    const layoutPath = op.layoutPath?.startsWith(BUILTIN_LAYOUT_PREFIX)
      ? ensureBuiltinLayout(
          session.opened.archive,
          session.opened.deck.size,
          op.layoutPath.slice(BUILTIN_LAYOUT_PREFIX.length),
        )
      : op.layoutPath
    const changed = layoutPath
      ? setSlideLayout(session.opened, op.slideIndex, layoutPath)
      : op.layoutPath
        ? null
        : resetSlideLayout(session.opened, op.slideIndex)
    if (!changed) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private setSlideSize(session: SlidesSession, op: SetSlideSizeOp): RenderSlide[] | null {
    pushHistory(session)
    if (!setSlideSize(session.opened, op.cx, op.cy)) {
      session.undo.pop()
      return null
    }
    session.metaDirty = true
    return renderSlides(session)
  }

  private async importGenerated(
    session: SlidesSession,
    op: {
      pptxBase64: string
      fitWidthPx: number
      mode?: 'replace' | 'append' | 'replace_at' | 'insert_at'
      atIndex?: number
      deckName?: string
    },
  ): Promise<
    OpenResult & { appendedFrom?: number; replacedIndex?: number; insertedIndex?: number }
  > {
    const generated = await openPptx(decodePptx(op.pptxBase64, this.maxPresentationBytes))
    if (generated.deck.slides.length === 0) throw new Error('生成的演示文稿没有幻灯片')
    const mode = op.mode ?? 'replace'
    session.fitWidthPx = safeFitWidth(op.fitWidthPx)
    pushHistory(session)
    try {
      if (mode === 'replace') {
        session.opened = generated
        session.name = safeName(op.deckName || 'Generated.pptx')
        session.webPath = `webfile://${randomUUID()}/${encodeURIComponent(session.name)}`
        session.metaDirty = true
        return openResult(session)
      }

      const beforeCount = session.opened.deck.slides.length
      for (let index = 0; index < generated.deck.slides.length; index += 1) {
        const bundle = copySlide(generated, index)
        if (!bundle) throw new Error(`无法读取生成的第 ${index + 1} 页`)
        const inserted = pasteSlide(session.opened, session.opened.deck.slides.length - 1, bundle, {
          keepSourceFormatting: true,
        })
        if (!inserted) throw new Error(`无法合并生成的第 ${index + 1} 页`)
        promoteSlideBackground(inserted, session.opened.deck.size)
      }

      if (mode === 'replace_at') {
        const atIndex = op.atIndex
        if (generated.deck.slides.length !== 1 || atIndex === undefined) {
          throw new Error('单页替换需要一页内容和有效页码')
        }
        if (atIndex < 0 || atIndex >= beforeCount) throw new Error('替换页码超出范围')
        if (
          !moveSlide(session.opened, beforeCount, atIndex) ||
          !deleteSlide(session.opened, atIndex + 1)
        ) {
          throw new Error('替换生成页失败')
        }
        session.metaDirty = true
        return { ...openResult(session), replacedIndex: atIndex }
      }

      if (mode === 'insert_at') {
        const atIndex = op.atIndex
        if (generated.deck.slides.length !== 1 || atIndex === undefined) {
          throw new Error('单页插入需要一页内容和有效页码')
        }
        if (atIndex < 0 || atIndex > beforeCount) throw new Error('插入页码超出范围')
        if (atIndex < beforeCount && !moveSlide(session.opened, beforeCount, atIndex)) {
          throw new Error('插入生成页失败')
        }
        session.metaDirty = true
        return { ...openResult(session), insertedIndex: atIndex }
      }

      session.metaDirty = true
      return { ...openResult(session), appendedFrom: beforeCount }
    } catch (error) {
      const previous = session.undo.pop()
      if (previous) restore(session, previous)
      throw error
    }
  }

  private editConnectorEndpoints(
    session: SlidesSession,
    op: EditConnectorEndpointsOp,
  ): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    const element = slide?.elements.find((candidate) => candidate.id === op.sourceId)
    if (!slide || !element) return null
    pushHistory(session)
    const first = {
      x: toEmu(session, op.fitWidthPx, op.x1Px),
      y: toEmu(session, op.fitWidthPx, op.y1Px),
    }
    const second = {
      x: toEmu(session, op.fitWidthPx, op.x2Px),
      y: toEmu(session, op.fitWidthPx, op.y2Px),
    }
    element.transform = {
      ...element.transform,
      offset: {
        x: Math.min(first.x, second.x),
        y: Math.min(first.y, second.y),
        cx: Math.abs(second.x - first.x),
        cy: Math.abs(second.y - first.y),
      },
      rot: 0,
      flipH: first.x > second.x,
      flipV: first.y > second.y,
    }
    element.dirtyTransform = true
    const connection = (value: { targetId: string; idx: number } | null | undefined) => {
      if (value === undefined || value === null) return value
      const target = slide.elements.find((candidate) => candidate.id === value.targetId)
      const id = target ? elementSpid(target) : null
      return id === null ? null : { id, idx: value.idx }
    }
    setElementConnection(slide, op.sourceId, {
      start: connection(op.start),
      end: connection(op.end),
    })
    return renderSlide(session, op.slideIndex)
  }

  private editPictureSrcRect(session: SlidesSession, op: EditPictureSrcRectOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!editPictureSrcRect(slide, op.sourceId, op.srcRect)) {
      session.undo.pop()
      return null
    }
    if (op.boxPx && op.fitWidthPx) {
      const element = slide.elements.find((candidate) => candidate.id === op.sourceId)
      if (element) {
        element.transform = {
          ...element.transform,
          offset: {
            x: toEmu(session, op.fitWidthPx, op.boxPx.x),
            y: toEmu(session, op.fitWidthPx, op.boxPx.y),
            cx: toEmu(session, op.fitWidthPx, op.boxPx.w),
            cy: toEmu(session, op.fitWidthPx, op.boxPx.h),
          },
        }
        element.dirtyTransform = true
        updateConnectorsForMoved(slide, [op.sourceId])
      }
    }
    return renderSlide(session, op.slideIndex)
  }

  private setTextAnchor(
    session: SlidesSession,
    op: { slideIndex: number; sourceId: string; anchor: 'top' | 'middle' | 'bottom' },
  ): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setElementTextAnchor(slide, op.sourceId, op.anchor)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private groupElements(session: SlidesSession, op: GroupElementsOp): unknown {
    pushHistory(session)
    const result = groupElements(session.opened, op.slideIndex, op.sourceIds)
    if (!result) {
      session.undo.pop()
      return null
    }
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, groupId: result.groupId } : null
  }

  private ungroupElement(session: SlidesSession, op: UngroupElementOp): RenderSlide | null {
    pushHistory(session)
    if (!ungroupElement(session.opened, op.slideIndex, op.sourceId)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private editText(session: SlidesSession, op: EditTextOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (op.groupId) {
      const found = findGroupChild(slide, op.groupId, op.sourceId)
      const child = found?.child
      if (!child || (child.type !== 'text' && child.type !== 'shape')) return null
      const textChild = child as TextElement
      if (!textChild.text) return null
      pushHistory(session)
      textChild.text.paragraphs = applyEditParagraphs(textChild.text.paragraphs, op.paragraphs)
      if (!patchGroupChildText(slide, op.groupId, textChild)) {
        const previous = session.undo.pop()
        if (previous) restore(session, previous)
        return null
      }
      for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
        setGroupChildParagraphFormat(slide, op.groupId, op.sourceId, patch, [index])
      }
      return renderSlide(session, op.slideIndex)
    }
    const element = findTextElement(slide, op.sourceId)
    if (!element?.text) return null
    pushHistory(session)
    element.text.paragraphs = applyEditParagraphs(element.text.paragraphs, op.paragraphs)
    element.dirty = true
    for (const { index, patch } of collectParagraphFormatPatches(op.paragraphs)) {
      setElementParagraphFormat(slide, op.sourceId, patch, [index])
    }
    return renderSlide(session, op.slideIndex)
  }

  private setElementFont(session: SlidesSession, op: SetElementFontOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    let changed = false
    for (const sourceId of op.sourceIds) {
      const updated = op.groupId
        ? setGroupChildFont(slide, op.groupId, sourceId, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
        : setElementFont(slide, sourceId, {
            fontFamily: op.fontFamily,
            fontSizePt: op.fontSizePt,
            strike: op.strike,
            bold: op.bold,
            italic: op.italic,
            underline: op.underline,
            color: op.color,
          })
      changed = updated || changed
    }
    if (!changed) session.undo.pop()
    return changed ? renderSlide(session, op.slideIndex) : null
  }

  private setElementParagraphFormat(
    session: SlidesSession,
    op: SetElementParagraphFormatOp,
  ): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const patch = {
      bullet: op.bullet,
      bulletChar: op.bulletChar,
      bulletHangEmu: op.bulletHangEmu,
      bulletSizePct: op.bulletSizePct,
      bulletColor: op.bulletColor,
      lineSpacingPct: op.lineSpacingPct,
      spaceBeforePt: op.spaceBeforePt,
      spaceAfterPt: op.spaceAfterPt,
      align: op.align,
      indentDelta: op.indentDelta,
    }
    let changed = false
    for (const sourceId of op.sourceIds) {
      const updated = op.groupId
        ? setGroupChildParagraphFormat(slide, op.groupId, sourceId, patch)
        : setElementParagraphFormat(slide, sourceId, patch)
      changed = updated || changed
    }
    if (!changed) session.undo.pop()
    return changed ? renderSlide(session, op.slideIndex) : null
  }

  private editTransform(session: SlidesSession, op: EditTransformOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const element = op.groupId
      ? null
      : slide.elements.find((candidate) => candidate.id === op.sourceId)
    const groupChild = op.groupId ? findGroupChild(slide, op.groupId, op.sourceId) : null
    if (!element && !groupChild) return null
    if (op.preview) {
      if (!session.transformPreview) {
        pushHistory(session)
        session.transformPreview = true
      }
    } else if (session.transformPreview) session.transformPreview = false
    else pushHistory(session)
    if (groupChild) {
      const childOffset = groupChild.grp.childOffset
      const childX = childOffset?.x ?? groupChild.grp.transform.offset.x
      const childY = childOffset?.y ?? groupChild.grp.transform.offset.y
      const groupOffset = groupChild.grp.transform.offset
      const scaleX = childOffset?.cx ? groupOffset.cx / childOffset.cx : 1
      const scaleY = childOffset?.cy ? groupOffset.cy / childOffset.cy : 1
      const changed = editGroupChildTransform(
        slide,
        op.groupId!,
        op.sourceId,
        {
          x: toEmu(session, op.fitWidthPx, op.xPx / scaleX) + childX,
          y: toEmu(session, op.fitWidthPx, op.yPx / scaleY) + childY,
          cx: toEmu(session, op.fitWidthPx, op.wPx / scaleX),
          cy: toEmu(session, op.fitWidthPx, op.hPx / scaleY),
        },
        op.rotationDeg,
      )
      if (!changed) {
        session.undo.pop()
        return null
      }
      return renderSlide(session, op.slideIndex)
    }
    const isTable = element!.type === 'table'
    const width = toEmu(session, op.fitWidthPx, op.wPx)
    const height = toEmu(session, op.fitWidthPx, op.hPx)
    if (isTable) resizeTable(slide, op.sourceId, width, height)
    element!.transform = {
      ...element!.transform,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: isTable ? element!.transform.offset.cx : width,
        cy: isTable ? element!.transform.offset.cy : height,
      },
      rot: Math.round(op.rotationDeg * 60000),
    }
    element!.dirtyTransform = true
    updateConnectorsForMoved(slide, [op.sourceId])
    return renderSlide(session, op.slideIndex)
  }

  private batchEditTransform(session: SlidesSession, op: BatchEditTransformOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const items = op.items.map((item) => ({
      item,
      element: slide.elements.find((candidate) => candidate.id === item.sourceId),
    }))
    if (items.some(({ element }) => !element)) return null
    pushHistory(session)
    for (const { item, element } of items) {
      const width = toEmu(session, op.fitWidthPx, item.wPx)
      const height = toEmu(session, op.fitWidthPx, item.hPx)
      const isTable = element!.type === 'table'
      if (isTable) resizeTable(slide, item.sourceId, width, height)
      element!.transform = {
        ...element!.transform,
        offset: {
          x: toEmu(session, op.fitWidthPx, item.xPx),
          y: toEmu(session, op.fitWidthPx, item.yPx),
          cx: isTable ? element!.transform.offset.cx : width,
          cy: isTable ? element!.transform.offset.cy : height,
        },
        rot: Math.round(item.rotationDeg * 60000),
      }
      element!.dirtyTransform = true
    }
    updateConnectorsForMoved(
      slide,
      op.items.map((item) => item.sourceId),
    )
    return renderSlide(session, op.slideIndex)
  }

  private editFill(session: SlidesSession, op: EditFillOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (op.groupId) {
      const fill =
        typeof op.fill === 'string'
          ? op.fill
          : {
              stops: [
                { pos: 0, color: op.fill.gradient.from },
                { pos: 1, color: op.fill.gradient.to },
              ],
              ...(op.fill.gradient.radial
                ? { radial: true }
                : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) }),
            }
      pushHistory(session)
      if (!editGroupChildFill(slide, op.groupId, op.sourceId, fill)) {
        session.undo.pop()
        return null
      }
      return renderSlide(session, op.slideIndex)
    }
    const element = slide ? findTextElement(slide, op.sourceId) : null
    if (!element) return null
    pushHistory(session)
    if (typeof op.fill === 'string') {
      element.fill = op.fill === 'none' ? { type: 'none' } : { type: 'solid', color: op.fill }
    } else {
      const gradient = op.fill.gradient
      element.fill = {
        type: 'gradient',
        stops: [
          { pos: 0, color: gradient.from },
          { pos: 1, color: gradient.to },
        ],
        ...(gradient.radial
          ? { path: 'circle' as const }
          : { angle: Math.round((gradient.angleDeg ?? 0) * 60000) }),
      }
    }
    element.dirtyFill = true
    return renderSlide(session, op.slideIndex)
  }

  private editStroke(session: SlidesSession, op: EditStrokeOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    if (op.groupId) {
      pushHistory(session)
      const stroke = op.stroke
        ? {
            color: op.stroke.color,
            widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
            ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
          }
        : null
      if (!editGroupChildStroke(slide, op.groupId, op.sourceId, stroke)) {
        session.undo.pop()
        return null
      }
      return renderSlide(session, op.slideIndex)
    }
    const element = slide?.elements.find((candidate) => candidate.id === op.sourceId)
    if (
      !element ||
      (element.type !== 'text' && element.type !== 'shape' && element.type !== 'picture')
    )
      return null
    pushHistory(session)
    const stroked = element as TextElement
    stroked.stroke = op.stroke
      ? {
          fill: { type: 'solid', color: op.stroke.color },
          width: Math.round(op.stroke.widthPt * EMU_PER_PT),
          ...(op.stroke.dash ? { dash: op.stroke.dash } : {}),
        }
      : undefined
    element.dirtyStroke = true
    return renderSlide(session, op.slideIndex)
  }

  private editPictureOpacity(session: SlidesSession, op: EditPictureOpacityOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setPictureOpacity(slide, op.sourceId, op.opacity)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private editBackground(session: SlidesSession, op: EditBackgroundOp): RenderSlide[] | null {
    const targets =
      op.slideIndex === -1
        ? session.opened.deck.slides
        : [session.opened.deck.slides[op.slideIndex]].filter((slide): slide is Slide => !!slide)
    if (targets.length === 0) return null
    pushHistory(session)
    for (const slide of targets) setSlideBackground(slide, op.color)
    session.fitWidthPx = op.fitWidthPx
    return renderSlides(session)
  }

  private addElement(
    session: SlidesSession,
    op: AddElementOp,
  ): { slide: RenderSlide; sourceId: string } | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const paragraphs: Paragraph[] | undefined = op.paragraphs?.length
      ? (op.paragraphs as Paragraph[])
      : op.text
        ? op.text.split('\n').map((text) => ({ runs: [{ text }] }))
        : undefined
    const element = addElement(slide, {
      kind: op.kind,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: toEmu(session, op.fitWidthPx, op.wPx),
        cy: toEmu(session, op.fitWidthPx, op.hPx),
      },
      ...(paragraphs ? { paragraphs } : {}),
      ...(op.fillColor ? { fillColor: op.fillColor } : {}),
      ...(op.stroke
        ? {
            stroke: {
              color: op.stroke.color,
              widthEmu: Math.round(op.stroke.widthPt * EMU_PER_PT),
            },
          }
        : {}),
    })
    const rendered = renderSlide(session, op.slideIndex)
    return rendered ? { slide: rendered, sourceId: element.id } : null
  }

  private deleteElement(session: SlidesSession, op: DeleteElementOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!deleteElement(slide, op.sourceId)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private reorderElement(session: SlidesSession, op: ReorderElementOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!reorderElement(slide, op.sourceId, op.dir)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private flipElements(session: SlidesSession, op: FlipElementOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const elements = op.sourceIds
      .map((sourceId) =>
        op.groupId
          ? findGroupChild(slide, op.groupId, sourceId)?.child
          : slide.elements.find((element) => element.id === sourceId),
      )
      .filter((element): element is (typeof slide.elements)[number] => !!element)
    if (elements.length === 0) return null
    pushHistory(session)
    for (const element of elements) {
      if (op.axis === 'h') element.transform.flipH = !element.transform.flipH
      else element.transform.flipV = !element.transform.flipV
      element.dirtyTransform = true
    }
    updateConnectorsForMoved(
      slide,
      elements.map((element) => element.id),
    )
    return renderSlide(session, op.slideIndex)
  }

  private copyElements(session: SlidesSession, op: CopyElementsOp): number {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return 0
    const items = op.sourceIds
      .map((id) => slide.elements.find((element) => element.id === id))
      .filter((element): element is NonNullable<typeof element> => !!element)
      .map((element) => copyElementData(session.opened, slide, element))
    if (items.length) session.clipboard = { items, pasteCount: 0 }
    return items.length
  }

  private pasteElements(session: SlidesSession, op: PasteElementsOp): unknown {
    const clipboard = session.clipboard
    if (!clipboard?.items.length || !session.opened.deck.slides[op.slideIndex]) return null
    const shift = toEmu(session, op.fitWidthPx, 16 * (clipboard.pasteCount + 1))
    pushHistory(session)
    const result = pasteElements(session.opened, op.slideIndex, clipboard.items, {
      dx: shift,
      dy: shift,
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    clipboard.pasteCount += 1
    session.fitWidthPx = op.fitWidthPx
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceIds: result.elementIds } : null
  }

  private duplicateElements(session: SlidesSession, op: DuplicateElementsOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const items = op.sourceIds
      .map((id) => slide.elements.find((element) => element.id === id))
      .filter((element): element is NonNullable<typeof element> => !!element)
      .map((element) => copyElementData(session.opened, slide, element))
    if (!items.length) return null
    pushHistory(session)
    const result = pasteElements(session.opened, op.slideIndex, items, {
      dx: toEmu(session, op.fitWidthPx, op.dxPx),
      dy: toEmu(session, op.fitWidthPx, op.dyPx),
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rendered = renderSlide(session, op.slideIndex)
    return rendered ? { slide: rendered, sourceIds: result.elementIds } : null
  }

  private addTable(session: SlidesSession, op: AddTableOp): unknown {
    if (!session.opened.deck.slides[op.slideIndex]) return null
    pushHistory(session)
    const result = addTable(session.opened, op.slideIndex, {
      rows: op.rows,
      cols: op.cols,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: toEmu(session, op.fitWidthPx, op.wPx),
        cy: toEmu(session, op.fitWidthPx, op.hPx),
      },
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private editTableCell(session: SlidesSession, op: EditTableCellOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!editTableCellText(slide, op.sourceId, op.row, op.col, op.paragraphs as Paragraph[])) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private editTableStyle(session: SlidesSession, op: EditTableStyleOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    const elementIndex = slide.elements.findIndex((element) => element.id === op.sourceId)
    pushHistory(session)
    const preset = op.styleName ? TABLE_STYLE_PRESETS[op.styleName] : undefined
    const edit = preset
      ? {
          tblPrXml: preset.tblPrXml,
          clearDirectFormatting: true,
          ...(preset.border
            ? {
                borderPreset: 'all' as const,
                borderColor: preset.border.color,
                borderWidthEmu: preset.border.widthEmu,
              }
            : {}),
        }
      : {
          ...(op.firstRow !== undefined ? { firstRow: op.firstRow } : {}),
          ...(op.bandRow !== undefined ? { bandRow: op.bandRow } : {}),
          ...(op.shadingColor !== undefined ? { shadingColor: op.shadingColor } : {}),
          ...(op.borderPreset !== undefined ? { borderPreset: op.borderPreset } : {}),
          ...(op.borderColor !== undefined ? { borderColor: op.borderColor } : {}),
          ...(op.borderWidthPt !== undefined
            ? {
                borderWidthEmu:
                  op.borderWidthPt === null ? null : Math.round(op.borderWidthPt * EMU_PER_PT),
              }
            : {}),
          ...(op.cells ? { cells: op.cells } : {}),
        }
    if (preset?.styleId && preset.styleDefXml) {
      ensureTableStylePart(session.opened, preset.styleId, preset.styleDefXml)
    }
    if (!editTableStyle(slide, op.sourceId, edit)) {
      session.undo.pop()
      return null
    }
    if (!materializeSlide(session.opened, op.slideIndex)) return null
    const rendered = renderSlide(session, op.slideIndex)
    return rendered
      ? {
          slide: rendered,
          sourceId: session.opened.deck.slides[op.slideIndex]?.elements[elementIndex]?.id ?? null,
        }
      : null
  }

  private tableStructure(session: SlidesSession, op: TableStructureIpcOp): unknown {
    pushHistory(session)
    const result = editTableStructure(session.opened, op.slideIndex, op.sourceId, {
      kind: op.kind,
      index: op.index,
      ...(op.before ? { before: true } : {}),
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private tableMerge(session: SlidesSession, op: TableMergeIpcOp): unknown {
    pushHistory(session)
    const result = mergeTableCells(session.opened, op.slideIndex, op.sourceId, {
      kind: op.kind,
      row: op.row,
      col: op.col,
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private setTableColWidth(session: SlidesSession, op: SetTableColWidthOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setTableColWidth(slide, op.sourceId, op.col, toEmu(session, op.fitWidthPx, op.wPx))) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private setTableRowHeight(session: SlidesSession, op: SetTableRowHeightOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setTableRowHeight(slide, op.sourceId, op.row, toEmu(session, op.fitWidthPx, op.hPx))) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private setTableCellAnchor(session: SlidesSession, op: SetTableCellAnchorOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!setTableCellAnchor(slide, op.sourceId, op.row, op.col, op.anchor)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private addInk(session: SlidesSession, op: AddInkOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const element = addPicture(session.opened, slide, {
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: 'png',
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: Math.max(1, toEmu(session, op.fitWidthPx, op.wPx)),
        cy: Math.max(1, toEmu(session, op.fitWidthPx, op.hPx)),
      },
      name: `genoffice-ink-${Date.now().toString(36)}`,
      descr: op.payload,
    })
    if (!element) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const rendered = renderSlide(session, op.slideIndex)
    return rendered ? { slide: rendered, sourceId: element.id } : null
  }

  private addChart(session: SlidesSession, op: AddChartOp): unknown {
    if (!session.opened.deck.slides[op.slideIndex]) return null
    pushHistory(session)
    const result = addChart(session.opened, op.slideIndex, {
      kind: op.kind === 'barH' ? 'bar' : op.kind,
      ...(op.kind === 'barH' ? { barDir: 'bar' as const } : {}),
      ...(op.title ? { title: op.title } : {}),
      categories: op.categories,
      series: op.series,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: toEmu(session, op.fitWidthPx, op.wPx),
        cy: toEmu(session, op.fitWidthPx, op.hPx),
      },
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private editChart(session: SlidesSession, op: EditChartOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const changed = editChartElement(session.opened, op.slideIndex, op.sourceId, {
      ...(op.kind ? { kind: op.kind === 'barH' ? 'bar' : op.kind } : {}),
      ...(op.kind === 'barH' ? { barDir: 'bar' as const } : {}),
      ...(op.categories ? { categories: op.categories } : {}),
      ...(op.series ? { series: op.series } : {}),
      ...(op.title !== undefined ? { title: op.title } : {}),
      ...(op.colorScheme
        ? {
            colorScheme: chartColorSchemes(session.opened).find(
              (scheme) => scheme.key === op.colorScheme,
            )?.colors,
          }
        : {}),
      ...(op.legendPos ? { legendPos: op.legendPos } : {}),
      ...(op.dataLabels !== undefined ? { dataLabels: op.dataLabels } : {}),
      ...(op.gridlines !== undefined ? { gridlines: op.gridlines } : {}),
      ...(op.catAxisTitle !== undefined ? { catAxisTitle: op.catAxisTitle } : {}),
      ...(op.valAxisTitle !== undefined ? { valAxisTitle: op.valAxisTitle } : {}),
      ...(op.gapWidthPct !== undefined ? { gapWidthPct: op.gapWidthPct } : {}),
      ...(op.switchRowCol ? { switchRowCol: true } : {}),
      ...(op.pointColors ? { pointColors: op.pointColors } : {}),
    })
    if (!changed) {
      session.undo.pop()
      return null
    }
    const index = slide.elements.findIndex((element) => element.id === op.sourceId)
    const fresh = reparseDeck(session.opened)
    session.opened = fresh
    const rendered = renderSlide(session, op.slideIndex)
    return rendered
      ? { slide: rendered, sourceId: fresh.deck.slides[op.slideIndex]?.elements[index]?.id ?? null }
      : null
  }

  private addSmartArt(session: SlidesSession, op: AddSmartArtOp): unknown {
    if (!session.opened.deck.slides[op.slideIndex]) return null
    pushHistory(session)
    const result = addSmartArt(session.opened, op.slideIndex, {
      layout: op.layout,
      items: op.items,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: toEmu(session, op.fitWidthPx, op.wPx),
        cy: toEmu(session, op.fitWidthPx, op.hPx),
      },
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private addImageBytes(session: SlidesSession, op: AddImageBytesOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    const element = addPicture(session.opened, slide, {
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: op.ext,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: Math.max(1, toEmu(session, op.fitWidthPx, op.wPx)),
        cy: Math.max(1, toEmu(session, op.fitWidthPx, op.hPx)),
      },
      ...(op.name ? { name: op.name } : {}),
    })
    if (!element) {
      session.undo.pop()
      return { error: 'unsupported', ext: op.ext }
    }
    session.fitWidthPx = op.fitWidthPx
    const rendered = renderSlide(session, op.slideIndex)
    return rendered ? { slide: rendered, sourceId: element.id } : null
  }

  private replacePictureBytes(session: SlidesSession, op: ReplacePictureBytesOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (
      !replacePictureBytes(
        session.opened,
        slide,
        op.sourceId,
        new Uint8Array(Buffer.from(op.base64, 'base64')),
        op.ext,
        op.keepSrcRect ? { keepSrcRect: true } : undefined,
      )
    ) {
      session.undo.pop()
      return { error: 'unsupported', ext: op.ext }
    }
    return renderSlide(session, op.slideIndex)
  }

  private addMediaBytes(session: SlidesSession, op: AddMediaBytesOp): unknown {
    if (!session.opened.deck.slides[op.slideIndex]) return null
    const deckSize = session.opened.deck.size
    const cx = Math.round(deckSize.cx * 0.6)
    const cy = Math.round((cx * 9) / 16)
    pushHistory(session)
    const result = addMedia(session.opened, op.slideIndex, {
      kind: op.kind,
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: op.ext,
      offset: {
        x: Math.round((deckSize.cx - cx) / 2),
        y: Math.round((deckSize.cy - cy) / 2),
        cx,
        cy,
      },
      ...(op.name ? { name: op.name } : {}),
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private editImageFillBytes(
    session: SlidesSession,
    op: { slideIndex: number; sourceId: string; base64: string; ext: string },
  ): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (
      !setElementImageFill(
        session.opened,
        slide,
        op.sourceId,
        new Uint8Array(Buffer.from(op.base64, 'base64')),
        op.ext,
      )
    ) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private addModel3dBytes(
    session: SlidesSession,
    op: { slideIndex: number; base64: string; ext: string; fitWidthPx: number; name?: string },
  ): unknown {
    if (!session.opened.deck.slides[op.slideIndex]) return null
    const deckSize = session.opened.deck.size
    const side = Math.round(deckSize.cy * 0.5)
    pushHistory(session)
    const result = addModel3d(session.opened, op.slideIndex, {
      bytes: new Uint8Array(Buffer.from(op.base64, 'base64')),
      ext: op.ext,
      offset: {
        x: Math.round((deckSize.cx - side) / 2),
        y: Math.round((deckSize.cy - side) / 2),
        cx: side,
        cy: side,
      },
      ...(op.name ? { name: op.name } : {}),
    })
    if (!result) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    const slide = renderSlide(session, op.slideIndex)
    return slide ? { slide, sourceId: result.elementId } : null
  }

  private getMediaData(session: SlidesSession, slideIndex: number, sourceId: string): unknown {
    const element = session.opened.deck.slides[slideIndex]?.elements.find(
      (candidate) => candidate.id === sourceId,
    )
    if (!element || element.type !== 'picture' || !element.media?.target) return null
    if (element.media.external) return { kind: element.media.kind, dataUrl: element.media.target }
    const bytes = session.opened.archive.readBytes(element.media.target)
    if (!bytes) return null
    const extension = element.media.target.split('.').pop()?.toLowerCase() ?? ''
    const mime: Record<string, string> = {
      mp4: 'video/mp4',
      m4v: 'video/mp4',
      mov: 'video/mp4',
      webm: 'video/webm',
      avi: 'video/x-msvideo',
      mp3: 'audio/mpeg',
      wav: 'audio/wav',
      m4a: 'audio/mp4',
      aac: 'audio/aac',
      ogg: 'audio/ogg',
    }
    return {
      kind: element.media.kind,
      dataUrl: `data:${mime[extension] ?? (element.media.kind === 'video' ? 'video/mp4' : 'audio/mpeg')};base64,${Buffer.from(bytes).toString('base64')}`,
    }
  }

  private addSlide(
    session: SlidesSession,
    op: AddSlideOp,
  ): { slides: RenderSlide[]; index: number } | null {
    pushHistory(session)
    if (!duplicateSlide(session.opened, op.sourceIndex, { clearText: !!op.clearText })) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return { slides: renderSlides(session), index: op.sourceIndex + 1 }
  }

  private addSlideWithLayout(session: SlidesSession, op: AddSlideWithLayoutOp): unknown {
    pushHistory(session)
    const layoutPath = op.layoutPath.startsWith(BUILTIN_LAYOUT_PREFIX)
      ? ensureBuiltinLayout(
          session.opened.archive,
          session.opened.deck.size,
          op.layoutPath.slice(BUILTIN_LAYOUT_PREFIX.length),
        )
      : op.layoutPath
    if (!layoutPath || !insertSlideWithLayout(session.opened, op.sourceIndex, layoutPath)) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return { slides: renderSlides(session), index: op.sourceIndex + 1 }
  }

  private addBlankSlide(
    session: SlidesSession,
    op: AddBlankSlideOp,
  ): { slides: RenderSlide[]; index: number } | null {
    pushHistory(session)
    if (!insertBlankSlide(session.opened, op.sourceIndex)) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return { slides: renderSlides(session), index: op.sourceIndex + 1 }
  }

  private deleteSlide(session: SlidesSession, slideIndex: number): RenderSlide[] | null {
    pushHistory(session)
    if (!deleteSlide(session.opened, slideIndex)) {
      session.undo.pop()
      return null
    }
    session.metaDirty = true
    return renderSlides(session)
  }

  private moveSlide(
    session: SlidesSession,
    op: MoveSlideOp,
  ): { slides: RenderSlide[]; sections: SectionInfo[] } | null {
    pushHistory(session)
    if (!moveSlide(session.opened, op.fromIndex, op.toIndex)) {
      session.undo.pop()
      return null
    }
    session.metaDirty = true
    return { slides: renderSlides(session), sections: getSections(session.opened) }
  }

  private copySlide(session: SlidesSession, slideIndex: number, png?: string): boolean {
    const bundle = copySlide(session.opened, slideIndex)
    if (!bundle) return false
    this.slideClipboard = { bundle, ...(png ? { png } : {}) }
    return true
  }

  private pasteSlide(session: SlidesSession, op: PasteSlideOp): unknown {
    const clipboard = this.slideClipboard
    if (!clipboard) return null
    pushHistory(session)
    if (op.mode === 'picture') {
      const index = Math.min(Math.max(op.afterIndex, 0), session.opened.deck.slides.length - 1)
      const slide = session.opened.deck.slides[index]
      if (!slide || !clipboard.png) {
        session.undo.pop()
        return null
      }
      const element = addPicture(session.opened, slide, {
        bytes: new Uint8Array(Buffer.from(clipboard.png, 'base64')),
        ext: 'png',
        offset: { x: 0, y: 0, ...session.opened.deck.size },
      })
      if (!element) {
        session.undo.pop()
        return null
      }
      session.fitWidthPx = op.fitWidthPx
      session.lastSlidePaste = { afterIndex: op.afterIndex, undoLen: session.undo.length }
      return { slides: renderSlides(session), index, sourceId: element.id }
    }
    const slide = pasteSlide(session.opened, op.afterIndex, clipboard.bundle, {
      keepSourceFormatting: op.mode === 'source',
    })
    if (!slide) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    session.lastSlidePaste = { afterIndex: op.afterIndex, undoLen: session.undo.length }
    return { slides: renderSlides(session), index: session.opened.deck.slides.indexOf(slide) }
  }

  private repasteSlide(session: SlidesSession, op: RepasteSlideOp): unknown {
    const record = session.lastSlidePaste
    if (!record || session.undo.length !== record.undoLen || !this.slideClipboard) return null
    const beforePaste = session.undo.pop()
    if (!beforePaste) return null
    restore(session, beforePaste)
    session.lastSlidePaste = null
    return this.pasteSlide(session, {
      afterIndex: record.afterIndex,
      fitWidthPx: op.fitWidthPx,
      mode: op.mode,
    })
  }

  private setLink(session: SlidesSession, op: SetLinkOp): RenderSlide | null {
    if (!session.opened.deck.slides[op.slideIndex]) return null
    pushHistory(session)
    if (!setElementLink(session.opened, op.slideIndex, op.sourceId, op.target)) {
      session.undo.pop()
      return null
    }
    return renderSlide(session, op.slideIndex)
  }

  private applyHeaderFooter(session: SlidesSession, op: HeaderFooterOp): RenderSlide[] | null {
    pushHistory(session)
    if (
      !applyHeaderFooter(session.opened, {
        footer: op.footer ?? null,
        slideNum: !!op.slideNum,
        date: op.date ?? null,
        ...(op.dateAuto ? { dateAuto: true } : {}),
      })
    ) {
      session.undo.pop()
      return null
    }
    session.fitWidthPx = op.fitWidthPx
    return renderSlides(session)
  }

  private applyTheme(
    session: SlidesSession,
    op: ApplyThemeOp,
  ): RenderSlide[] | { error: string } | null {
    pushHistory(session)
    const spec = {
      name: op.name,
      colors: op.colors,
      ...(op.majorFont ? { majorFont: op.majorFont } : {}),
      ...(op.minorFont ? { minorFont: op.minorFont } : {}),
    }
    try {
      commitSaved(session.opened)
      const patched = applyThemeToArchive(session.opened, spec)
      const remapped = remapDeckColors(session.opened, spec)
      if (patched === 0 && remapped === 0) {
        session.undo.pop()
        return null
      }
      session.opened = reparseDeck(session.opened)
      session.fitWidthPx = op.fitWidthPx
      session.metaDirty = true
      return renderSlides(session)
    } catch (error) {
      const previous = session.undo.pop()
      if (previous) restore(session, previous)
      return { error: error instanceof Error ? error.message : String(error) }
    }
  }

  private renderMaster(session: SlidesSession): RenderSlide | null {
    const master = session.masterEdit
    return master
      ? buildRenderSlide(master.slide, session.opened.deck.size, {
          fitWidthPx: session.fitWidthPx,
          media: mediaResolver(session.opened),
        })
      : null
  }

  private commitMaster(session: SlidesSession): void {
    const master = session.masterEdit
    if (!master) return
    session.opened.archive.entries.set(
      master.partPath,
      Buffer.from(patchSlideXml(master.slide), 'utf8'),
    )
    for (let index = 0; index < session.opened.deck.slides.length; index += 1) {
      materializeSlide(session.opened, index)
    }
    session.metaDirty = true
  }

  private masterEnter(session: SlidesSession, fitWidthPx: number): MasterEnterResult | null {
    session.fitWidthPx = safeFitWidth(fitWidthPx)
    const items: MasterEnterResult['items'] = []
    for (const part of listMasterParts(session.opened.archive)) {
      const slide = parseMasterPart(session.opened.archive, part.partPath)
      if (!slide) continue
      const rendered = buildRenderSlide(slide, session.opened.deck.size, {
        fitWidthPx: session.fitWidthPx,
        media: mediaResolver(session.opened),
      })
      items.push({ partPath: part.partPath, kind: part.kind, name: part.name, slide: rendered })
      if (!session.masterEdit) session.masterEdit = { partPath: part.partPath, slide }
    }
    return items.length ? { items } : null
  }

  private masterOpen(session: SlidesSession, partPath: string): RenderSlide | null {
    const slide = parseMasterPart(session.opened.archive, partPath)
    if (!slide) return null
    session.masterEdit = { partPath, slide }
    return this.renderMaster(session)
  }

  private masterEditText(session: SlidesSession, op: MasterEditTextOp): RenderSlide | null {
    const master = session.masterEdit
    const element = master ? findTextElement(master.slide, op.sourceId) : null
    if (!master || !element?.text) return null
    pushHistory(session)
    element.text.paragraphs = applyEditParagraphs(element.text.paragraphs, op.paragraphs)
    element.dirty = true
    this.commitMaster(session)
    return this.renderMaster(session)
  }

  private masterEditTransform(
    session: SlidesSession,
    op: MasterEditTransformOp,
  ): RenderSlide | null {
    const master = session.masterEdit
    const element = master?.slide.elements.find((candidate) => candidate.id === op.sourceId)
    if (!master || !element) return null
    if (op.preview) {
      if (!session.transformPreview) {
        pushHistory(session)
        session.transformPreview = true
      }
    } else if (session.transformPreview) session.transformPreview = false
    else pushHistory(session)
    element.transform = {
      ...element.transform,
      offset: {
        x: toEmu(session, op.fitWidthPx, op.xPx),
        y: toEmu(session, op.fitWidthPx, op.yPx),
        cx: toEmu(session, op.fitWidthPx, op.wPx),
        cy: toEmu(session, op.fitWidthPx, op.hPx),
      },
      rot: Math.round(op.rotationDeg * 60000),
    }
    element.dirtyTransform = true
    if (!op.preview) this.commitMaster(session)
    return this.renderMaster(session)
  }

  private masterEditFill(session: SlidesSession, op: MasterEditFillOp): RenderSlide | null {
    const master = session.masterEdit
    const element = master?.slide.elements.find((candidate) => candidate.id === op.sourceId)
    if (!master || !element || (element.type !== 'text' && element.type !== 'shape')) return null
    pushHistory(session)
    if (typeof op.fill === 'string') {
      element.fill = op.fill === 'none' ? { type: 'none' } : { type: 'solid', color: op.fill }
    } else {
      element.fill = {
        type: 'gradient',
        stops: [
          { pos: 0, color: op.fill.gradient.from },
          { pos: 1, color: op.fill.gradient.to },
        ],
        ...(op.fill.gradient.radial
          ? { path: 'circle' as const }
          : { angle: Math.round((op.fill.gradient.angleDeg ?? 0) * 60000) }),
      }
    }
    element.dirtyFill = true
    this.commitMaster(session)
    return this.renderMaster(session)
  }

  private masterEditStroke(session: SlidesSession, op: MasterEditStrokeOp): RenderSlide | null {
    const master = session.masterEdit
    const element = master?.slide.elements.find((candidate) => candidate.id === op.sourceId)
    if (!master || !element || (element.type !== 'text' && element.type !== 'shape')) return null
    pushHistory(session)
    element.stroke = op.stroke
      ? {
          fill: { type: 'solid', color: op.stroke.color },
          width: Math.round(op.stroke.widthPt * EMU_PER_PT),
        }
      : undefined
    element.dirtyStroke = true
    this.commitMaster(session)
    return this.renderMaster(session)
  }

  private masterDeleteElement(
    session: SlidesSession,
    op: MasterDeleteElementOp,
  ): RenderSlide | null {
    const master = session.masterEdit
    if (!master) return null
    pushHistory(session)
    if (!deleteElement(master.slide, op.sourceId)) {
      session.undo.pop()
      return null
    }
    this.commitMaster(session)
    return this.renderMaster(session)
  }

  private setNotes(session: SlidesSession, op: SetNotesOp): boolean {
    pushHistory(session)
    const result = setSlideNotes(session.opened, op.slideIndex, op.text)
    if (!result) session.undo.pop()
    else session.metaDirty = true
    return result
  }

  private addComment(session: SlidesSession, op: AddCommentOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!addSlideComment(session.opened, op.slideIndex, { author: 'Web User', text: op.text })) {
      session.undo.pop()
      return null
    }
    session.metaDirty = true
    return getSlideComments(session.opened.archive, slide.path)
  }

  private deleteComment(session: SlidesSession, op: DeleteCommentOp): unknown {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    if (!deleteSlideComment(session.opened, op.slideIndex, op)) {
      session.undo.pop()
      return null
    }
    session.metaDirty = true
    return getSlideComments(session.opened.archive, slide.path)
  }

  private setTransition(session: SlidesSession, op: SetTransitionOp): boolean {
    const targets =
      op.slideIndex === -1
        ? session.opened.deck.slides
        : [session.opened.deck.slides[op.slideIndex]].filter((slide): slide is Slide => !!slide)
    if (targets.length === 0) return false
    pushHistory(session)
    for (const slide of targets) setSlideTransition(slide, op.kind)
    return true
  }

  private setAnimations(session: SlidesSession, op: SetAnimationsOp): boolean {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return false
    const animations: SlideAnimation[] = []
    for (const item of op.items) {
      const element = slide.elements.find((candidate) => candidate.id === item.sourceId)
      const spid = element ? elementSpid(element) : null
      if (spid === null) continue
      animations.push({
        spid,
        effect: item.effect,
        trigger: item.trigger,
        durationMs: Math.max(0, Math.round(item.durationMs)),
        delayMs: Math.max(0, Math.round(item.delayMs)),
        ...(item.motionPath === undefined ? {} : { motionPath: item.motionPath }),
        ...(item.paragraph === undefined ? {} : { paragraph: item.paragraph }),
      })
    }
    pushHistory(session)
    setSlideAnimations(slide, animations)
    return true
  }

  private setAdvanceTimes(session: SlidesSession, op: SetAdvanceTimesOp): boolean {
    const valid = op.times.filter(({ slideIndex }) => session.opened.deck.slides[slideIndex])
    if (valid.length === 0) return false
    pushHistory(session)
    for (const { slideIndex, ms } of valid) {
      const slide = session.opened.deck.slides[slideIndex]!
      setSlideAdvanceTime(slide, ms)
    }
    return true
  }

  private setSlideHidden(session: SlidesSession, op: SetSlideHiddenOp): RenderSlide | null {
    const slide = session.opened.deck.slides[op.slideIndex]
    if (!slide) return null
    pushHistory(session)
    setSlideHidden(slide, op.hidden)
    return renderSlide(session, op.slideIndex)
  }

  private getAnimations(session: SlidesSession, slideIndex: number): AnimationItem[] {
    const slide = session.opened.deck.slides[slideIndex]
    if (!slide) return []
    const bySpid = new Map<number, (typeof slide.elements)[number]>()
    for (const element of slide.elements) {
      const spid = elementSpid(element)
      if (spid !== null && !bySpid.has(spid)) bySpid.set(spid, element)
    }
    return getSlideAnimations(slide).flatMap((animation) => {
      const element = bySpid.get(animation.spid)
      return element
        ? [
            {
              sourceId: element.id,
              targetName: element.name || element.type,
              effect: animation.effect,
              trigger: animation.trigger,
              durationMs: animation.durationMs,
              delayMs: animation.delayMs,
              ...(animation.motionPath === undefined ? {} : { motionPath: animation.motionPath }),
              ...(animation.paragraph === undefined ? {} : { paragraph: animation.paragraph }),
            },
          ]
        : []
    })
  }

  private setSections(session: SlidesSession, sections: SectionInfo[]): SectionInfo[] {
    pushHistory(session)
    setSections(session.opened, sections)
    session.metaDirty = true
    return getSections(session.opened)
  }

  private addSection(session: SlidesSession, op: AddSectionOp): SectionInfo[] | null {
    pushHistory(session)
    const result = addSection(session.opened, op.atSlideIndex, op.name)
    if (!result) session.undo.pop()
    else session.metaDirty = true
    return result
  }

  private renameSection(session: SlidesSession, op: RenameSectionOp): SectionInfo[] | null {
    pushHistory(session)
    const result = renameSection(session.opened, op.id, op.name)
    if (!result) session.undo.pop()
    else session.metaDirty = true
    return result
  }

  private removeSection(session: SlidesSession, op: RemoveSectionOp): SectionInfo[] | null {
    pushHistory(session)
    const result = removeSection(session.opened, op.id, { keepSlides: true })
    if (!result) session.undo.pop()
    else session.metaDirty = true
    return result
  }

  private moveSection(
    session: SlidesSession,
    op: MoveSectionOp,
  ): { slides: RenderSlide[]; sections: SectionInfo[] } | null {
    pushHistory(session)
    const sections = moveSection(session.opened, op.id, op.dir)
    if (!sections) {
      session.undo.pop()
      return null
    }
    session.metaDirty = true
    return { slides: renderSlides(session), sections }
  }

  private undo(session: SlidesSession): RenderSlide[] | null {
    if (session.masterEdit) return null
    if (session.historyBatch) endHistoryBatch(session)
    const previous = session.undo.pop()
    if (!previous) return null
    session.redo.push(snapshot(session))
    restore(session, previous)
    return renderSlides(session)
  }

  private redo(session: SlidesSession): RenderSlide[] | null {
    if (session.masterEdit || session.historyBatch) return null
    const next = session.redo.pop()
    if (!next) return null
    session.undo.push(snapshot(session))
    restore(session, next)
    return renderSlides(session)
  }

  private restoreAiSnapshot(session: SlidesSession, id: number): RenderSlide[] | null {
    if (session.masterEdit || session.historyBatch) return null
    const value = session.aiSnapshots.get(id)
    if (!value) return null
    pushHistory(session)
    restore(session, value)
    session.aiSnapshots.delete(id)
    return renderSlides(session)
  }

  private requireSession(sessionId: string): SlidesSession {
    this.removeExpiredSessions()
    const session = this.sessions.get(sessionId)
    if (!session) throw new Error('演示文稿会话不存在或已过期')
    session.lastAccessAt = Date.now()
    return session
  }

  private prepareSessionSlot(): void {
    this.removeExpiredSessions()
    if (this.sessions.size >= this.maxSessions) throw new Error('演示文稿服务繁忙，请稍后重试')
  }

  private removeExpiredSessions(): void {
    const cutoff = Date.now() - this.sessionTtlMs
    for (const [id, session] of this.sessions) {
      if (session.lastAccessAt < cutoff) this.sessions.delete(id)
    }
  }
}
