import { useEffect, useRef, useState } from 'react'
import type { PointerEvent as ReactPointerEvent, ReactElement } from 'react'
import { AgentLoop } from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { AiComposer, AiTypingIndicator } from '@genoffice/ui'
import { aiLangDirective, t as tGlobal, useI18n } from '../i18n/locale'
import { Markdown } from '@genoffice/ui'
import sendEnterOn from '../assets/send-enter-on.png'
import sendEnterOff from '../assets/send-enter-off.png'
import sendStop from '../assets/send-stop.png'
import { createPdfSkill } from './pdf-skill'
import { createElectronTransport } from './transport'
import { executePdfTool, type PdfAiDeps } from './tools'
import {
  auditNumbersInSearchIndex,
  type NumberAuditFinding,
  type NumberAuditReport,
} from '../number-audit'
import { auditConsistencyInSearchIndex } from '../consistency-audit'
import { consistencyAuditCopy } from './consistency-audit-copy'
import { auditDocumentInSearchIndex } from '../document-audit'
import { documentAuditCopy } from './document-audit-copy'

const PANEL_WIDTH_KEY = 'pdf-ai-panel-width'
const PANEL_WIDTH_DEFAULT = 360
const PANEL_WIDTH_MIN = 280

function clampPanelWidth(w: number): number {
  return Math.min(Math.max(w, PANEL_WIDTH_MIN), Math.min(720, Math.round(window.innerWidth * 0.6)))
}

function loadPanelWidth(): number {
  const saved = Number(localStorage.getItem(PANEL_WIDTH_KEY))
  return Number.isFinite(saved) && saved > 0 ? clampPanelWidth(saved) : PANEL_WIDTH_DEFAULT
}

interface ToolActivity {
  name: string
  summary: string
  isError?: boolean
  output?: string
}

interface ChatEntry {
  role: 'user' | 'assistant'
  text: string
  streaming?: boolean
  isError?: boolean
  /** the run failed and this user message was rolled back out of the model context */
  undelivered?: boolean
  tools?: ToolActivity[]
}

type Phase = 'thinking' | 'replying' | 'working'

interface NumberAuditCopy {
  action: string
  scanSummary(report: NumberAuditReport): string
  commentSummary(count: number): string
  result(report: NumberAuditReport, commentsAdded: number, readOnly: boolean): string
  subject(finding: NumberAuditFinding): string
  comment(finding: NumberAuditFinding): string
}

function compactFindingText(value: string): string {
  return value.replace(/\s+/gu, ' ').trim().slice(0, 160).replace(/`/gu, '\\`')
}

function zhFindingLine(finding: NumberAuditFinding, traditional: boolean): string {
  const page = finding.pageIndex + 1
  if (finding.kind === 'arithmetic') {
    return `- 第 ${page} ${traditional ? '頁' : '页'}：\`${compactFindingText(finding.anchorText)}\` ${traditional ? '的結果為' : '的结果为'} ${finding.stated}，${traditional ? '應為' : '应为'} ${finding.expected}。`
  }
  if (finding.kind === 'consistency') {
    return `- 第 ${page} ${traditional ? '頁' : '页'}："${finding.label}" ${traditional ? '為' : '为'} ${finding.stated}，${traditional ? '與' : '与'}第 ${(finding.canonicalPageIndex ?? 0) + 1} ${traditional ? '頁的' : '页的'} ${finding.expected} ${traditional ? '不一致' : '不一致'}。`
  }
  const location = `${traditional ? '表格' : '表格'} ${finding.tableNumber} ${traditional ? '第' : '第'} ${finding.rowNumber} ${traditional ? '列' : '行'}`
  return finding.kind === 'tableFormula'
    ? `- 第 ${page} ${traditional ? '頁' : '页'} ${location}：${finding.label ?? (traditional ? '金額' : '金额')} ${traditional ? '為' : '为'} ${finding.stated}，${traditional ? '依數量與單價計算應為' : '按数量与单价计算应为'} ${finding.expected}。`
    : `- 第 ${page} ${traditional ? '頁' : '页'} ${location}：${finding.label ?? (traditional ? '合計' : '合计')} ${traditional ? '為' : '为'} ${finding.stated}，${traditional ? '依明細加總應為' : '按明细加总应为'} ${finding.expected}。`
}

function enFindingLine(finding: NumberAuditFinding): string {
  const page = finding.pageIndex + 1
  if (finding.kind === 'arithmetic') {
    return `- Page ${page}: \`${compactFindingText(finding.anchorText)}\` states ${finding.stated}; expected ${finding.expected}.`
  }
  if (finding.kind === 'consistency') {
    return `- Page ${page}: "${finding.label}" is ${finding.stated}, conflicting with ${finding.expected} on page ${(finding.canonicalPageIndex ?? 0) + 1}.`
  }
  const location = `table ${finding.tableNumber}, row ${finding.rowNumber}`
  return finding.kind === 'tableFormula'
    ? `- Page ${page}, ${location}: ${finding.label ?? 'amount'} is ${finding.stated}; quantity × unit price gives ${finding.expected}.`
    : `- Page ${page}, ${location}: ${finding.label ?? 'total'} is ${finding.stated}; the detail rows sum to ${finding.expected}.`
}

function zhFindingSubject(finding: NumberAuditFinding, traditional: boolean): string {
  if (finding.kind === 'arithmetic') return traditional ? '數字計算不一致' : '数字计算不一致'
  if (finding.kind === 'consistency') return traditional ? '跨頁數字不一致' : '跨页数字不一致'
  if (finding.kind === 'tableFormula')
    return traditional ? '表格金額計算不一致' : '表格金额计算不一致'
  return traditional ? '表格合計不一致' : '表格合计不一致'
}

function zhFindingComment(finding: NumberAuditFinding, traditional: boolean): string {
  if (finding.kind === 'arithmetic') {
    return traditional
      ? `此算式的明確結果為 ${finding.stated}，按原式計算應為 ${finding.expected}。請核對並修正。`
      : `该算式的明示结果为 ${finding.stated}，按原式计算应为 ${finding.expected}。请核对并更正。`
  }
  if (finding.kind === 'consistency') {
    return traditional
      ? `「${finding.label}」在第 ${(finding.canonicalPageIndex ?? 0) + 1} 頁為 ${finding.expected}，本頁為 ${finding.stated}。請確認版本和統計口徑。`
      : `“${finding.label}”在第 ${(finding.canonicalPageIndex ?? 0) + 1} 页为 ${finding.expected}，本页为 ${finding.stated}。请确认版本和统计口径。`
  }
  if (finding.kind === 'tableFormula') {
    return traditional
      ? `表格第 ${finding.rowNumber} 列的${finding.label ?? '金額'}為 ${finding.stated}，依數量與單價計算應為 ${finding.expected}。請核對明細。`
      : `表格第 ${finding.rowNumber} 行的${finding.label ?? '金额'}为 ${finding.stated}，按数量与单价计算应为 ${finding.expected}。请核对明细。`
  }
  return traditional
    ? `表格第 ${finding.rowNumber} 列的${finding.label ?? '合計'}為 ${finding.stated}，依明細加總應為 ${finding.expected}。請核對合計公式。`
    : `表格第 ${finding.rowNumber} 行的${finding.label ?? '合计'}为 ${finding.stated}，按明细加总应为 ${finding.expected}。请核对合计公式。`
}

function enFindingSubject(finding: NumberAuditFinding): string {
  if (finding.kind === 'arithmetic') return 'Arithmetic mismatch'
  if (finding.kind === 'consistency') return 'Cross-page number mismatch'
  if (finding.kind === 'tableFormula') return 'Table amount mismatch'
  return 'Table total mismatch'
}

function enFindingComment(finding: NumberAuditFinding): string {
  if (finding.kind === 'arithmetic') {
    return `This expression states ${finding.stated}, but evaluates to ${finding.expected}. Verify and correct the result.`
  }
  if (finding.kind === 'consistency') {
    return `“${finding.label}” is ${finding.expected} on page ${(finding.canonicalPageIndex ?? 0) + 1} and ${finding.stated} here. Confirm the version and reporting basis.`
  }
  if (finding.kind === 'tableFormula') {
    return `Row ${finding.rowNumber} states ${finding.stated} for ${finding.label ?? 'amount'}, but quantity × unit price gives ${finding.expected}. Verify the detail row.`
  }
  return `Row ${finding.rowNumber} states ${finding.stated} for ${finding.label ?? 'total'}, but the detail rows sum to ${finding.expected}. Verify the total formula.`
}

function numberAuditCopy(lang: string): NumberAuditCopy {
  if (lang === 'zh') {
    return {
      action: '数据核验',
      scanSummary: (report) => `本地核验 ${report.pagesExamined} 页`,
      commentSummary: (count) => `添加 ${count} 条核验批注`,
      result: (report, commentsAdded, readOnly) => {
        const header = `### 数据核验完成\n\n已检查 ${report.pagesExamined} 页、${report.expressionsChecked} 个显式算式、${report.namedFigures} 个命名指标和 ${report.tablesExamined} 个表格（${report.tableChecks} 项表格校验）。`
        if (report.findings.length === 0) return `${header}\n\n未发现可确定验证的数字异常。`
        const mutation = readOnly
          ? '当前文档为只读，未写入批注。'
          : `已添加 ${commentsAdded} 条便签批注。`
        const lines = report.findings.slice(0, 30).map((finding) => zhFindingLine(finding, false))
        if (report.truncated || report.findings.length > 30) lines.push('- 其余异常已省略。')
        return `${header}\n\n发现 ${report.findings.length} 个异常，${mutation}\n\n${lines.join('\n')}`
      },
      subject: (finding) => zhFindingSubject(finding, false),
      comment: (finding) => zhFindingComment(finding, false),
    }
  }
  if (lang === 'zh-TW') {
    return {
      action: '數據核驗',
      scanSummary: (report) => `本機核驗 ${report.pagesExamined} 頁`,
      commentSummary: (count) => `新增 ${count} 則核驗註解`,
      result: (report, commentsAdded, readOnly) => {
        const header = `### 數據核驗完成\n\n已檢查 ${report.pagesExamined} 頁、${report.expressionsChecked} 個明確算式、${report.namedFigures} 個命名指標和 ${report.tablesExamined} 個表格（${report.tableChecks} 項表格核驗）。`
        if (report.findings.length === 0) return `${header}\n\n未發現可確定驗證的數字異常。`
        const mutation = readOnly
          ? '目前文件為唯讀，未寫入註解。'
          : `已新增 ${commentsAdded} 則便利貼註解。`
        const lines = report.findings.slice(0, 30).map((finding) => zhFindingLine(finding, true))
        if (report.truncated || report.findings.length > 30) lines.push('- 其餘異常已省略。')
        return `${header}\n\n發現 ${report.findings.length} 個異常，${mutation}\n\n${lines.join('\n')}`
      },
      subject: (finding) => zhFindingSubject(finding, true),
      comment: (finding) => zhFindingComment(finding, true),
    }
  }
  return {
    action: 'Check numbers',
    scanSummary: (report) => `Checked ${report.pagesExamined} pages locally`,
    commentSummary: (count) => `Added ${count} audit comments`,
    result: (report, commentsAdded, readOnly) => {
      const header = `### Number check complete\n\nChecked ${report.pagesExamined} pages, ${report.expressionsChecked} explicit expressions, ${report.namedFigures} named figures, and ${report.tablesExamined} tables (${report.tableChecks} table checks).`
      if (report.findings.length === 0)
        return `${header}\n\nNo deterministically verifiable number issues were found.`
      const mutation = readOnly
        ? 'The document is read-only, so no comments were added.'
        : `Added ${commentsAdded} sticky-note comments.`
      const lines = report.findings.slice(0, 30).map(enFindingLine)
      if (report.truncated || report.findings.length > 30)
        lines.push('- Additional findings omitted.')
      return `${header}\n\nFound ${report.findings.length} issues. ${mutation}\n\n${lines.join('\n')}`
    },
    subject: enFindingSubject,
    comment: enFindingComment,
  }
}

export function AiPanel({
  api,
  onCollapse,
  preset,
}: {
  api: PdfAiDeps
  onCollapse: () => void
  /** Ribbon AI buttons push a one-shot prompt; a new nonce triggers an auto-run */
  preset?: { text: string; nonce: number } | null
}): ReactElement {
  const { lang, t } = useI18n()
  const [chat, setChat] = useState<ChatEntry[]>([])
  const [prompt, setPrompt] = useState('')
  const [busy, setBusy] = useState(false)
  const [phase, setPhase] = useState<Phase>('thinking')
  const chatRef = useRef<HTMLDivElement>(null)
  const stickToBottomRef = useRef(true)
  const [panelWidth, setPanelWidth] = useState(loadPanelWidth)
  const [resizing, setResizing] = useState(false)
  const asideRef = useRef<HTMLElement>(null)
  const localAuditAbortRef = useRef<AbortController | null>(null)

  useEffect(() => () => localAuditAbortRef.current?.abort(), [])

  // The .ai-dock wrapper owns the animated width (docs-style 180ms slide);
  // it tracks the resizable panel width through this variable
  useEffect(() => {
    const dock = asideRef.current?.closest('.ai-dock') as HTMLElement | null
    dock?.style.setProperty('--ai-panel-width', `${panelWidth}px`)
  }, [panelWidth])
  const settingsRef = useRef<AiSettings | null>(null)
  const langRef = useRef(lang)
  langRef.current = lang
  const apiRef = useRef(api)
  apiRef.current = api

  const patchLast = (patch: Partial<ChatEntry> | ((last: ChatEntry) => Partial<ChatEntry>)) => {
    setChat((prev) => {
      const next = [...prev]
      const last = next[next.length - 1]
      if (!last || last.role !== 'assistant') return prev
      next[next.length - 1] = { ...last, ...(typeof patch === 'function' ? patch(last) : patch) }
      return next
    })
  }

  // The loop is built once; every mutable value goes through a ref getter
  const loopRef = useRef<AgentLoop | null>(null)
  if (!loopRef.current) {
    const deps: PdfAiDeps = {
      doc: () => apiRef.current.doc(),
      fileName: () => apiRef.current.fileName(),
      pageCount: () => apiRef.current.pageCount(),
      currentPage: () => apiRef.current.currentPage(),
      readOnly: () => apiRef.current.readOnly(),
      outline: () => apiRef.current.outline(),
      searchIndex: () => apiRef.current.searchIndex(),
      isDeleted: (i) => apiRef.current.isDeleted(i),
      gotoPage: (p) => apiRef.current.gotoPage(p),
      addMarkup: (type, idx, rects) => apiRef.current.addMarkup(type, idx, rects),
      addReviewComment: (comment) => apiRef.current.addReviewComment(comment),
      setClassification: (metadata) => apiRef.current.setClassification(metadata),
      createPdfDocument: (document) => apiRef.current.createPdfDocument(document),
      editText: (input) => apiRef.current.editText(input),
      editFonts: () => apiRef.current.editFonts(),
      formEdits: () => apiRef.current.formEdits(),
      applyFormEdit: (v) => apiRef.current.applyFormEdit(v),
      rotatePage: (idx, dir) => apiRef.current.rotatePage(idx, dir),
      deletePage: (idx) => apiRef.current.deletePage(idx),
      pageGeom: (idx) => apiRef.current.pageGeom(idx),
      listImages: () => apiRef.current.listImages(),
      isImageClaimed: (ref) => apiRef.current.isImageClaimed(ref),
      insertImage: (idx, png, rect, layer) => apiRef.current.insertImage(idx, png, rect, layer),
      transformImage: (ref, rect, layer, quarterTurns) =>
        apiRef.current.transformImage(ref, rect, layer, quarterTurns),
      replaceImage: (ref, png) => apiRef.current.replaceImage(ref, png),
      deleteImage: (ref) => apiRef.current.deleteImage(ref),
      searchImages: (query, max) => apiRef.current.searchImages(query, max),
      generateImage: (op) => apiRef.current.generateImage(op),
      fetchImage: (url) => apiRef.current.fetchImage(url),
    }
    loopRef.current = new AgentLoop({
      transport: createElectronTransport(() => settingsRef.current!),
      skill: createPdfSkill(deps),
      systemSuffix: () => aiLangDirective(langRef.current),
      events: {
        onText: (text) => {
          setPhase('replying')
          patchLast({ text })
        },
        onToolExecuted: ({ call, execution }) => {
          setPhase('working')
          patchLast((last) => ({
            tools: [
              ...(last.tools ?? []),
              {
                name: call.name,
                summary: execution.summary,
                isError: execution.isError,
                output: execution.output?.slice(0, 2000),
              },
            ],
          }))
        },
        onTurnEnd: () => {
          setPhase('thinking')
          patchLast({ streaming: false })
          setChat((prev) => [...prev, { role: 'assistant', text: '', streaming: true }])
        },
        onDone: ({ text, cancelled, turnLimit }) => {
          const final = turnLimit
            ? [text, tGlobal('aiTurnLimit')].filter(Boolean).join('\n\n')
            : text || (cancelled ? tGlobal('aiStopped') : '')
          patchLast((last) => ({
            streaming: false,
            text: final || (last.tools?.length ? last.text : tGlobal('aiNoReply')),
          }))
          setBusy(false)
        },
        onError: (error) => {
          setChat((prev) => {
            const next = [...prev]
            // the loop rolled this run's user message out of the model context — surface that
            for (let i = next.length - 1; i >= 0; i--) {
              const entry = next[i]!
              if (entry.role === 'user') {
                next[i] = { ...entry, undelivered: true }
                break
              }
            }
            const last = next.at(-1)
            if (last?.role === 'assistant') {
              next[next.length - 1] = { ...last, streaming: false, text: error, isError: true }
            }
            return next
          })
          setBusy(false)
        },
      },
    })
  }

  useEffect(() => {
    if (stickToBottomRef.current) {
      chatRef.current?.scrollTo({ top: chatRef.current.scrollHeight })
    }
  }, [chat, busy])

  const onChatScroll = (): void => {
    const el = chatRef.current
    if (!el) return
    stickToBottomRef.current = el.scrollHeight - el.scrollTop - el.clientHeight < 48
  }

  const send = (text: string): void => {
    const instruction = text.trim()
    const loop = loopRef.current
    if (!instruction || !loop || loop.busy) return
    stickToBottomRef.current = true
    setChat((prev) => [
      ...prev,
      { role: 'user', text: instruction },
      { role: 'assistant', text: '', streaming: true },
    ])
    setPrompt('')
    setBusy(true)
    setPhase('thinking')
    void (async () => {
      try {
        settingsRef.current = await window.pdfApi.getAiSettings()
        await loop.run(instruction)
      } catch (err) {
        patchLast({
          streaming: false,
          text: err instanceof Error ? err.message : String(err),
          isError: true,
        })
        setBusy(false)
      }
    })()
  }

  const runNumberAudit = (): void => {
    const indexPromise = apiRef.current.searchIndex()
    if (!indexPromise || busy) return
    const copy = numberAuditCopy(langRef.current)
    const controller = new AbortController()
    localAuditAbortRef.current = controller
    stickToBottomRef.current = true
    setChat((prev) => [
      ...prev,
      { role: 'user', text: copy.action },
      { role: 'assistant', text: '', streaming: true },
    ])
    setBusy(true)
    setPhase('working')
    void (async () => {
      try {
        const index = await indexPromise
        if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError')
        const report = auditNumbersInSearchIndex(index)
        const tools: ToolActivity[] = [
          {
            name: 'local_number_audit',
            summary: copy.scanSummary(report),
            output: JSON.stringify(report, null, 2).slice(0, 2_000),
          },
        ]
        let commentsAdded = 0
        if (!apiRef.current.readOnly()) {
          for (let index = 0; index < report.findings.length; index += 1) {
            if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError')
            const finding = report.findings[index]!
            const execution = await executePdfTool(
              apiRef.current,
              {
                id: `local-number-audit-${index}`,
                name: 'add_review_comment',
                input: {
                  page: finding.pageIndex + 1,
                  anchor_text: finding.anchorText,
                  occurrence: finding.occurrence,
                  subject: copy.subject(finding),
                  comment: copy.comment(finding),
                  author: 'GenOffice Number Audit',
                },
              },
              controller.signal,
            )
            if (!execution.isError) commentsAdded += 1
          }
          if (commentsAdded > 0) {
            tools.push({
              name: 'add_review_comment',
              summary: copy.commentSummary(commentsAdded),
            })
            apiRef.current.gotoPage(report.findings[0]!.pageIndex + 1)
          }
        }
        patchLast({
          streaming: false,
          text: copy.result(report, commentsAdded, apiRef.current.readOnly()),
          tools,
        })
      } catch (error) {
        patchLast({
          streaming: false,
          text:
            error instanceof DOMException && error.name === 'AbortError'
              ? tGlobal('aiStopped')
              : error instanceof Error
                ? error.message
                : String(error),
          isError: !(error instanceof DOMException && error.name === 'AbortError'),
        })
      } finally {
        if (localAuditAbortRef.current === controller) localAuditAbortRef.current = null
        setBusy(false)
        setPhase('thinking')
      }
    })()
  }

  const runConsistencyAudit = (): void => {
    const indexPromise = apiRef.current.searchIndex()
    if (!indexPromise || busy) return
    const copy = consistencyAuditCopy(langRef.current)
    const controller = new AbortController()
    localAuditAbortRef.current = controller
    stickToBottomRef.current = true
    setChat((prev) => [
      ...prev,
      { role: 'user', text: copy.action },
      { role: 'assistant', text: '', streaming: true },
    ])
    setBusy(true)
    setPhase('working')
    void (async () => {
      try {
        const index = await indexPromise
        if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError')
        const report = auditConsistencyInSearchIndex(index)
        const tools: ToolActivity[] = [
          {
            name: 'local_consistency_audit',
            summary: copy.scanSummary(report),
            output: JSON.stringify(report, null, 2).slice(0, 2_000),
          },
        ]
        let commentsAdded = 0
        if (!apiRef.current.readOnly()) {
          for (let findingIndex = 0; findingIndex < report.findings.length; findingIndex += 1) {
            const finding = report.findings[findingIndex]!
            for (const [claimIndex, claim] of [finding.first, finding.second].entries()) {
              if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError')
              const execution = await executePdfTool(
                apiRef.current,
                {
                  id: `local-consistency-audit-${findingIndex}-${claimIndex}`,
                  name: 'add_review_comment',
                  input: {
                    page: claim.pageIndex + 1,
                    anchor_text: claim.anchorText,
                    occurrence: claim.occurrence,
                    subject: copy.subject(finding),
                    comment: copy.comment(finding, claim),
                    author: 'GenOffice Consistency Audit',
                  },
                },
                controller.signal,
              )
              if (!execution.isError) commentsAdded += 1
            }
          }
          if (commentsAdded > 0) {
            tools.push({
              name: 'add_review_comment',
              summary: copy.commentSummary(commentsAdded),
            })
            apiRef.current.gotoPage(report.findings[0]!.first.pageIndex + 1)
          }
        }
        patchLast({
          streaming: false,
          text: copy.result(report, commentsAdded, apiRef.current.readOnly()),
          tools,
        })
      } catch (error) {
        patchLast({
          streaming: false,
          text:
            error instanceof DOMException && error.name === 'AbortError'
              ? tGlobal('aiStopped')
              : error instanceof Error
                ? error.message
                : String(error),
          isError: !(error instanceof DOMException && error.name === 'AbortError'),
        })
      } finally {
        if (localAuditAbortRef.current === controller) localAuditAbortRef.current = null
        setBusy(false)
        setPhase('thinking')
      }
    })()
  }

  const runDocumentAudit = (): void => {
    const indexPromise = apiRef.current.searchIndex()
    if (!indexPromise || busy) return
    const copy = documentAuditCopy(langRef.current)
    const controller = new AbortController()
    localAuditAbortRef.current = controller
    stickToBottomRef.current = true
    setChat((prev) => [
      ...prev,
      { role: 'user', text: copy.action },
      { role: 'assistant', text: '', streaming: true },
    ])
    setBusy(true)
    setPhase('working')
    void (async () => {
      try {
        const index = await indexPromise
        if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError')
        const report = auditDocumentInSearchIndex(index)
        const tools: ToolActivity[] = [
          {
            name: 'local_document_audit',
            summary: copy.scanSummary(report),
            output: JSON.stringify(
              {
                ...report,
                sensitiveFindings: report.sensitiveFindings.map(
                  ({ anchorText: _anchorText, ...finding }) => finding,
                ),
              },
              null,
              2,
            ).slice(0, 2_000),
          },
        ]
        let commentsAdded = 0
        if (!apiRef.current.readOnly()) {
          apiRef.current.setClassification({
            labels: report.classifications.map(({ id, name }) => ({ id, name })),
            sensitivity: report.sensitivity,
          })
          tools.push({
            name: 'set_document_classification',
            summary: copy.classificationSummary(report),
          })
          for (
            let findingIndex = 0;
            findingIndex < report.sensitiveFindings.length;
            findingIndex += 1
          ) {
            if (controller.signal.aborted) throw new DOMException('Aborted', 'AbortError')
            const finding = report.sensitiveFindings[findingIndex]!
            const execution = await executePdfTool(
              apiRef.current,
              {
                id: `local-document-audit-${findingIndex}`,
                name: 'add_review_comment',
                input: {
                  page: finding.pageIndex + 1,
                  anchor_text: finding.anchorText,
                  occurrence: finding.occurrence,
                  subject: copy.subject(finding),
                  comment: copy.comment(finding),
                  author: 'GenOffice Document Audit',
                },
              },
              controller.signal,
            )
            if (!execution.isError) commentsAdded += 1
          }
          if (commentsAdded > 0) {
            tools.push({ name: 'add_review_comment', summary: copy.commentSummary(commentsAdded) })
            apiRef.current.gotoPage(report.sensitiveFindings[0]!.pageIndex + 1)
          }
        }
        patchLast({
          streaming: false,
          text: copy.result(report, commentsAdded, apiRef.current.readOnly()),
          tools,
        })
      } catch (error) {
        patchLast({
          streaming: false,
          text:
            error instanceof DOMException && error.name === 'AbortError'
              ? tGlobal('aiStopped')
              : error instanceof Error
                ? error.message
                : String(error),
          isError: !(error instanceof DOMException && error.name === 'AbortError'),
        })
      } finally {
        if (localAuditAbortRef.current === controller) localAuditAbortRef.current = null
        setBusy(false)
        setPhase('thinking')
      }
    })()
  }

  const stop = (): void => {
    loopRef.current?.cancel()
    localAuditAbortRef.current?.abort()
  }

  // One-click AI actions from the ribbon (same pattern as the docs ribbon presets)
  useEffect(() => {
    if (preset) send(preset.text)
    // eslint-disable-next-line react-hooks/exhaustive-deps -- run once per nonce
  }, [preset?.nonce])

  // Re-clamp the persisted width when the window shrinks (max is 60% of the window)
  useEffect(() => {
    const onResize = (): void => setPanelWidth((w) => clampPanelWidth(w))
    window.addEventListener('resize', onResize)
    return () => window.removeEventListener('resize', onResize)
  }, [])

  const resizeCleanupRef = useRef<(() => void) | null>(null)
  useEffect(() => () => resizeCleanupRef.current?.(), [])

  /** Drag the right edge to resize: the panel is flush with the window's left edge, so width = clientX */
  const startResize = (e: ReactPointerEvent<HTMLDivElement>): void => {
    e.preventDefault()
    const resizer = e.currentTarget
    setResizing(true)
    document.body.style.cursor = 'col-resize'
    document.body.style.userSelect = 'none'
    const onMove = (ev: PointerEvent): void => {
      setPanelWidth(clampPanelWidth(ev.clientX))
    }
    let done = false
    const cleanup = (): void => {
      if (done) return
      done = true
      resizeCleanupRef.current = null
      window.removeEventListener('pointermove', onMove)
      window.removeEventListener('pointerup', cleanup)
      window.removeEventListener('pointercancel', cleanup)
      resizer.removeEventListener('lostpointercapture', cleanup)
      document.body.style.cursor = ''
      document.body.style.userSelect = ''
      setResizing(false)
      setPanelWidth((w) => {
        localStorage.setItem(PANEL_WIDTH_KEY, String(Math.round(w)))
        return w
      })
    }
    resizeCleanupRef.current = cleanup
    window.addEventListener('pointermove', onMove)
    window.addEventListener('pointerup', cleanup)
    window.addEventListener('pointercancel', cleanup)
    // lostpointercapture also fires if the resizer is unmounted mid-drag (panel collapse)
    resizer.addEventListener('lostpointercapture', cleanup)
    resizer.setPointerCapture(e.pointerId)
  }

  const typingLabel =
    phase === 'replying' ? t('aiReplying') : phase === 'working' ? t('aiWorking') : t('aiThinking')

  return (
    <aside
      ref={asideRef}
      className={`copilot${resizing ? ' ai-panel-resizing' : ''}`}
      style={{ width: '100%' }}
    >
      <div
        className="ai-panel-resizer"
        onPointerDown={startResize}
        role="separator"
        aria-orientation="vertical"
        aria-label="Genspark"
      />
      <header className="ai-panel-header">
        <span className="ai-panel-title">
          <GensparkMark size={22} />
          Genspark
        </span>
        <div className="ai-panel-header-actions">
          {chat.length > 0 && (
            <button
              className="ai-header-btn"
              onClick={() => {
                stop()
                loopRef.current?.reset()
                setBusy(false)
                setChat([])
              }}
              data-tip={t('aiNewChat')}
              aria-label={t('aiNewChat')}
            >
              <IconNewChat />
            </button>
          )}
          <button
            className="ai-header-btn"
            onClick={onCollapse}
            data-tip={t('aiCollapsePanel')}
            aria-label={t('aiCollapsePanel')}
          >
            <IconCollapse />
          </button>
        </div>
      </header>

      <div className="ai-chat" ref={chatRef} onScroll={onChatScroll}>
        {chat.length === 0 && (
          <div className="ai-chat-empty">
            <div className="ai-chat-empty-title">{t('aiEmptyTitle')}</div>
            <div className="ai-chat-empty-body">{t('aiEmptyBody')}</div>
            <div className="ai-quick-actions">
              <button
                className="ai-quick-btn"
                onClick={() =>
                  setPrompt(
                    lang === 'zh'
                      ? '请创建一份新的 PDF：'
                      : lang === 'zh-TW'
                        ? '請建立一份新的 PDF：'
                        : 'Create a new PDF: ',
                  )
                }
              >
                {lang === 'zh' ? '创建 PDF' : lang === 'zh-TW' ? '建立 PDF' : 'Create PDF'}
              </button>
              <button className="ai-quick-btn" onClick={() => send(t('aiQuickSummaryPrompt'))}>
                {t('aiQuickSummary')}
              </button>
              <button className="ai-quick-btn" onClick={() => send(t('aiQuickKeyPointsPrompt'))}>
                {t('aiQuickKeyPoints')}
              </button>
              <button className="ai-quick-btn" onClick={runNumberAudit}>
                {numberAuditCopy(lang).action}
              </button>
              <button className="ai-quick-btn" onClick={runConsistencyAudit}>
                {consistencyAuditCopy(lang).action}
              </button>
              <button className="ai-quick-btn" onClick={runDocumentAudit}>
                {documentAuditCopy(lang).action}
              </button>
              {!api.readOnly() && (
                <button
                  className="ai-quick-btn"
                  onClick={() =>
                    send(
                      lang === 'zh'
                        ? '请全面审阅这份 PDF，找出矛盾、事实风险、表述不清和待办事项。每个有明确原文依据的问题都添加一条便签批注，并在回复中按页码汇总。'
                        : lang === 'zh-TW'
                          ? '請全面審閱這份 PDF，找出矛盾、事實風險、表述不清和待辦事項。每個有明確原文依據的問題都新增一則便利貼註解，並在回覆中按頁碼彙總。'
                          : 'Review this PDF for contradictions, factual risks, unclear wording, and action items. Add one sticky-note comment for each finding supported by an exact passage, then summarize the findings by page.',
                    )
                  }
                >
                  {lang === 'zh'
                    ? '审阅并批注'
                    : lang === 'zh-TW'
                      ? '審閱並註解'
                      : 'Review & comment'}
                </button>
              )}
            </div>
          </div>
        )}
        {chat.map((entry, i) => {
          if (entry.role === 'user') {
            return (
              <div key={i} className="ai-msg ai-msg-user">
                {entry.text}
                {entry.undelivered && (
                  <div className="ai-msg-undelivered">
                    {t('aiUndelivered')}
                    {!busy && (
                      <button className="ai-retry-btn" onClick={() => send(entry.text)}>
                        {t('aiRetry')}
                      </button>
                    )}
                  </div>
                )}
              </div>
            )
          }
          const hasTools = (entry.tools?.length ?? 0) > 0
          if (!entry.text && !hasTools) return null
          return (
            <div
              key={i}
              className={`ai-msg ai-msg-assistant${entry.isError ? ' ai-msg-error' : ''}`}
            >
              {hasTools && <ToolChipList tools={entry.tools!} />}
              {entry.text && <Markdown text={entry.text} />}
            </div>
          )
        })}
        {/* In-progress state: a standalone three-dot row at the end of the stream, kept until done */}
        {busy && <AiTypingIndicator label={typingLabel} />}
      </div>

      <div className="ai-composer">
        <AiComposer
          value={prompt}
          busy={busy}
          placeholder={t('aiComposerPlaceholder')}
          hintIdle={t('aiHintIdle')}
          hintBusy={t('aiHintBusy')}
          sendLabel={t('aiSend')}
          stopLabel={t('aiStop')}
          iconOnly
          sendIconEnabled={<img src={sendEnterOn} alt="" aria-hidden />}
          sendIconDisabled={<img src={sendEnterOff} alt="" aria-hidden />}
          stopIcon={<img src={sendStop} alt="" aria-hidden />}
          onChange={setPrompt}
          onSend={() => send(prompt)}
          onStop={stop}
        />
      </div>
    </aside>
  )
}

/** Tool row list (unified with docs/slides/sheets): dot + summary, expandable details when there's output */
/** Step-row status icons (timeline glyphs: 14px in a 20px slot, 1.6 stroke) */
function StepIcon({ status }: { status: 'running' | 'done' | 'error' }) {
  if (status === 'running') {
    return (
      <svg
        viewBox="0 0 24 24"
        width="14"
        height="14"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden
      >
        <path d="M6.5 3.5h11M6.5 20.5h11M8 3.5v3.2c0 2.6 4 4.2 4 5.3 0 1.1 4 2.7 4 5.3v3.2M16 3.5v3.2c0 2.6-4 4.2-4 5.3 0 1.1-4 2.7-4 5.3v3.2" />
      </svg>
    )
  }
  if (status === 'error') {
    return (
      <svg
        viewBox="0 0 24 24"
        width="14"
        height="14"
        fill="none"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
        aria-hidden
      >
        <circle cx="12" cy="12" r="9" />
        <path d="m9.2 9.2 5.6 5.6M14.8 9.2l-5.6 5.6" />
      </svg>
    )
  }
  return (
    <svg
      viewBox="0 0 24 24"
      width="14"
      height="14"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.6"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden
    >
      <circle cx="12" cy="12" r="9" />
      <path d="m8.5 12.4 2.4 2.4 4.6-5" />
    </svg>
  )
}

/** Tool activity group: a single quiet summary row
 *  that auto-opens while tools run, auto-collapses into "Worked · N steps" when they finish,
 *  and a manual toggle that always wins. Rows inside are step rows with 1px connectors. */
function ToolChipList({ tools }: { tools: ToolActivity[] }) {
  const { t: tr } = useI18n()
  const [expanded, setExpanded] = useState<Set<number>>(new Set())
  const [userOpen, setUserOpen] = useState<boolean | null>(null)

  const toggle = (j: number) => {
    setExpanded((prev) => {
      const next = new Set(prev)
      if (next.has(j)) next.delete(j)
      else next.add(j)
      return next
    })
  }

  const open = userOpen ?? false
  const label = tr('aiWorkedSteps', { n: tools.length })

  return (
    <div className="ai-work-group">
      <button
        type="button"
        className={`ai-work-group-summary`}
        aria-expanded={open}
        onClick={() => setUserOpen(!open)}
      >
        <span className="ai-work-group-label">{label}</span>
        <span className={`ai-tool-chip-caret${open ? ' open' : ''}`} aria-hidden>
          ›
        </span>
      </button>
      <div className={`ai-work-group-body${open ? ' open' : ''}`}>
        <div className="ai-work-group-body-inner">
          {tools.map((tool, j) => {
            const hasOutput = !!tool.output
            const isOpen = expanded.has(j)
            const stepStatus = tool.isError ? 'error' : 'done'
            return (
              <div key={j} className="ai-step-row">
                <span className={`ai-step-icon ${stepStatus}`} aria-hidden>
                  <StepIcon status={stepStatus} />
                </span>
                <div className="ai-step-content">
                  {hasOutput ? (
                    <button
                      type="button"
                      className="ai-step-title clickable"
                      data-tip={tool.name}
                      aria-expanded={isOpen}
                      onClick={() => toggle(j)}
                    >
                      {tool.summary}
                    </button>
                  ) : (
                    <span className="ai-step-title" data-tip={tool.name}>
                      {tool.summary}
                    </span>
                  )}
                  {hasOutput && isOpen && (
                    <div className="ai-step-detail">
                      <div className="ai-tool-output">
                        <div className="ai-tool-output-pre">{tool.output}</div>
                      </div>
                    </div>
                  )}
                </div>
              </div>
            )
          })}
        </div>
      </div>
    </div>
  )
}

function Svg({ children }: { children: React.ReactNode }): ReactElement {
  return (
    <svg
      width="15"
      height="15"
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      {children}
    </svg>
  )
}

function IconNewChat(): ReactElement {
  return (
    <Svg>
      <path
        d="M13.5 7.2v-3A1.7 1.7 0 0 0 11.8 2.5H4.2a1.7 1.7 0 0 0-1.7 1.7v6.1a1.7 1.7 0 0 0 1.7 1.7h1.1v2l2.6-2h1.3"
        strokeLinejoin="round"
      />
      <path d="M12.2 9.4v4M10.2 11.4h4" />
    </Svg>
  )
}

/* Same glyph as the sheets IconCollapse (16×16 viewBox, 1.2/1.3 stroke), rendered at 15px */
function IconCollapse(): ReactElement {
  return (
    <svg
      width={15}
      height={15}
      viewBox="0 0 16 16"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.2"
      strokeLinecap="round"
      aria-hidden
    >
      {/* Mirrored: the AI panel docks on the LEFT, so the divider and arrow point left */}
      <rect x="1.5" y="2.5" width="13" height="11" rx="1" />
      <path d="M5.5 2.5v11" />
      <path d="M12.5 8H8.1M9.8 5.9 7.7 8l2.1 2.1" strokeWidth="1.3" strokeLinejoin="round" />
    </svg>
  )
}

/** Genspark brand mark (rounded-square sparkle badge), inline so it renders
 * crisply at device resolution instead of going through <img> rasterization */
export function GensparkMark({ size = 18 }: { size?: number }): React.JSX.Element {
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 130 130.025"
      fill="none"
      xmlns="http://www.w3.org/2000/svg"
      aria-hidden
    >
      <path
        d="M105.115 0H24.6428C11.0443 0 0 11.0686 0 24.6915V105.334C0 118.981 11.0199 130.025 24.6428 130.025H105.115C118.714 130.025 129.758 118.957 129.758 105.334V24.6915C129.758 11.0443 118.714 0 105.115 0ZM71.5201 35.2735C85.5078 33.1571 86.7729 31.9164 88.865 17.88C88.938 17.4421 89.3028 17.1259 89.7407 17.1259C90.1786 17.1259 90.5435 17.4421 90.6164 17.88C92.7328 31.8921 93.9735 33.1571 107.961 35.2735C108.399 35.3465 108.715 35.7114 108.715 36.1493C108.715 36.5871 108.399 36.952 107.961 37.025C93.9249 39.1414 92.7085 40.4064 90.5677 54.6131C90.5191 54.9537 90.2516 55.197 89.911 55.197C89.5704 55.197 89.3028 54.9537 89.2542 54.6131C87.1134 40.4064 85.5565 39.1658 71.4958 37.025C71.0579 36.952 70.7417 36.5871 70.7417 36.1493C70.7417 35.7114 71.0579 35.3465 71.4958 35.2735H71.5201ZM101.758 78.5261C101.758 78.8181 101.563 79.037 101.271 79.0856C92.3193 80.4236 91.5652 81.2264 90.2029 90.2759C90.1786 90.4948 89.9839 90.6408 89.7893 90.6408C89.5703 90.6408 89.4001 90.4948 89.3758 90.2759C88.0135 81.2507 87.0161 80.4479 78.0883 79.0856C77.7964 79.037 77.6017 78.7937 77.6017 78.5261C77.6017 78.2342 77.7964 78.0153 78.0883 77.9666C86.9918 76.6287 87.7703 75.8259 89.1326 66.898C89.1812 66.6061 89.4244 66.4115 89.692 66.4115C89.9839 66.4115 90.2028 66.6061 90.2515 66.898C91.5894 75.8259 92.3923 76.6043 101.296 77.9666C101.588 78.0153 101.782 78.2585 101.782 78.5261H101.758ZM16.5178 54.8077C16.5178 54.1023 17.0286 53.4941 17.7341 53.3968C40.1388 50.0154 42.1093 47.9963 45.4907 25.5672C45.588 24.8861 46.1961 24.3509 46.9016 24.3509C47.6071 24.3509 48.191 24.8617 48.3126 25.5672C51.694 47.9963 53.6887 50.0154 76.0691 53.3968C76.7503 53.4941 77.2855 54.1023 77.2855 54.8077C77.2855 55.5132 76.7746 56.1214 76.0691 56.2187C53.5914 59.6244 51.6696 61.6192 48.2639 84.3645C48.1909 84.8754 47.7287 85.2889 47.2179 85.2889C46.707 85.2889 46.2448 84.8997 46.1718 84.3645C42.7418 61.6435 40.2604 59.6244 17.7584 56.2187C17.0772 56.1214 16.542 55.5132 16.542 54.8077H16.5178ZM112.097 109.591C112.097 111.416 110.613 112.9 108.813 112.9H21.2614C19.4369 112.9 17.9774 111.416 17.9774 109.591V102.658C17.9774 100.834 19.4612 99.3497 21.2614 99.3497H108.813C110.637 99.3497 112.097 100.834 112.097 102.658V109.591Z"
        fill="currentColor"
      />
    </svg>
  )
}
