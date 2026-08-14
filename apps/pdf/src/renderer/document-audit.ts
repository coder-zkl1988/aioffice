import type { SearchIndex } from './search'

export type DocumentSensitivityLevel = 'standard' | 'internal' | 'confidential' | 'restricted'

export type SensitiveFindingKind =
  'email' | 'phone' | 'cnIdentity' | 'paymentCard' | 'ssn' | 'passport' | 'iban'

export interface DocumentClassification {
  id: string
  name: string
  evidenceCount: number
}

export interface SensitiveFinding {
  kind: SensitiveFindingKind
  pageIndex: number
  anchorText: string
  occurrence: number
  maskedValue: string
}

export interface DocumentAuditReport {
  pagesExamined: number
  classificationPageIndexes: number[]
  classifications: DocumentClassification[]
  sensitivity: DocumentSensitivityLevel
  explicitSensitivityMarkers: number
  sensitiveFindings: SensitiveFinding[]
  truncated: boolean
}

interface DocumentAuditOptions {
  maxFindings?: number
}

interface ClassificationRule {
  id: string
  name: string
  patterns: RegExp[]
  parents?: Array<Pick<DocumentClassification, 'id' | 'name'>>
}

interface SensitiveCandidate {
  kind: SensitiveFindingKind
  raw: string
  start: number
}

interface SensitiveDetector {
  kind: SensitiveFindingKind
  expression: RegExp
  value(match: RegExpExecArray): string | null
  valid(value: string): boolean
}

const CLASSIFICATION_RULES: ClassificationRule[] = [
  {
    id: 'proforma-invoice',
    name: 'Proforma invoice',
    patterns: [/^(?:proforma invoice|形式发票|形式發票)(?:\s*(?:no\.?|number|#|编号|編號).*)?$/iu],
    parents: [{ id: 'invoice', name: 'Invoice' }],
  },
  {
    id: 'invoice',
    name: 'Invoice',
    patterns: [
      /^(?:(?:tax|commercial)\s+)?invoice(?:\s*(?:no\.?|number|#).*)?$/iu,
      /^(?:发票|發票)(?:\s*(?:号码|號碼|编号|編號).*)?$/u,
    ],
  },
  {
    id: 'receipt',
    name: 'Receipt',
    patterns: [
      /^(?:(?:payment|sales|cash)\s+)?receipt(?:\s*(?:no\.?|number|#).*)?$/iu,
      /^(?:收据|收據)$/u,
    ],
  },
  {
    id: 'purchase-order',
    name: 'Purchase order',
    patterns: [/^(?:purchase order|采购订单|採購訂單)(?:\s*(?:no\.?|number|#|编号|編號).*)?$/iu],
  },
  {
    id: 'bank-statement',
    name: 'Bank statement',
    patterns: [/^(?:bank statement|银行对账单|銀行對賬單|银行流水|銀行流水)$/iu],
  },
  {
    id: 'financial-statement',
    name: 'Financial statement',
    patterns: [/^(?:financial statements?|财务报表|財務報表)$/iu],
  },
  {
    id: 'nda',
    name: 'NDA',
    patterns: [/^(?:mutual\s+)?non[- ]disclosure agreement$/iu, /^(?:保密协议|保密協議)$/u],
    parents: [{ id: 'contract', name: 'Contract' }],
  },
  {
    id: 'service-agreement',
    name: 'Service agreement',
    patterns: [/^(?:master\s+)?service(?:s)? agreement$/iu, /^(?:服务协议|服務協議)$/u],
    parents: [{ id: 'contract', name: 'Contract' }],
  },
  {
    id: 'employment-contract',
    name: 'Employment contract',
    patterns: [/^(?:employment contract|employment agreement|劳动合同|勞動合同)$/iu],
    parents: [{ id: 'contract', name: 'Contract' }],
  },
  {
    id: 'lease-agreement',
    name: 'Lease agreement',
    patterns: [/^(?:lease agreement|rental agreement|租赁合同|租賃合同)$/iu],
    parents: [{ id: 'contract', name: 'Contract' }],
  },
  {
    id: 'contract',
    name: 'Contract',
    patterns: [/^(?:contract|agreement|合同|合约|合約|协议|協議)$/iu],
  },
  {
    id: 'resume',
    name: 'Resume',
    patterns: [/^(?:resume|résumé|curriculum vitae|个人简历|個人簡歷)$/iu],
  },
  {
    id: 'meeting-minutes',
    name: 'Meeting minutes',
    patterns: [/^(?:minutes of (?:the )?.+ meeting|meeting minutes|会议纪要|會議紀要)$/iu],
  },
  {
    id: 'meeting-agenda',
    name: 'Meeting agenda',
    patterns: [/^(?:meeting agenda|会议议程|會議議程)$/iu],
  },
  {
    id: 'press-release',
    name: 'Press release',
    patterns: [/^(?:press release|news release|新闻稿|新聞稿)$/iu],
  },
  {
    id: 'annual-report',
    name: 'Annual report',
    patterns: [/^(?:\d{4}\s*)?(?:annual report|年度报告|年度報告)$/iu],
  },
  {
    id: 'audit-report',
    name: 'Audit report',
    patterns: [/^(?:independent auditor'?s report|audit report|审计报告|審計報告)$/iu],
  },
  {
    id: 'standard-operating-procedure',
    name: 'Standard operating procedure',
    patterns: [/^(?:standard operating procedure|sop|标准作业程序|標準作業程序)$/iu],
  },
  {
    id: 'safety-data-sheet',
    name: 'Safety data sheet',
    patterns: [/^(?:safety data sheet|material safety data sheet|msds|安全数据表|安全資料表)$/iu],
  },
  {
    id: 'statement-of-work',
    name: 'Statement of work',
    patterns: [/^(?:statement of work|sow|工作说明书|工作說明書)$/iu],
  },
  {
    id: 'business-proposal',
    name: 'Business proposal',
    patterns: [/^(?:business proposal|商业提案|商業提案)$/iu],
    parents: [{ id: 'proposal', name: 'Proposal' }],
  },
  {
    id: 'proposal',
    name: 'Proposal',
    patterns: [/^(?:proposal|提案书|提案書)$/iu],
  },
  {
    id: 'payslip',
    name: 'Payslip',
    patterns: [/^(?:payslip|pay slip|salary slip|工资单|工資單)$/iu],
  },
  {
    id: 'tax-form',
    name: 'Tax form',
    patterns: [/^(?:tax form|form\s+(?:w-?2|w-?9|1040)|纳税申报表|納稅申報表)$/iu],
  },
  {
    id: 'user-guide',
    name: 'User guide',
    patterns: [/^(?:user guide|user manual|用户指南|使用手册|使用手冊)$/iu],
  },
  {
    id: 'release-notes',
    name: 'Release notes',
    patterns: [/^(?:release notes|发行说明|發行說明|发布说明|發佈說明)$/iu],
  },
]

const LEVEL_RANK: Record<DocumentSensitivityLevel, number> = {
  standard: 0,
  internal: 1,
  confidential: 2,
  restricted: 3,
}

const SENSITIVITY_MARKERS: Array<{ level: DocumentSensitivityLevel; pattern: RegExp }> = [
  { level: 'restricted', pattern: /^(?:strictly confidential|高度机密|高度機密|绝密|絕密)$/iu },
  { level: 'confidential', pattern: /^(?:confidential|机密|機密|秘密)$/iu },
  { level: 'internal', pattern: /^(?:internal use only|internal|内部使用|內部使用)$/iu },
]

function pageWindowIndexes(pageCount: number): number[] {
  if (pageCount <= 4) return Array.from({ length: pageCount }, (_, index) => index)
  return [0, 1, pageCount - 2, pageCount - 1]
}

function pageLines(page: SearchIndex[number]): string[] {
  const lines = new Set<string>()
  for (const item of page.items) {
    const text = page.text.slice(item.start, item.end).replace(/\s+/gu, ' ').trim()
    if (text && text.length <= 240) lines.add(text)
  }
  for (const match of page.text.matchAll(/[^\n]+/gu)) {
    const text = match[0].replace(/\s+/gu, ' ').trim()
    if (text && text.length <= 240) lines.add(text)
  }
  return [...lines]
}

function classify(index: SearchIndex): {
  pageIndexes: number[]
  classifications: DocumentClassification[]
} {
  const pageIndexes = pageWindowIndexes(index.length)
  const lines = pageIndexes.flatMap((pageIndex) => pageLines(index[pageIndex]!))
  const results = new Map<string, DocumentClassification>()
  const add = (value: Pick<DocumentClassification, 'id' | 'name'>, evidenceCount: number): void => {
    const current = results.get(value.id)
    if (current) current.evidenceCount += evidenceCount
    else results.set(value.id, { ...value, evidenceCount })
  }

  for (const rule of CLASSIFICATION_RULES) {
    const evidenceCount = lines.filter((line) =>
      rule.patterns.some((pattern) => {
        pattern.lastIndex = 0
        return pattern.test(line)
      }),
    ).length
    if (evidenceCount === 0) continue
    add(rule, evidenceCount)
    for (const parent of rule.parents ?? []) add(parent, evidenceCount)
  }

  return {
    pageIndexes,
    classifications: [...results.values()]
      .sort((left, right) => right.evidenceCount - left.evidenceCount)
      .slice(0, 5),
  }
}

function digits(value: string): string {
  return value.replace(/\D/gu, '')
}

function validLuhn(value: string): boolean {
  const normalized = digits(value)
  if (normalized.length < 13 || normalized.length > 19 || /^(\d)\1+$/u.test(normalized))
    return false
  let sum = 0
  let double = false
  for (let index = normalized.length - 1; index >= 0; index -= 1) {
    let digit = Number(normalized[index])
    if (double) {
      digit *= 2
      if (digit > 9) digit -= 9
    }
    sum += digit
    double = !double
  }
  return sum % 10 === 0
}

function validChineseIdentity(value: string): boolean {
  const normalized = value.toLocaleUpperCase()
  if (!/^\d{17}[\dX]$/u.test(normalized)) return false
  const year = Number(normalized.slice(6, 10))
  const month = Number(normalized.slice(10, 12))
  const day = Number(normalized.slice(12, 14))
  const date = new Date(Date.UTC(year, month - 1, day))
  if (
    year < 1900 ||
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return false
  }
  const weights = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
  const checks = ['1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2']
  const sum = weights.reduce(
    (total, weight, index) => total + Number(normalized[index]) * weight,
    0,
  )
  return checks[sum % 11] === normalized[17]
}

function validSsn(value: string): boolean {
  const match = /^(\d{3})-(\d{2})-(\d{4})$/u.exec(value)
  if (!match) return false
  return (
    match[1] !== '000' &&
    match[1] !== '666' &&
    !match[1]!.startsWith('9') &&
    match[2] !== '00' &&
    match[3] !== '0000'
  )
}

function validPassport(value: string): boolean {
  const normalized = value.toLocaleUpperCase()
  return /^[A-Z0-9]{5,12}$/u.test(normalized) && !/^(.)\1+$/u.test(normalized)
}

function validIban(value: string): boolean {
  const normalized = value.replace(/\s+/gu, '').toLocaleUpperCase()
  if (!/^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$/u.test(normalized)) return false
  const rearranged = normalized.slice(4) + normalized.slice(0, 4)
  let remainder = 0
  for (const character of rearranged) {
    const part = /\d/u.test(character) ? character : String(character.charCodeAt(0) - 55)
    for (const digit of part) remainder = (remainder * 10 + Number(digit)) % 97
  }
  return remainder === 1
}

const SENSITIVE_DETECTORS: SensitiveDetector[] = [
  {
    kind: 'email',
    expression: /\b[A-Z0-9._%+-]{1,64}@[A-Z0-9.-]+\.[A-Z]{2,24}\b/giu,
    value: (match) => match[0],
    valid: () => true,
  },
  {
    kind: 'phone',
    expression:
      /(?:\b(?:phone|telephone|mobile|tel\.?|contact number)|(?:电话|電話|手机|手機|联系电话|聯繫電話))\s*[:：]?\s*(\+?\d[\d\s().-]{6,20}\d)/giu,
    value: (match) => match[1] ?? null,
    valid: (value) => {
      const count = digits(value).length
      return count >= 7 && count <= 15
    },
  },
  {
    kind: 'cnIdentity',
    expression: /(?<!\d)\d{17}[\dXx](?![\dXx])/gu,
    value: (match) => match[0],
    valid: validChineseIdentity,
  },
  {
    kind: 'paymentCard',
    expression:
      /(?:\b(?:credit card|debit card|bank card|card number|card no\.?)|(?:信用卡|借记卡|借記卡|银行卡|銀行卡|卡号|卡號))\s*[:：]?\s*((?:\d[ -]?){12,18}\d)/giu,
    value: (match) => match[1] ?? null,
    valid: validLuhn,
  },
  {
    kind: 'ssn',
    expression: /(?:\b(?:SSN|social security(?: number)?)\b)\s*[:：]?\s*(\d{3}-\d{2}-\d{4})/giu,
    value: (match) => match[1] ?? null,
    valid: validSsn,
  },
  {
    kind: 'passport',
    expression:
      /(?:\bpassport(?:\s*(?:number|no\.?))?|(?:护照号码|護照號碼|护照号|護照號))\s*[:：]?\s*([A-Z0-9]{5,12})/giu,
    value: (match) => match[1] ?? null,
    valid: validPassport,
  },
  {
    kind: 'iban',
    expression: /\b[A-Z]{2}\d{2}(?:\s?[A-Z0-9]){11,30}\b/giu,
    value: (match) => match[0],
    valid: validIban,
  },
]

function occurrenceAt(text: string, anchor: string, targetIndex: number): number {
  const lower = text.toLocaleLowerCase()
  const query = anchor.toLocaleLowerCase()
  let occurrence = 0
  let from = 0
  while (from <= targetIndex) {
    const index = lower.indexOf(query, from)
    if (index < 0 || index > targetIndex) break
    occurrence += 1
    from = index + query.length
  }
  return Math.max(1, occurrence)
}

function maskValue(kind: SensitiveFindingKind, value: string): string {
  if (kind === 'email') {
    const [local = '', domain = ''] = value.split('@')
    return `${local.slice(0, 1)}***@${domain}`
  }
  const normalized = kind === 'passport' ? value.toLocaleUpperCase() : digits(value)
  if (kind === 'iban') {
    const compact = value.replace(/\s+/gu, '').toLocaleUpperCase()
    return `${compact.slice(0, 4)} **** **** ${compact.slice(-4)}`
  }
  if (kind === 'ssn') return `***-**-${normalized.slice(-4)}`
  if (kind === 'passport') return `${normalized.slice(0, 1)}***${normalized.slice(-3)}`
  return `${'*'.repeat(Math.min(12, Math.max(4, normalized.length - 4)))}${normalized.slice(-4)}`
}

function sensitiveCandidates(pageText: string): SensitiveCandidate[] {
  const candidates = new Map<string, SensitiveCandidate>()
  for (const detector of SENSITIVE_DETECTORS) {
    detector.expression.lastIndex = 0
    let match: RegExpExecArray | null
    while ((match = detector.expression.exec(pageText))) {
      const raw = detector.value(match)?.trim()
      if (!raw || !detector.valid(raw)) continue
      const offset = match[0].indexOf(raw)
      if (offset < 0) continue
      const start = match.index + offset
      candidates.set(`${detector.kind}:${start}:${raw.length}`, { kind: detector.kind, raw, start })
    }
  }
  return [...candidates.values()].sort((left, right) => left.start - right.start)
}

function strongerLevel(
  current: DocumentSensitivityLevel,
  candidate: DocumentSensitivityLevel,
): DocumentSensitivityLevel {
  return LEVEL_RANK[candidate] > LEVEL_RANK[current] ? candidate : current
}

export function auditDocumentInSearchIndex(
  index: SearchIndex,
  options: DocumentAuditOptions = {},
): DocumentAuditReport {
  const maxFindings = options.maxFindings ?? 30
  if (!Number.isInteger(maxFindings) || maxFindings < 1) {
    throw new Error('maxFindings must be a positive integer')
  }

  const { pageIndexes, classifications } = classify(index)
  const sensitiveFindings: SensitiveFinding[] = []
  let truncated = false
  let sensitivity: DocumentSensitivityLevel = classifications.some(({ id }) => id === 'nda')
    ? 'confidential'
    : 'standard'
  let explicitSensitivityMarkers = 0

  for (let pageIndex = 0; pageIndex < index.length; pageIndex += 1) {
    const page = index[pageIndex]!
    for (const line of pageLines(page)) {
      for (const marker of SENSITIVITY_MARKERS) {
        marker.pattern.lastIndex = 0
        if (!marker.pattern.test(line)) continue
        explicitSensitivityMarkers += 1
        sensitivity = strongerLevel(sensitivity, marker.level)
      }
    }
    for (const candidate of sensitiveCandidates(page.text)) {
      sensitivity = strongerLevel(
        sensitivity,
        candidate.kind === 'email' || candidate.kind === 'phone' ? 'confidential' : 'restricted',
      )
      if (sensitiveFindings.length >= maxFindings) {
        truncated = true
        continue
      }
      sensitiveFindings.push({
        kind: candidate.kind,
        pageIndex,
        anchorText: candidate.raw,
        occurrence: occurrenceAt(page.text, candidate.raw, candidate.start),
        maskedValue: maskValue(candidate.kind, candidate.raw),
      })
    }
  }

  return {
    pagesExamined: index.length,
    classificationPageIndexes: pageIndexes,
    classifications,
    sensitivity,
    explicitSensitivityMarkers,
    sensitiveFindings,
    truncated,
  }
}
