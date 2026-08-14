import type { SearchIndex } from './search'

export type ConsistencyClaimKind =
  'deadline' | 'effectiveDate' | 'expiryDate' | 'status' | 'version'

export interface ConsistencyClaim {
  kind: ConsistencyClaimKind
  subject: string
  label: string
  pageIndex: number
  anchorText: string
  occurrence: number
  displayValue: string
  normalizedValue: string
  precision: 'year' | 'monthDay' | 'exact'
}

export interface ConsistencyFinding {
  kind: ConsistencyClaimKind
  subject: string
  first: ConsistencyClaim
  second: ConsistencyClaim
}

export interface ConsistencyAuditReport {
  pagesExamined: number
  claimsExamined: number
  findings: ConsistencyFinding[]
  truncated: boolean
}

interface ConsistencyAuditOptions {
  maxFindings?: number
}

interface TextSegment {
  text: string
  start: number
}

interface ParsedClaimValue {
  displayValue: string
  normalizedValue: string
  precision: ConsistencyClaim['precision']
}

interface ClaimPattern {
  kind: ConsistencyClaimKind
  expression: RegExp
  parse(value: string): ParsedClaimValue | null
}

interface LocatedClaim extends ConsistencyClaim {
  sourceStart: number
  sourceEnd: number
  subjectKey: string
}

const MONTHS: Record<string, number> = {
  jan: 1,
  january: 1,
  feb: 2,
  february: 2,
  mar: 3,
  march: 3,
  apr: 4,
  april: 4,
  may: 5,
  jun: 6,
  june: 6,
  jul: 7,
  july: 7,
  aug: 8,
  august: 8,
  sep: 9,
  sept: 9,
  september: 9,
  oct: 10,
  october: 10,
  nov: 11,
  november: 11,
  dec: 12,
  december: 12,
}

const MONTH_SOURCE =
  '(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)'
const DATE_SOURCE = [
  '\\d{4}[ \\t]*[年.\\/-][ \\t]*\\d{1,2}[ \\t]*[月.\\/-][ \\t]*\\d{1,2}[ \\t]*日?',
  `${MONTH_SOURCE}\\.?[ \\t]+\\d{1,2}(?:st|nd|rd|th)?(?:(?:[ \\t]*,[ \\t]*|[ \\t]+)\\d{4})?`,
  `\\d{1,2}(?:st|nd|rd|th)?[ \\t]+${MONTH_SOURCE}\\.?(?:[ \\t]+\\d{4})?`,
  '\\d{1,2}[ \\t]*月[ \\t]*\\d{1,2}[ \\t]*日',
].join('|')

const DATE_LABELS: Record<'deadline' | 'effectiveDate' | 'expiryDate', string> = {
  deadline:
    '(?:[A-Za-z0-9][A-Za-z0-9 &()_\\/-]{0,48}[ \\t]+)?(?:deadline|due[ \\t]+date)|(?:[\\p{Script=Han}A-Za-z0-9]{0,24})(?:截止日期|截止時間|截止时间|交付日期)',
  effectiveDate:
    '(?:[A-Za-z0-9][A-Za-z0-9 &()_\\/-]{0,48}[ \\t]+)?(?:effective[ \\t]+date|commencement[ \\t]+date)|(?:[\\p{Script=Han}A-Za-z0-9]{0,24})(?:生效日期|開始日期|开始日期|施行日期)',
  expiryDate:
    '(?:[A-Za-z0-9][A-Za-z0-9 &()_\\/-]{0,48}[ \\t]+)?(?:expiration[ \\t]+date|expiry[ \\t]+date|expiry)|(?:[\\p{Script=Han}A-Za-z0-9]{0,24})(?:到期日期|失效日期|終止日期|终止日期)',
}

const STATUS_LABEL_SOURCE =
  '(?:(?:project|document|contract|agreement|policy|release|proposal|approval)[ \\t]+status)|(?:(?:项目|項目|文档|文件|文檔|合同|合約|协议|協議|政策|发布|發佈|提案|审批|審批)状态)|(?:(?:项目|項目|文档|文件|文檔|合同|合約|协议|協議|政策|发布|發佈|提案|审批|審批)狀態)'
const STATUS_VALUE_SOURCE =
  '(?:draft|approved|rejected|cancelled|canceled|completed|pending|on[ \\t]+hold|active|inactive|suspended|草稿|已批准|批准|已拒绝|已拒絕|拒绝|拒絕|已取消|取消|已完成|完成|待定|待审批|待審批|进行中|進行中|暂停|暫停|生效|失效)'
const VERSION_LABEL_SOURCE =
  '(?:(?:document|contract|release|proposal|policy)[ \\t]+version|version|revision)|(?:(?:文档|文件|文檔|合同|合約|发布|發佈|提案|政策)?版本|修订版|修訂版)'
const VERSION_VALUE_SOURCE = 'v?\\d+(?:\\.\\d+){0,3}(?:[-_][A-Za-z0-9]+)?'
const LEADING_MARKER_SOURCE = '[ \\t]*(?:[•·▪◦*\\-–—][ \\t]*)?'

function labeledPattern(labelSource: string, valueSource: string): RegExp {
  return new RegExp(
    `^${LEADING_MARKER_SOURCE}(${labelSource})[ \\t]*[:：][ \\t]*(${valueSource})`,
    'iu',
  )
}

function declarativeDatePattern(labelSource: string): RegExp {
  return new RegExp(
    `^${LEADING_MARKER_SOURCE}(?:the[ \\t]+)?(${labelSource})[ \\t]+(?:is|remains|will[ \\t]+be)[ \\t]+(${DATE_SOURCE})|^${LEADING_MARKER_SOURCE}(${labelSource})[ \\t]*(?:为|為|是|定于|定於)[ \\t]*(${DATE_SOURCE})`,
    'iu',
  )
}

function parseDate(value: string): ParsedClaimValue | null {
  const displayValue = value.trim().replace(/[.。]$/u, '')
  let year: number | null = null
  let month: number | null = null
  let day: number | null = null

  let match = /^(\d{4})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})\s*日?$/u.exec(displayValue)
  if (match) {
    year = Number(match[1])
    month = Number(match[2])
    day = Number(match[3])
  } else {
    match = /^(\d{1,2})\s*月\s*(\d{1,2})\s*日$/u.exec(displayValue)
    if (match) {
      month = Number(match[1])
      day = Number(match[2])
    } else {
      match = new RegExp(
        `^(${MONTH_SOURCE})\\.?[ \\t]+(\\d{1,2})(?:st|nd|rd|th)?(?:[ \\t]*,[ \\t]*|[ \\t]+)?(\\d{4})?$`,
        'iu',
      ).exec(displayValue)
      if (match) {
        month = MONTHS[match[1]!.replace(/\./gu, '').toLocaleLowerCase()] ?? null
        day = Number(match[2])
        year = match[3] ? Number(match[3]) : null
      } else {
        match = new RegExp(
          `^(\\d{1,2})(?:st|nd|rd|th)?[ \\t]+(${MONTH_SOURCE})\\.?(?:[ \\t]+(\\d{4}))?$`,
          'iu',
        ).exec(displayValue)
        if (match) {
          day = Number(match[1])
          month = MONTHS[match[2]!.replace(/\./gu, '').toLocaleLowerCase()] ?? null
          year = match[3] ? Number(match[3]) : null
        }
      }
    }
  }

  if (!month || !day || month > 12 || day > 31) return null
  const validationYear = year ?? 2000
  const candidate = new Date(Date.UTC(validationYear, month - 1, day))
  if (
    candidate.getUTCFullYear() !== validationYear ||
    candidate.getUTCMonth() !== month - 1 ||
    candidate.getUTCDate() !== day
  ) {
    return null
  }

  const pad = (part: number): string => String(part).padStart(2, '0')
  return {
    displayValue,
    normalizedValue: `${year === null ? '----' : String(year).padStart(4, '0')}-${pad(month)}-${pad(day)}`,
    precision: year === null ? 'monthDay' : 'year',
  }
}

const STATUS_VALUES = new Map<string, string>([
  ['draft', 'draft'],
  ['草稿', 'draft'],
  ['approved', 'approved'],
  ['已批准', 'approved'],
  ['批准', 'approved'],
  ['rejected', 'rejected'],
  ['已拒绝', 'rejected'],
  ['已拒絕', 'rejected'],
  ['拒绝', 'rejected'],
  ['拒絕', 'rejected'],
  ['cancelled', 'cancelled'],
  ['canceled', 'cancelled'],
  ['已取消', 'cancelled'],
  ['取消', 'cancelled'],
  ['completed', 'completed'],
  ['已完成', 'completed'],
  ['完成', 'completed'],
  ['pending', 'pending'],
  ['待定', 'pending'],
  ['待审批', 'pending'],
  ['待審批', 'pending'],
  ['on hold', 'on-hold'],
  ['暂停', 'on-hold'],
  ['暫停', 'on-hold'],
  ['active', 'active'],
  ['进行中', 'active'],
  ['進行中', 'active'],
  ['生效', 'active'],
  ['inactive', 'inactive'],
  ['suspended', 'inactive'],
  ['失效', 'inactive'],
])

function parseStatus(value: string): ParsedClaimValue | null {
  const displayValue = value.trim()
  const key = displayValue.normalize('NFKC').toLocaleLowerCase().replace(/\s+/gu, ' ')
  const normalizedValue = STATUS_VALUES.get(key)
  return normalizedValue ? { displayValue, normalizedValue, precision: 'exact' } : null
}

function parseVersion(value: string): ParsedClaimValue | null {
  const displayValue = value.trim()
  const normalizedValue = displayValue
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/^v/u, '')
    .replace(/_/gu, '-')
  return normalizedValue ? { displayValue, normalizedValue, precision: 'exact' } : null
}

const CLAIM_PATTERNS: ClaimPattern[] = [
  ...(['deadline', 'effectiveDate', 'expiryDate'] as const).flatMap((kind) => [
    { kind, expression: labeledPattern(DATE_LABELS[kind], DATE_SOURCE), parse: parseDate },
    { kind, expression: declarativeDatePattern(DATE_LABELS[kind]), parse: parseDate },
  ]),
  {
    kind: 'status',
    expression: labeledPattern(STATUS_LABEL_SOURCE, STATUS_VALUE_SOURCE),
    parse: parseStatus,
  },
  {
    kind: 'version',
    expression: labeledPattern(VERSION_LABEL_SOURCE, VERSION_VALUE_SOURCE),
    parse: parseVersion,
  },
]

function segmentsForPage(page: SearchIndex[number]): TextSegment[] {
  const segments = new Map<string, TextSegment>()
  for (const item of page.items) {
    if (item.end <= item.start) continue
    const text = page.text.slice(item.start, item.end)
    if (text.trim()) segments.set(`${item.start}:${item.end}`, { text, start: item.start })
  }
  const lines = /[^\n]+/gu
  let match: RegExpExecArray | null
  while ((match = lines.exec(page.text))) {
    if (match[0].trim() && match[0].length <= 500) {
      segments.set(`${match.index}:${match.index + match[0].length}`, {
        text: match[0],
        start: match.index,
      })
    }
  }
  return [...segments.values()].sort((left, right) => left.start - right.start)
}

function normalizeLabel(label: string): string {
  return label
    .normalize('NFKC')
    .toLocaleLowerCase()
    .replace(/^[\s•·▪◦*\-–—]+/u, '')
    .replace(/[()_\-/]+/gu, ' ')
    .replace(/\s+/gu, ' ')
    .trim()
}

function subjectFor(kind: ConsistencyClaimKind, label: string): string {
  const normalized = normalizeLabel(label).replace(/^the\s+/u, '')
  const aliases: Record<ConsistencyClaimKind, RegExp> = {
    deadline: /(?:deadline|due date|截止日期|截止時間|截止时间|交付日期)$/u,
    effectiveDate: /(?:effective date|commencement date|生效日期|開始日期|开始日期|施行日期)$/u,
    expiryDate: /(?:expiration date|expiry date|expiry|到期日期|失效日期|終止日期|终止日期)$/u,
    status: /(?:status|状态|狀態)$/u,
    version: /(?:version|revision|版本|修订版|修訂版)$/u,
  }
  const prefix = normalized.replace(aliases[kind], '').trim()
  return prefix || 'document'
}

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

function matchParts(match: RegExpExecArray): { label: string; value: string } | null {
  if (match[1] && match[2]) return { label: match[1], value: match[2] }
  if (match[3] && match[4]) return { label: match[3], value: match[4] }
  return null
}

function claimsForPage(page: SearchIndex[number], pageIndex: number): LocatedClaim[] {
  const claims = new Map<string, LocatedClaim>()
  for (const segment of segmentsForPage(page)) {
    for (const pattern of CLAIM_PATTERNS) {
      pattern.expression.lastIndex = 0
      const match = pattern.expression.exec(segment.text)
      const parts = match ? matchParts(match) : null
      if (!match || !parts) continue
      const parsed = pattern.parse(parts.value)
      if (!parsed) continue
      const leading = match[0].length - match[0].trimStart().length
      const anchorText = match[0].trim()
      const sourceStart = segment.start + match.index + leading
      const sourceEnd = sourceStart + anchorText.length
      const subject = subjectFor(pattern.kind, parts.label)
      const key = `${pageIndex}:${sourceStart}:${sourceEnd}:${pattern.kind}`
      claims.set(key, {
        kind: pattern.kind,
        subject,
        label: parts.label.trim(),
        pageIndex,
        anchorText,
        occurrence: occurrenceAt(page.text, anchorText, sourceStart),
        ...parsed,
        sourceStart,
        sourceEnd,
        subjectKey: `${pattern.kind}:${subject}:${parsed.precision}`,
      })
    }
  }
  return [...claims.values()].sort((left, right) => left.sourceStart - right.sourceStart)
}

function publicClaim(claim: LocatedClaim): ConsistencyClaim {
  return {
    kind: claim.kind,
    subject: claim.subject,
    label: claim.label,
    pageIndex: claim.pageIndex,
    anchorText: claim.anchorText,
    occurrence: claim.occurrence,
    displayValue: claim.displayValue,
    normalizedValue: claim.normalizedValue,
    precision: claim.precision,
  }
}

export function auditConsistencyInSearchIndex(
  index: SearchIndex,
  options: ConsistencyAuditOptions = {},
): ConsistencyAuditReport {
  const maxFindings = options.maxFindings ?? 25
  if (!Number.isInteger(maxFindings) || maxFindings < 1) {
    throw new Error('maxFindings must be a positive integer')
  }

  const claims = index.flatMap((page, pageIndex) => claimsForPage(page, pageIndex))
  const groups = new Map<string, LocatedClaim[]>()
  for (const claim of claims) {
    const group = groups.get(claim.subjectKey) ?? []
    group.push(claim)
    groups.set(claim.subjectKey, group)
  }

  const findings: ConsistencyFinding[] = []
  let truncated = false
  for (const group of groups.values()) {
    const first = group[0]
    if (!first) continue
    for (const claim of group.slice(1)) {
      if (claim.normalizedValue === first.normalizedValue) continue
      if (findings.length >= maxFindings) {
        truncated = true
        break
      }
      findings.push({
        kind: first.kind,
        subject: first.subject,
        first: publicClaim(first),
        second: publicClaim(claim),
      })
    }
    if (truncated) break
  }

  return {
    pagesExamined: index.length,
    claimsExamined: claims.length,
    findings,
    truncated,
  }
}
