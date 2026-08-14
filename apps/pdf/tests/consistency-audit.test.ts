import { describe, expect, it } from 'vitest'
import { auditConsistencyInSearchIndex } from '../src/renderer/consistency-audit'
import type { SearchIndex } from '../src/renderer/search'

function indexFor(...pages: string[]): SearchIndex {
  return pages.map((text) => {
    const items: SearchIndex[number]['items'] = []
    let from = 0
    for (const line of text.split('\n')) {
      if (line) {
        items.push({
          start: from,
          end: from + line.length,
          x: 20,
          y: 700 - items.length * 20,
          w: line.length * 6,
          h: 12,
        })
      }
      from += line.length + 1
    }
    return { text, lower: text.toLocaleLowerCase(), items }
  })
}

describe('auditConsistencyInSearchIndex', () => {
  it('finds cross-page deadline conflicts while accepting equivalent date formats', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        'Project Deadline: March 5, 2026',
        'Project Deadline: 2026-03-05',
        'Project Deadline: April 10, 2026',
      ),
    )
    expect(report).toMatchObject({ pagesExamined: 3, claimsExamined: 3, truncated: false })
    expect(report.findings).toEqual([
      expect.objectContaining({
        kind: 'deadline',
        subject: 'project',
        first: expect.objectContaining({ pageIndex: 0, normalizedValue: '2026-03-05' }),
        second: expect.objectContaining({ pageIndex: 2, normalizedValue: '2026-04-10' }),
      }),
    ])
  })

  it('compares month-day dates only with claims of the same precision', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        'The project deadline is March 5.',
        'Project Deadline: March 5, 2026',
        'The project deadline will be April 10.',
      ),
    )
    expect(report.claimsExamined).toBe(3)
    expect(report.findings).toEqual([
      expect.objectContaining({
        first: expect.objectContaining({ displayValue: 'March 5' }),
        second: expect.objectContaining({ displayValue: 'April 10' }),
      }),
    ])
  })

  it('keeps effective and expiry dates in separate ledgers', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        'Contract Effective Date: 2026/03/01\nContract Expiry Date: 2027/03/01',
        'Contract Effective Date: 2026年3月2日\nContract Expiry Date: 2027年3月1日',
      ),
    )
    expect(report.claimsExamined).toBe(4)
    expect(report.findings).toEqual([expect.objectContaining({ kind: 'effectiveDate' })])
  })

  it('finds explicit status and version conflicts', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        'Document Status: Draft\nDocument Version: v1.2',
        'Document Status: Approved\nDocument Version: 1.3',
      ),
    )
    expect(report.findings.map((finding) => finding.kind)).toEqual(['status', 'version'])
  })

  it('supports explicit Chinese labels and declarative dates', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        '项目截止日期为2026年3月5日\n文档版本：2.0',
        '项目截止日期：2026年4月10日\n文档版本：2.1',
      ),
    )
    expect(report.claimsExamined).toBe(4)
    expect(report.findings.map((finding) => finding.kind)).toEqual(['deadline', 'version'])
  })

  it('ignores unlabeled prose and generic record statuses', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        'We discussed March 5, 2026 and may revisit April 10, 2026.',
        'Status: Active\nAlice status: pending\nVersion control is important.',
      ),
    )
    expect(report.claimsExamined).toBe(0)
    expect(report.findings).toEqual([])
  })

  it('reports same-page conflicts and exact repeated occurrences', () => {
    const statement = 'Release Deadline: March 5, 2026'
    const report = auditConsistencyInSearchIndex(
      indexFor(`${statement}\n${statement}\nRelease Deadline: April 10, 2026`),
    )
    expect(report.claimsExamined).toBe(3)
    expect(report.findings).toHaveLength(1)
    expect(report.findings[0]!.first.occurrence).toBe(1)
    expect(report.findings[0]!.second.occurrence).toBe(1)
  })

  it('caps findings and validates options', () => {
    const report = auditConsistencyInSearchIndex(
      indexFor(
        'Project Deadline: March 5, 2026',
        'Project Deadline: April 10, 2026',
        'Project Deadline: May 11, 2026',
      ),
      { maxFindings: 1 },
    )
    expect(report.findings).toHaveLength(1)
    expect(report.truncated).toBe(true)
    expect(() => auditConsistencyInSearchIndex([], { maxFindings: 0 })).toThrow('positive integer')
  })
})
