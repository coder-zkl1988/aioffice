import type {
  ConsistencyAuditReport,
  ConsistencyClaim,
  ConsistencyClaimKind,
  ConsistencyFinding,
} from '../consistency-audit'

export interface ConsistencyAuditCopy {
  action: string
  scanSummary(report: ConsistencyAuditReport): string
  commentSummary(count: number): string
  result(report: ConsistencyAuditReport, commentsAdded: number, readOnly: boolean): string
  subject(finding: ConsistencyFinding): string
  comment(finding: ConsistencyFinding, claim: ConsistencyClaim): string
}

function kindName(kind: ConsistencyClaimKind, lang: string): string {
  if (lang === 'zh') {
    return {
      deadline: '截止日期',
      effectiveDate: '生效日期',
      expiryDate: '到期日期',
      status: '状态',
      version: '版本',
    }[kind]
  }
  if (lang === 'zh-TW') {
    return {
      deadline: '截止日期',
      effectiveDate: '生效日期',
      expiryDate: '到期日期',
      status: '狀態',
      version: '版本',
    }[kind]
  }
  return {
    deadline: 'deadline',
    effectiveDate: 'effective date',
    expiryDate: 'expiry date',
    status: 'status',
    version: 'version',
  }[kind]
}

function findingLine(finding: ConsistencyFinding, lang: string): string {
  const firstPage = finding.first.pageIndex + 1
  const secondPage = finding.second.pageIndex + 1
  const name = kindName(finding.kind, lang)
  const subject = finding.subject === 'document' ? '' : `“${finding.subject}”`
  if (lang === 'zh') {
    return `- ${subject ? `${subject}的` : ''}${name}：第 ${firstPage} 页为 ${finding.first.displayValue}，第 ${secondPage} 页为 ${finding.second.displayValue}。`
  }
  if (lang === 'zh-TW') {
    return `- ${subject ? `${subject}的` : ''}${name}：第 ${firstPage} 頁為 ${finding.first.displayValue}，第 ${secondPage} 頁為 ${finding.second.displayValue}。`
  }
  const prefix = subject ? `${subject} ` : ''
  return `- ${prefix}${name}: ${finding.first.displayValue} on page ${firstPage}, but ${finding.second.displayValue} on page ${secondPage}.`
}

export function consistencyAuditCopy(lang: string): ConsistencyAuditCopy {
  if (lang === 'zh') {
    return {
      action: '一致性检查',
      scanSummary: (report) =>
        `本地检查 ${report.pagesExamined} 页、${report.claimsExamined} 项明确声明`,
      commentSummary: (count) => `添加 ${count} 条一致性批注`,
      result: (report, commentsAdded, readOnly) => {
        const header = `### 一致性检查完成\n\n已检查 ${report.pagesExamined} 页中的 ${report.claimsExamined} 项明确日期、状态和版本声明。`
        if (report.findings.length === 0) return `${header}\n\n未发现可确定验证的矛盾。`
        const mutation = readOnly
          ? '当前文档为只读，未写入批注。'
          : `已在冲突原文处添加 ${commentsAdded} 条便签批注。`
        const lines = report.findings.slice(0, 20).map((finding) => findingLine(finding, lang))
        if (report.truncated || report.findings.length > 20) lines.push('- 其余冲突已省略。')
        return `${header}\n\n发现 ${report.findings.length} 组冲突，${mutation}\n\n${lines.join('\n')}`
      },
      subject: (finding) => `${kindName(finding.kind, lang)}冲突`,
      comment: (finding, claim) => {
        const counterpart = claim === finding.first ? finding.second : finding.first
        return `此处${kindName(finding.kind, lang)}为 ${claim.displayValue}，但第 ${counterpart.pageIndex + 1} 页的对应声明为 ${counterpart.displayValue}。请确认正确版本并统一。`
      },
    }
  }
  if (lang === 'zh-TW') {
    return {
      action: '一致性檢查',
      scanSummary: (report) =>
        `本機檢查 ${report.pagesExamined} 頁、${report.claimsExamined} 項明確聲明`,
      commentSummary: (count) => `新增 ${count} 則一致性註解`,
      result: (report, commentsAdded, readOnly) => {
        const header = `### 一致性檢查完成\n\n已檢查 ${report.pagesExamined} 頁中的 ${report.claimsExamined} 項明確日期、狀態和版本聲明。`
        if (report.findings.length === 0) return `${header}\n\n未發現可確定驗證的矛盾。`
        const mutation = readOnly
          ? '目前文件為唯讀，未寫入註解。'
          : `已在衝突原文處新增 ${commentsAdded} 則便利貼註解。`
        const lines = report.findings.slice(0, 20).map((finding) => findingLine(finding, lang))
        if (report.truncated || report.findings.length > 20) lines.push('- 其餘衝突已省略。')
        return `${header}\n\n發現 ${report.findings.length} 組衝突，${mutation}\n\n${lines.join('\n')}`
      },
      subject: (finding) => `${kindName(finding.kind, lang)}衝突`,
      comment: (finding, claim) => {
        const counterpart = claim === finding.first ? finding.second : finding.first
        return `此處${kindName(finding.kind, lang)}為 ${claim.displayValue}，但第 ${counterpart.pageIndex + 1} 頁的對應聲明為 ${counterpart.displayValue}。請確認正確版本並統一。`
      },
    }
  }
  return {
    action: 'Check consistency',
    scanSummary: (report) =>
      `Checked ${report.pagesExamined} pages and ${report.claimsExamined} explicit claims locally`,
    commentSummary: (count) => `Added ${count} consistency comments`,
    result: (report, commentsAdded, readOnly) => {
      const header = `### Consistency check complete\n\nChecked ${report.claimsExamined} explicit date, status, and version claims across ${report.pagesExamined} pages.`
      if (report.findings.length === 0)
        return `${header}\n\nNo deterministically verifiable contradictions were found.`
      const mutation = readOnly
        ? 'The document is read-only, so no comments were added.'
        : `Added ${commentsAdded} sticky-note comments beside the conflicting passages.`
      const lines = report.findings.slice(0, 20).map((finding) => findingLine(finding, lang))
      if (report.truncated || report.findings.length > 20)
        lines.push('- Additional contradictions omitted.')
      return `${header}\n\nFound ${report.findings.length} conflicts. ${mutation}\n\n${lines.join('\n')}`
    },
    subject: (finding) => `${kindName(finding.kind, lang)} conflict`,
    comment: (finding, claim) => {
      const counterpart = claim === finding.first ? finding.second : finding.first
      return `This ${kindName(finding.kind, lang)} is ${claim.displayValue}, but the corresponding claim on page ${counterpart.pageIndex + 1} is ${counterpart.displayValue}. Confirm the correct version and make the claims consistent.`
    },
  }
}
