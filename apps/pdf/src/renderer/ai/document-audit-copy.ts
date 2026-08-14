import type {
  DocumentAuditReport,
  DocumentSensitivityLevel,
  SensitiveFinding,
  SensitiveFindingKind,
} from '../document-audit'

export interface DocumentAuditCopy {
  action: string
  scanSummary(report: DocumentAuditReport): string
  classificationSummary(report: DocumentAuditReport): string
  commentSummary(count: number): string
  result(report: DocumentAuditReport, commentsAdded: number, readOnly: boolean): string
  subject(finding: SensitiveFinding): string
  comment(finding: SensitiveFinding): string
}

function sensitivityName(level: DocumentSensitivityLevel, lang: string): string {
  if (lang === 'zh') {
    return { standard: '普通', internal: '内部', confidential: '机密', restricted: '严格受限' }[
      level
    ]
  }
  if (lang === 'zh-TW') {
    return { standard: '一般', internal: '內部', confidential: '機密', restricted: '嚴格受限' }[
      level
    ]
  }
  return {
    standard: 'Standard',
    internal: 'Internal',
    confidential: 'Confidential',
    restricted: 'Restricted',
  }[level]
}

function findingName(kind: SensitiveFindingKind, lang: string): string {
  if (lang === 'zh') {
    return {
      email: '电子邮箱',
      phone: '电话号码',
      cnIdentity: '中国居民身份证号',
      paymentCard: '支付卡号',
      ssn: '社会安全号码',
      passport: '护照号码',
      iban: '国际银行账号',
    }[kind]
  }
  if (lang === 'zh-TW') {
    return {
      email: '電子郵件',
      phone: '電話號碼',
      cnIdentity: '中國居民身分證號',
      paymentCard: '支付卡號',
      ssn: '社會安全號碼',
      passport: '護照號碼',
      iban: '國際銀行帳號',
    }[kind]
  }
  return {
    email: 'email address',
    phone: 'phone number',
    cnIdentity: 'Chinese resident identity number',
    paymentCard: 'payment card number',
    ssn: 'Social Security number',
    passport: 'passport number',
    iban: 'IBAN',
  }[kind]
}

function labels(report: DocumentAuditReport): string {
  return report.classifications.length > 0
    ? report.classifications.map((classification) => classification.name).join('、')
    : '—'
}

function findingLines(report: DocumentAuditReport, lang: string): string[] {
  return report.sensitiveFindings.slice(0, 20).map((finding) => {
    if (lang === 'zh') {
      return `- 第 ${finding.pageIndex + 1} 页：${findingName(finding.kind, lang)} \`${finding.maskedValue}\``
    }
    if (lang === 'zh-TW') {
      return `- 第 ${finding.pageIndex + 1} 頁：${findingName(finding.kind, lang)} \`${finding.maskedValue}\``
    }
    return `- Page ${finding.pageIndex + 1}: ${findingName(finding.kind, lang)} \`${finding.maskedValue}\``
  })
}

export function documentAuditCopy(lang: string): DocumentAuditCopy {
  if (lang === 'zh') {
    return {
      action: '文档审查',
      scanSummary: (report) =>
        `本地审查 ${report.pagesExamined} 页、${report.sensitiveFindings.length} 处敏感信息`,
      classificationSummary: (report) =>
        `保存文档分类：${labels(report)}（${sensitivityName(report.sensitivity, lang)}）`,
      commentSummary: (count) => `添加 ${count} 条敏感信息批注`,
      result: (report, commentsAdded, readOnly) => {
        const header = `### 文档审查完成\n\n- 文档类型：${labels(report)}\n- 敏感等级：${sensitivityName(report.sensitivity, lang)}\n- 明确密级标记：${report.explicitSensitivityMarkers} 处`
        if (report.sensitiveFindings.length === 0) {
          return `${header}\n\n未发现可确定识别的敏感字段。类型判断仅采用首尾页明确标题。`
        }
        const mutation = readOnly
          ? '当前文档为只读，未写入批注。'
          : `已添加 ${commentsAdded} 条便签批注；未自动删除或涂黑原文。`
        const lines = findingLines(report, lang)
        if (report.truncated || report.sensitiveFindings.length > 20)
          lines.push('- 其余结果已省略。')
        return `${header}\n\n发现 ${report.sensitiveFindings.length} 处敏感信息，${mutation}\n\n${lines.join('\n')}`
      },
      subject: (finding) => `敏感信息：${findingName(finding.kind, lang)}`,
      comment: (finding) =>
        `检测到${findingName(finding.kind, lang)} ${finding.maskedValue}。请确认共享范围；如需移除，请使用“安全涂黑”工具。`,
    }
  }
  if (lang === 'zh-TW') {
    return {
      action: '文件審查',
      scanSummary: (report) =>
        `本機審查 ${report.pagesExamined} 頁、${report.sensitiveFindings.length} 處敏感資訊`,
      classificationSummary: (report) =>
        `儲存文件分類：${labels(report)}（${sensitivityName(report.sensitivity, lang)}）`,
      commentSummary: (count) => `新增 ${count} 則敏感資訊註解`,
      result: (report, commentsAdded, readOnly) => {
        const header = `### 文件審查完成\n\n- 文件類型：${labels(report)}\n- 敏感等級：${sensitivityName(report.sensitivity, lang)}\n- 明確密級標記：${report.explicitSensitivityMarkers} 處`
        if (report.sensitiveFindings.length === 0) {
          return `${header}\n\n未發現可確定識別的敏感欄位。類型判斷僅採用首尾頁明確標題。`
        }
        const mutation = readOnly
          ? '目前文件為唯讀，未寫入註解。'
          : `已新增 ${commentsAdded} 則便利貼註解；未自動刪除或塗黑原文。`
        const lines = findingLines(report, lang)
        if (report.truncated || report.sensitiveFindings.length > 20)
          lines.push('- 其餘結果已省略。')
        return `${header}\n\n發現 ${report.sensitiveFindings.length} 處敏感資訊，${mutation}\n\n${lines.join('\n')}`
      },
      subject: (finding) => `敏感資訊：${findingName(finding.kind, lang)}`,
      comment: (finding) =>
        `偵測到${findingName(finding.kind, lang)} ${finding.maskedValue}。請確認分享範圍；如需移除，請使用「安全塗黑」工具。`,
    }
  }
  return {
    action: 'Audit document',
    scanSummary: (report) =>
      `Audited ${report.pagesExamined} pages and found ${report.sensitiveFindings.length} sensitive fields locally`,
    classificationSummary: (report) =>
      `Saved document classification: ${labels(report)} (${sensitivityName(report.sensitivity, lang)})`,
    commentSummary: (count) => `Added ${count} sensitive-data comments`,
    result: (report, commentsAdded, readOnly) => {
      const classification =
        report.classifications.length > 0
          ? report.classifications.map(({ name }) => name).join(', ')
          : 'Unclassified'
      const header = `### Document audit complete\n\n- Document type: ${classification}\n- Sensitivity: ${sensitivityName(report.sensitivity, lang)}\n- Explicit sensitivity markings: ${report.explicitSensitivityMarkers}`
      if (report.sensitiveFindings.length === 0) {
        return `${header}\n\nNo deterministically identifiable sensitive fields were found. Classification uses explicit titles on the first and last pages only.`
      }
      const mutation = readOnly
        ? 'The document is read-only, so no comments were added.'
        : `Added ${commentsAdded} sticky-note comments. No content was removed or redacted automatically.`
      const lines = findingLines(report, lang)
      if (report.truncated || report.sensitiveFindings.length > 20)
        lines.push('- Additional results omitted.')
      return `${header}\n\nFound ${report.sensitiveFindings.length} sensitive fields. ${mutation}\n\n${lines.join('\n')}`
    },
    subject: (finding) => `Sensitive data: ${findingName(finding.kind, lang)}`,
    comment: (finding) =>
      `Detected ${findingName(finding.kind, lang)} ${finding.maskedValue}. Confirm the sharing scope; use Secure redaction if it must be removed.`,
  }
}
