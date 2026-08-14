import { describe, expect, it } from 'vitest'
import { generatedBinaryFileKind, WEB_BINARY_FILE_MIMES } from './files'

describe('generated browser files', () => {
  it('recognizes editable and binary outputs by extension and MIME', () => {
    expect(generatedBinaryFileKind('.DOCX', WEB_BINARY_FILE_MIMES.docx)).toBe('docx')
    expect(generatedBinaryFileKind('.md', 'text/markdown;charset=utf-8')).toBe('markdown')
    expect(generatedBinaryFileKind('.MARKDOWN', WEB_BINARY_FILE_MIMES.markdown)).toBe('markdown')
    expect(generatedBinaryFileKind('.pdf', WEB_BINARY_FILE_MIMES.pdf)).toBe('pdf')
    expect(generatedBinaryFileKind('.XLSX', WEB_BINARY_FILE_MIMES.xlsx)).toBe('xlsx')
    expect(generatedBinaryFileKind('.PPTX', WEB_BINARY_FILE_MIMES.pptx)).toBe('pptx')
  })

  it('does not persist ambiguous or mismatched generated output types', () => {
    expect(generatedBinaryFileKind('.pptx', 'application/octet-stream')).toBeNull()
    expect(generatedBinaryFileKind('.zip', WEB_BINARY_FILE_MIMES.pptx)).toBeNull()
  })
})
