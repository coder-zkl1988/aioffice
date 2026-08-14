export const PDF_MOBILE_SCANNER_CHANNELS = {
  createSession: 'pdf:mobile-scanner:create-session',
  pollSession: 'pdf:mobile-scanner:poll-session',
  closeSession: 'pdf:mobile-scanner:close-session',
} as const

export interface PdfMobileScannerSession {
  sessionId: string
  uploadUrl: string
  expiresAt: number
}

export interface PdfMobileScannerFile {
  id: string
  name: string
  type: string
  size: number
  bytes: Uint8Array
}

export interface PdfMobileScannerPollResult {
  files: PdfMobileScannerFile[]
  expiresAt: number
}

export interface PdfMobileScannerApi {
  createSession(): Promise<PdfMobileScannerSession>
  pollSession(sessionId: string): Promise<PdfMobileScannerPollResult>
  closeSession(sessionId: string): Promise<void>
}

export function isPdfMobileScannerSessionId(value: unknown): value is string {
  return typeof value === 'string' && /^[A-Za-z0-9_-]{24,64}$/u.test(value)
}

declare global {
  interface Window {
    pdfMobileScanner?: PdfMobileScannerApi
  }
}
