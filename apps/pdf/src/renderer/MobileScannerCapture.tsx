import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import { QRCodeSVG } from 'qrcode.react'
import type { PdfMobileScannerSession } from '../shared/mobile-scanner'

export interface MobileScannerCaptureLabels {
  start: string
  stop: string
  title: string
  hint: string
  waiting: string
  received: string
  copyLink: string
  copied: string
  expiresIn: string
  expired: string
  unavailable: string
}

function scannerFile(file: { name: string; type: string; bytes: Uint8Array }): File {
  const data = file.bytes.buffer.slice(
    file.bytes.byteOffset,
    file.bytes.byteOffset + file.bytes.byteLength,
  ) as ArrayBuffer
  return new File([data], file.name, { type: file.type, lastModified: Date.now() })
}

export function MobileScannerCapture({
  disabled,
  labels,
  onFiles,
}: {
  disabled: boolean
  labels: MobileScannerCaptureLabels
  onFiles: (files: File[]) => void
}): ReactElement | null {
  const api = window.pdfMobileScanner
  const sessionRef = useRef<PdfMobileScannerSession | null>(null)
  const onFilesRef = useRef(onFiles)
  const mountedRef = useRef(true)
  const [session, setSession] = useState<PdfMobileScannerSession | null>(null)
  const [starting, setStarting] = useState(false)
  const [error, setError] = useState('')
  const [copied, setCopied] = useState(false)
  const [receivedCount, setReceivedCount] = useState(0)
  const [remainingSeconds, setRemainingSeconds] = useState(0)

  onFilesRef.current = onFiles
  sessionRef.current = session

  const stop = useCallback(async (): Promise<void> => {
    const active = sessionRef.current
    sessionRef.current = null
    if (mountedRef.current) {
      setSession(null)
      setCopied(false)
    }
    if (active && api) await api.closeSession(active.sessionId).catch(() => undefined)
  }, [api])

  const start = async (): Promise<void> => {
    if (!api || starting || disabled) return
    await stop()
    setStarting(true)
    setError('')
    setReceivedCount(0)
    try {
      const created = await api.createSession()
      if (!mountedRef.current) {
        await api.closeSession(created.sessionId).catch(() => undefined)
        return
      }
      setSession(created)
      setRemainingSeconds(Math.max(0, Math.ceil((created.expiresAt - Date.now()) / 1000)))
    } catch (scannerError) {
      setError(scannerError instanceof Error ? scannerError.message : String(scannerError))
    } finally {
      if (mountedRef.current) setStarting(false)
    }
  }

  useEffect(() => {
    mountedRef.current = true
    return () => {
      mountedRef.current = false
      const active = sessionRef.current
      if (active && api) void api.closeSession(active.sessionId).catch(() => undefined)
    }
  }, [api])

  useEffect(() => {
    if (!api || !session) return
    let disposed = false
    let polling = false

    const poll = async (): Promise<void> => {
      if (disposed || polling) return
      polling = true
      try {
        const result = await api.pollSession(session.sessionId)
        if (disposed) return
        if (result.files.length > 0) {
          onFilesRef.current(result.files.map(scannerFile))
          setReceivedCount((count) => count + result.files.length)
          setError('')
        }
      } catch (pollError) {
        if (!disposed) setError(pollError instanceof Error ? pollError.message : String(pollError))
      } finally {
        polling = false
      }
    }

    const updateRemaining = (): void => {
      const remaining = Math.max(0, Math.ceil((session.expiresAt - Date.now()) / 1000))
      setRemainingSeconds(remaining)
      if (remaining === 0) {
        setError(labels.expired)
        void stop()
      }
    }

    void poll()
    const pollTimer = window.setInterval(() => void poll(), 1500)
    const clockTimer = window.setInterval(updateRemaining, 1000)
    return () => {
      disposed = true
      window.clearInterval(pollTimer)
      window.clearInterval(clockTimer)
    }
  }, [api, labels.expired, session, stop])

  if (!api) return null

  const copyLink = async (): Promise<void> => {
    if (!session) return
    try {
      await navigator.clipboard.writeText(session.uploadUrl)
      setCopied(true)
      window.setTimeout(() => mountedRef.current && setCopied(false), 1600)
    } catch (copyError) {
      setError(copyError instanceof Error ? copyError.message : String(copyError))
    }
  }

  return (
    <section className={`pdf-mobile-scanner${session ? ' active' : ''}`}>
      <div className="pdf-mobile-scanner-toolbar">
        <div>
          <strong>{labels.title}</strong>
          <span>{labels.hint}</span>
        </div>
        {!session ? (
          <button
            className="pdf-modal-btn"
            type="button"
            disabled={disabled || starting}
            onClick={() => void start()}
          >
            {starting ? `${labels.start}…` : labels.start}
          </button>
        ) : (
          <button
            className="pdf-modal-btn"
            type="button"
            disabled={disabled}
            onClick={() => void stop()}
          >
            {labels.stop}
          </button>
        )}
      </div>
      {session && (
        <div className="pdf-mobile-scanner-session">
          <div className="pdf-mobile-scanner-qr" aria-label={labels.title}>
            <QRCodeSVG value={session.uploadUrl} size={176} level="M" marginSize={2} />
          </div>
          <div className="pdf-mobile-scanner-details">
            <span>
              {receivedCount > 0 ? `${labels.received} ${receivedCount}` : labels.waiting}
            </span>
            <span>
              {labels.expiresIn} {Math.floor(remainingSeconds / 60)}:
              {String(remainingSeconds % 60).padStart(2, '0')}
            </span>
            <button
              className="pdf-modal-btn"
              type="button"
              disabled={disabled}
              onClick={() => void copyLink()}
            >
              {copied ? labels.copied : labels.copyLink}
            </button>
          </div>
        </div>
      )}
      {error && <div className="pdf-tools-error">{error || labels.unavailable}</div>}
    </section>
  )
}
