import { useCallback, useEffect, useRef, useState } from 'react'
import type { ReactElement } from 'react'
import {
  cameraCaptureSize,
  cameraFailureKind,
  cameraPageFileName,
  type CameraFailureKind,
} from './camera-capture'

export interface CameraCaptureLabels {
  start: string
  stop: string
  switchCamera: string
  capture: string
  readyHint: string
  privacyHint: string
  captureFailed: string
  failures: Record<CameraFailureKind, string>
}

export function CameraCapture({
  disabled,
  labels,
  onCapture,
}: {
  disabled: boolean
  labels: CameraCaptureLabels
  onCapture: (file: File) => void
}): ReactElement {
  const videoRef = useRef<HTMLVideoElement>(null)
  const streamRef = useRef<MediaStream | null>(null)
  const requestIdRef = useRef(0)
  const captureCountRef = useRef(0)
  const [active, setActive] = useState(false)
  const [starting, setStarting] = useState(false)
  const [facingMode, setFacingMode] = useState<'environment' | 'user'>('environment')
  const [error, setError] = useState('')

  const stop = useCallback((): void => {
    requestIdRef.current += 1
    for (const track of streamRef.current?.getTracks() ?? []) track.stop()
    streamRef.current = null
    if (videoRef.current) videoRef.current.srcObject = null
    setActive(false)
    setStarting(false)
  }, [])

  const start = useCallback(
    async (nextFacingMode = facingMode): Promise<void> => {
      stop()
      const requestId = requestIdRef.current
      setError('')
      setStarting(true)
      try {
        if (!navigator.mediaDevices?.getUserMedia) throw new Error('Camera API unavailable')
        const stream = await navigator.mediaDevices.getUserMedia({
          audio: false,
          video: {
            facingMode: { ideal: nextFacingMode },
            width: { ideal: 2560 },
            height: { ideal: 1920 },
          },
        })
        if (requestId !== requestIdRef.current) {
          for (const track of stream.getTracks()) track.stop()
          return
        }
        streamRef.current = stream
        setActive(true)
        requestAnimationFrame(() => {
          if (!videoRef.current || streamRef.current !== stream) return
          videoRef.current.srcObject = stream
          void videoRef.current.play().catch(() => undefined)
        })
      } catch (cameraError) {
        if (requestId === requestIdRef.current) {
          setError(labels.failures[cameraFailureKind(cameraError)])
        }
      } finally {
        if (requestId === requestIdRef.current) setStarting(false)
      }
    },
    [facingMode, labels.failures, stop],
  )

  const switchCamera = (): void => {
    const next = facingMode === 'environment' ? 'user' : 'environment'
    setFacingMode(next)
    void start(next)
  }

  const capture = (): void => {
    const video = videoRef.current
    if (!video || video.videoWidth <= 0 || video.videoHeight <= 0) {
      setError(labels.captureFailed)
      return
    }
    try {
      const { width, height } = cameraCaptureSize(video.videoWidth, video.videoHeight)
      const canvas = document.createElement('canvas')
      canvas.width = width
      canvas.height = height
      const context = canvas.getContext('2d')
      if (!context) throw new Error(labels.captureFailed)
      context.drawImage(video, 0, 0, width, height)
      canvas.toBlob(
        (blob) => {
          if (!blob) {
            setError(labels.captureFailed)
            return
          }
          captureCountRef.current += 1
          onCapture(
            new File([blob], cameraPageFileName(captureCountRef.current), {
              type: 'image/jpeg',
              lastModified: Date.now(),
            }),
          )
          setError('')
        },
        'image/jpeg',
        0.92,
      )
    } catch {
      setError(labels.captureFailed)
    }
  }

  useEffect(() => () => stop(), [stop])
  useEffect(() => {
    if (disabled) stop()
  }, [disabled, stop])

  return (
    <section className={`pdf-camera-capture${active ? ' active' : ''}`}>
      <div className="pdf-camera-toolbar">
        <span>{labels.privacyHint}</span>
        {!active ? (
          <button
            className="pdf-modal-btn"
            type="button"
            disabled={disabled || starting}
            onClick={() => void start()}
          >
            {starting ? `${labels.start}…` : labels.start}
          </button>
        ) : (
          <button className="pdf-modal-btn" type="button" disabled={disabled} onClick={stop}>
            {labels.stop}
          </button>
        )}
      </div>
      {active && (
        <div className="pdf-camera-stage">
          <video ref={videoRef} playsInline muted aria-label={labels.readyHint} />
          <div className="pdf-camera-actions">
            <button
              className="pdf-modal-btn"
              type="button"
              disabled={disabled}
              onClick={switchCamera}
            >
              {labels.switchCamera}
            </button>
            <button
              className="pdf-modal-btn primary"
              type="button"
              disabled={disabled}
              onClick={capture}
            >
              {labels.capture}
            </button>
          </div>
        </div>
      )}
      {error && <div className="pdf-tools-error">{error}</div>}
    </section>
  )
}
