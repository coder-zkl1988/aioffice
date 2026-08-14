export type CameraFailureKind = 'denied' | 'unavailable' | 'unsupported'

export function cameraFailureKind(error: unknown): CameraFailureKind {
  if (typeof navigator === 'undefined' || !navigator.mediaDevices?.getUserMedia) {
    return 'unsupported'
  }
  const name = error instanceof DOMException ? error.name : ''
  if (name === 'NotAllowedError' || name === 'SecurityError') return 'denied'
  return 'unavailable'
}

export function cameraCaptureSize(
  sourceWidth: number,
  sourceHeight: number,
  maxLongEdge = 2560,
): { width: number; height: number } {
  if (
    !Number.isFinite(sourceWidth) ||
    !Number.isFinite(sourceHeight) ||
    sourceWidth <= 0 ||
    sourceHeight <= 0 ||
    !Number.isFinite(maxLongEdge) ||
    maxLongEdge <= 0
  ) {
    throw new Error('Camera frame dimensions are invalid')
  }
  const scale = Math.min(1, maxLongEdge / Math.max(sourceWidth, sourceHeight))
  return {
    width: Math.max(1, Math.round(sourceWidth * scale)),
    height: Math.max(1, Math.round(sourceHeight * scale)),
  }
}

export function cameraPageFileName(captureNumber: number, now = new Date()): string {
  const number = Number.isInteger(captureNumber) && captureNumber > 0 ? captureNumber : 1
  const stamp = now.toISOString().replace(/[:.]/g, '-').replace('T', '_').replace('Z', '')
  return `scan-${stamp}-${String(number).padStart(2, '0')}.jpg`
}
