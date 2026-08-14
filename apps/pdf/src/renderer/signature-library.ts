import type { SignatureData } from './SignatureDialog'

const DATABASE_NAME = 'genoffice-pdf-signatures'
const DATABASE_VERSION = 1
const STORE_NAME = 'signatures'

export const MAX_SAVED_SIGNATURES = 20
export const MAX_SIGNATURE_BYTES = 2 * 1024 * 1024
export const MAX_SIGNATURE_LIBRARY_BYTES = 20 * 1024 * 1024

export type SignatureLibraryErrorCode = 'item-limit' | 'item-too-large' | 'total-limit' | 'storage'

export interface SavedSignature {
  id: string
  label: string
  data: SignatureData
  byteSize: number
  createdAt: number
  updatedAt: number
}

export class SignatureLibraryError extends Error {
  constructor(readonly code: SignatureLibraryErrorCode) {
    super(code)
    this.name = 'SignatureLibraryError'
  }
}

export function signatureDataByteSize(data: SignatureData): number {
  return new TextEncoder().encode(JSON.stringify(data)).byteLength
}

export function validateSignatureLibraryAddition(
  saved: readonly Pick<SavedSignature, 'byteSize'>[],
  data: SignatureData,
): SignatureLibraryErrorCode | null {
  if (saved.length >= MAX_SAVED_SIGNATURES) return 'item-limit'
  const byteSize = signatureDataByteSize(data)
  if (byteSize > MAX_SIGNATURE_BYTES) return 'item-too-large'
  const totalBytes = saved.reduce((sum, item) => sum + Math.max(0, item.byteSize), 0)
  if (totalBytes + byteSize > MAX_SIGNATURE_LIBRARY_BYTES) return 'total-limit'
  return null
}

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new Error('IndexedDB request failed'))
  })
}

function transactionDone(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve()
    transaction.onerror = () =>
      reject(transaction.error ?? new Error('IndexedDB transaction failed'))
    transaction.onabort = () =>
      reject(transaction.error ?? new Error('IndexedDB transaction aborted'))
  })
}

function openDatabase(): Promise<IDBDatabase> {
  if (typeof indexedDB === 'undefined') return Promise.reject(new SignatureLibraryError('storage'))
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION)
    request.onupgradeneeded = () => {
      if (!request.result.objectStoreNames.contains(STORE_NAME)) {
        request.result.createObjectStore(STORE_NAME, { keyPath: 'id' })
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error ?? new SignatureLibraryError('storage'))
    request.onblocked = () => reject(new SignatureLibraryError('storage'))
  })
}

function isSignatureData(value: unknown): value is SignatureData {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<SignatureData>
  if (candidate.kind === 'strokes') {
    return (
      Array.isArray(candidate.paths) &&
      candidate.paths.every((path) => Array.isArray(path) && path.every(Number.isFinite)) &&
      Number.isFinite(candidate.width) &&
      Number.isFinite(candidate.height)
    )
  }
  return (
    candidate.kind === 'image' &&
    typeof candidate.image === 'string' &&
    Number.isFinite(candidate.width) &&
    Number.isFinite(candidate.height)
  )
}

function isSavedSignature(value: unknown): value is SavedSignature {
  if (!value || typeof value !== 'object') return false
  const candidate = value as Partial<SavedSignature>
  return (
    typeof candidate.id === 'string' &&
    typeof candidate.label === 'string' &&
    typeof candidate.byteSize === 'number' &&
    typeof candidate.createdAt === 'number' &&
    typeof candidate.updatedAt === 'number' &&
    isSignatureData(candidate.data)
  )
}

function makeId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID()
  }
  return `signature-${Date.now()}-${Math.random().toString(36).slice(2)}`
}

export async function listSavedSignatures(): Promise<SavedSignature[]> {
  let database: IDBDatabase | null = null
  try {
    database = await openDatabase()
    const transaction = database.transaction(STORE_NAME, 'readonly')
    const done = transactionDone(transaction)
    const values = await requestResult(transaction.objectStore(STORE_NAME).getAll())
    await done
    return values.filter(isSavedSignature).sort((left, right) => right.updatedAt - left.updatedAt)
  } catch (error) {
    if (error instanceof SignatureLibraryError) throw error
    throw new SignatureLibraryError('storage')
  } finally {
    database?.close()
  }
}

export async function saveSignature(label: string, data: SignatureData): Promise<SavedSignature> {
  const saved = await listSavedSignatures()
  const limitError = validateSignatureLibraryAddition(saved, data)
  if (limitError) throw new SignatureLibraryError(limitError)

  const now = Date.now()
  const record: SavedSignature = {
    id: makeId(),
    label: label.trim().slice(0, 80),
    data,
    byteSize: signatureDataByteSize(data),
    createdAt: now,
    updatedAt: now,
  }

  let database: IDBDatabase | null = null
  try {
    database = await openDatabase()
    const transaction = database.transaction(STORE_NAME, 'readwrite')
    transaction.objectStore(STORE_NAME).put(record)
    await transactionDone(transaction)
    return record
  } catch (error) {
    if (error instanceof SignatureLibraryError) throw error
    throw new SignatureLibraryError('storage')
  } finally {
    database?.close()
  }
}

export async function deleteSavedSignature(id: string): Promise<void> {
  let database: IDBDatabase | null = null
  try {
    database = await openDatabase()
    const transaction = database.transaction(STORE_NAME, 'readwrite')
    transaction.objectStore(STORE_NAME).delete(id)
    await transactionDone(transaction)
  } catch (error) {
    if (error instanceof SignatureLibraryError) throw error
    throw new SignatureLibraryError('storage')
  } finally {
    database?.close()
  }
}
