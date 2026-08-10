export type WebFileKind = 'docx' | 'markdown' | 'pdf'

export interface StoredWebFile {
  path: string
  name: string
  kind: WebFileKind
  mime: string
  updatedAt: number
  data: ArrayBuffer | string
}

const DB_NAME = 'genoffice-web'
const STORE_NAME = 'files'
const PENDING_KEY = 'genoffice.web.pending'
const handles = new Map<string, FileSystemFileHandle>()

function openDatabase(): Promise<IDBDatabase> {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1)
    request.onupgradeneeded = () => {
      const database = request.result
      if (!database.objectStoreNames.contains(STORE_NAME)) {
        database.createObjectStore(STORE_NAME, { keyPath: 'path' })
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })
}

async function transact<T>(
  mode: IDBTransactionMode,
  run: (store: IDBObjectStore) => IDBRequest<T>,
): Promise<T> {
  const database = await openDatabase()
  return new Promise((resolve, reject) => {
    const transaction = database.transaction(STORE_NAME, mode)
    const request = run(transaction.objectStore(STORE_NAME))
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
    transaction.oncomplete = () => database.close()
    transaction.onerror = () => reject(transaction.error)
  })
}

export function makeWebPath(name: string): string {
  return `webfile://${crypto.randomUUID()}/${encodeURIComponent(name)}`
}

export async function putStoredFile(file: StoredWebFile): Promise<void> {
  await transact('readwrite', (store) => store.put(file))
}

export async function getStoredFile(path: string): Promise<StoredWebFile | null> {
  return (await transact('readonly', (store) => store.get(path))) ?? null
}

export async function listStoredFiles(kind?: WebFileKind): Promise<StoredWebFile[]> {
  const all = await transact('readonly', (store) => store.getAll())
  return all
    .filter((file) => !kind || file.kind === kind)
    .sort((left, right) => right.updatedAt - left.updatedAt)
}

export function queuePendingFile(path: string): void {
  sessionStorage.setItem(PENDING_KEY, path)
}

export function consumePendingPath(): string | null {
  const path = sessionStorage.getItem(PENDING_KEY)
  sessionStorage.removeItem(PENDING_KEY)
  return path
}

function pickWithInput(accept: string): Promise<File | null> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.accept = accept
    input.onchange = () => resolve(input.files?.[0] ?? null)
    input.oncancel = () => resolve(null)
    input.click()
  })
}

export async function pickBrowserFile(
  kind: WebFileKind,
): Promise<{ file: File; path: string } | null> {
  const accept =
    kind === 'docx'
      ? '.docx'
      : kind === 'pdf'
        ? '.pdf,application/pdf'
        : '.md,.markdown,text/markdown,text/plain'
  let file: File | null = null
  let handle: FileSystemFileHandle | undefined

  if (window.showOpenFilePicker) {
    try {
      const [picked] = await window.showOpenFilePicker({
        multiple: false,
        types: [
          {
            description:
              kind === 'docx' ? 'Word 文档' : kind === 'pdf' ? 'PDF 文档' : 'Markdown 文档',
            accept:
              kind === 'docx'
                ? {
                    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': [
                      '.docx',
                    ],
                  }
                : kind === 'pdf'
                  ? { 'application/pdf': ['.pdf'] }
                  : { 'text/markdown': ['.md', '.markdown'] },
          },
        ],
      })
      handle = picked
      file = await picked.getFile()
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') return null
      throw error
    }
  } else {
    file = await pickWithInput(accept)
  }

  if (!file) return null
  const path = makeWebPath(file.name)
  if (handle) handles.set(path, handle)
  return { file, path }
}

export function getFileHandle(path: string): FileSystemFileHandle | undefined {
  return handles.get(path)
}

function extensionName(name: string, extension: string): string {
  return name.toLowerCase().endsWith(extension) ? name : `${name}${extension}`
}

function downloadBlob(name: string, blob: Blob): void {
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = name
  link.click()
  window.setTimeout(() => URL.revokeObjectURL(url), 1000)
}

export async function writeBrowserFile(options: {
  path?: string | null
  name: string
  extension: string
  mime: string
  blob: Blob
  forcePicker?: boolean
}): Promise<{ path: string; name: string }> {
  let path = options.path ?? null
  let handle = path ? handles.get(path) : undefined
  const name = extensionName(options.name, options.extension)

  if ((!handle || options.forcePicker) && window.showSaveFilePicker) {
    try {
      const pickedHandle = await window.showSaveFilePicker({
        suggestedName: name,
        types: [{ description: name, accept: { [options.mime]: [options.extension] } }],
      })
      handle = pickedHandle
      path = makeWebPath(pickedHandle.name)
      handles.set(path, pickedHandle)
    } catch (error) {
      if ((error as DOMException).name === 'AbortError') throw error
      handle = undefined
    }
  }

  if (handle) {
    const writable = await handle.createWritable()
    await writable.write(options.blob)
    await writable.close()
    return { path: path!, name: handle.name }
  }

  downloadBlob(name, options.blob)
  return { path: path ?? makeWebPath(name), name }
}

export function readTheme(): 'light' | 'dark' | 'system' {
  const value = localStorage.getItem('genoffice.web.theme')
  return value === 'light' || value === 'dark' ? value : 'system'
}

export function readLanguage(): 'zh' | 'en' {
  return navigator.language.toLowerCase().startsWith('zh') ? 'zh' : 'en'
}
