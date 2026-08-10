import type { ProjectApi } from '@genoffice/project-store'
import type { DesktopApi } from '../../docs/src/shared/ipc'
import type { MarkdownApi } from '../../markdown/src/shared/ipc'

declare global {
  interface FileSystemWritableFileStream {
    write(data: Blob | BufferSource | string): Promise<void>
    close(): Promise<void>
  }

  interface FileSystemFileHandle {
    name: string
    createWritable(): Promise<FileSystemWritableFileStream>
    getFile(): Promise<File>
  }

  interface Window {
    showOpenFilePicker?: (options?: unknown) => Promise<FileSystemFileHandle[]>
    showSaveFilePicker?: (options?: unknown) => Promise<FileSystemFileHandle>
    desktop: DesktopApi
    markdownApi: MarkdownApi
    projectApi: ProjectApi
  }
}

export {}
