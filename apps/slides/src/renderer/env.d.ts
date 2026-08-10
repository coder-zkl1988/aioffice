/// <reference types="vite/client" />
import type { DesktopFilesApi, SlidesApi } from '../shared/ipc'
import type { ProjectApi } from '@genoffice/project-store'

declare global {
  interface Window {
    slidesApi: SlidesApi
    desktop: DesktopFilesApi
    projectApi: ProjectApi
  }
}

export {}
