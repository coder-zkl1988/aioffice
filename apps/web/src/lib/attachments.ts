const MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024
const MAX_IMAGE_BYTES = 5 * 1024 * 1024
const IMAGE_MIMES: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  webp: 'image/webp',
}
const SUPPORTED_EXTENSIONS = new Set([
  'txt',
  'md',
  'markdown',
  'csv',
  'tsv',
  'json',
  'yaml',
  'yml',
  'xml',
  'html',
  'htm',
  'log',
  'js',
  'ts',
  'tsx',
  'jsx',
  'py',
  'java',
  'c',
  'h',
  'cpp',
  'go',
  'rs',
  'rb',
  'sh',
  'sql',
  'css',
  'docx',
  'pdf',
  'pptx',
  'xlsx',
  ...Object.keys(IMAGE_MIMES),
])

export interface WebAttachmentMeta {
  path: string
  name: string
  ext: string
  sizeBytes: number
}

export interface WebAttachmentAddResult {
  accepted: WebAttachmentMeta[]
  rejected: string[]
}

export interface WebAttachmentReadResult {
  ok: boolean
  error?: string
  name?: string
  totalChars?: number
  text?: string
  offset?: number
}

export interface WebAttachmentImageResult {
  ok: boolean
  base64?: string
  mime?: string
  error?: string
}

function extension(name: string): string {
  return name.split('.').pop()?.toLowerCase() ?? ''
}

function base64FromBytes(data: ArrayBuffer): string {
  const bytes = new Uint8Array(data)
  let binary = ''
  for (let offset = 0; offset < bytes.length; offset += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + 0x8000))
  }
  return btoa(binary)
}

function pickFiles(): Promise<File[]> {
  return new Promise((resolve) => {
    const input = document.createElement('input')
    input.type = 'file'
    input.multiple = true
    input.accept = [...SUPPORTED_EXTENSIONS].map((ext) => `.${ext}`).join(',')
    input.onchange = () => resolve([...input.files!])
    input.oncancel = () => resolve([])
    input.click()
  })
}

export function createWebAttachments() {
  const files = new Map<string, File>()

  function addFiles(selected: File[]): WebAttachmentAddResult {
    const accepted: WebAttachmentMeta[] = []
    const rejected: string[] = []
    for (const file of selected) {
      const ext = extension(file.name)
      if (!SUPPORTED_EXTENSIONS.has(ext)) {
        rejected.push(`${file.name}: 不支持 .${ext || 'unknown'} 文件`)
        continue
      }
      const maxBytes = IMAGE_MIMES[ext] ? MAX_IMAGE_BYTES : MAX_ATTACHMENT_BYTES
      if (file.size > maxBytes) {
        rejected.push(`${file.name}: 文件不能超过 ${Math.round(maxBytes / 1024 / 1024)} MB`)
        continue
      }
      const path = `attachment://${crypto.randomUUID()}/${encodeURIComponent(file.name)}`
      files.set(path, file)
      accepted.push({ path, name: file.name, ext, sizeBytes: file.size })
    }
    return { accepted, rejected }
  }

  return {
    pickAttachments: async () => addFiles(await pickFiles()),
    addAttachmentPaths: async () => ({
      accepted: [],
      rejected: ['Web 版不接受系统绝对路径，请使用附件选择按钮'],
    }),
    addPastedImage: async (data: ArrayBuffer, ext: string) => {
      const cleanExt = ext.toLowerCase()
      const mime = IMAGE_MIMES[cleanExt]
      if (!mime) return { accepted: [], rejected: [`不支持 .${cleanExt || 'unknown'} 图片`] }
      return addFiles([new File([data], `pasted-image.${cleanExt}`, { type: mime })])
    },
    readAttachment: async (
      path: string,
      offset: number,
      maxChars: number,
    ): Promise<WebAttachmentReadResult> => {
      const file = files.get(path)
      if (!file) return { ok: false, error: '附件不存在' }
      if (IMAGE_MIMES[extension(file.name)]) {
        return { ok: false, error: '图片附件已作为多模态内容发送，无可提取文字' }
      }
      const response = await fetch(new URL('./api/attachments/text', document.baseURI), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          name: file.name,
          base64: base64FromBytes(await file.arrayBuffer()),
          offset,
          maxChars,
        }),
      })
      const result = (await response.json().catch(() => ({}))) as WebAttachmentReadResult
      if (!response.ok)
        return { ok: false, error: result.error || `附件解析返回 HTTP ${response.status}` }
      return result
    },
    readAttachmentImage: async (path: string): Promise<WebAttachmentImageResult> => {
      const file = files.get(path)
      if (!file) return { ok: false, error: '附件不存在' }
      const mime = IMAGE_MIMES[extension(file.name)]
      if (!mime) return { ok: false, error: '附件不是支持的图片' }
      return { ok: true, base64: base64FromBytes(await file.arrayBuffer()), mime }
    },
    getPathForFile: (file: File) => {
      const result = addFiles([file])
      return result.accepted[0]?.path ?? file.name
    },
  }
}
