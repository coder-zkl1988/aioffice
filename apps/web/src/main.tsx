import { useEffect, useMemo, useRef, useState } from 'react'
import { createRoot } from 'react-dom/client'
import {
  ArrowRight,
  Check,
  Clock3,
  Eye,
  EyeOff,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Home,
  Menu,
  Moon,
  PanelLeftClose,
  Presentation,
  Search,
  Settings2,
  Sparkles,
  Sun,
  Upload,
  X,
} from 'lucide-react'
import '@genoffice/ui/tokens.css'
import './workspace.css'
import {
  listStoredFiles,
  makeWebPath,
  putStoredFile,
  queuePendingFile,
  type StoredWebFile,
  type WebFileKind,
} from './lib/files'
import { getWebAiSettings, saveWebAiSettings, testWebAiConnection } from './lib/ai'

type Theme = 'light' | 'dark'
type AiDraft = { baseUrl: string; apiKey: string; model: string }

const modules = [
  { name: 'Docs', detail: 'DOCX', icon: FileText, accent: 'docs', enabled: true },
  { name: 'Markdown', detail: 'MD', icon: Menu, accent: 'markdown', enabled: true },
  {
    name: 'Sheets',
    detail: 'XLSX',
    icon: FileSpreadsheet,
    accent: 'sheets',
    enabled: true,
  },
  { name: 'Slides', detail: 'PPTX · 迁移中', icon: Presentation, accent: 'slides', enabled: false },
  { name: 'PDF', detail: 'PDF', icon: FileText, accent: 'pdf', enabled: true },
] as const

function routeFor(kind: WebFileKind): string {
  if (kind === 'docx') return './docs.html'
  if (kind === 'xlsx') return './sheets.html'
  if (kind === 'pdf') return './pdf.html'
  return './markdown.html'
}

function kindForName(name: string): WebFileKind | null {
  if (/\.docx$/i.test(name)) return 'docx'
  if (/\.xlsx$/i.test(name)) return 'xlsx'
  if (/\.(md|markdown)$/i.test(name)) return 'markdown'
  if (/\.pdf$/i.test(name)) return 'pdf'
  return null
}

async function importFile(file: File): Promise<StoredWebFile> {
  const kind = kindForName(file.name)
  if (!kind) throw new Error('仅支持 .docx、.xlsx、.md、.markdown 和 .pdf 文件')
  const stored: StoredWebFile = {
    path: makeWebPath(file.name),
    name: file.name,
    kind,
    mime:
      file.type ||
      (kind === 'docx'
        ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        : kind === 'xlsx'
          ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          : kind === 'pdf'
            ? 'application/pdf'
            : 'text/markdown'),
    updatedAt: Date.now(),
    data: kind === 'markdown' ? await file.text() : await file.arrayBuffer(),
  }
  await putStoredFile(stored)
  return stored
}

function App() {
  const [theme, setTheme] = useState<Theme>(() =>
    localStorage.getItem('genoffice.web.theme') === 'dark' ? 'dark' : 'light',
  )
  const [recent, setRecent] = useState<StoredWebFile[]>([])
  const [query, setQuery] = useState('')
  const [dragging, setDragging] = useState(false)
  const [error, setError] = useState('')
  const [sidebarOpen, setSidebarOpen] = useState(false)
  const [aiSettingsOpen, setAiSettingsOpen] = useState(false)
  const [showApiKey, setShowApiKey] = useState(false)
  const [aiDraft, setAiDraft] = useState<AiDraft>(() => {
    const config = getWebAiSettings().providers.custom
    return { baseUrl: config.baseUrl || '', apiKey: config.apiKey, model: config.model }
  })
  const [aiStatus, setAiStatus] = useState<{
    kind: 'idle' | 'testing' | 'ok' | 'error'
    text: string
  }>({
    kind: 'idle',
    text: '',
  })
  const inputRef = useRef<HTMLInputElement>(null)

  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('genoffice.web.theme', theme)
  }, [theme])

  useEffect(() => {
    void listStoredFiles().then(setRecent)
  }, [])

  useEffect(() => {
    if (!aiSettingsOpen) return
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') setAiSettingsOpen(false)
    }
    window.addEventListener('keydown', onKeyDown)
    return () => window.removeEventListener('keydown', onKeyDown)
  }, [aiSettingsOpen])

  const visibleRecent = useMemo(() => {
    const normalized = query.trim().toLowerCase()
    return normalized
      ? recent.filter((file) => file.name.toLowerCase().includes(normalized))
      : recent
  }, [query, recent])

  const openStored = (file: StoredWebFile) => {
    queuePendingFile(file.path)
    window.location.href = routeFor(file.kind)
  }

  const handleFile = async (file: File | null) => {
    if (!file) return
    setError('')
    try {
      openStored(await importFile(file))
    } catch (reason) {
      setError(reason instanceof Error ? reason.message : String(reason))
    }
  }

  const newFile = async (kind: WebFileKind) => {
    if (kind === 'xlsx') {
      const response = await fetch(new URL('./api/sheets/blank', document.baseURI))
      if (!response.ok) throw new Error('无法创建空白工作簿')
      const result = (await response.json()) as { name: string; xlsxBase64: string }
      const binary = atob(result.xlsxBase64)
      const data = Uint8Array.from(binary, (character) => character.charCodeAt(0)).buffer
      const file: StoredWebFile = {
        path: makeWebPath(result.name),
        name: result.name,
        kind: 'xlsx',
        mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        updatedAt: Date.now(),
        data,
      }
      await putStoredFile(file)
      openStored(file)
      return
    }
    sessionStorage.removeItem('genoffice.web.pending')
    window.location.href = routeFor(kind)
  }

  const openAiSettings = () => {
    const config = getWebAiSettings().providers.custom
    setAiDraft({ baseUrl: config.baseUrl || '', apiKey: config.apiKey, model: config.model })
    setAiStatus({ kind: 'idle', text: '' })
    setAiSettingsOpen(true)
  }

  const settingsFromDraft = () => {
    const settings = getWebAiSettings()
    return {
      ...settings,
      provider: 'custom' as const,
      providers: {
        ...settings.providers,
        custom: {
          baseUrl: aiDraft.baseUrl.trim(),
          apiKey: aiDraft.apiKey.trim(),
          model: aiDraft.model.trim(),
        },
      },
    }
  }

  const saveAiSettings = () => {
    if (!aiDraft.baseUrl.trim() || !aiDraft.apiKey.trim() || !aiDraft.model.trim()) {
      setAiStatus({ kind: 'error', text: '请完整填写 Base URL、API Key 和模型名称' })
      return
    }
    saveWebAiSettings(settingsFromDraft())
    setAiStatus({ kind: 'ok', text: '设置已保存' })
  }

  const testAiSettings = async () => {
    setAiStatus({ kind: 'testing', text: '正在测试连接...' })
    const result = await testWebAiConnection(settingsFromDraft())
    setAiStatus(
      result.ok
        ? { kind: 'ok', text: '连接成功' }
        : { kind: 'error', text: result.error || '连接失败' },
    )
  }

  return (
    <div className="workspace-shell">
      <aside className={sidebarOpen ? 'sidebar is-open' : 'sidebar'}>
        <div className="brand-row">
          <div className="brand-mark">G</div>
          <div>
            <strong>GenOffice</strong>
            <span>Web Workspace</span>
          </div>
          <button
            className="icon-button sidebar-close"
            onClick={() => setSidebarOpen(false)}
            aria-label="关闭导航"
          >
            <PanelLeftClose size={17} />
          </button>
        </div>
        <nav className="main-nav" aria-label="主导航">
          <button className="nav-item is-active">
            <Home size={17} />
            工作台
          </button>
          <button
            className="nav-item"
            onClick={() => document.getElementById('recent')?.scrollIntoView()}
          >
            <Clock3 size={17} />
            最近文件
          </button>
        </nav>
        <div className="sidebar-status">
          <span>
            <Check size={14} />
            浏览器本地存储
          </span>
          <small>文件默认不上传服务器</small>
        </div>
      </aside>

      <main className="workspace-main">
        <header className="topbar">
          <button
            className="icon-button mobile-menu"
            onClick={() => setSidebarOpen(true)}
            aria-label="打开导航"
          >
            <Menu size={18} />
          </button>
          <label className="search-field">
            <Search size={17} />
            <input
              value={query}
              onChange={(event) => setQuery(event.target.value)}
              placeholder="搜索最近文件"
            />
          </label>
          <span className="edition-badge">Web Preview</span>
          <button
            className="icon-button"
            onClick={openAiSettings}
            aria-label="AI 设置"
            title="AI 设置"
          >
            <Settings2 size={18} />
          </button>
          <button
            className="icon-button"
            onClick={() => setTheme(theme === 'light' ? 'dark' : 'light')}
            aria-label="切换主题"
            title="切换主题"
          >
            {theme === 'light' ? <Moon size={18} /> : <Sun size={18} />}
          </button>
        </header>

        <div className="workspace-content">
          <section className="section-block create-section">
            <div className="section-heading">
              <div>
                <p className="eyebrow">
                  <Sparkles size={14} />
                  创建
                </p>
                <h1>开始处理文件</h1>
              </div>
              <button className="open-command" onClick={() => inputRef.current?.click()}>
                <FolderOpen size={17} />
                打开文件
              </button>
              <input
                ref={inputRef}
                className="visually-hidden"
                type="file"
                accept=".docx,.xlsx,.md,.markdown,.pdf"
                onChange={(event) => void handleFile(event.target.files?.[0] ?? null)}
              />
            </div>

            <div className="module-grid">
              {modules.map((module) => {
                const Icon = module.icon
                const kind =
                  module.name === 'Docs'
                    ? 'docx'
                    : module.name === 'Sheets'
                      ? 'xlsx'
                      : module.name === 'PDF'
                        ? 'pdf'
                        : 'markdown'
                return (
                  <button
                    key={module.name}
                    className={`module-tile module-${module.accent}`}
                    disabled={!module.enabled}
                    onClick={() => module.enabled && void newFile(kind as WebFileKind)}
                  >
                    <span className="module-icon">
                      <Icon size={23} />
                    </span>
                    <span>
                      <strong>{module.name}</strong>
                      <small>{module.detail}</small>
                    </span>
                    {module.enabled && <ArrowRight className="module-arrow" size={17} />}
                  </button>
                )
              })}
            </div>
          </section>

          <section
            className={dragging ? 'drop-zone is-dragging' : 'drop-zone'}
            onDragEnter={(event) => {
              event.preventDefault()
              setDragging(true)
            }}
            onDragOver={(event) => event.preventDefault()}
            onDragLeave={(event) => {
              if (event.currentTarget === event.target) setDragging(false)
            }}
            onDrop={(event) => {
              event.preventDefault()
              setDragging(false)
              void handleFile(event.dataTransfer.files[0] ?? null)
            }}
            onClick={() => inputRef.current?.click()}
          >
            <Upload size={22} />
            <div>
              <strong>拖入文件</strong>
              <span>DOCX · XLSX · MD · MARKDOWN · PDF</span>
            </div>
          </section>

          {error && (
            <div className="error-banner" role="alert">
              {error}
            </div>
          )}

          <section className="section-block recent-section" id="recent">
            <div className="section-heading compact">
              <div>
                <p className="eyebrow">
                  <Clock3 size={14} />
                  最近
                </p>
                <h2>浏览器文件</h2>
              </div>
              <span className="count-label">{visibleRecent.length} 个</span>
            </div>
            {visibleRecent.length > 0 ? (
              <div className="recent-list">
                {visibleRecent.map((file) => (
                  <button key={file.path} className="recent-row" onClick={() => openStored(file)}>
                    <span className={`file-type file-${file.kind}`}>
                      {file.kind === 'markdown' ? (
                        <Menu size={18} />
                      ) : file.kind === 'xlsx' ? (
                        <FileSpreadsheet size={18} />
                      ) : (
                        <FileText size={18} />
                      )}
                    </span>
                    <span className="file-info">
                      <strong>{file.name}</strong>
                      <small>
                        {file.kind === 'docx'
                          ? 'Word 文档'
                          : file.kind === 'xlsx'
                            ? 'Excel 工作簿'
                            : file.kind === 'pdf'
                              ? 'PDF 文档'
                              : 'Markdown 文档'}
                      </small>
                    </span>
                    <time>
                      {new Intl.DateTimeFormat('zh-CN', {
                        month: '2-digit',
                        day: '2-digit',
                        hour: '2-digit',
                        minute: '2-digit',
                      }).format(file.updatedAt)}
                    </time>
                    <ArrowRight size={16} />
                  </button>
                ))}
              </div>
            ) : (
              <div className="empty-state">
                <FolderOpen size={24} />
                <span>{query ? '没有匹配文件' : '暂无最近文件'}</span>
              </div>
            )}
          </section>
        </div>
      </main>

      {aiSettingsOpen && (
        <div
          className="modal-backdrop"
          role="presentation"
          onMouseDown={() => setAiSettingsOpen(false)}
        >
          <section
            className="settings-dialog"
            role="dialog"
            aria-modal="true"
            aria-labelledby="ai-settings-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <header className="dialog-header">
              <div>
                <span>OpenAI Compatible</span>
                <h2 id="ai-settings-title">AI 模型设置</h2>
              </div>
              <button
                className="icon-button"
                onClick={() => setAiSettingsOpen(false)}
                aria-label="关闭"
                title="关闭"
              >
                <X size={18} />
              </button>
            </header>

            <div className="settings-fields">
              <label>
                <span>Base URL</span>
                <input
                  type="url"
                  value={aiDraft.baseUrl}
                  onChange={(event) =>
                    setAiDraft((current) => ({ ...current, baseUrl: event.target.value }))
                  }
                  placeholder="https://api.example.com/v1"
                  autoComplete="url"
                />
              </label>
              <label>
                <span>API Key</span>
                <div className="secret-field">
                  <input
                    type={showApiKey ? 'text' : 'password'}
                    value={aiDraft.apiKey}
                    onChange={(event) =>
                      setAiDraft((current) => ({ ...current, apiKey: event.target.value }))
                    }
                    placeholder="sk-..."
                    autoComplete="off"
                  />
                  <button
                    type="button"
                    onClick={() => setShowApiKey((visible) => !visible)}
                    aria-label={showApiKey ? '隐藏 API Key' : '显示 API Key'}
                    title={showApiKey ? '隐藏 API Key' : '显示 API Key'}
                  >
                    {showApiKey ? <EyeOff size={17} /> : <Eye size={17} />}
                  </button>
                </div>
              </label>
              <label>
                <span>模型名称</span>
                <input
                  value={aiDraft.model}
                  onChange={(event) =>
                    setAiDraft((current) => ({ ...current, model: event.target.value }))
                  }
                  placeholder="gpt-4.1-mini"
                  autoComplete="off"
                />
              </label>
            </div>

            <p className="settings-note">API Key 仅保存在当前浏览器，并通过本站代理转发。</p>
            {aiStatus.text && (
              <p className={`settings-status is-${aiStatus.kind}`} role="status">
                {aiStatus.text}
              </p>
            )}
            <footer className="dialog-actions">
              <button
                className="secondary-command"
                onClick={() => void testAiSettings()}
                disabled={aiStatus.kind === 'testing'}
              >
                测试连接
              </button>
              <button className="primary-command" onClick={saveAiSettings}>
                <Check size={16} />
                保存设置
              </button>
            </footer>
          </section>
        </div>
      )}
    </div>
  )
}

createRoot(document.getElementById('root')!).render(<App />)
