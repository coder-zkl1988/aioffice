import { randomUUID } from 'node:crypto'
import { mkdtemp, readFile, rm, unlink, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, join } from 'node:path'

import { blankXlsxBuffer } from '../../sheets/src/gateway/csv-import'
import { readArchiveEntryText } from '../../sheets/src/gateway/xlsx-package-io'
import { parsePivotDefinition } from '../../sheets/src/gateway/xlsx-pivot'
import {
  openWorkbookSession,
  sha256File,
  writeWorkbookTo,
  type SessionInfo,
} from '../../sheets/src/main/workbook-session'
import { XlsxSidecarClient } from '../../sheets/src/main/xlsx-sidecar-client'
import {
  workbookFileSchema,
  workbookFormulaCellsRequestSchema,
  workbookFormulaCellsResultSchema,
  workbookMediaRequestSchema,
  workbookMediaResultSchema,
  workbookPivotDefinitionSchema,
  workbookPivotRequestSchema,
  workbookRangeRequestSchema,
  workbookRangeResultSchema,
  workbookRecalcRequestSchema,
  workbookRecalcResultSchema,
  workbookSaveRequestSchema,
  type WorkbookFile,
  type WorkbookRecalcResult,
} from '../../sheets/src/shared/desktop-api'

interface WebWorkbookSession {
  directory: string
  name: string
  webPath: string
}

interface OpenWorkbookInput {
  name: string
  webPath?: string
  xlsxBase64: string
}

interface SaveWorkbookInput {
  request: unknown
  name?: string
  webPath?: string
}

interface SidecarRecalcCell {
  sheet: string
  row: number
  column: number
  formatted: string
  number?: number
  isFormula: boolean
}

function safeName(value: string): string {
  const name = basename(value)
    .replace(/[\u0000-\u001f]/g, '')
    .trim()
  if (!name || name.length > 255) throw new Error('工作簿名称无效')
  return name.toLowerCase().endsWith('.xlsx') ? name : `${name}.xlsx`
}

function decodeWorkbook(value: string, maxBytes: number): Buffer {
  if (typeof value !== 'string') throw new Error('工作簿内容无效')
  const bytes = Buffer.from(value, 'base64')
  if (bytes.length < 4 || bytes.length > maxBytes || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error('XLSX 文件无效或过大')
  }
  return bytes
}

function publicFile(file: WorkbookFile, meta: WebWorkbookSession): WorkbookFile {
  return workbookFileSchema.parse({ ...file, name: meta.name, path: meta.webPath })
}

export class SheetsWebService {
  private readonly client: XlsxSidecarClient
  private readonly sessions = new Map<string, SessionInfo>()
  private readonly webSessions = new Map<string, WebWorkbookSession>()

  constructor(
    binaryPath: string,
    private readonly maxWorkbookBytes = 64 * 1024 * 1024,
  ) {
    this.client = new XlsxSidecarClient(binaryPath)
  }

  async blank(): Promise<{ name: string; xlsxBase64: string }> {
    const name = 'Untitled.xlsx'
    return { name, xlsxBase64: (await blankXlsxBuffer()).toString('base64') }
  }

  async open(input: OpenWorkbookInput): Promise<WorkbookFile> {
    const name = safeName(input.name)
    const directory = await mkdtemp(join(tmpdir(), 'genoffice-sheets-'))
    const path = join(directory, name)
    let sessionId: string | undefined
    try {
      await writeFile(path, decodeWorkbook(input.xlsxBase64, this.maxWorkbookBytes), {
        mode: 0o600,
      })
      const file = await openWorkbookSession(this.client, path, this.sessions)
      sessionId = file.sessionId
      const meta = {
        directory,
        name,
        webPath: input.webPath || `webfile://${randomUUID()}/${encodeURIComponent(name)}`,
      }
      this.webSessions.set(file.sessionId, meta)
      return publicFile(file, meta)
    } catch (error) {
      if (sessionId) {
        this.sessions.delete(sessionId)
        this.webSessions.delete(sessionId)
        await this.client.close(sessionId).catch(() => undefined)
      }
      await rm(directory, { recursive: true, force: true })
      throw error
    }
  }

  async readRange(input: unknown): Promise<unknown> {
    const request = workbookRangeRequestSchema.parse(input)
    this.requireSession(request.sessionId)
    return workbookRangeResultSchema.parse(await this.client.readRange(request))
  }

  async readFormulas(input: unknown): Promise<unknown> {
    const request = workbookFormulaCellsRequestSchema.parse(input)
    this.requireSession(request.sessionId)
    return workbookFormulaCellsResultSchema.parse(await this.client.readFormulaCells(request))
  }

  async recalc(input: unknown): Promise<WorkbookRecalcResult> {
    const request = workbookRecalcRequestSchema.parse(input)
    const session = this.requireSession(request.sessionId)
    const fileSheetName = (sheetId: string): string => {
      const name = session.sheetNames.get(sheetId)
      if (!name) throw new Error(`Unknown sheet for recalculation: ${sheetId}`)
      return name
    }
    const raw = (await this.client.recalcCells({
      path: session.path,
      edits: request.edits.map((edit) => ({
        sheet: fileSheetName(edit.sheetId),
        row: edit.row,
        column: edit.column,
        input: edit.input,
      })),
      reads: request.reads.map((read) => ({
        sheet: fileSheetName(read.sheetId),
        range: read.range,
      })),
    })) as { cells?: SidecarRecalcCell[] }
    if (!Array.isArray(raw.cells)) throw new Error('工作簿重算响应无效')
    const idsByName = new Map([...session.sheetNames].map(([id, name]) => [name, id]))
    return workbookRecalcResultSchema.parse({
      cells: raw.cells.flatMap((cell) => {
        const sheetId = idsByName.get(cell.sheet)
        return sheetId
          ? [
              {
                sheetId,
                row: cell.row,
                column: cell.column,
                formatted: cell.formatted,
                ...(cell.number === undefined ? {} : { number: cell.number }),
                isFormula: cell.isFormula,
              },
            ]
          : []
      }),
    })
  }

  async readMedia(input: unknown): Promise<unknown> {
    const request = workbookMediaRequestSchema.parse(input)
    this.requireSession(request.sessionId)
    return workbookMediaResultSchema.parse(await this.client.readMedia(request))
  }

  async readPivot(input: unknown): Promise<unknown> {
    const request = workbookPivotRequestSchema.parse(input)
    const session = this.requireSession(request.sessionId)
    const [pivotXml, cacheXml] = await Promise.all([
      readArchiveEntryText(this.client, session.path, request.path),
      readArchiveEntryText(this.client, session.path, request.cachePath),
    ])
    return workbookPivotDefinitionSchema.parse(parsePivotDefinition(pivotXml, cacheXml))
  }

  async save(input: unknown): Promise<{
    canceled: false
    file: WorkbookFile
    touchedEntries: string[]
    xlsxBase64: string
  }> {
    const body = input as SaveWorkbookInput
    const request = workbookSaveRequestSchema.parse(body.request)
    const session = this.requireSession(request.sessionId)
    const meta = this.requireWebSession(request.sessionId)
    if ((await sha256File(session.path)) !== session.sha256) {
      throw new Error('工作簿会话文件已被外部修改')
    }
    const targetPath = join(meta.directory, `${randomUUID()}.xlsx`)
    const mutation = await writeWorkbookTo(this.client, session, request, targetPath)
    this.sessions.delete(request.sessionId)
    this.webSessions.delete(request.sessionId)
    await this.client.close(request.sessionId).catch(() => undefined)
    const file = await openWorkbookSession(this.client, targetPath, this.sessions)
    const nextMeta = {
      ...meta,
      name: body.name ? safeName(body.name) : meta.name,
      webPath: typeof body.webPath === 'string' && body.webPath ? body.webPath : meta.webPath,
    }
    this.webSessions.set(file.sessionId, nextMeta)
    await unlink(session.path).catch(() => undefined)
    return {
      canceled: false,
      file: publicFile(file, nextMeta),
      touchedEntries: [...mutation.touchedEntries],
      xlsxBase64: (await readFile(targetPath)).toString('base64'),
    }
  }

  async writeRecovery(input: unknown): Promise<{ ok: boolean }> {
    const request = workbookSaveRequestSchema.parse(input)
    const session = this.requireSession(request.sessionId)
    const meta = this.requireWebSession(request.sessionId)
    try {
      await writeWorkbookTo(this.client, session, request, join(meta.directory, 'recovery.xlsx'))
      return { ok: true }
    } catch {
      return { ok: false }
    }
  }

  async close(sessionId: unknown): Promise<void> {
    if (typeof sessionId !== 'string') return
    const session = this.sessions.get(sessionId)
    const meta = this.webSessions.get(sessionId)
    this.sessions.delete(sessionId)
    this.webSessions.delete(sessionId)
    await this.client.close(sessionId).catch(() => undefined)
    if (session) await unlink(session.path).catch(() => undefined)
    if (meta) await rm(meta.directory, { recursive: true, force: true })
  }

  private requireSession(sessionId: string): SessionInfo {
    const session = this.sessions.get(sessionId)
    if (!session) throw new Error('工作簿会话不存在或已过期')
    return session
  }

  private requireWebSession(sessionId: string): WebWorkbookSession {
    const session = this.webSessions.get(sessionId)
    if (!session) throw new Error('工作簿会话不存在或已过期')
    return session
  }
}
