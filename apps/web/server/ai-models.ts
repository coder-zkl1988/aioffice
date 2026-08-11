function modelId(value: unknown): string | null {
  if (typeof value === 'string') {
    const id = value.trim()
    return id && id.length <= 256 ? id : null
  }
  if (!value || typeof value !== 'object') return null
  const source = value as { id?: unknown; name?: unknown; model?: unknown }
  return modelId(source.id) || modelId(source.name) || modelId(source.model)
}

export function normalizeProviderModels(payload: unknown): string[] {
  const source = payload as { data?: unknown; models?: unknown } | Array<unknown> | null | undefined
  const candidates = Array.isArray(source)
    ? source
    : Array.isArray(source?.data)
      ? source.data
      : Array.isArray(source?.models)
        ? source.models
        : []

  return [...new Set(candidates.map(modelId).filter((id): id is string => Boolean(id)))]
    .sort((left, right) => left.localeCompare(right, 'en'))
    .slice(0, 500)
}
