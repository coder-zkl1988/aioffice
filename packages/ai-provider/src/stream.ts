import type { AgentMessage, AgentToolDef } from '@genoffice/agent-core'
import { withOutputCapFallback } from './output-cap'
import { streamAnthropic } from './protocols/anthropic'
import { streamGemini } from './protocols/gemini'
import { streamOpenAiCompatible } from './protocols/openai-compatible'
import { streamOpenAiResponses } from './protocols/openai-responses'
import { streamCodexAppServer } from './codex-app-server'
import type { StreamCallbacks } from './protocols/shared'
import { getProviderAdapter, type AiProtocol } from './registry'
import {
  anthropicThinkingFields,
  geminiThinkingLevelFor,
  openAiEffortFor,
  toggleThinkingBodyFor,
} from './thinking-effort'
import type { AiProviderConfig, AiProviderId } from './types'

export { streamAnthropic } from './protocols/anthropic'
export { streamGemini } from './protocols/gemini'
export { streamOpenAiCompatible } from './protocols/openai-compatible'
export { streamOpenAiResponses } from './protocols/openai-responses'
export { AiCreditsError, sseLines } from './protocols/shared'
export type { StreamCallbacks } from './protocols/shared'

/**
 * The effort an OpenAI-shaped request carries. The per-model pick from
 * `config.thinking` wins; when a custom endpoint's model id is unknown to the
 * shipped family heuristics it resolves to nothing, so the legacy
 * per-provider `reasoningEffort` pick is forwarded verbatim instead — the
 * web workspace's settings UI still writes that field. `legacy` marks that
 * fallback, whose reasoning-model request shape (no temperature,
 * `max_completion_tokens`) the fork relied on.
 */
function openAiEffort(
  provider: AiProviderId,
  config: AiProviderConfig,
  protocol: AiProtocol,
): { effort: string | undefined; legacy: boolean } {
  const native = openAiEffortFor(config, protocol)
  if (native !== undefined) return { effort: native, legacy: false }
  const pick = provider === 'custom' ? config.reasoningEffort : undefined
  if (pick !== undefined && pick !== 'default') return { effort: pick, legacy: true }
  return { effort: undefined, legacy: false }
}

/** route a streaming, tool-calling-capable turn by provider id */
export async function streamForProvider(
  provider: AiProviderId,
  config: AiProviderConfig,
  system: string,
  messages: AgentMessage[],
  tools: AgentToolDef[],
  maxTokens: number,
  cb: StreamCallbacks,
): Promise<void> {
  const endpoint = getProviderAdapter(provider).resolveEndpoint(config)
  const { baseUrl } = endpoint
  if (endpoint.model) config = { ...config, model: endpoint.model }
  if (endpoint.protocol === 'codex-app-server') {
    return streamCodexAppServer(config, system, messages, tools, maxTokens, cb)
  }
  // A model the user declared a ceiling for outranks the caller's budget; the
  // endpoint's own 400 would otherwise cost a billed round trip.
  const cap = endpoint.maxTokens !== undefined ? Math.min(maxTokens, endpoint.maxTokens) : maxTokens
  const protocol: Exclude<AiProtocol, 'codex-app-server'> = endpoint.protocol
  const effort = openAiEffort(provider, config, protocol)
  return withOutputCapFallback(baseUrl, config.model, cap, (tokens) => {
    switch (protocol) {
      case 'anthropic':
        return streamAnthropic(config, system, messages, tools, tokens, cb, baseUrl, {
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
          ...anthropicThinkingFields(config),
        })
      case 'gemini':
        return streamGemini(config, system, messages, tools, tokens, cb, baseUrl, {
          omitTemperature: endpoint.omitTemperature,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
          thinkingLevel: geminiThinkingLevelFor(config),
        })
      case 'openai-compatible':
        return streamOpenAiCompatible(baseUrl, config, system, messages, tools, tokens, cb, {
          omitTemperature: effort.legacy || endpoint.omitTemperature,
          useMaxCompletionTokens: effort.legacy || endpoint.useMaxCompletionTokens,
          // a stored pick outranks the profile's own body extras
          bodyExtras: { ...endpoint.bodyExtras, ...toggleThinkingBodyFor(config) },
          echoReasoning: endpoint.echoReasoning,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
          reasoningEffort: effort.effort,
        })
      case 'openai-responses':
        return streamOpenAiResponses(baseUrl, config, system, messages, tools, tokens, cb, {
          bodyExtras: endpoint.bodyExtras,
          headers: endpoint.headers,
          idleTimeoutMs: endpoint.idleTimeoutMs,
          reasoningEffort: effort.effort,
        })
    }
  })
}
