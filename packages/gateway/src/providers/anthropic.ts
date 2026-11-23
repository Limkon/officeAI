import type { ConnectionTestResult, ProviderConfig } from '@officeai/shared';
import { normalizeBaseUrl, readLines, type ProviderAdapter, type ProviderStreamParams } from './common.js';

function buildHeaders(provider: ProviderConfig): Record<string, string> {
  return {
    'Content-Type': 'application/json',
    'x-api-key': provider.apiKey.trim(),
    'anthropic-version': '2023-06-01',
    ...(provider.extraHeaders ?? {})
  };
}

export const anthropicAdapter: ProviderAdapter = {
  async streamChat({
    provider,
    model,
    messages,
    temperature,
    maxTokens,
    signal,
    onEvent
  }: ProviderStreamParams) {
    const base = normalizeBaseUrl(provider.baseUrl);
    const systemText = messages
      .filter((m) => m.role === 'system')
      .map((m) => m.content)
      .join('\n\n')
      .trim();

    const chatMessages = messages
      .filter((m) => m.role !== 'system')
      .map((m) => ({
        role: m.role === 'assistant' ? 'assistant' : 'user',
        content: m.content
      }));

    const payload: Record<string, unknown> = {
      model,
      messages: chatMessages,
      max_tokens: maxTokens > 0 ? maxTokens : 4096,
      temperature,
      stream: true
    };
    if (systemText) {
      payload.system = systemText;
    }

    const res = await fetch(`${base}/v1/messages`, {
      method: 'POST',
      headers: buildHeaders(provider),
      body: JSON.stringify(payload),
      signal
    });

    if (!res.ok || !res.body) {
      const errText = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}: ${errText || res.statusText}`);
    }

    let promptTokens: number | undefined;
    let completionTokens: number | undefined;

    for await (const line of readLines(res.body)) {
      if (!line.startsWith('data:')) continue;
      const data = line.slice(5).trim();
      if (!data) continue;

      let parsed: Record<string, unknown>;
      try {
        parsed = JSON.parse(data);
      } catch {
        continue;
      }

      const type = parsed.type;
      if (type === 'message_start') {
        const msg = parsed.message as { usage?: { input_tokens?: number } } | undefined;
        promptTokens = msg?.usage?.input_tokens;
      } else if (type === 'content_block_delta') {
        const delta = parsed.delta as Record<string, unknown> | undefined;
        if (delta?.type === 'text_delta' && typeof delta.text === 'string') {
          onEvent({ type: 'delta', content: delta.text });
        } else if (delta?.type === 'thinking_delta' && typeof delta.thinking === 'string') {
          onEvent({ type: 'reasoning', content: delta.thinking });
        }
      } else if (type === 'message_delta') {
        const u = parsed.usage as { output_tokens?: number } | undefined;
        if (u?.output_tokens !== undefined) {
          completionTokens = u.output_tokens;
        }
      } else if (type === 'error') {
        const err = parsed.error as { message?: string } | undefined;
        throw new Error(err?.message || 'Anthropic stream error');
      }
    }

    onEvent({
      type: 'done',
      usage: { promptTokens, completionTokens }
    });
  },

  async fetchModels(provider: ProviderConfig): Promise<string[]> {
    const base = normalizeBaseUrl(provider.baseUrl);
    const res = await fetch(`${base}/v1/models`, {
      method: 'GET',
      headers: buildHeaders(provider),
      signal: AbortSignal.timeout(8000)
    });
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}: ${txt || res.statusText}`);
    }
    const json = (await res.json()) as { data?: Array<{ id?: string }> };
    const models = (json.data ?? [])
      .map((m) => m.id)
      .filter((id): id is string => Boolean(id));
    return models;
  },

  async testConnection(provider: ProviderConfig): Promise<ConnectionTestResult> {
    const start = Date.now();
    try {
      const models = await this.fetchModels(provider);
      const latencyMs = Date.now() - start;
      return {
        ok: true,
        latencyMs,
        models: models.length > 0 ? models : provider.models,
        message: `Claude 接口连接成功，发现 ${models.length} 个模型 (${latencyMs} ms)`
      };
    } catch (err) {
      return {
        ok: false,
        latencyMs: Date.now() - start,
        models: [],
        message: err instanceof Error ? err.message : String(err)
      };
    }
  }
};
