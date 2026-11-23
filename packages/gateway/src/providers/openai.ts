import { ThinkTagSplitter, type ConnectionTestResult, type ProviderConfig } from '@officeai/shared';
import { normalizeBaseUrl, readLines, type ProviderAdapter, type ProviderStreamParams } from './common.js';

function buildHeaders(provider: ProviderConfig): Record<string, string> {
  const headers: Record<string, string> = {
    'Content-Type': 'application/json',
    ...(provider.extraHeaders ?? {})
  };
  if (provider.apiKey.trim()) {
    headers.Authorization = `Bearer ${provider.apiKey.trim()}`;
  }
  return headers;
}

export const openaiAdapter: ProviderAdapter = {
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
    const url = `${base}/chat/completions`;
    const isReasoningModel = /^o[134]/i.test(model) || model.includes('reasoner');

    const payload: Record<string, unknown> = {
      model,
      messages,
      stream: true,
      stream_options: { include_usage: true }
    };
    if (!isReasoningModel) {
      payload.temperature = temperature;
    }
    if (maxTokens > 0) {
      payload.max_tokens = maxTokens;
    }

    const res = await fetch(url, {
      method: 'POST',
      headers: buildHeaders(provider),
      body: JSON.stringify(payload),
      signal
    });

    if (!res.ok || !res.body) {
      const errText = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}: ${errText || res.statusText}`);
    }

    const splitter = new ThinkTagSplitter();
    let usage: { promptTokens?: number; completionTokens?: number } | undefined;

    for await (const line of readLines(res.body)) {
      if (!line.startsWith('data:')) continue;
      const data = line.slice(5).trim();
      if (!data) continue;
      if (data === '[DONE]') break;

      let parsed: Record<string, unknown>;
      try {
        parsed = JSON.parse(data);
      } catch {
        continue;
      }

      const rawUsage = parsed.usage as Record<string, number> | undefined;
      if (rawUsage) {
        usage = {
          promptTokens: rawUsage.prompt_tokens,
          completionTokens: rawUsage.completion_tokens
        };
      }

      const choices = parsed.choices as Array<Record<string, unknown>> | undefined;
      const delta = choices?.[0]?.delta as Record<string, unknown> | undefined;
      if (!delta) continue;

      const reasoningContent = delta.reasoning_content;
      if (typeof reasoningContent === 'string' && reasoningContent.length > 0) {
        onEvent({ type: 'reasoning', content: reasoningContent });
      }

      const content = delta.content;
      if (typeof content === 'string' && content.length > 0) {
        const split = splitter.push(content);
        if (split.reasoning) onEvent({ type: 'reasoning', content: split.reasoning });
        if (split.content) onEvent({ type: 'delta', content: split.content });
      }
    }

    const tail = splitter.flush();
    if (tail.reasoning) onEvent({ type: 'reasoning', content: tail.reasoning });
    if (tail.content) onEvent({ type: 'delta', content: tail.content });
    onEvent({ type: 'done', usage });
  },

  async fetchModels(provider: ProviderConfig): Promise<string[]> {
    const base = normalizeBaseUrl(provider.baseUrl);
    const res = await fetch(`${base}/models`, {
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
    return Array.from(new Set(models));
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
        message: `连接成功，发现 ${models.length} 个模型 (${latencyMs} ms)`
      };
    } catch (err) {
      // Some OpenAI-compatible endpoints (e.g. Doubao Ark) don't expose GET /models,
      // so fallback to a minimal chat ping if fetchModels returns 404/405.
      const msg = err instanceof Error ? err.message : String(err);
      if (msg.includes('404') || msg.includes('405')) {
        try {
          const base = normalizeBaseUrl(provider.baseUrl);
          const pingRes = await fetch(`${base}/chat/completions`, {
            method: 'POST',
            headers: buildHeaders(provider),
            body: JSON.stringify({
              model: provider.selectedModel || provider.models[0],
              messages: [{ role: 'user', content: 'hi' }],
              max_tokens: 4
            }),
            signal: AbortSignal.timeout(10000)
          });
          const latencyMs = Date.now() - start;
          if (pingRes.ok) {
            return {
              ok: true,
              latencyMs,
              models: provider.models,
              message: `对话接口连接成功 (${latencyMs} ms)`
            };
          }
          const txt = await pingRes.text().catch(() => '');
          return {
            ok: false,
            latencyMs,
            models: [],
            message: `连接失败: HTTP ${pingRes.status} ${txt}`
          };
        } catch (pingErr) {
          return {
            ok: false,
            latencyMs: Date.now() - start,
            models: [],
            message: pingErr instanceof Error ? pingErr.message : String(pingErr)
          };
        }
      }
      return {
        ok: false,
        latencyMs: Date.now() - start,
        models: [],
        message: msg
      };
    }
  }
};
