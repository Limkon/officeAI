import type { ConnectionTestResult, ProviderConfig } from '@officeai/shared';
import { normalizeBaseUrl, readLines, type ProviderAdapter, type ProviderStreamParams } from './common.js';

export const geminiAdapter: ProviderAdapter = {
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
    const cleanModel = model.replace(/^models\//, '');
    const url = `${base}/v1beta/models/${cleanModel}:streamGenerateContent?alt=sse`;

    const systemText = messages
      .filter((m) => m.role === 'system')
      .map((m) => m.content)
      .join('\n\n')
      .trim();

    const contents = messages
      .filter((m) => m.role !== 'system' && m.content.trim().length > 0)
      .map((m) => ({
        role: m.role === 'assistant' ? 'model' : 'user',
        parts: [{ text: m.content }]
      }));

    if (contents.length === 0) {
      throw new Error('消息列表不能为空');
    }

    const payload: Record<string, unknown> = {
      contents,
      generationConfig: {
        temperature,
        maxOutputTokens: maxTokens > 0 ? maxTokens : 4096
      }
    };
    if (systemText) {
      payload.systemInstruction = {
        parts: [{ text: systemText }]
      };
    }

    const res = await fetch(url, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-goog-api-key': provider.apiKey.trim(),
        ...(provider.extraHeaders ?? {})
      },
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

      if (parsed.error && typeof parsed.error === 'object') {
        const errMsg = (parsed.error as { message?: string }).message;
        throw new Error(errMsg || 'Gemini stream error');
      }

      const usageMeta = parsed.usageMetadata as
        | { promptTokenCount?: number; candidatesTokenCount?: number }
        | undefined;
      if (usageMeta) {
        promptTokens = usageMeta.promptTokenCount;
        completionTokens = usageMeta.candidatesTokenCount;
      }

      const candidates = parsed.candidates as
        | Array<{ content?: { parts?: Array<{ text?: string; thought?: boolean }> } }>
        | undefined;
      const parts = candidates?.[0]?.content?.parts ?? [];
      for (const part of parts) {
        if (!part.text) continue;
        if (part.thought) {
          onEvent({ type: 'reasoning', content: part.text });
        } else {
          onEvent({ type: 'delta', content: part.text });
        }
      }
    }

    onEvent({
      type: 'done',
      usage: { promptTokens, completionTokens }
    });
  },

  async fetchModels(provider: ProviderConfig): Promise<string[]> {
    const base = normalizeBaseUrl(provider.baseUrl);
    const res = await fetch(`${base}/v1beta/models`, {
      method: 'GET',
      headers: {
        'x-goog-api-key': provider.apiKey.trim(),
        ...(provider.extraHeaders ?? {})
      },
      signal: AbortSignal.timeout(8000)
    });
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}: ${txt || res.statusText}`);
    }
    const json = (await res.json()) as {
      models?: Array<{ name?: string; supportedGenerationMethods?: string[] }>;
    };
    return (json.models ?? [])
      .filter((m) => m.supportedGenerationMethods?.includes('generateContent'))
      .map((m) => (m.name ?? '').replace(/^models\//, ''))
      .filter(Boolean);
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
        message: `Gemini 接口连接成功，发现 ${models.length} 个模型 (${latencyMs} ms)`
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
