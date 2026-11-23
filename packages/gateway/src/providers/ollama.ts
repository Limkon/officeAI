import { ThinkTagSplitter, type ConnectionTestResult, type ProviderConfig } from '@officeai/shared';
import { normalizeBaseUrl, readLines, type ProviderAdapter, type ProviderStreamParams } from './common.js';

export const ollamaAdapter: ProviderAdapter = {
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
    const res = await fetch(`${base}/api/chat`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        ...(provider.extraHeaders ?? {})
      },
      body: JSON.stringify({
        model,
        messages,
        stream: true,
        options: {
          temperature,
          ...(maxTokens > 0 ? { num_predict: maxTokens } : {})
        }
      }),
      signal
    });

    if (!res.ok || !res.body) {
      const errText = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}: ${errText || res.statusText}`);
    }

    const splitter = new ThinkTagSplitter();
    let promptTokens: number | undefined;
    let completionTokens: number | undefined;

    for await (const line of readLines(res.body)) {
      if (!line.trim()) continue;
      let parsed: Record<string, unknown>;
      try {
        parsed = JSON.parse(line);
      } catch {
        continue;
      }

      if (typeof parsed.error === 'string') {
        throw new Error(parsed.error);
      }

      const msg = parsed.message as
        | { content?: string; thinking?: string }
        | undefined;
      if (msg?.thinking) {
        onEvent({ type: 'reasoning', content: msg.thinking });
      }
      if (msg?.content) {
        const split = splitter.push(msg.content);
        if (split.reasoning) onEvent({ type: 'reasoning', content: split.reasoning });
        if (split.content) onEvent({ type: 'delta', content: split.content });
      }

      if (parsed.done === true) {
        promptTokens = parsed.prompt_eval_count as number | undefined;
        completionTokens = parsed.eval_count as number | undefined;
      }
    }

    const tail = splitter.flush();
    if (tail.reasoning) onEvent({ type: 'reasoning', content: tail.reasoning });
    if (tail.content) onEvent({ type: 'delta', content: tail.content });

    onEvent({
      type: 'done',
      usage: { promptTokens, completionTokens }
    });
  },

  async fetchModels(provider: ProviderConfig): Promise<string[]> {
    const base = normalizeBaseUrl(provider.baseUrl);
    const res = await fetch(`${base}/api/tags`, {
      method: 'GET',
      signal: AbortSignal.timeout(4000)
    });
    if (!res.ok) {
      const txt = await res.text().catch(() => '');
      throw new Error(`HTTP ${res.status}: ${txt || res.statusText}`);
    }
    const json = (await res.json()) as { models?: Array<{ name?: string }> };
    return (json.models ?? [])
      .map((m) => m.name)
      .filter((name): name is string => Boolean(name));
  },

  async testConnection(provider: ProviderConfig): Promise<ConnectionTestResult> {
    const start = Date.now();
    try {
      const models = await this.fetchModels(provider);
      const latencyMs = Date.now() - start;
      return {
        ok: true,
        latencyMs,
        models,
        message:
          models.length > 0
            ? `Ollama 运行正常，已安装 ${models.length} 个本地模型 (${latencyMs} ms)`
            : 'Ollama 服务已连接，但尚未拉取模型（可运行 ollama pull qwen2.5:7b）'
      };
    } catch (err) {
      const detail = err instanceof Error ? err.message : String(err);
      return {
        ok: false,
        latencyMs: Date.now() - start,
        models: [],
        message: `无法连接 Ollama (${provider.baseUrl})：请确认 Ollama 已启动 (${detail})`
      };
    }
  }
};
