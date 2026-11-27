import { ThinkTagSplitter } from '@officeai/shared';
import type { ProviderConfig, StreamEvent } from '@officeai/shared';

function normalizeBaseUrl(raw: string): string {
  return raw.trim().replace(/\/+$/, '').replace(/\/chat\/completions$/, '').replace(/\/v1\/messages$/, '').replace(/\/v1beta\/.*$/, '');
}

function readLinesStream(body: ReadableStream<Uint8Array>): AsyncIterable<string> {
  const reader = body.getReader();
  const decoder = new TextDecoder();
  let buf = '';
  return {
    [Symbol.asyncIterator]() {
      return {
        async next() {
          while (true) {
            const { value, done } = await reader.read();
            if (done) {
              if (buf.length > 0) {
                const line = buf;
                buf = '';
                return { value: line, done: false };
              }
              return { value: undefined as any, done: true };
            }
            buf += decoder.decode(value, { stream: true });
            const idx = buf.indexOf('\n');
            if (idx !== -1) {
              const line = buf.slice(0, idx);
              buf = buf.slice(idx + 1);
              return { value: line, done: false };
            }
          }
        },
      };
    },
  };
}

function buildOpenAIHeaders(p: ProviderConfig): Record<string, string> {
  const h: Record<string, string> = { 'Content-Type': 'application/json', ...(p.extraHeaders ?? {}) };
  if (p.apiKey.trim()) h['Authorization'] = `Bearer ${p.apiKey.trim()}`;
  return h;
}

async function streamOpenAI(params: {
  provider: ProviderConfig;
  model: string;
  messages: any[];
  temperature: number;
  maxTokens: number;
  signal?: AbortSignal;
  onEvent: (e: StreamEvent) => void;
}) {
  const { provider, model, messages, temperature, maxTokens, signal, onEvent } = params;
  const base = normalizeBaseUrl(provider.baseUrl);
  const url = `${base}/chat/completions`;
  const isReasoning = /^o[134](?:-|$)/i.test(model) || model.includes('reasoner');
  const buildPayload = (withStreamOptions: boolean) => {
    const p: any = { model, messages, stream: true };
    if (withStreamOptions) p.stream_options = { include_usage: true };
    if (!isReasoning) p.temperature = temperature;
    if (maxTokens > 0) {
      if (/^o[134]/i.test(model)) p.max_completion_tokens = maxTokens;
      else p.max_tokens = maxTokens;
    }
    return p;
  };
  const headers = buildOpenAIHeaders(provider);
  const supports = provider.category === 'cloud' && provider.presetId !== 'custom-cloud';
  let res: Response;
  try {
    res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(buildPayload(supports)), signal });
  } catch (e) {
    // CORS / network error from browser direct mode
    const msg = e instanceof Error ? e.message : String(e);
    if (msg.includes('Failed to fetch') || msg.includes('NetworkError')) {
      throw new Error(
        `浏览器直连失败（CORS/网络）：${msg}。云端模型若不支持浏览器跨域，请改用本地网关代理（127.0.0.1:18080）或在网关的“配置中心”中配置。本地模型（Ollama）若为 http 而任务窗格为 https，会被浏览器以“混合内容”拦截，请用 http 任务窗格或网关。`
      );
    }
    throw e;
  }
  if (res.status === 400 && supports) {
    const t = await res.text().catch(() => '');
    if (/stream_options|unknown|unrecognized/i.test(t)) {
      res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(buildPayload(false)), signal });
    } else throw new Error(`HTTP 400: ${t || res.statusText}`);
  }
  if (!res.ok || !res.body) {
    const t = await res.text().catch(() => '');
    throw new Error(`HTTP ${res.status}: ${t || res.statusText}`);
  }
  const splitter = new ThinkTagSplitter();
  let usage: any;
  for await (const line of readLinesStream(res.body as any)) {
    if (!line.startsWith('data:')) continue;
    const data = line.slice(5).trim();
    if (!data || data === '[DONE]') continue;
    let j: any;
    try { j = JSON.parse(data); } catch { continue; }
    if (j.usage) usage = { promptTokens: j.usage.prompt_tokens, completionTokens: j.usage.completion_tokens };
    const delta = j.choices?.[0]?.delta;
    if (!delta) continue;
    if (delta.reasoning_content) onEvent({ type: 'reasoning', content: delta.reasoning_content });
    if (delta.content) {
      const s = splitter.push(delta.content);
      if (s.reasoning) onEvent({ type: 'reasoning', content: s.reasoning });
      if (s.content) onEvent({ type: 'delta', content: s.content });
    }
  }
  const tail = splitter.flush();
  if (tail.reasoning) onEvent({ type: 'reasoning', content: tail.reasoning });
  if (tail.content) onEvent({ type: 'delta', content: tail.content });
  onEvent({ type: 'done', usage });
}

async function streamAnthropic(params: any) {
  const { provider, model, messages, temperature, maxTokens, signal, onEvent } = params;
  const base = normalizeBaseUrl(provider.baseUrl);
  // Anthropic via OpenAI compat or native
  // Try native: /v1/messages
  const url = base.includes('anthropic.com') ? `${base}/v1/messages` : `${base}/chat/completions`;
  const headers: Record<string, string> = { 'Content-Type': 'application/json', ...(provider.extraHeaders ?? {}) };
  if (provider.apiKey.trim()) {
    if (base.includes('anthropic.com')) {
      headers['x-api-key'] = provider.apiKey.trim();
      headers['anthropic-version'] = '2023-06-01';
    } else headers['Authorization'] = `Bearer ${provider.apiKey.trim()}`;
  }
  // 简化：走 OpenAI 兼容路径
  return streamOpenAI(params);
}

async function streamGemini(params: any) {
  const { provider, model, messages, signal, onEvent } = params;
  const base = normalizeBaseUrl(provider.baseUrl);
  // Gemini via OpenAI compat endpoint or native
  // 若 base 为 generativelanguage.googleapis.com，拼 /v1beta/models/{model}:streamGenerateContent
  if (base.includes('generativelanguage.googleapis.com')) {
    const url = `${base}/v1beta/models/${encodeURIComponent(model)}:streamGenerateContent?alt=sse`;
    const headers: Record<string, string> = { 'Content-Type': 'application/json' };
    if (provider.apiKey.trim()) headers['x-goog-api-key'] = provider.apiKey.trim();
    const gemMessages = messages.filter((m: any) => m.role !== 'system').map((m: any) => ({ role: m.role === 'assistant' ? 'model' : 'user', parts: [{ text: m.content }] }));
    const sys = messages.find((m: any) => m.role === 'system');
    const body: any = { contents: gemMessages };
    if (sys) body.systemInstruction = { parts: [{ text: sys.content }] };
    let res: Response;
    try {
      res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body), signal });
    } catch (e) {
      throw new Error(`Gemini 直连失败（CORS）：${e instanceof Error ? e.message : String(e)}`);
    }
    if (!res.ok || !res.body) throw new Error(`HTTP ${res.status}: ${await res.text().catch(() => res.statusText)}`);
    for await (const line of readLinesStream(res.body as any)) {
      if (!line.startsWith('data:')) continue;
      const d = line.slice(5).trim();
      if (!d) continue;
      let j: any;
      try { j = JSON.parse(d); } catch { continue; }
      const text = j.candidates?.[0]?.content?.parts?.[0]?.text;
      if (text) onEvent({ type: 'delta', content: text });
      if (j.usageMetadata) onEvent({ type: 'done', usage: { promptTokens: j.usageMetadata.promptTokenCount, completionTokens: j.usageMetadata.candidatesTokenCount } });
    }
    onEvent({ type: 'done' });
    return;
  }
  return streamOpenAI(params);
}

async function streamOllama(params: any) {
  const { provider, model, messages, signal, onEvent } = params;
  const base = normalizeBaseUrl(provider.baseUrl);
  const url = `${base}/api/chat`;
  let res: Response;
  try {
    res = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ model, messages, stream: true }),
      signal,
    });
  } catch (e) {
    throw new Error(`Ollama 直连失败：${e instanceof Error ? e.message : String(e)}。若任务窗格为 https 而 Ollama 为 http，会被“混合内容”拦截，请用 http 任务窗格或经由本地网关。`);
  }
  if (!res.ok || !res.body) throw new Error(`HTTP ${res.status}: ${await res.text().catch(() => res.statusText)}`);
  for await (const line of readLinesStream(res.body as any)) {
    if (!line) continue;
    let j: any;
    try { j = JSON.parse(line); } catch { continue; }
    if (j.message?.content) onEvent({ type: 'delta', content: j.message.content });
    if (j.done) onEvent({ type: 'done', usage: { promptTokens: j.prompt_eval_count, completionTokens: j.eval_count } });
  }
}

export async function directStreamChat(params: {
  provider: ProviderConfig;
  model: string;
  messages: any[];
  temperature: number;
  maxTokens: number;
  signal?: AbortSignal;
  onEvent: (e: StreamEvent) => void;
}) {
  const p = params.provider.protocol;
  if (p === 'anthropic') return streamAnthropic(params);
  if (p === 'gemini') return streamGemini(params);
  if (p === 'ollama') return streamOllama(params);
  return streamOpenAI(params);
}

export async function directTestConnection(provider: ProviderConfig) {
  const start = Date.now();
  try {
    const base = normalizeBaseUrl(provider.baseUrl);
    const headers: Record<string, string> = provider.protocol === 'gemini' && provider.apiKey.trim()
      ? { 'x-goog-api-key': provider.apiKey.trim() }
      : provider.apiKey.trim() ? { Authorization: `Bearer ${provider.apiKey.trim()}` } : {};
    // 尝试 models
    let res: Response;
    try {
      res = await fetch(`${base}/models`, { headers, signal: AbortSignal.timeout(8000) as any });
    } catch (e) {
      throw new Error(`直连失败（CORS/网络）：${e instanceof Error ? e.message : String(e)}`);
    }
    if (res.ok) {
      const j: any = await res.json().catch(() => ({}));
      const models = (j.data || []).map((m: any) => m.id).filter(Boolean);
      return { ok: true, latencyMs: Date.now() - start, models: models.length ? models : provider.models, message: `直连成功，发现 ${models.length} 个模型` };
    }
    // 回退 ping chat
    const ping = await fetch(`${base}/chat/completions`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', ...headers },
      body: JSON.stringify({ model: provider.selectedModel || provider.models[0], messages: [{ role: 'user', content: 'hi' }], max_tokens: 2, stream: false }),
      signal: AbortSignal.timeout(10000) as any,
    });
    const t = await ping.text().catch(() => '');
    if (ping.ok) return { ok: true, latencyMs: Date.now() - start, models: provider.models, message: `直连成功` };
    return { ok: false, latencyMs: Date.now() - start, models: [], message: `HTTP ${ping.status}: ${t}` };
  } catch (e) {
    return { ok: false, latencyMs: Date.now() - start, models: [], message: e instanceof Error ? e.message : String(e) };
  }
}
