import type { LocalRuntimeStatus, ProviderConfig } from '@officeai/shared';
import { ollamaAdapter } from './ollama.js';
import { openaiAdapter } from './openai.js';

export async function probeLocalRuntimes(
  providers: ProviderConfig[]
): Promise<LocalRuntimeStatus[]> {
  const ollamaCfg = providers.find((p) => p.id === 'ollama');
  const lmCfg = providers.find((p) => p.id === 'lmstudio');
  const vllmCfg = providers.find((p) => p.id === 'vllm');

  const checks: Array<Promise<LocalRuntimeStatus>> = [
    probeSingle('ollama', 'Ollama', ollamaCfg?.baseUrl ?? 'http://127.0.0.1:11434', async (url) => {
      const res = await fetch(`${url.replace(/\/+$/, '')}/api/tags`, {
        signal: AbortSignal.timeout(1500)
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = (await res.json()) as { models?: Array<{ name?: string }> };
      return (data.models ?? []).map((m) => m.name).filter((x): x is string => Boolean(x));
    }),
    probeSingle('lmstudio', 'LM Studio', lmCfg?.baseUrl ?? 'http://127.0.0.1:1234/v1', async (url) => {
      const res = await fetch(`${url.replace(/\/+$/, '')}/models`, {
        signal: AbortSignal.timeout(1500)
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = (await res.json()) as { data?: Array<{ id?: string }> };
      return (data.data ?? []).map((m) => m.id).filter((x): x is string => Boolean(x));
    }),
    probeSingle('vllm', 'vLLM / llama.cpp', vllmCfg?.baseUrl ?? 'http://127.0.0.1:8000/v1', async (url) => {
      const res = await fetch(`${url.replace(/\/+$/, '')}/models`, {
        signal: AbortSignal.timeout(1500)
      });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const data = (await res.json()) as { data?: Array<{ id?: string }> };
      return (data.data ?? []).map((m) => m.id).filter((x): x is string => Boolean(x));
    })
  ];

  return Promise.all(checks);
}

async function probeSingle(
  id: LocalRuntimeStatus['id'],
  name: string,
  baseUrl: string,
  fetcher: (url: string) => Promise<string[]>
): Promise<LocalRuntimeStatus> {
  try {
    const models = await fetcher(baseUrl);
    return { id, name, baseUrl, running: true, models };
  } catch (err) {
    return {
      id,
      name,
      baseUrl,
      running: false,
      models: [],
      error: err instanceof Error ? err.message : String(err)
    };
  }
}

export function getAdapter(protocol: ProviderConfig['protocol']) {
  switch (protocol) {
    case 'anthropic':
      return anthropicAdapterFallback();
    case 'gemini':
      return geminiAdapterFallback();
    case 'ollama':
      return ollamaAdapter;
    case 'openai':
    default:
      return openaiAdapter;
  }
}

// Imported lazily or re-exported cleanly in index.ts
import { anthropicAdapter } from './anthropic.js';
import { geminiAdapter } from './gemini.js';

function anthropicAdapterFallback() {
  return anthropicAdapter;
}
function geminiAdapterFallback() {
  return geminiAdapter;
}
