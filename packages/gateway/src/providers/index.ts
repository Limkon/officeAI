import type { ProviderConfig } from '@officeai/shared';
import { anthropicAdapter } from './anthropic.js';
import { geminiAdapter } from './gemini.js';
import { ollamaAdapter } from './ollama.js';
import { openaiAdapter } from './openai.js';
import type { ProviderAdapter } from './common.js';

export { probeLocalRuntimes } from './local-discovery.js';
export type { ProviderAdapter, ProviderStreamParams } from './common.js';

export function resolveProviderAdapter(protocol: ProviderConfig['protocol']): ProviderAdapter {
  switch (protocol) {
    case 'anthropic':
      return anthropicAdapter;
    case 'gemini':
      return geminiAdapter;
    case 'ollama':
      return ollamaAdapter;
    case 'openai':
    default:
      return openaiAdapter;
  }
}
