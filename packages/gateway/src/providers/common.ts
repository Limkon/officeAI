import type { ChatMessage, ConnectionTestResult, ProviderConfig, StreamEvent } from '@officeai/shared';

export interface ProviderStreamParams {
  provider: ProviderConfig;
  model: string;
  messages: ChatMessage[];
  temperature: number;
  maxTokens: number;
  signal?: AbortSignal;
  onEvent: (event: StreamEvent) => void;
}

export interface ProviderAdapter {
  streamChat(params: ProviderStreamParams): Promise<void>;
  fetchModels(provider: ProviderConfig): Promise<string[]>;
  testConnection(provider: ProviderConfig): Promise<ConnectionTestResult>;
}

export async function* readLines(
  body: ReadableStream<Uint8Array>
): AsyncGenerator<string> {
  const reader = body.getReader();
  const decoder = new TextDecoder('utf-8');
  let buffer = '';

  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      let newlineIdx = buffer.indexOf('\n');
      while (newlineIdx !== -1) {
        const line = buffer.slice(0, newlineIdx).replace(/\r$/, '');
        buffer = buffer.slice(newlineIdx + 1);
        yield line;
        newlineIdx = buffer.indexOf('\n');
      }
    }
    buffer += decoder.decode();
    if (buffer.length > 0) {
      yield buffer.replace(/\r$/, '');
    }
  } finally {
    reader.releaseLock();
  }
}

export function normalizeBaseUrl(url: string): string {
  return url.trim().replace(/\/+$/, '');
}
