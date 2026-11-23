/**
 * Incrementally extracts <think>...</think> blocks from streaming text
 * (used by DeepSeek-R1, QwQ, GLM-Z1, Ollama reasoning models).
 */
export class ThinkTagSplitter {
  private buffer = '';
  private inThink = false;

  push(chunk: string): { content: string; reasoning: string } {
    this.buffer += chunk;
    let content = '';
    let reasoning = '';

    while (this.buffer.length > 0) {
      if (!this.inThink) {
        const openIdx = this.buffer.indexOf('<think>');
        if (openIdx === -1) {
          // Keep up to 6 trailing chars in case '<think>' spans chunks
          const safeLen = Math.max(0, this.buffer.length - 6);
          if (safeLen === 0) break;
          content += this.buffer.slice(0, safeLen);
          this.buffer = this.buffer.slice(safeLen);
          break;
        }
        content += this.buffer.slice(0, openIdx);
        this.buffer = this.buffer.slice(openIdx + '<think>'.length);
        this.inThink = true;
      } else {
        const closeIdx = this.buffer.indexOf('</think>');
        if (closeIdx === -1) {
          const safeLen = Math.max(0, this.buffer.length - 7);
          if (safeLen === 0) break;
          reasoning += this.buffer.slice(0, safeLen);
          this.buffer = this.buffer.slice(safeLen);
          break;
        }
        reasoning += this.buffer.slice(0, closeIdx);
        this.buffer = this.buffer.slice(closeIdx + '</think>'.length);
        this.inThink = false;
      }
    }

    return { content, reasoning };
  }

  flush(): { content: string; reasoning: string } {
    const rest = this.buffer;
    this.buffer = '';
    if (this.inThink) {
      return { content: '', reasoning: rest };
    }
    return { content: rest, reasoning: '' };
  }
}

export function maskApiKey(key: string): string {
  const trimmed = key.trim();
  if (!trimmed) return '';
  if (trimmed.length <= 8) return '****';
  return `${trimmed.slice(0, 4)}••••${trimmed.slice(-4)}`;
}

export function renderPromptTemplate(
  template: string,
  vars: { selection?: string; context?: string }
): string {
  return template
    .replace(/\{\{selection\}\}/g, vars.selection ?? '')
    .replace(/\{\{context\}\}/g, vars.context ?? '');
}

/** Custom URL protocol used to launch the OfficeAI gateway on demand. */
export const OFFICEAI_PROTOCOL_SCHEME = 'officeai';

/**
 * Builds the "start the gateway" deep link used by Office / WPS add-ins.
 * Windows requires the `officeai://activate/<path>` form, where the path is
 * URL-encoded (both the command and the query string end up encoded).
 */
export function buildGatewayLaunchUrl(gatewayHttpUrl: string, app?: string): string {
  const query = app ? `?app=${encodeURIComponent(app)}` : '';
  return `${OFFICEAI_PROTOCOL_SCHEME}://activate/${encodeURIComponent(
    `${gatewayHttpUrl}/start${query}`
  )}`;
}

/**
 * Extracts the local gateway HTTP URL from a deep link like
 * `officeai://activate/http%3A%2F%2F127.0.0.1%3A18080%2Fstart`.
 * Returns undefined when the link is malformed or not loopback-only.
 */
export function parseGatewayLaunchUrl(deepLink: string): string | undefined {
  const match = /^officeai:\/\/activate\/(.+)$/i.exec(deepLink.trim());
  if (!match) return undefined;
  let target: string;
  try {
    target = decodeURIComponent(match[1]);
  } catch {
    return undefined;
  }
  try {
    const parsed = new URL(target);
    const loopback =
      parsed.hostname === '127.0.0.1' || parsed.hostname === 'localhost' || parsed.hostname === '[::1]';
    if (!loopback || parsed.protocol !== 'http:') return undefined;
    return `${parsed.protocol}//${parsed.host}`;
  } catch {
    return undefined;
  }
}
